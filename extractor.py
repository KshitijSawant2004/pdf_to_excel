"""PDF table extraction utilities for native and scanned engineering PDFs."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
import shutil
import tempfile

import pandas as pd


NUMBER_WITH_UNIT_RE = re.compile(
    r"^\s*([-+]?\d[\d,]*(?:\.\d+)?)\s*([A-Za-z%][A-Za-z0-9/%^.\-]*)\s*$"
)


@dataclass
class ExtractedTable:
    """A table plus basic provenance details."""

    dataframe: pd.DataFrame
    source: str
    page: int | str


@dataclass
class ExtractionResult:
    """All tables and messages produced while reading a PDF."""

    tables: list[ExtractedTable] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)
    used_ocr: bool = False


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize empty cells, promote headers, and drop blank rows/columns."""

    cleaned = df.copy()
    cleaned = cleaned.map(lambda value: str(value).strip() if value is not None else "")
    cleaned = cleaned.replace(r"^\s*$", pd.NA, regex=True)
    cleaned = cleaned.dropna(how="all").dropna(axis=1, how="all")
    if cleaned.empty:
        return cleaned

    cleaned = cleaned.fillna("")
    first_row = cleaned.iloc[0].astype(str).str.strip()
    non_empty = int((first_row != "").sum())
    default_headers = all(str(col).isdigit() for col in cleaned.columns)
    if non_empty >= max(2, len(cleaned.columns) // 2) and default_headers:
        cleaned = cleaned.iloc[1:].copy()
        cleaned.columns = _dedupe_headers(first_row.tolist())
    else:
        cleaned.columns = _dedupe_headers([str(col).strip() or f"Column {i + 1}" for i, col in enumerate(cleaned.columns)])

    cleaned = _separate_number_units(cleaned.reset_index(drop=True))
    return cleaned


def _separate_number_units(df: pd.DataFrame) -> pd.DataFrame:
    """Split values like '203.10mAHD' into numeric and unit/comment cells.

    PDF table extractors often attach the unit printed just outside a numeric
    grid to the nearest numeric cell. Keeping that text in the numeric cell
    breaks formulas and makes cells appear populated when they should be blank.
    """

    cleaned = df.copy()
    for row_index in range(len(cleaned)):
        for col_index in range(len(cleaned.columns)):
            value = str(cleaned.iat[row_index, col_index]).strip()
            match = NUMBER_WITH_UNIT_RE.match(value)
            if not match:
                continue

            number, unit = match.groups()
            cleaned.iat[row_index, col_index] = number.replace(",", "")

            next_col = col_index + 1
            if next_col < len(cleaned.columns):
                next_value = str(cleaned.iat[row_index, next_col]).strip()
                if not next_value:
                    cleaned.iat[row_index, next_col] = unit

    return cleaned


def _dedupe_headers(headers: list[str]) -> list[str]:
    """Ensure Excel-friendly unique column names."""

    seen: dict[str, int] = {}
    result = []
    for index, header in enumerate(headers):
        name = header.strip() or f"Column {index + 1}"
        count = seen.get(name, 0)
        seen[name] = count + 1
        result.append(name if count == 0 else f"{name}_{count + 1}")
    return result


def _looks_useful(df: pd.DataFrame) -> bool:
    """Keep only tables with at least two rows and columns after cleanup."""

    return not df.empty and df.shape[0] >= 1 and df.shape[1] >= 2


def _extract_with_pdfplumber(pdf_path: Path, result: ExtractionResult) -> None:
    """Extract native PDF tables using pdfplumber."""

    try:
        import pdfplumber
    except ImportError as exc:
        result.messages.append(f"pdfplumber is not installed: {exc}")
        return

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                for raw_table in page.extract_tables() or []:
                    df = _clean_dataframe(pd.DataFrame(raw_table))
                    if _looks_useful(df):
                        result.tables.append(ExtractedTable(df, "pdfplumber", page_number))
    except Exception as exc:
        result.messages.append(f"pdfplumber extraction skipped: {exc}")


def _extract_with_camelot(pdf_path: Path, result: ExtractionResult) -> None:
    """Extract native PDF tables using Camelot lattice and stream modes."""

    try:
        import camelot
    except ImportError as exc:
        result.messages.append(f"Camelot is not installed: {exc}")
        return

    for flavor in ("lattice", "stream"):
        try:
            tables = camelot.read_pdf(str(pdf_path), pages="all", flavor=flavor)
            for table in tables:
                df = _clean_dataframe(table.df)
                if _looks_useful(df):
                    page = getattr(table, "page", "unknown")
                    result.tables.append(ExtractedTable(df, f"camelot-{flavor}", page))
        except Exception as exc:
            result.messages.append(f"Camelot {flavor} extraction skipped: {exc}")


def _pdf_has_text(pdf_path: Path) -> bool:
    """Detect whether a PDF has enough embedded text to avoid OCR."""

    try:
        import pdfplumber

        with pdfplumber.open(str(pdf_path)) as pdf:
            text = "\n".join((page.extract_text() or "") for page in pdf.pages[:3])
        return len(text.strip()) > 80
    except Exception:
        return False


def _ocr_lines_to_dataframe(text: str) -> pd.DataFrame:
    """Convert OCR text into a best-effort table by splitting aligned whitespace."""

    rows: list[list[str]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        cells = [cell.strip() for cell in re.split(r"\s{2,}|\t+", line) if cell.strip()]
        if len(cells) >= 2:
            rows.append(cells)

    if not rows:
        return pd.DataFrame()

    width = max(len(row) for row in rows)
    padded = [row + [""] * (width - len(row)) for row in rows]
    return _clean_dataframe(pd.DataFrame(padded))


def _extract_with_ocr(pdf_path: Path, result: ExtractionResult) -> None:
    """Render scanned PDF pages and extract table-like text with pytesseract."""

    try:
        import fitz
        import pytesseract
        from PIL import Image
    except ImportError as exc:
        result.messages.append(f"OCR dependencies are not installed: {exc}")
        return

    tesseract_path = shutil.which("tesseract") or r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if Path(tesseract_path).exists():
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    try:
        document = fitz.open(str(pdf_path))
        result.used_ocr = True
        for page_index in range(len(document)):
            page = document.load_page(page_index)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
            text = pytesseract.image_to_string(image, config="--psm 6")
            df = _ocr_lines_to_dataframe(text)
            if _looks_useful(df):
                result.tables.append(ExtractedTable(df, "pytesseract-ocr", page_index + 1))
    except Exception as exc:
        result.messages.append(
            "OCR extraction skipped. Ensure the Tesseract executable is installed "
            f"and on PATH. Details: {exc}"
        )


def extract_tables_from_pdf(pdf_path: str | Path) -> ExtractionResult:
    """Extract tables from a PDF using native parsers first, then OCR if needed."""

    path = Path(pdf_path)
    result = ExtractionResult()

    _extract_with_pdfplumber(path, result)
    _extract_with_camelot(path, result)

    if not result.tables and not _pdf_has_text(path):
        result.messages.append("No native tables found; scanned PDF suspected, running OCR.")
        _extract_with_ocr(path, result)
    elif not result.tables:
        result.messages.append("No tables were found in the text-based PDF.")

    result.messages.append(f"Extracted {len(result.tables)} table(s).")
    return result


def save_uploaded_pdf(uploaded_file) -> Path:
    """Persist a Streamlit upload to a temporary PDF file for parser libraries."""

    suffix = Path(uploaded_file.name).suffix or ".pdf"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return Path(tmp.name)
