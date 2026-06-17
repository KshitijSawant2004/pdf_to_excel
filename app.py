"""Streamlit UI for the Engineering PDF to Excel Converter."""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path


REQUIRED_PACKAGES = {
    "streamlit": "streamlit",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "pdfplumber": "pdfplumber",
    "camelot": "camelot-py",
    "pytesseract": "pytesseract",
    "fitz": "pymupdf",
    "PIL": "Pillow",
    "numpy": "numpy",
}


def ensure_required_packages() -> None:
    """Install missing Python packages automatically before the app imports them."""

    missing = []
    for module_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(package_name)

    if missing:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


ensure_required_packages()

import streamlit as st

from excel_generator import generate_workbook
from extractor import extract_tables_from_pdf, save_uploaded_pdf
from formula_detector import describe_rules


st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="PDF",
    layout="wide",
)


def main() -> None:
    """Render the Streamlit app and coordinate PDF extraction to Excel export."""

    st.title("PDF to Excel Converter")
    st.caption("Extract scanned PDF tables, detect common formulas, and export XLSX.")

    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
    if uploaded_file is None:
        st.info("Upload a PDF to begin.")
        return

    pdf_path: Path | None = None
    try:
        pdf_path = save_uploaded_pdf(uploaded_file)
        with st.spinner("Extracting tables with pdfplumber, Camelot, and OCR fallback..."):
            result = extract_tables_from_pdf(pdf_path)

        for message in result.messages:
            st.write(message)

        if not result.tables:
            st.warning(
                "No tables were extracted. For scanned files, confirm Tesseract OCR is installed "
                "and available on your PATH."
            )
            return

        tables_for_excel = []
        st.subheader("Extracted tables")
        for index, table in enumerate(result.tables, start=1):
            source_name = f"{table.source} page {table.page}"
            tables_for_excel.append((table.dataframe, source_name))
            with st.expander(f"Table {index} - {source_name}", expanded=index == 1):
                st.dataframe(table.dataframe, use_container_width=True)

        with st.spinner("Detecting formulas and generating Excel workbook..."):
            workbook_bytes, formula_summary = generate_workbook(tables_for_excel)

        st.subheader("Detected formulas")
        any_rules = False
        for table_name, rules in formula_summary.items():
            if not rules:
                continue
            any_rules = True
            matching_table = result.tables[int(table_name.split()[1]) - 1]
            st.markdown(f"**{table_name}**")
            for description in describe_rules(matching_table.dataframe, rules):
                st.write(f"- {description}")

        if not any_rules:
            st.write("No reliable formula patterns were detected, so values were preserved as-is.")

        output_name = f"{Path(uploaded_file.name).stem}_converted.xlsx"
        st.download_button(
            label="Download XLSX",
            data=workbook_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    finally:
        if pdf_path and pdf_path.exists():
            pdf_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
