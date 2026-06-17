"""Flask UI for the Engineering PDF to Excel Converter."""

from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path
from uuid import uuid4

from flask import Flask, redirect, render_template, request, send_file, url_for

from excel_generator import generate_workbook
from extractor import extract_tables_from_pdf
from formula_detector import describe_rules


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

WORKBOOK_CACHE: dict[str, dict[str, bytes | str]] = {}
CACHE_LIMIT = 8


def _cleanup_cache() -> None:
    """Keep only a small number of generated workbooks in memory."""

    while len(WORKBOOK_CACHE) > CACHE_LIMIT:
        WORKBOOK_CACHE.pop(next(iter(WORKBOOK_CACHE)))


def _save_uploaded_pdf(uploaded_file) -> Path:
    """Persist an uploaded PDF to a temporary file for parser libraries."""

    suffix = Path(uploaded_file.filename or "").suffix or ".pdf"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.read())
        return Path(tmp.name)


def _process_pdf(uploaded_file) -> dict[str, object]:
    """Extract tables, build workbook bytes, and prepare HTML-friendly results."""

    pdf_path = _save_uploaded_pdf(uploaded_file)
    try:
        result = extract_tables_from_pdf(pdf_path)
        tables_for_excel: list[tuple[object, str]] = []
        tables: list[dict[str, object]] = []

        for index, table in enumerate(result.tables, start=1):
            source_name = f"{table.source} page {table.page}"
            tables_for_excel.append((table.dataframe, source_name))
            tables.append(
                {
                    "title": f"Table {index} - {source_name}",
                    "html": table.dataframe.to_html(
                        classes="data-table",
                        index=False,
                        border=0,
                        escape=True,
                    ),
                }
            )

        workbook_bytes, formula_summary = generate_workbook(tables_for_excel)

        formula_sections: list[dict[str, object]] = []
        for table_name, rules in formula_summary.items():
            if not rules:
                continue
            matching_table = result.tables[int(table_name.split()[1]) - 1]
            formula_sections.append(
                {
                    "title": table_name,
                    "descriptions": describe_rules(matching_table.dataframe, rules),
                }
            )

        return {
            "messages": result.messages,
            "tables": tables,
            "formula_sections": formula_sections,
            "workbook_bytes": workbook_bytes,
        }
    finally:
        pdf_path.unlink(missing_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():
    """Render the HTML upload form and processing results."""

    context = {
        "messages": [],
        "tables": [],
        "formula_sections": [],
        "download_url": None,
        "output_name": None,
        "error": None,
    }

    if request.method == "POST":
        uploaded_file = request.files.get("pdf_file")
        if uploaded_file is None or not uploaded_file.filename:
            context["error"] = "Choose a PDF file to continue."
            return render_template("index.html", **context)

        if not uploaded_file.filename.lower().endswith(".pdf"):
            context["error"] = "Please upload a PDF file."
            return render_template("index.html", **context)

        try:
            processed = _process_pdf(uploaded_file)
        except Exception as exc:
            context["error"] = f"Processing failed: {exc}"
            return render_template("index.html", **context)

        token = uuid4().hex
        output_name = f"{Path(uploaded_file.filename).stem}_converted.xlsx"
        WORKBOOK_CACHE[token] = {
            "bytes": processed["workbook_bytes"],
            "filename": output_name,
        }
        _cleanup_cache()

        context.update(
            {
                "messages": processed["messages"],
                "tables": processed["tables"],
                "formula_sections": processed["formula_sections"],
                "download_url": url_for("download", token=token),
                "output_name": output_name,
            }
        )
        return render_template("index.html", **context)

    return render_template("index.html", **context)


@app.route("/download/<token>")
def download(token: str):
    """Return the generated workbook as a file download."""

    cached = WORKBOOK_CACHE.pop(token, None)
    if cached is None:
        return redirect(url_for("index"))

    return send_file(
        BytesIO(cached["bytes"]),
        as_attachment=True,
        download_name=str(cached["filename"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
