"""Detect spreadsheet-style engineering formulas in extracted tables."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class FormulaRule:
    """Represents one formula pattern that can be written to Excel."""

    target_col: int
    formula_type: str
    source_cols: tuple[int, ...]
    confidence: float


AMOUNT_WORDS = ("amount", "total", "value", "cost", "net")
QTY_WORDS = ("qty", "quantity", "nos", "pcs", "units")
RATE_WORDS = ("rate", "price", "unit rate", "unit_rate")
GST_WORDS = ("gst", "tax", "cgst", "sgst", "igst")
PERCENT_WORDS = ("%", "percent", "percentage")
SUBTOTAL_WORDS = ("subtotal", "sub total", "taxable")
PAID_WORDS = ("paid", "advance", "received")
BALANCE_WORDS = ("balance", "due", "outstanding")


def normalize_number(value: object) -> float | None:
    """Convert PDF/OCR cell text into a float when possible."""

    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    cleaned = (
        text.replace(",", "")
        .replace("\u20b9", "")
        .replace("Rs.", "")
        .replace("INR", "")
        .replace("%", "")
        .replace("(", "-")
        .replace(")", "")
        .strip()
    )
    try:
        return float(cleaned)
    except ValueError:
        return None


def _numeric_series(series: pd.Series) -> pd.Series:
    """Return a numeric version of a DataFrame column."""

    return series.map(normalize_number).astype("float64")


def _is_close(left: pd.Series, right: pd.Series, tolerance: float = 0.03) -> float:
    """Measure how often two numeric series are approximately equal."""

    valid = left.notna() & right.notna() & (right.abs() > 1e-9)
    if valid.sum() == 0:
        return 0.0
    relative_error = ((left[valid] - right[valid]).abs() / right[valid].abs()).replace(
        [np.inf, -np.inf], np.nan
    )
    return float((relative_error <= tolerance).mean())


def _name_contains(name: object, words: Iterable[str]) -> bool:
    """Check whether a column header looks like a known engineering field."""

    lowered = str(name).strip().lower()
    return any(word in lowered for word in words)


def _excel_col(col_index: int) -> str:
    """Convert a zero-based DataFrame column index to an Excel column name."""

    name = ""
    index = col_index + 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _prefer_named_columns(
    columns: list[object], candidates: list[int], words: Iterable[str]
) -> list[int]:
    """Put columns with matching headers first while keeping all candidates."""

    named = [idx for idx in candidates if _name_contains(columns[idx], words)]
    unnamed = [idx for idx in candidates if idx not in named]
    return named + unnamed


def detect_formula_rules(df: pd.DataFrame) -> list[FormulaRule]:
    """Find common engineering invoice and estimate formulas in a DataFrame."""

    if df.empty or len(df.columns) < 2:
        return []

    columns = list(df.columns)
    numeric = {idx: _numeric_series(df.iloc[:, idx]) for idx in range(len(columns))}
    minimum_numeric_cells = max(2, int(len(df) * 0.55))
    numeric_cols = [
        idx for idx, series in numeric.items() if series.notna().sum() >= minimum_numeric_cells
    ]
    rules: list[FormulaRule] = []

    qty_candidates = _prefer_named_columns(columns, numeric_cols, QTY_WORDS)
    rate_candidates = _prefer_named_columns(columns, numeric_cols, RATE_WORDS)
    amount_candidates = _prefer_named_columns(columns, numeric_cols, AMOUNT_WORDS)

    for target in amount_candidates:
        best: FormulaRule | None = None
        for qty in qty_candidates:
            for rate in rate_candidates:
                if len({target, qty, rate}) < 3:
                    continue
                confidence = _is_close(numeric[target], numeric[qty] * numeric[rate])
                if confidence >= 0.65 and (best is None or confidence > best.confidence):
                    best = FormulaRule(target, "multiply", (qty, rate), confidence)
        if best:
            rules.append(best)

    gst_amount_cols = [
        idx
        for idx in numeric_cols
        if _name_contains(columns[idx], GST_WORDS)
        and not _name_contains(columns[idx], PERCENT_WORDS)
    ]
    gst_percent_cols = [
        idx
        for idx in numeric_cols
        if _name_contains(columns[idx], GST_WORDS + PERCENT_WORDS)
        and (
            _name_contains(columns[idx], PERCENT_WORDS)
            or numeric[idx].dropna().between(0, 100).mean() > 0.8
        )
    ]
    base_cols = _prefer_named_columns(columns, numeric_cols, SUBTOTAL_WORDS + AMOUNT_WORDS)

    for target in gst_amount_cols:
        best = None
        for base in base_cols:
            for pct in gst_percent_cols:
                if len({target, base, pct}) < 3:
                    continue
                confidence = _is_close(numeric[target], numeric[base] * numeric[pct] / 100)
                if confidence >= 0.65 and (best is None or confidence > best.confidence):
                    best = FormulaRule(target, "gst_percent", (base, pct), confidence)
        if best:
            rules.append(best)

    for target in amount_candidates:
        best = None
        for left in numeric_cols:
            for right in numeric_cols:
                if len({target, left, right}) < 3:
                    continue
                confidence = _is_close(numeric[target], numeric[left] + numeric[right])
                if confidence >= 0.7 and (best is None or confidence > best.confidence):
                    best = FormulaRule(target, "add", (left, right), confidence)
        if best and not any(rule.target_col == target for rule in rules):
            rules.append(best)

    balance_candidates = _prefer_named_columns(columns, numeric_cols, BALANCE_WORDS)
    paid_candidates = _prefer_named_columns(columns, numeric_cols, PAID_WORDS)
    total_candidates = _prefer_named_columns(columns, numeric_cols, AMOUNT_WORDS)

    for target in balance_candidates:
        best = None
        for total in total_candidates:
            for paid in paid_candidates:
                if len({target, total, paid}) < 3:
                    continue
                confidence = _is_close(numeric[target], numeric[total] - numeric[paid])
                if confidence >= 0.65 and (best is None or confidence > best.confidence):
                    best = FormulaRule(target, "subtract", (total, paid), confidence)
        if best:
            rules.append(best)

    deduped: dict[int, FormulaRule] = {}
    for rule in rules:
        current = deduped.get(rule.target_col)
        if current is None or rule.confidence > current.confidence:
            deduped[rule.target_col] = rule
    return list(deduped.values())


def formula_for_row(rule: FormulaRule, excel_row: int) -> str:
    """Build an Excel formula string for a detected rule and row."""

    refs = [f"{_excel_col(col)}{excel_row}" for col in rule.source_cols]
    if rule.formula_type == "multiply":
        return f"={refs[0]}*{refs[1]}"
    if rule.formula_type == "gst_percent":
        return f"={refs[0]}*{refs[1]}/100"
    if rule.formula_type == "add":
        return f"={refs[0]}+{refs[1]}"
    if rule.formula_type == "subtract":
        return f"={refs[0]}-{refs[1]}"
    return ""


def describe_rules(df: pd.DataFrame, rules: list[FormulaRule]) -> list[str]:
    """Return human-readable descriptions for Streamlit status output."""

    descriptions = []
    for rule in rules:
        target = str(df.columns[rule.target_col])
        sources = ", ".join(str(df.columns[col]) for col in rule.source_cols)
        descriptions.append(
            f"{target}: {rule.formula_type.replace('_', ' ')} from {sources} "
            f"({rule.confidence:.0%} match)"
        )
    return descriptions
