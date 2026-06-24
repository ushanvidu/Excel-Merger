"""Shared helpers: header normalization and per-column value coercion."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

import pandas as pd

# ---------------------------------------------------------------------------
# Header / text normalization
# ---------------------------------------------------------------------------

_WS_RE = re.compile(r"\s+")
_NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")


def clean_cell(value: Any) -> str:
    """Turn a raw extracted cell into a clean single-line string.

    Collapses internal whitespace/newlines (common in PDF cells that wrap
    across lines) and strips surrounding spaces. ``None`` becomes "".
    """
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return _WS_RE.sub(" ", text).strip()


def normalize_header(name: Any) -> str:
    """Normalize a column header for fuzzy comparison.

    Lowercases, strips punctuation/spaces so that e.g. "Customer Name",
    "customer_name" and "Customer  Name" all compare equal.
    """
    text = clean_cell(name).lower()
    text = _NON_ALNUM_RE.sub(" ", text)
    return _WS_RE.sub(" ", text).strip()


def dedupe_headers(headers: list[str]) -> list[str]:
    """Ensure column headers are unique and non-empty.

    Blank headers become ``Column_N``; duplicates get a numeric suffix so a
    DataFrame can be built without collisions.
    """
    seen: dict[str, int] = {}
    result: list[str] = []
    for idx, raw in enumerate(headers, start=1):
        name = clean_cell(raw) or f"Column_{idx}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result


# ---------------------------------------------------------------------------
# Value coercion
# ---------------------------------------------------------------------------

# Comma/space group separators and a leading currency symbol are tolerated.
_NUMERIC_RE = re.compile(r"^[-+]?[\d,\s]*\.?\d+$")
_CURRENCY_RE = re.compile(r"[^\d.\-+]")


def _try_number(text: str) -> Any:
    candidate = text.strip()
    stripped = candidate.lstrip("$£€ ").strip()
    if _NUMERIC_RE.match(stripped):
        cleaned = _CURRENCY_RE.sub("", stripped)
        if cleaned in ("", "-", "+", "."):
            return None
        try:
            num = float(cleaned)
        except ValueError:
            return None
        if num.is_integer() and "." not in cleaned:
            return int(num)
        return num
    return None


def coerce_value(value: Any, kind: str = "auto") -> Any:
    """Coerce a single value toward a target ``kind``.

    kind is one of: ``"auto"``, ``"text"``, ``"number"``, ``"date"``.
    On failure the original cleaned string is returned, so data is never lost.
    """
    if value is None:
        return None
    if isinstance(value, (int, float, datetime, date)) and not isinstance(value, bool):
        return value

    text = clean_cell(value)
    if text == "":
        return None

    if kind == "text":
        return text

    if kind in ("number", "auto"):
        num = _try_number(text)
        if num is not None:
            return num
        if kind == "number":
            return text  # leave as-is rather than drop

    if kind in ("date", "auto"):
        parsed = _try_date(text)
        if parsed is not None:
            return parsed
        if kind == "date":
            return text

    return text


def _try_date(text: str) -> datetime | None:
    # Only attempt on strings that look date-ish to avoid mangling codes/ids.
    if not re.search(r"\d", text):
        return None
    if not re.search(r"[/\-.]|\d{4}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec",
                     text, re.IGNORECASE):
        return None
    try:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=False)
    except (ValueError, TypeError):
        return None
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def infer_column_kind(series: pd.Series, sample: int = 50) -> str:
    """Guess a column's dominant type from its values: number/date/text."""
    values = [v for v in series.head(sample).tolist() if clean_cell(v) != ""]
    if not values:
        return "text"
    numbers = sum(1 for v in values if _try_number(clean_cell(v)) is not None)
    if numbers / len(values) >= 0.8:
        return "number"
    dates = sum(1 for v in values if _try_date(clean_cell(v)) is not None)
    if dates / len(values) >= 0.8:
        return "date"
    return "text"
