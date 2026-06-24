"""Extract tabular data from digital (text-based) PDFs into DataFrames.

Designed for PDFs whose text is selectable (exported from software, invoices,
reports) and whose data is laid out as tables with rows of records. Multiple
PDFs are combined into one DataFrame, with a ``__source_file__`` column added
for traceability.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, BinaryIO

import pandas as pd
import pdfplumber

from .utils import clean_cell, dedupe_headers

SOURCE_COLUMN = "__source_file__"

# Two strategies cover the vast majority of digital PDFs:
#  - "lines": tables drawn with ruling lines (borders) -> use line geometry.
#  - "text" : borderless tables aligned by whitespace -> infer columns from text.
_TABLE_SETTINGS = {
    "lines": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 4,
    },
    "text": {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 4,
        "join_tolerance": 4,
    },
}


@dataclass
class ExtractionResult:
    """Outcome of extracting one or more PDFs."""

    dataframe: pd.DataFrame
    per_file_rows: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def _rows_from_pdf(pdf: pdfplumber.PDF, strategy: str) -> list[list[str]]:
    """Return all table rows (as lists of clean strings) across every page."""
    settings = _TABLE_SETTINGS[strategy]
    rows: list[list[str]] = []
    for page in pdf.pages:
        for table in page.extract_tables(table_settings=settings):
            for raw_row in table:
                cleaned = [clean_cell(c) for c in raw_row]
                if any(cell != "" for cell in cleaned):  # skip blank rows
                    rows.append(cleaned)
    return rows


def _normalize_width(rows: list[list[str]]) -> list[list[str]]:
    """Pad/trim every row to the modal column count for a rectangular table."""
    if not rows:
        return rows
    widths = [len(r) for r in rows]
    target = max(set(widths), key=widths.count)  # most common width
    fixed = []
    for row in rows:
        if len(row) < target:
            row = row + [""] * (target - len(row))
        elif len(row) > target:
            row = row[:target]
        fixed.append(row)
    return fixed


def _strip_repeated_headers(rows: list[list[str]], header: list[str]) -> list[list[str]]:
    """Drop body rows that exactly repeat the header (multi-page tables)."""
    return [r for r in rows if r != header]


def extract_single_pdf(
    file: str | BinaryIO,
    *,
    strategy: str = "auto",
    has_header: bool = True,
) -> tuple[pd.DataFrame, list[str]]:
    """Extract one PDF into a DataFrame.

    Parameters
    ----------
    file : path or binary file-like object
    strategy : "auto" | "lines" | "text"
        "auto" tries line-based extraction first and falls back to text-based
        if it yields nothing.
    has_header : if True the first table row is treated as column names.
    """
    warnings: list[str] = []
    tried = ["lines", "text"] if strategy == "auto" else [strategy]

    rows: list[list[str]] = []
    used = ""
    with pdfplumber.open(file) as pdf:
        for strat in tried:
            rows = _rows_from_pdf(pdf, strat)
            if rows:
                used = strat
                break

    if not rows:
        warnings.append("No tables detected in this PDF.")
        return pd.DataFrame(), warnings

    if strategy == "auto" and used == "text":
        warnings.append("No bordered table found; used text-alignment extraction.")

    rows = _normalize_width(rows)

    if has_header:
        header = dedupe_headers(rows[0])
        body = _strip_repeated_headers(rows[1:], rows[0])
    else:
        width = len(rows[0])
        header = [f"Column_{i + 1}" for i in range(width)]
        body = rows

    df = pd.DataFrame(body, columns=header)
    return df, warnings


def extract_pdfs(
    files: list[tuple[str, Any]],
    *,
    strategy: str = "auto",
    has_header: bool = True,
) -> ExtractionResult:
    """Extract and vertically combine several PDFs.

    Parameters
    ----------
    files : list of ``(filename, file_or_path)`` tuples.

    Returns
    -------
    ExtractionResult with the combined DataFrame (``__source_file__`` column
    appended), per-file row counts and any warnings.
    """
    frames: list[pd.DataFrame] = []
    per_file: dict[str, int] = {}
    warnings: list[str] = []

    for name, handle in files:
        df, file_warnings = extract_single_pdf(
            handle, strategy=strategy, has_header=has_header
        )
        for w in file_warnings:
            warnings.append(f"[{name}] {w}")
        per_file[name] = len(df)
        if not df.empty:
            df = df.copy()
            df[SOURCE_COLUMN] = name
            frames.append(df)

    if not frames:
        return ExtractionResult(pd.DataFrame(), per_file, warnings)

    combined = combine_frames(frames)
    return ExtractionResult(combined, per_file, warnings)


def combine_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Concatenate frames, unioning columns so differing PDFs still combine.

    Columns are ordered by first appearance; ``__source_file__`` is moved last.
    """
    ordered: list[str] = []
    for df in frames:
        for col in df.columns:
            if col not in ordered:
                ordered.append(col)
    if SOURCE_COLUMN in ordered:
        ordered = [c for c in ordered if c != SOURCE_COLUMN] + [SOURCE_COLUMN]

    aligned = [df.reindex(columns=ordered, fill_value="") for df in frames]
    combined = pd.concat(aligned, ignore_index=True)
    return combined
