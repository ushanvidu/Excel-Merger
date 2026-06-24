"""Unified loader for source documents — PDF *or* already-converted Excel/CSV.

Each source document describes ONE network site (though a single file may also
contain a small table). PDFs are converted to a table here (the "PDF -> Excel"
step); Excel/CSV files are read directly. All sources are normalized to the same
DataFrame shape and combined, tagged with ``__source_file__``.

Orientation
-----------
A single-site document often comes as a two-column **key/value form**
(``Parameter | Value``). ``orientation="key_value"`` pivots that into a single
wide row so it flows through the same matching/merge pipeline as a normal record.
``"records"`` keeps rows as-is. ``"auto"`` treats any 2-column source as
key/value and everything else as records.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

import pandas as pd

from .pdf_extractor import SOURCE_COLUMN, combine_frames, extract_single_pdf
from .utils import clean_cell, dedupe_headers

PDF_EXTS = {"pdf"}
EXCEL_EXTS = {"xlsx", "xlsm"}
CSV_EXTS = {"csv"}
SUPPORTED_EXTS = PDF_EXTS | EXCEL_EXTS | CSV_EXTS | {"xls"}


@dataclass
class LoadedSources:
    """Result of loading several source documents."""

    combined: pd.DataFrame                       # all sources stacked, tagged
    converted: dict[str, pd.DataFrame] = field(default_factory=dict)  # per file
    per_file_rows: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def file_kind(name: str) -> str:
    """Return 'pdf' | 'excel' | 'csv' | 'unsupported' for a filename."""
    ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
    if ext in PDF_EXTS:
        return "pdf"
    if ext in EXCEL_EXTS or ext == "xls":
        return "excel"
    if ext in CSV_EXTS:
        return "csv"
    return "unsupported"


def _read_excel(data: bytes) -> tuple[pd.DataFrame, list[str]]:
    # Read everything as text so site IDs like "COL-001" aren't coerced to numbers.
    try:
        df = pd.read_excel(io.BytesIO(data), dtype=str)
    except Exception as exc:  # noqa: BLE001 - surface to the UI
        return pd.DataFrame(), [f"Could not read Excel file: {exc}"]
    return df.fillna(""), []


def _read_csv(data: bytes) -> tuple[pd.DataFrame, list[str]]:
    try:
        df = pd.read_csv(io.BytesIO(data), dtype=str)
    except Exception as exc:  # noqa: BLE001
        return pd.DataFrame(), [f"Could not read CSV file: {exc}"]
    return df.fillna(""), []


def convert_to_dataframe(
    name: str,
    data: bytes,
    *,
    pdf_strategy: str = "auto",
    has_header: bool = True,
) -> tuple[pd.DataFrame, list[str]]:
    """Load any supported file into a raw DataFrame (before orientation)."""
    kind = file_kind(name)
    if kind == "pdf":
        return extract_single_pdf(io.BytesIO(data), strategy=pdf_strategy, has_header=has_header)
    if kind == "excel":
        return _read_excel(data)
    if kind == "csv":
        return _read_csv(data)
    return pd.DataFrame(), [f"Unsupported file type: {name}"]


def _pivot_key_value(df: pd.DataFrame) -> pd.DataFrame:
    """Turn a 2-column key/value table into a single wide row."""
    if df.shape[1] < 2 or df.empty:
        return df
    keys = [clean_cell(k) for k in df.iloc[:, 0].tolist()]
    values = [clean_cell(v) for v in df.iloc[:, 1].tolist()]
    pairs = [(k, v) for k, v in zip(keys, values) if k != ""]
    if not pairs:
        return df
    cols = dedupe_headers([k for k, _ in pairs])
    row = {c: v for c, (_, v) in zip(cols, pairs)}
    return pd.DataFrame([row], columns=cols)


def apply_orientation(df: pd.DataFrame, orientation: str) -> pd.DataFrame:
    """Apply 'records' | 'key_value' | 'auto' orientation to a source frame."""
    if df.empty:
        return df
    if orientation == "records":
        return df
    if orientation == "key_value":
        return _pivot_key_value(df)
    # auto: a 2-column source is almost always a key/value form.
    if df.shape[1] == 2:
        return _pivot_key_value(df)
    return df


def load_sources(
    files: list[tuple[str, bytes]],
    *,
    orientation: str = "auto",
    pdf_strategy: str = "auto",
    has_header: bool = True,
) -> LoadedSources:
    """Load and combine several source documents into one DataFrame.

    ``files`` is a list of ``(filename, raw_bytes)``. Returns the combined frame
    plus the per-file converted frames (so each can be previewed/downloaded as
    its own Excel) and any warnings.
    """
    converted: dict[str, pd.DataFrame] = {}
    per_file: dict[str, int] = {}
    warnings: list[str] = []
    frames: list[pd.DataFrame] = []

    for name, data in files:
        raw, file_warnings = convert_to_dataframe(
            name, data, pdf_strategy=pdf_strategy, has_header=has_header
        )
        df = apply_orientation(raw, orientation)
        converted[name] = df
        per_file[name] = len(df)
        for w in file_warnings:
            warnings.append(f"[{name}] {w}")
        if not df.empty:
            tagged = df.copy()
            tagged[SOURCE_COLUMN] = name
            frames.append(tagged)

    combined = combine_frames(frames) if frames else pd.DataFrame()
    return LoadedSources(combined, converted, per_file, warnings)


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Converted") -> bytes:
    """Serialize a DataFrame to .xlsx bytes (for the 'download converted' button)."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31] or "Converted")
    return out.getvalue()
