"""Template-aware extractor for InfraMS TSSR PDFs (5G Expansion survey reports).

These reports follow a fixed template. The data we care about lives in four
tables, each a list of ``Sector N / Antenna M`` rows with a known column schema:

  - **Antenna Details**           (existing physical antenna config)
  - **Sector Details**            (existing RF: tilts, RRU, jumper, feeder …)
  - **Solution Antenna Details**  (proposed antenna config)
  - **Solution Sector Details**   (proposed RF config)

pdfplumber wraps cell text across lines (``Sect\\nor 1`` -> ``"Sect or 1"`` and
``L850\\n_1`` -> ``"L850 _1"``). Because the columns are codes/numbers, we clean
by collapsing internal whitespace, then map the cleaned fields onto the fixed
schema. The result is one tidy DataFrame per section, ready to write to topic
sheets or to fill the master sheet.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pandas as pd
import pdfplumber

# ---- Fixed column schemas (in PDF order) ----------------------------------
ANTENNA_COLS = [
    "Sector No", "Antenna No", "Antenna Type", "Tower Leg", "Bracket Type",
    "Antenna Height (m)", "Antenna Direction", "Mechanical Tilt",
    "Spare Port Condition", "Port 01", "Port 02", "Port 03", "Port 04", "Port 05",
]
SECTOR_COLS = [
    "Sector No", "Antenna No", "Sector", "Electrical Tilt 01", "Electrical Tilt 02",
    "Mechanical Tilt", "Jumper Length", "Jumper Connector Type", "RRU Type",
    "Filter/Combiner Type", "Feeder Type", "Feeder Connector Type", "Feeder Length",
    "Fiber Condition", "Power Cable Condition", "Remarks",
]

SECTION_SCHEMAS = {
    "Antenna Details": ANTENNA_COLS,
    "Sector Details": SECTOR_COLS,
    "Solution Antenna Details": ANTENNA_COLS,
    "Solution Sector Details": SECTOR_COLS,
}

META_COLS = ["Source File", "Site ID", "Source Page"]

# Space-tolerant: geometry reconstruction preserves real spaces ("Sector 2").
_SECTOR_RE = re.compile(r"(?i)^sect\w*\s*0*(\d+)$")
_ANTENNA_RE = re.compile(r"(?i)^anten\w*\s*0*(\d+)$")
# A "Sector" (band/cell) value rather than an antenna type.
_BAND_RE = re.compile(
    r"(?i)(l\s?850|1800|2100|2600|l2\.?[36]|umts|gsm|900|n\d{2,3}|w$|_[pqrs]$|sec\s?\d)"
)


@dataclass
class TssrResult:
    site_id: str
    sections: dict[str, pd.DataFrame] = field(default_factory=dict)
    log: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# Vertical distance (pt) within which words are treated as the same text line.
_LINE_TOL = 4.0


def _reconstruct_cell(words: list[dict]) -> str:
    """Rebuild a cell's text from the words inside its bbox, using geometry.

    Words on the *same* line keep their real space; words *stacked* across lines
    are a line-wrap and are joined with no space — so "Sector" split as
    "Secto"/"r" rejoins correctly, while "Head"/"Frame" keeps its space, and a
    value wrapped over many lines (e.g. "1800|L1800_P") is never truncated.
    """
    if not words:
        return ""
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines: list[list[dict]] = [[ordered[0]]]
    top = ordered[0]["top"]
    for w in ordered[1:]:
        if abs(w["top"] - top) <= _LINE_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])
            top = w["top"]
    line_texts = [
        " ".join(w["text"] for w in sorted(ln, key=lambda w: w["x0"]))
        for ln in lines
    ]
    return re.sub(r"\s+", " ", "".join(line_texts)).strip()


def _fields(row: list) -> list[str]:
    """Strip cells and drop structurally-empty ones (no whitespace collapse)."""
    return [c for c in (str(x or "").strip() for x in row) if c]


def _is_data_row(fields: list[str]) -> bool:
    return bool(fields) and _SECTOR_RE.match(fields[0]) is not None


def _fmt_sector(value: str) -> str:
    m = _SECTOR_RE.match(value)
    return f"Sector {int(m.group(1))}" if m else value


def _fmt_antenna(value: str) -> str:
    m = _ANTENNA_RE.match(value)
    return f"Antenna {int(m.group(1))}" if m else value


_ANTENNA_TYPE_RE = re.compile(r"(?i)(port|beam|tb|sb|frame|anten|gain|dual)")


def _classify(fields: list[str]) -> str:
    """Return 'antenna' or 'sector' from the 3rd field (type) and field count.

    Antenna-type detection is checked first because a type such as
    "SB 1800 High Gain" contains a band-looking token ("1800").
    """
    third = fields[2] if len(fields) > 2 else ""
    if _ANTENNA_TYPE_RE.search(third):
        return "antenna"
    if _BAND_RE.search(third):
        return "sector"
    # Fall back to width: sector schema is wider.
    return "sector" if len(fields) >= 15 else "antenna"


def _row_to_schema(fields: list[str], cols: list[str]) -> tuple[dict, str | None]:
    """Map cleaned fields onto a fixed schema, padding/truncating as needed."""
    warn = None
    vals = list(fields)
    if len(vals) < len(cols):
        warn = f"{len(vals)} fields < {len(cols)} expected (padded)"
        vals += [""] * (len(cols) - len(vals))
    elif len(vals) > len(cols):
        warn = f"{len(vals)} fields > {len(cols)} expected (truncated)"
        vals = vals[: len(cols)]
    record = dict(zip(cols, vals))
    record["Sector No"] = _fmt_sector(record["Sector No"])
    record["Antenna No"] = _fmt_antenna(record["Antenna No"])
    return record, warn


def parse_site_id(filename: str, page_text: str = "") -> str:
    """Best-effort site ID from the filename (…- GMALG1 (1).pdf -> GMALG1)."""
    base = filename.rsplit("/", 1)[-1]
    base = re.sub(r"\.(pdf|PDF)$", "", base)
    # Tokens that look like a site code: letters followed by digits.
    candidates = re.findall(r"[A-Z]{2,}[A-Z0-9]*\d+", base.upper())
    if candidates:
        return candidates[-1]
    m = re.search(r"site\s*(?:id|name)\s*[:\-]?\s*([A-Z0-9]+)", page_text, re.I)
    return m.group(1) if m else base


def _extract_rows_geometry(page) -> list[list[str]]:
    """Reconstruct table rows from word geometry.

    Uses the ruling-line table finder to get the column/row grid, then fills
    each cell from the words whose centre lies inside its bbox — so nothing is
    truncated and words land in the right column even when the cell text is
    wrapped. Falls back to pdfplumber's text-based ``extract_tables`` if no
    ruled table is found on the page.
    """
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    rows: list[list[str]] = []
    for table in page.find_tables():
        for trow in table.rows:
            cells: list[str] = []
            for bbox in trow.cells:
                if bbox is None:
                    cells.append("")
                    continue
                x0, top, x1, bottom = bbox
                inside = [
                    w for w in words
                    if x0 <= (w["x0"] + w["x1"]) / 2 <= x1
                    and top <= (w["top"] + w["bottom"]) / 2 <= bottom
                ]
                cells.append(_reconstruct_cell(inside))
            rows.append(cells)
    if not rows:
        for table in page.extract_tables():
            rows.extend([str(c or "") for c in raw] for raw in table)
    return rows


def extract_tssr(file: str | bytes, *, filename: str = "", site_id: str | None = None) -> TssrResult:
    """Parse one TSSR PDF into its four data sections."""
    handle = io.BytesIO(file) if isinstance(file, (bytes, bytearray)) else file
    rows_by_section: dict[str, list[dict]] = {k: [] for k in SECTION_SCHEMAS}
    log: list[dict] = []
    warnings: list[str] = []

    with pdfplumber.open(handle) as pdf:
        first_text = pdf.pages[0].extract_text() or "" if pdf.pages else ""
        sid = site_id or parse_site_id(filename or getattr(pdf, "stream", "") or "", first_text)

        is_solution = False
        for page_no, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if "Solution Antenna Details" in text or "Solution Sector Details" in text:
                is_solution = True

            for raw in _extract_rows_geometry(page):
                fields = _fields(raw)
                if not _is_data_row(fields):
                    continue
                kind = _classify(fields)
                if kind == "antenna":
                    section = "Solution Antenna Details" if is_solution else "Antenna Details"
                else:
                    section = "Solution Sector Details" if is_solution else "Sector Details"
                record, warn = _row_to_schema(fields, SECTION_SCHEMAS[section])
                record = {"Source File": filename, "Site ID": sid,
                          "Source Page": page_no, **record}
                rows_by_section[section].append(record)
                log.append({
                    "Source Page": page_no, "Section": section,
                    "Sector No": record["Sector No"], "Antenna No": record["Antenna No"],
                    "Fields": len(fields), "Note": warn or "ok",
                })
                if warn:
                    warnings.append(f"[p{page_no} {section}] {warn}: {fields[:4]}")

    sections: dict[str, pd.DataFrame] = {}
    for name, schema in SECTION_SCHEMAS.items():
        cols = META_COLS + schema
        rows = rows_by_section[name]
        sections[name] = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)

    return TssrResult(site_id=sid, sections=sections, log=log, warnings=warnings)


# ---------------------------------------------------------------------------
# Multi-PDF combine + workbook export (the "convert PDF -> Excel, topic sheets")
# ---------------------------------------------------------------------------

def extract_many(files: list[tuple[str, bytes]]) -> list[TssrResult]:
    """Extract several TSSR PDFs (one site each)."""
    return [extract_tssr(data, filename=name) for name, data in files]


def results_from_workbook(xlsx_bytes: bytes) -> list[TssrResult]:
    """Build TssrResults from an already-converted multi-sheet workbook.

    Reads the four topic sheets (if present), groups rows by site, and rebuilds
    per-site sections. The Site ID is re-derived from the 'Source File' column
    (more reliable than a possibly-broken 'Site ID' column), falling back to the
    'Site ID' column. Lets the fill step consume a pre-converted Excel instead of
    re-parsing PDFs.
    """
    from collections import defaultdict

    sheets = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=None, dtype=str)
    per_site: dict[str, dict[str, list[dict]]] = defaultdict(
        lambda: {k: [] for k in SECTION_SCHEMAS}
    )

    for section, schema in SECTION_SCHEMAS.items():
        df = sheets.get(section)
        if df is None:
            continue
        df = df.fillna("")
        for _, row in df.iterrows():
            src = str(row.get("Source File", "")).strip()
            sid = parse_site_id(src) if src else str(row.get("Site ID", "")).strip()
            if not sid:
                continue
            rec = {"Source File": src, "Site ID": sid,
                   "Source Page": str(row.get("Source Page", row.get("Source Pages", "")))}
            for c in schema:
                rec[c] = str(row.get(c, "")).strip()
            per_site[sid][section].append(rec)

    results: list[TssrResult] = []
    for sid, secs in per_site.items():
        sections: dict[str, pd.DataFrame] = {}
        for name, schema in SECTION_SCHEMAS.items():
            cols = META_COLS + schema
            rows = secs[name]
            sections[name] = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
        results.append(TssrResult(site_id=sid, sections=sections))
    return results


def merge_results(results: list[TssrResult]) -> list[TssrResult]:
    """De-duplicate results by site, keeping the richer table per section.

    When the same site is supplied from more than one source (e.g. a PDF *and* a
    pre-converted Excel), each section takes whichever copy has more rows — so
    you get the best Solution Antenna Details and the best Solution Sector
    Details regardless of which source they came from.
    """
    by_site: dict[str, TssrResult] = {}
    for r in results:
        key = r.site_id.upper()
        cur = by_site.get(key)
        if cur is None:
            by_site[key] = r
            continue
        merged = {}
        for name in SECTION_SCHEMAS:
            a, b = cur.sections[name], r.sections[name]
            merged[name] = a if len(a) >= len(b) else b
        by_site[key] = TssrResult(
            site_id=cur.site_id, sections=merged,
            log=cur.log + r.log, warnings=cur.warnings + r.warnings,
        )
    return list(by_site.values())


def combine_sections(results: list[TssrResult]) -> dict[str, pd.DataFrame]:
    """Stack each section across all extracted PDFs into one DataFrame."""
    combined: dict[str, pd.DataFrame] = {}
    for name, schema in SECTION_SCHEMAS.items():
        cols = META_COLS + schema
        frames = [r.sections[name] for r in results if not r.sections[name].empty]
        combined[name] = (
            pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=cols)
        )
    return combined


def build_log(results: list[TssrResult]) -> pd.DataFrame:
    rows: list[dict] = []
    for r in results:
        for entry in r.log:
            rows.append({"Site ID": r.site_id, **entry})
    return pd.DataFrame(rows, columns=["Site ID", "Source Page", "Section",
                                       "Sector No", "Antenna No", "Fields", "Note"])


def to_workbook_bytes(results: list[TssrResult]) -> bytes:
    """Write the combined sections to a multi-sheet .xlsx (one sheet per topic)."""
    sections = combine_sections(results)
    log = build_log(results)
    summary = pd.DataFrame(
        [{"Site ID": r.site_id,
          **{name: len(r.sections[name]) for name in SECTION_SCHEMAS}}
         for r in results],
        columns=["Site ID", *SECTION_SCHEMAS.keys()],
    )

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        for name, df in sections.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        log.to_excel(writer, sheet_name="Extraction Log", index=False)
    return out.getvalue()
