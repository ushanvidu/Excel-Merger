"""Fill the master 'RF Parameter' sheet's TSSR columns from TSSR PDF data.

Decisions confirmed with the user:
  - Source     : the **Solution** tables (Solution Antenna/Sector Details).
  - Match key  : **Sector + Band** (band->PDF-token mapping is configurable).
  - TSSR Status: left blank.

Join, per master row (one Sector x Tech x Band):
  1. Determine the sector number from the master 'Sector' (S1->1); if that cell
     is broken (#REF!/blank), recover it from 'Antenna azimuth_Mbitel' which
     matches the PDF 'Antenna Direction'.
  2. Match a Solution **Sector Details** row by sector + band token  -> tilts
     (Electrical Tilt 01, Mechanical Tilt) and the Antenna No.
  3. Look up the Solution **Antenna Details** row by sector (+ that Antenna No,
     preferring one whose ports mention the band)  -> azimuth, height, type.
Only mapped cells are written; everything else in the workbook is untouched.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

from openpyxl import load_workbook

from .tssr_extractor import TssrResult, merge_results
from .utils import clean_cell, coerce_value

# Master column header -> the TSSR field it should receive.
MASTER_TSSR_COLUMNS = [
    "Antenna_TSSR", "Antenna height_TSSR", "Antenna azimuth_TSSR",
    "Mechanical tilt_TSSR", "Electrical tilt_TSSR", "TSSR Status",
]

# Default master-Band -> regex matched against the Solution Sector 'Sector' token
# (e.g. "L850_1", "1800|L1800_P", "UMTS_W", "900_A", "L2.3_sec1", "L2.6_sec1").
DEFAULT_BAND_MAP = {
    "700": r"700|n28",
    "850": r"l?\s*850",
    "900": r"900",
    "1800": r"1800|l1800",
    "2100": r"umts|2100|l2100|w2100",
    "2300": r"l2\.?3|2300",
    "2600": r"l2\.?6|2600",
    "3500": r"5g|n78|3500|n3500",
}

_SECTOR_NUM_RE = re.compile(r"(?i)s(?:ector)?\s*0*(\d+)")


@dataclass
class FillReport:
    rows_filled: int = 0
    rows_unmatched: int = 0
    cells_written: dict[str, int] = field(default_factory=dict)
    details: list[dict] = field(default_factory=list)          # filled rows
    unmatched_details: list[dict] = field(default_factory=list)  # rows not filled
    sites_seen: set[str] = field(default_factory=set)
    sites_no_pdf: set[str] = field(default_factory=set)


def sector_num(value) -> int | None:
    m = _SECTOR_NUM_RE.search(clean_cell(value))
    return int(m.group(1)) if m else None


def _band_token(band) -> str:
    return clean_cell(band)


def _direction_to_sector(result: TssrResult) -> dict[str, int]:
    """Map a PDF 'Antenna Direction' value -> sector number (for #REF! recovery).

    Built from both Existing and Solution antenna tables, since the master's
    'Antenna azimuth_Mbitel' tracks the existing on-site direction.
    """
    out: dict[str, int] = {}
    for section in ("Antenna Details", "Solution Antenna Details"):
        df = result.sections.get(section)
        if df is None:
            continue
        for _, r in df.iterrows():
            sn = sector_num(r["Sector No"])
            d = clean_cell(r["Antenna Direction"])
            if sn and d and d not in out:
                out[d] = sn
    return out


def _pdf_sector_set(result: TssrResult) -> set[int]:
    """All sector numbers present in a site's PDF (to detect 0- vs 1-based)."""
    out: set[int] = set()
    for section in result.sections.values():
        if "Sector No" in section.columns:
            for v in section["Sector No"]:
                n = sector_num(v)
                if n is not None:
                    out.add(n)
    return out


def _match_sector_row(df, sn: int, band_rx: re.Pattern):
    """First Solution Sector Details row for this sector whose token matches band."""
    for _, r in df.iterrows():
        if sector_num(r["Sector No"]) != sn:
            continue
        if band_rx.search(clean_cell(r["Sector"])):
            return r
    return None


def _match_antenna_row(df, sn: int, antenna_no: str | None, band: str):
    """Solution Antenna Details row for this sector, preferring antenna_no and a
    port that mentions the band; falls back to the first sector match."""
    candidates = [r for _, r in df.iterrows() if sector_num(r["Sector No"]) == sn]
    if not candidates:
        return None
    pool = candidates
    if antenna_no:
        an = sector_num(antenna_no)
        named = [r for r in pool if sector_num(r["Antenna No"]) == an]
        if named:
            pool = named
    if band:
        ports = [r for r in pool
                 if any(band in clean_cell(r[f"Port 0{i}"]) for i in range(1, 6))]
        if ports:
            return ports[0]
    return pool[0]


def fill_master(
    master_bytes: bytes,
    results: list[TssrResult],
    *,
    sheet_name: str = "RF Parameter",
    header_row: int = 1,
    site_col: str = "Site Name",
    source: str = "Solution",
    band_map: dict[str, str] | None = None,
) -> tuple[bytes, FillReport]:
    """Return (filled_xlsx_bytes, report)."""
    band_map = band_map or DEFAULT_BAND_MAP
    band_rx = {b: re.compile(rx, re.I) for b, rx in band_map.items()}
    by_site = {r.site_id.upper(): r for r in merge_results(results)}
    report = FillReport()

    wb = load_workbook(io.BytesIO(master_bytes))
    ws = wb[sheet_name]

    # Locate columns by header.
    header = {clean_cell(c.value): c.column for c in ws[header_row]}
    needed = [site_col, "Sector", "Band", "Antenna azimuth_Mbitel", *MASTER_TSSR_COLUMNS]
    missing = [c for c in needed if c not in header]
    if missing:
        raise ValueError(f"Master sheet missing columns: {missing}")
    col = {name: header[name] for name in needed}

    # Pre-pass: build a per-site azimuth -> sector map. Seeded from the PDF
    # directions, then augmented from master rows whose 'Sector' is valid (with
    # the 0/1-based offset applied) — so #REF! rows that share an azimuth with a
    # resolved sibling row still resolve, even when the surveyed direction value
    # differs from the master's azimuth.
    site_rows: dict[str, list[tuple[int | None, str]]] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        s = clean_cell(ws.cell(r, col[site_col]).value)
        if not s or s.upper() not in by_site:
            continue
        site_rows.setdefault(s, []).append((
            sector_num(ws.cell(r, col["Sector"]).value),
            clean_cell(ws.cell(r, col["Antenna azimuth_Mbitel"]).value),
        ))

    az_maps: dict[str, dict[str, int]] = {}
    offsets: dict[str, int] = {}
    for s, rows in site_rows.items():
        result = by_site[s.upper()]
        dmap = _direction_to_sector(result)
        sset = _pdf_sector_set(result)
        offset = -1 if (sset and min(sset) == 0) else 0
        for raw, az in rows:
            if az and az not in dmap and raw is not None:
                sec = raw + offset
                if not sset or sec in sset:
                    dmap[az] = sec
        az_maps[s], offsets[s] = dmap, offset

    fills = {c: 0 for c in MASTER_TSSR_COLUMNS if c != "TSSR Status"}

    for row in range(header_row + 1, ws.max_row + 1):
        site = clean_cell(ws.cell(row, col[site_col]).value)
        if not site:
            continue
        result = by_site.get(site.upper())
        report.sites_seen.add(site)
        if result is None:
            report.sites_no_pdf.add(site)
            continue

        # 1. Resolve the PDF sector number via the augmented azimuth map, then
        #    fall back to the master 'Sector' number with the site's offset.
        azimuth = clean_cell(ws.cell(row, col["Antenna azimuth_Mbitel"]).value)
        sn = az_maps.get(site, {}).get(azimuth)
        if sn is None:
            raw = sector_num(ws.cell(row, col["Sector"]).value)
            if raw is not None:
                sn = raw + offsets.get(site, 0)
        band = _band_token(ws.cell(row, col["Band"]).value)
        if sn is None or not band:
            report.rows_unmatched += 1
            report.unmatched_details.append({
                "row": row, "site": site, "sector": sn or "", "band": band,
                "reason": "sector not resolved" if sn is None else "no band value",
            })
            continue

        sec_df = result.sections[f"{source} Sector Details"]
        ant_df = result.sections[f"{source} Antenna Details"]

        rx = band_rx.get(band)
        sec_row = _match_sector_row(sec_df, sn, rx) if rx is not None else None
        antenna_no = clean_cell(sec_row["Antenna No"]) if sec_row is not None else None
        ant_row = _match_antenna_row(ant_df, sn, antenna_no, band)

        values: dict[str, object] = {}
        if ant_row is not None:
            values["Antenna_TSSR"] = clean_cell(ant_row["Antenna Type"])
            values["Antenna height_TSSR"] = coerce_value(ant_row["Antenna Height (m)"], "number")
            values["Antenna azimuth_TSSR"] = coerce_value(ant_row["Antenna Direction"], "number")
        if sec_row is not None:
            values["Mechanical tilt_TSSR"] = coerce_value(sec_row["Mechanical Tilt"], "number")
            values["Electrical tilt_TSSR"] = coerce_value(sec_row["Electrical Tilt 01"], "number")

        if not values:
            report.rows_unmatched += 1
            report.unmatched_details.append({
                "row": row, "site": site, "sector": sn, "band": band,
                "reason": "no Solution row for this sector+band",
            })
            continue

        for name, val in values.items():
            if val is None or (isinstance(val, str) and val == ""):
                continue
            ws.cell(row, col[name], val)
            fills[name] += 1
        report.rows_filled += 1
        report.details.append({
            "row": row, "site": site, "sector": sn, "band": band,
            "matched_sector_token": clean_cell(sec_row["Sector"]) if sec_row is not None else "",
            "antenna": antenna_no or "",
            "azimuth": values.get("Antenna azimuth_TSSR", ""),
            "height": values.get("Antenna height_TSSR", ""),
            "antenna_type": values.get("Antenna_TSSR", ""),
            "mech_tilt": values.get("Mechanical tilt_TSSR", ""),
            "elec_tilt": values.get("Electrical tilt_TSSR", ""),
            "status": "filled",
        })

    report.cells_written = {k: v for k, v in fills.items() if v}
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), report
