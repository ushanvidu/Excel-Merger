"""Fill Master Sheet — fill a master sheet's columns from TSSR PDFs / Excel.

Upload the master and the site documents, choose which columns to fill and where
each one comes from, then download the filled sheet. Works with any master that
has a site key + sector + band, and any TSSR PDFs (not just the samples).
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

from pdf_excel_merger import excel_reader
from pdf_excel_merger.tssr_extractor import extract_many, results_from_workbook
from pdf_excel_merger.tssr_fill import (
    ANTENNA_FIELDS,
    DEFAULT_BAND_MAP,
    DEFAULT_FIELD_MAP,
    SECTOR_FIELDS,
    fill_master,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SS = st.session_state
SS.setdefault("fill_bytes", None)
SS.setdefault("fill_report", None)

# "Antenna · Field" / "Sector · Field" labels for the source-field dropdowns.
SOURCE_OPTIONS = (
    [f"Antenna · {f}" for f in ANTENNA_FIELDS]
    + [f"Sector · {f}" for f in SECTOR_FIELDS]
)
DEFAULT_BY_TARGET = {s["target"]: s for s in DEFAULT_FIELD_MAP}


def _parse_source(label: str) -> tuple[str, str]:
    section, field = label.split(" · ", 1)
    return section.strip().lower(), field.strip()


def _suggest_source(target: str) -> str:
    """Best source-field label for a master column (defaults + fuzzy fallback)."""
    if target in DEFAULT_BY_TARGET:
        spec = DEFAULT_BY_TARGET[target]
        return f"{spec['section'].capitalize()} · {spec['field']}"
    best, best_score = SOURCE_OPTIONS[0], -1.0
    for opt in SOURCE_OPTIONS:
        score = fuzz.token_set_ratio(target.lower(), _parse_source(opt)[1].lower())
        if score > best_score:
            best, best_score = opt, score
    return best


def _suggest_kind(target: str) -> str:
    t = target.lower()
    if any(k in t for k in ("tilt", "height", "azimuth", "direction", "angle")):
        return "number"
    return "auto"


def render() -> None:
    st.title("🗂️ Fill Master Sheet")
    st.caption(
        "Match each site document to the master sheet by **Site + Sector + Band** "
        "and fill the columns you choose. Only those cells are written — existing "
        "formatting, formulas and other sheets are preserved."
    )

    # ---- Step 1: uploads ----
    st.subheader("1 · Upload files")
    c1, c2 = st.columns(2)
    with c1:
        master_file = st.file_uploader("Master sheet (.xlsx)", type=["xlsx", "xlsm"])
    with c2:
        pdf_files = st.file_uploader("TSSR PDFs (one site each, many at once)",
                                     type=["pdf"], accept_multiple_files=True)
    conv_files = st.file_uploader(
        "Already-converted Excel (optional — combined best-of-both with PDFs)",
        type=["xlsx", "xlsm"], accept_multiple_files=True)

    if not master_file:
        st.info("Upload a master sheet to begin.")
        return
    master_bytes = master_file.getvalue()

    try:
        sheets = excel_reader.list_sheets(io.BytesIO(master_bytes))
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read the workbook: {exc}")
        return

    # ---- Step 2: settings ----
    st.subheader("2 · Sheet & matching")
    idx = sheets.index("RF Parameter") if "RF Parameter" in sheets else 0
    sc1, sc2 = st.columns([2, 1])
    sheet_name = sc1.selectbox("Sheet to fill", sheets, index=idx)
    header_row = sc2.number_input("Header row", 1, 100, 1, 1)

    try:
        info = excel_reader.read_template(io.BytesIO(master_bytes), sheet_name, int(header_row))
        master_cols = info.column_names
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read sheet '{sheet_name}': {exc}")
        return
    if not master_cols:
        st.warning(f"No column headers found on row {header_row}. Adjust the header row.")
        return

    def pick(label, options, prefer):
        default = next((o for o in options if o == prefer), options[0])
        return st.selectbox(label, options, index=options.index(default))

    with st.expander("Key columns & data source", expanded=False):
        k1, k2, k3 = st.columns(3)
        with k1:
            site_col = pick("Site key column", master_cols, "Site Name")
        with k2:
            sector_col = pick("Sector column", master_cols, "Sector")
        with k3:
            band_col = pick("Band column", master_cols, "Band")
        a1, a2 = st.columns(2)
        with a1:
            azimuth_col = pick("Existing azimuth column", master_cols, "Antenna azimuth_Mbitel")
        with a2:
            source = st.selectbox("Source tables", ["Solution", "Existing"], index=0)

    # ---- Step 3: columns to fill + mapping ----
    st.subheader("3 · Columns to fill")
    detected = [c for c in master_cols if "tssr" in c.lower()] or list(DEFAULT_BY_TARGET)
    detected = [c for c in detected if c in master_cols]
    fill_cols = st.multiselect(
        "Master columns to fill", master_cols, default=detected,
        help="Auto-detected the TSSR columns. Add or remove any column.")

    field_map: list[dict] = []
    if fill_cols:
        st.caption("Map each column to the document field that fills it:")
        h = st.columns([3, 4, 2])
        h[0].markdown("**Master column**")
        h[1].markdown("**Document field**")
        h[2].markdown("**Value type**")
        for tgt in fill_cols:
            c0, c1, c2 = st.columns([3, 4, 2])
            c0.write(tgt)
            default_src = _suggest_source(tgt)
            src = c1.selectbox(tgt, SOURCE_OPTIONS, index=SOURCE_OPTIONS.index(default_src),
                               label_visibility="collapsed", key=f"src_{tgt}")
            kinds = ["auto", "number", "text"]
            kind = c2.selectbox(tgt, kinds, index=kinds.index(_suggest_kind(tgt)),
                                label_visibility="collapsed", key=f"kind_{tgt}")
            section, field = _parse_source(src)
            field_map.append({"target": tgt, "section": section, "field": field, "kind": kind})

    with st.expander("Band → document-token mapping (advanced)", expanded=False):
        st.caption("Regex matched against the Solution 'Sector' token (e.g. `L850_1`).")
        band_df = st.data_editor(
            pd.DataFrame([{"Band": b, "Pattern": p} for b, p in DEFAULT_BAND_MAP.items()]),
            num_rows="dynamic", use_container_width=True, key="band_map")

    # ---- Step 4: run ----
    st.subheader("4 · Generate")
    has_source = bool(pdf_files) or bool(conv_files)
    if not has_source:
        st.info("Add at least one source (PDF or converted Excel).")
        return
    if not field_map:
        st.info("Select at least one column to fill.")
        return

    if st.button("Fill master sheet", type="primary", use_container_width=True):
        band_map = {str(r["Band"]).strip(): str(r["Pattern"]).strip()
                    for _, r in band_df.iterrows()
                    if str(r.get("Band", "")).strip() and str(r.get("Pattern", "")).strip()}
        try:
            with st.spinner("Reading documents and filling…"):
                results = []
                if pdf_files:
                    results += extract_many([(f.name, f.getvalue()) for f in pdf_files])
                for f in conv_files or []:
                    results += results_from_workbook(f.getvalue())
                out_bytes, report = fill_master(
                    master_bytes, results, sheet_name=sheet_name, header_row=int(header_row),
                    site_col=site_col, sector_col=sector_col, azimuth_col=azimuth_col,
                    band_col=band_col, source=source, band_map=band_map, field_map=field_map)
            SS["fill_bytes"], SS["fill_report"] = out_bytes, report
        except Exception as exc:  # noqa: BLE001
            st.error(f"Fill failed: {exc}")

    _render_result()


def _render_result() -> None:
    report = SS.get("fill_report")
    if report is None:
        return
    st.divider()
    st.subheader("Result")
    total = report.rows_filled + report.rows_unmatched
    m = st.columns(4)
    m[0].metric("Rows filled", f"{report.rows_filled}/{total}" if total else "0")
    m[1].metric("Unmatched", report.rows_unmatched)
    m[2].metric("Sites with data", len(report.sites_seen) - len(report.sites_no_pdf))
    m[3].metric("Flagged for review", len(report.conflicts))

    st.download_button("⬇️ Download filled master", data=SS["fill_bytes"],
                       file_name="master_filled.xlsx", mime=XLSX_MIME,
                       type="primary", use_container_width=True)

    if report.cells_written:
        st.write("**Cells written per column**")
        st.dataframe(pd.DataFrame([{"Column": k, "Filled": v}
                                   for k, v in report.cells_written.items()]),
                     hide_index=True, use_container_width=True)
    if report.conflicts:
        with st.expander(f"🚩 Sector conflicts to review ({len(report.conflicts)})", expanded=True):
            st.caption("The band-token suffix disagrees with the matched sector here — "
                       "the source data is ambiguous. Verify these rows against the PDF.")
            st.dataframe(pd.DataFrame(report.conflicts), hide_index=True, use_container_width=True)
    if report.details:
        with st.expander(f"Filled rows ({len(report.details)})"):
            st.dataframe(pd.DataFrame(report.details), hide_index=True, use_container_width=True)
    if report.unmatched_details:
        with st.expander(f"Unmatched rows ({len(report.unmatched_details)})"):
            st.dataframe(pd.DataFrame(report.unmatched_details), hide_index=True,
                         use_container_width=True)
    if report.sites_no_pdf:
        with st.expander(f"Sites in the sheet with no document ({len(report.sites_no_pdf)})"):
            st.write(", ".join(sorted(report.sites_no_pdf)))


render()
