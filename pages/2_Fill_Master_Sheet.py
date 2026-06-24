"""Fill master sheet — fills the RF Parameter TSSR columns from TSSR data.

Sources: TSSR PDFs (parsed live) and/or an already-converted multi-sheet Excel
(e.g. merged_antenna_sector_details). When both supply the same site, each
section keeps the richer table (best-of-both). Each master row (Sector x Band)
is matched to the site's Solution antenna/sector data and the TSSR columns are
filled: Antenna_TSSR / height / azimuth (from Solution Antenna Details) and
Mechanical/Electrical tilt (from Solution Sector Details). TSSR Status is left
blank. Only those cells are written; formatting and other sheets are preserved.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from pdf_excel_merger import excel_reader
from pdf_excel_merger.tssr_extractor import extract_many, results_from_workbook
from pdf_excel_merger.tssr_fill import DEFAULT_BAND_MAP, fill_master

st.set_page_config(page_title="Fill Master Sheet", page_icon="🗂️", layout="wide")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

SS = st.session_state
SS.setdefault("fill_bytes", None)
SS.setdefault("fill_report", None)

st.title("🗂️ Fill master sheet from TSSR data")
st.caption(
    "Upload the master workbook plus TSSR sources — PDFs and/or an "
    "already-converted Excel. Each master row is matched by **Site + Sector + "
    "Band** to its Solution data and the five TSSR columns are filled."
)

master_file = st.file_uploader("Master workbook (.xlsx)", type=["xlsx", "xlsm"])
c1, c2 = st.columns(2)
with c1:
    pdf_files = st.file_uploader("TSSR PDFs (one site each)", type=["pdf"],
                                 accept_multiple_files=True)
with c2:
    conv_files = st.file_uploader("Already-converted Excel (optional)",
                                  type=["xlsx", "xlsm"], accept_multiple_files=True)

master_bytes = master_file.getvalue() if master_file else None

if master_bytes:
    try:
        sheets = excel_reader.list_sheets(io.BytesIO(master_bytes))
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read workbook: {exc}")
        sheets = []

    if sheets:
        st.subheader("Settings")
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            idx = sheets.index("RF Parameter") if "RF Parameter" in sheets else 0
            sheet_name = st.selectbox("Sheet to fill", sheets, index=idx)
        with s2:
            header_row = st.number_input("Header row", 1, 100, 1, 1)
        with s3:
            source = st.selectbox("Source tables", ["Solution", "Existing"], index=0)
        with s4:
            site_col = st.text_input("Site key column", value="Site Name")

        st.markdown("**Band → PDF token mapping** (regex matched against the "
                    "Solution 'Sector' value — edit for your data)")
        band_df = st.data_editor(
            pd.DataFrame([{"Band": b, "Pattern": p} for b, p in DEFAULT_BAND_MAP.items()]),
            num_rows="dynamic", use_container_width=True, key="band_map",
        )

        has_source = bool(pdf_files) or bool(conv_files)
        if has_source and st.button("🔗 Match & fill", type="primary"):
            band_map = {
                str(r["Band"]).strip(): str(r["Pattern"]).strip()
                for _, r in band_df.iterrows()
                if str(r.get("Band", "")).strip() and str(r.get("Pattern", "")).strip()
            }
            try:
                with st.spinner("Reading sources and filling…"):
                    results = []
                    if pdf_files:
                        results += extract_many([(f.name, f.getvalue()) for f in pdf_files])
                    for f in conv_files or []:
                        results += results_from_workbook(f.getvalue())
                    out_bytes, report = fill_master(
                        master_bytes, results,
                        sheet_name=sheet_name, header_row=int(header_row),
                        site_col=site_col, source=source, band_map=band_map,
                    )
                SS["fill_bytes"] = out_bytes
                SS["fill_report"] = report
            except Exception as exc:  # noqa: BLE001
                st.error(f"Fill failed: {exc}")
        elif not has_source:
            st.info("Add at least one source (PDF or converted Excel) to fill.")

report = SS.get("fill_report")
if report is not None:
    st.header("Result")
    total = report.rows_filled + report.rows_unmatched
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Rows filled", f"{report.rows_filled}/{total}" if total else "0")
    m2.metric("Rows unmatched", report.rows_unmatched)
    m3.metric("Sites with data", len(report.sites_seen) - len(report.sites_no_pdf))
    m4.metric("Sites w/o data", len(report.sites_no_pdf))

    if report.cells_written:
        st.write("**Cells written per column:**")
        st.table(pd.DataFrame(
            [{"Column": k, "Filled": v} for k, v in report.cells_written.items()]
        ))

    st.download_button(
        "⬇️ Download filled master", data=SS["fill_bytes"],
        file_name="Noding_Config_filled.xlsx", mime=XLSX_MIME, type="primary",
    )

    if report.details:
        with st.expander(f"✅ Filled rows ({len(report.details)})"):
            st.dataframe(pd.DataFrame(report.details), use_container_width=True,
                         hide_index=True)
    if report.unmatched_details:
        with st.expander(f"⚠️ Unmatched rows ({len(report.unmatched_details)}) — "
                         "with reason", expanded=True):
            st.dataframe(pd.DataFrame(report.unmatched_details),
                         use_container_width=True, hide_index=True)
    if report.sites_no_pdf:
        with st.expander(f"Sites in the sheet with no TSSR data ({len(report.sites_no_pdf)})"):
            st.write(", ".join(sorted(report.sites_no_pdf)))
