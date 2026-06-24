"""TSSR PDF → Excel (topic sheets) — converts InfraMS TSSR survey PDFs into a
clean multi-sheet workbook (Antenna Details, Sector Details, Solution Antenna
Details, Solution Sector Details, Extraction Log).

This is the explicit "convert PDF to Excel, divide sheets by topic" step.
"""

from __future__ import annotations

import streamlit as st

from pdf_excel_merger.tssr_extractor import (
    SECTION_SCHEMAS,
    build_log,
    combine_sections,
    extract_many,
    to_workbook_bytes,
)

st.set_page_config(page_title="TSSR → Excel", page_icon="📡", layout="wide")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

SS = st.session_state
SS.setdefault("tssr_results", None)

st.title("📡 TSSR PDF → Excel (topic sheets)")
st.caption(
    "Upload one or more InfraMS TSSR survey PDFs. Each is parsed into clean, "
    "topic-divided tables and combined into a single downloadable workbook."
)

uploads = st.file_uploader(
    "TSSR PDF documents (one site each, many at once)",
    type=["pdf"], accept_multiple_files=True,
)

if uploads and st.button("🔄 Extract & convert", type="primary"):
    with st.spinner(f"Extracting {len(uploads)} PDF(s)…"):
        files = [(f.name, f.getvalue()) for f in uploads]
        SS["tssr_results"] = extract_many(files)

results = SS["tssr_results"]
if results:
    st.header("Summary")
    summary_rows = [
        {"Site ID": r.site_id, **{name: len(r.sections[name]) for name in SECTION_SCHEMAS}}
        for r in results
    ]
    st.dataframe(summary_rows, use_container_width=True)

    total_warnings = sum(len(r.warnings) for r in results)
    if total_warnings:
        with st.expander(f"⚠️ {total_warnings} extraction warning(s)"):
            for r in results:
                for w in r.warnings:
                    st.write(f"**{r.site_id}** — {w}")

    st.download_button(
        "⬇️ Download multi-sheet Excel",
        data=to_workbook_bytes(results),
        file_name="tssr_converted.xlsx",
        mime=XLSX_MIME, type="primary",
    )

    st.header("Topic sheets")
    combined = combine_sections(results)
    tabs = st.tabs(list(SECTION_SCHEMAS.keys()) + ["Extraction Log"])
    for tab, name in zip(tabs, SECTION_SCHEMAS):
        with tab:
            df = combined[name]
            st.caption(f"{len(df)} row(s)")
            st.dataframe(df, use_container_width=True, hide_index=True)
    with tabs[-1]:
        log = build_log(results)
        st.caption(f"{len(log)} extracted row(s)")
        st.dataframe(log, use_container_width=True, hide_index=True)
