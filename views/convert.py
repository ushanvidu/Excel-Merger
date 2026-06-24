"""Convert TSSR to Excel — turn TSSR PDFs into clean, topic-divided sheets."""

from __future__ import annotations

import streamlit as st

from pdf_excel_merger.tssr_extractor import (
    SECTION_SCHEMAS,
    build_log,
    combine_sections,
    extract_many,
    to_workbook_bytes,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SS = st.session_state
SS.setdefault("tssr_results", None)


def render() -> None:
    st.title("📄 Convert TSSR to Excel")
    st.caption(
        "Upload TSSR survey PDFs. Each is parsed into clean, topic-divided tables "
        "(Antenna / Sector / Solution details) and combined into one workbook."
    )

    uploads = st.file_uploader("TSSR PDF documents", type=["pdf"],
                               accept_multiple_files=True)
    if uploads and st.button("Extract & convert", type="primary", use_container_width=True):
        with st.spinner(f"Extracting {len(uploads)} PDF(s)…"):
            SS["tssr_results"] = extract_many([(f.name, f.getvalue()) for f in uploads])

    results = SS["tssr_results"]
    if not results:
        return

    st.divider()
    st.subheader("Summary")
    st.dataframe(
        [{"Site ID": r.site_id, **{n: len(r.sections[n]) for n in SECTION_SCHEMAS}}
         for r in results],
        hide_index=True, use_container_width=True)

    warnings = sum(len(r.warnings) for r in results)
    if warnings:
        with st.expander(f"⚠️ {warnings} extraction warning(s)"):
            for r in results:
                for w in r.warnings:
                    st.write(f"**{r.site_id}** — {w}")

    st.download_button("⬇️ Download multi-sheet Excel", data=to_workbook_bytes(results),
                       file_name="tssr_converted.xlsx", mime=XLSX_MIME,
                       type="primary", use_container_width=True)

    st.subheader("Topic sheets")
    combined = combine_sections(results)
    tabs = st.tabs(list(SECTION_SCHEMAS) + ["Extraction Log"])
    for tab, name in zip(tabs, SECTION_SCHEMAS):
        with tab:
            st.caption(f"{len(combined[name])} row(s)")
            st.dataframe(combined[name], use_container_width=True, hide_index=True)
    with tabs[-1]:
        log = build_log(results)
        st.caption(f"{len(log)} extracted row(s)")
        st.dataframe(log, use_container_width=True, hide_index=True)


render()
