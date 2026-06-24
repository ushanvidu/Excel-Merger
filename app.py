"""PDF/Excel → Master Sheet site filler — Streamlit UI.

Workflow: upload a master sheet (one row per network site) + per-site documents
(PDFs are converted to Excel automatically; Excel/CSV are read directly). The app
identifies each document's site, matches it to the right master-sheet row, and
fills the mapped fields into that row.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from pdf_excel_merger import excel_reader, mapping as mapping_mod, merger, site_matching
from pdf_excel_merger.pdf_extractor import SOURCE_COLUMN
from pdf_excel_merger.source_loader import (
    dataframe_to_excel_bytes,
    file_kind,
    load_sources,
)

st.set_page_config(page_title="Site Sheet Filler", page_icon="📡", layout="wide")

SS = st.session_state
SS.setdefault("loaded", None)        # LoadedSources
SS.setdefault("source_df", None)     # editable combined frame
SS.setdefault("result_bytes", None)
SS.setdefault("result_report", None)


def reset_results() -> None:
    SS["result_bytes"] = None
    SS["result_report"] = None


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------------------
st.title("📡 Network Site Sheet Filler")
st.caption(
    "Upload your master sheet (one row per site) and per-site documents. PDFs are "
    "converted to Excel automatically; the app finds each document's site and fills "
    "its row in the master sheet."
)
st.info(
    "**TSSR workflow?** Use the sidebar pages: **TSSR to Excel** (convert TSSR PDFs "
    "into topic-divided sheets) and **Fill Master Sheet** (fill the RF Parameter "
    "TSSR columns from TSSR PDFs). This Home page is the generic key-based merge."
)

# ===========================================================================
# Step 1 — Master sheet
# ===========================================================================
st.header("1 · Master sheet (all sites)")
master_file = st.file_uploader(
    "The Excel sheet to fill — contains one row per site",
    type=["xlsx", "xlsm"], accept_multiple_files=False,
)

master_bytes = master_file.getvalue() if master_file else None
master_info = None
sheet_name = None

if master_bytes:
    try:
        sheets = excel_reader.list_sheets(io.BytesIO(master_bytes))
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read the master sheet: {exc}")
        sheets = []
    if sheets:
        c1, c2 = st.columns(2)
        with c1:
            sheet_name = st.selectbox("Sheet", sheets, on_change=reset_results)
        with c2:
            suggested_hr = excel_reader.detect_header_row(io.BytesIO(master_bytes), sheet_name)
            header_row = st.number_input(
                "Header row", min_value=1, max_value=100, value=suggested_hr, step=1,
                on_change=reset_results,
            )
        try:
            master_info = excel_reader.read_template(
                io.BytesIO(master_bytes), sheet_name, int(header_row)
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not read sheet '{sheet_name}': {exc}")
        if master_info and master_info.columns:
            st.success(
                f"{len(master_info.columns)} columns · "
                f"{master_info.last_filled_row - master_info.header_row} existing site rows: "
                + ", ".join(master_info.column_names)
            )
        elif master_info:
            st.warning(f"No headers found on row {header_row}. Adjust the header row.")

# ===========================================================================
# Step 2 — Site documents (PDF and/or Excel/CSV)
# ===========================================================================
st.header("2 · Site documents (one site per file)")
d1, d2 = st.columns(2)
with d1:
    pdf_uploads = st.file_uploader(
        "📄 PDF documents  —  converted to Excel automatically",
        type=["pdf"], accept_multiple_files=True,
    )
with d2:
    excel_uploads = st.file_uploader(
        "📊 Already-converted Excel / CSV",
        type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
    )

all_uploads = list(pdf_uploads or []) + list(excel_uploads or [])

# ===========================================================================
# Step 3 — Conversion settings + run
# ===========================================================================
if all_uploads:
    st.header("3 · Convert & load")
    s1, s2 = st.columns(2)
    with s1:
        orientation = st.selectbox(
            "Document layout",
            ["auto", "key_value", "records"],
            format_func=lambda v: {
                "auto": "Auto-detect",
                "key_value": "Key/value form (Parameter | Value)",
                "records": "Table of records (header + rows)",
            }[v],
            help="Most single-site forms are key/value. 'auto' treats any "
                 "2-column document as key/value.",
        )
    with s2:
        strategy = st.selectbox(
            "PDF table style", ["auto", "lines", "text"],
            help="auto: bordered tables, fall back to text alignment.",
        )

    if st.button("🔄 Convert & load documents", type="primary"):
        with st.spinner("Converting PDFs and reading files…"):
            files = [(f.name, f.getvalue()) for f in all_uploads]
            SS["loaded"] = load_sources(files, orientation=orientation, pdf_strategy=strategy)
            SS["source_df"] = SS["loaded"].combined
        reset_results()

loaded = SS["loaded"]

# ===========================================================================
# Step 4 — Converted data (per file) + combined review
# ===========================================================================
if loaded is not None:
    st.header("4 · Converted data")
    for w in loaded.warnings:
        st.warning(w)

    if not loaded.converted:
        st.error("No data could be read from the uploaded documents.")
    else:
        st.caption("Each PDF has been converted to a table. Download any as Excel:")
        for name, conv in loaded.converted.items():
            kind = file_kind(name)
            icon = "📄" if kind == "pdf" else "📊"
            with st.expander(f"{icon} {name}  ({len(conv)} row(s))"):
                if conv.empty:
                    st.warning("No table detected — try a different PDF table style.")
                else:
                    st.dataframe(conv, use_container_width=True)
                    st.download_button(
                        "⬇️ Download as Excel", key=f"dl_{name}",
                        data=dataframe_to_excel_bytes(conv),
                        file_name=f"{name.rsplit('.', 1)[0]}_converted.xlsx",
                        mime=XLSX_MIME,
                    )

        src_df = SS["source_df"]
        if src_df is not None and not src_df.empty:
            st.subheader("Combined records — edit if needed before matching")
            st.caption(
                "One row per site document. Fix any wrong values or site keys here."
            )
            edited = st.data_editor(
                src_df, use_container_width=True, num_rows="dynamic", key="editor"
            )
            SS["source_df"] = edited

# ===========================================================================
# Steps 5-7 — Matching, mapping, generate
# ===========================================================================
source_df: pd.DataFrame | None = SS["source_df"]

if (
    source_df is not None and not source_df.empty
    and master_info is not None and master_info.columns
):
    source_cols = [c for c in source_df.columns if c != SOURCE_COLUMN]

    # ---- Step 5: key columns + match preview ----
    st.header("5 · Match documents to sites")
    k1, k2 = st.columns(2)
    with k1:
        master_key_default = site_matching.detect_key_column(master_info.column_names)
        main_key_col = st.selectbox(
            "Master sheet — site key column",
            master_info.column_names,
            index=master_info.column_names.index(master_key_default)
            if master_key_default in master_info.column_names else 0,
            help="The column that uniquely identifies a site (Site ID / Name).",
        )
    with k2:
        source_key_default = site_matching.detect_key_column(source_cols)
        source_key_col = st.selectbox(
            "Source documents — site key field",
            source_cols,
            index=source_cols.index(source_key_default)
            if source_key_default in source_cols else 0,
            help="The field in each document that holds its site ID / name.",
        )

    fz1, fz2 = st.columns([1, 2])
    with fz1:
        allow_fuzzy = st.checkbox(
            "Allow fuzzy name matching", value=True,
            help="For site names with small spelling/format differences. "
                 "IDs always match exactly first.",
        )
    with fz2:
        threshold = st.slider(
            "Fuzzy match threshold", min_value=70, max_value=100,
            value=int(site_matching.DEFAULT_FUZZY_THRESHOLD), disabled=not allow_fuzzy,
        )

    match_preview = merger.preview_key_matches(
        master_bytes, master_info, source_df,
        main_key_col=main_key_col, source_key_col=source_key_col,
        allow_fuzzy=allow_fuzzy, fuzzy_threshold=threshold,
    )
    matched_n = int((match_preview["Match"].isin(["exact", "fuzzy"])).sum())
    unmatched_n = len(match_preview) - matched_n
    st.dataframe(match_preview, use_container_width=True)
    if unmatched_n:
        st.warning(f"{matched_n} matched · {unmatched_n} not found in the master sheet.")
    else:
        st.success(f"All {matched_n} documents matched a site in the master sheet.")

    # ---- Step 6: field mapping ----
    st.header("6 · Map fields to fill")
    st.caption(
        "For each master column choose which document field fills it. The site "
        "key column is used for matching and isn't filled here."
    )
    target_cols = [c for c in master_info.column_names if c != main_key_col]
    map_source_options = [c for c in source_cols if c != source_key_col]
    suggestions = mapping_mod.suggest_mapping(target_cols, map_source_options)

    NONE_LABEL = "— leave unchanged —"
    options = [NONE_LABEL, *map_source_options]
    field_map: dict[str, str | None] = {}
    kinds: dict[str, str] = {}

    head = st.columns([3, 4, 2, 2])
    head[0].markdown("**Master column**")
    head[1].markdown("**Document field**")
    head[2].markdown("**Match**")
    head[3].markdown("**Value type**")
    for col in target_cols:
        c0, c1, c2, c3 = st.columns([3, 4, 2, 2])
        c0.write(col)
        sug = suggestions.get(col)
        default = sug.source if (sug and sug.source in options) else NONE_LABEL
        choice = c1.selectbox(
            col, options, index=options.index(default),
            label_visibility="collapsed", key=f"map_{col}",
        )
        field_map[col] = None if choice == NONE_LABEL else choice
        c2.write(f"{sug.score:.0f}%" if (sug and sug.source and choice == sug.source) else "—")
        kinds[col] = c3.selectbox(
            col, ["auto", "text", "number", "date"],
            label_visibility="collapsed", key=f"kind_{col}",
        )

    mapped_n = sum(1 for v in field_map.values() if v)
    st.info(f"{mapped_n} field(s) mapped to fill.")

    # ---- Step 7: generate ----
    st.header("7 · Generate filled master sheet")
    g1, g2 = st.columns([2, 3])
    with g1:
        on_unmatched = st.radio(
            "Sites not found in the master sheet",
            ["skip", "append"],
            format_func=lambda m: "Skip them (report only)" if m == "skip"
            else "Add them as new rows",
        )
    with g2:
        out_name = st.text_input("Output file name", value="sites_master_filled.xlsx")

    if st.button("⚙️ Fill master sheet", type="primary"):
        if mapped_n == 0:
            st.error("Map at least one field before generating.")
        else:
            with st.spinner("Filling master sheet…"):
                data_bytes, report = merger.merge_by_key(
                    master_bytes, master_info, source_df, field_map,
                    main_key_col=main_key_col, source_key_col=source_key_col,
                    on_unmatched=on_unmatched, allow_fuzzy=allow_fuzzy,
                    fuzzy_threshold=threshold, column_kinds=kinds,
                )
            SS["result_bytes"] = data_bytes
            SS["result_report"] = report
            SS["result_name"] = out_name or "sites_master_filled.xlsx"

# ===========================================================================
# Result
# ===========================================================================
if SS.get("result_bytes") is not None:
    report = SS["result_report"]
    st.success(
        f"Updated **{report.rows_updated}** site row(s)"
        + (f", added **{report.rows_appended}** new" if report.rows_appended else "")
        + (f", **{len(report.unmatched)}** unmatched" if report.unmatched else "")
        + (f"  ·  {report.fuzzy_matches} fuzzy" if report.fuzzy_matches else "")
        + "."
    )
    st.download_button(
        "⬇️ Download filled master sheet",
        data=SS["result_bytes"], file_name=SS.get("result_name", "sites_master_filled.xlsx"),
        mime=XLSX_MIME, type="primary",
    )
    with st.expander("Report", expanded=bool(report.unmatched)):
        if report.fields_filled:
            st.write("**Values filled per column:**")
            st.table(pd.DataFrame(
                [{"Column": k, "Filled": v} for k, v in report.fields_filled.items()]
            ))
        if report.unmatched:
            st.write("**Documents with no matching site:**")
            st.table(pd.DataFrame(report.unmatched))
        if report.fuzzy_matches:
            st.caption(
                f"{report.fuzzy_matches} match(es) were fuzzy (not exact) — "
                "double-check those rows."
            )
