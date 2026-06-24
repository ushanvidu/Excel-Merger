"""Network Site Sheet Filler — Streamlit app entry point.

Two tools:
  • Fill Master Sheet     — fill a master sheet's columns from TSSR PDFs/Excel.
  • Convert TSSR to Excel — turn TSSR PDFs into clean, topic-divided sheets.

Run:  streamlit run app.py
"""

from __future__ import annotations

import streamlit as st

st.set_page_config(page_title="Site Sheet Filler", page_icon="📡", layout="wide")

fill = st.Page("views/fill.py", title="Fill Master Sheet", icon="🗂️", default=True)
convert = st.Page("views/convert.py", title="Convert TSSR to Excel", icon="📄")

with st.sidebar:
    st.markdown("### 📡 Site Sheet Filler")
    st.caption("Fill a master site sheet from TSSR survey documents.")

st.navigation([fill, convert]).run()
