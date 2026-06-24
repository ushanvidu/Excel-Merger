# Network Site Sheet Filler

Fill a **master Excel sheet** (one row per network site across Sri Lanka) from
**per-site documents**. Each document describes ONE site and may be a **PDF**
(converted to Excel automatically) or an **already-converted Excel/CSV**. The app
identifies each document's site, finds its row in the master sheet by **Site ID
or Site Name**, and fills the mapped fields into that row.

## TSSR workflow (InfraMS survey PDFs → master RF Parameter sheet)

The app has two purpose-built pages (sidebar) for the TSSR / network-sites use case:

**1. TSSR to Excel** — convert one or more InfraMS TSSR PDFs into a clean,
topic-divided workbook: `Summary`, `Antenna Details`, `Sector Details`,
`Solution Antenna Details`, `Solution Sector Details`, `Extraction Log`. Wrapped
cells (`Sect/or 1` → `Sector 1`, `L850 _1` → `L850_1`) are de-wrapped and mapped
onto the fixed template schema. Download the multi-sheet `.xlsx`.

**2. Fill Master Sheet** — fill the master `RF Parameter` sheet's TSSR columns
(`Antenna_TSSR`, `Antenna height_TSSR`, `Antenna azimuth_TSSR`,
`Mechanical tilt_TSSR`, `Electrical tilt_TSSR`) from the PDFs' **Solution**
tables. Each master row (Sector × Band) is matched by **Sector + Band**; the
sector is recovered from `Antenna azimuth_Mbitel` when the `Sector` cell is
`#REF!`. The band → PDF-token mapping is editable in the UI. `TSSR Status` is
left blank. Only the mapped cells are written — formatting and other sheets are
preserved. A per-row report shows exactly what filled each row.

Code: `pdf_excel_merger/tssr_extractor.py` (extraction + topic-sheet export) and
`pdf_excel_merger/tssr_fill.py` (master fill). Tested end-to-end against the real
GMALG1 PDF + `Noding Config` master in `tests/test_tssr.py`.

## What it does (generic key-based merge)

1. **Upload the master sheet** — one row per site, with a key column (Site ID /
   Site Name) and the data columns to fill.
2. **Upload per-site documents** — PDFs *and/or* Excel/CSV, mixed, many at once.
   PDFs are **converted to Excel** (downloadable) as an explicit step.
3. **Match** each document to the correct site row:
   - **Exact** match on a normalized key (`COL-001` = `col 001` = `COL001`) —
     ideal for site IDs.
   - **Fuzzy** fallback for site *names* with small spelling/format differences
     (gated by a threshold and reported, so you can verify).
4. **Map fields** — choose which document field fills each master column
   (auto-suggested via fuzzy + business-term synonyms; you confirm).
5. **Generate** the filled master sheet — only the matched rows' mapped cells are
   written; existing formatting, formulas and other sheets are preserved.
   Unmatched sites are reported (and optionally appended as new rows).

## Key features

- **Dual input** — PDF (auto-converted) or Excel/CSV, in one run.
- **One-site forms** — two-column `Parameter | Value` documents are pivoted into a
  record automatically (`Document layout = Auto / Key-value`).
- **Update-by-key** — fills the *existing* site's row instead of appending blindly.
- **Editable review** — fix any extracted value or site key before matching.
- **Match preview + report** — see exactly which document mapped to which row,
  by exact/fuzzy, with confidence, before and after generating.
- **Type-aware** — numbers, dates and text are coerced to land correctly.

## Quick start

```bash
cd "Excel App"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# (optional) generate the sample network-sites scenario
pip install -r requirements-dev.txt
python make_samples.py

# launch
streamlit run app.py        # or: ./run.sh
```

Opens at <http://localhost:8501>.

## How to use

1. Upload the **master sheet**; pick the sheet + confirm the header row.
2. Upload **PDF** documents (left) and/or **Excel/CSV** documents (right).
3. Pick **Document layout** (`auto` works for most) and **PDF table style**, then
   **Convert & load**.
4. Review each converted document (download as Excel if you like) and fix any
   values in the combined table.
5. Confirm the **master key column** and **source key field** (auto-detected) and
   check the **match preview**.
6. Confirm the **field mapping**, choose how to handle **unmatched** sites, and
   click **Fill master sheet**, then **download**.

## Project layout

```
Excel App/
├── app.py                       # Streamlit UI (entry point)
├── pdf_excel_merger/
│   ├── source_loader.py         # PDF->Excel conversion + Excel/CSV loading + orientation
│   ├── pdf_extractor.py         # PDF -> table extraction
│   ├── excel_reader.py          # read master: sheets, header row, columns, key values
│   ├── site_matching.py         # detect key column, normalize + match site keys
│   ├── mapping.py               # fuzzy + synonym field auto-mapping
│   ├── merger.py                # match-by-key update (+ append) into the master sheet
│   └── utils.py                 # header normalization + value coercion
├── make_samples.py              # generate the network-sites sample scenario
├── tests/test_pipeline.py       # headless end-to-end test
├── requirements.txt             # runtime deps
└── requirements-dev.txt         # test/sample-generation deps (reportlab)
```

## Run the test

```bash
source .venv/bin/activate
pip install -r requirements-dev.txt
python -m tests.test_pipeline
```

## Notes & limits

- Built for **text-based** PDFs. Scanned/image PDFs need OCR — not included yet
  (can be added as a fallback).
- Tune for your data: extend the site-key hints in `site_matching.py`
  (`_KEY_HINTS`) and the field synonyms in `mapping.py` (`_SYNONYM_GROUPS`).
- If a PDF's table isn't detected, switch **PDF table style** between `lines`
  (bordered) and `text` (borderless), or set **Document layout** explicitly.
