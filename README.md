# Network Site Sheet Filler

A Streamlit app that fills a **master site sheet** from **TSSR survey documents**.
Upload the master plus the site PDFs (or already-converted Excel), choose which
columns to fill and where each comes from, and download the filled sheet. Each
site row is matched by **Site + Sector + Band**; only the chosen cells are
written, so existing formatting, formulas and other sheets are preserved.

Two tools (sidebar):

- **🗂 Fill Master Sheet** — match documents to master rows and fill the columns
  you select. Columns and their source fields are fully configurable, so it
  works with any master and any TSSR PDFs (not just samples).
- **📄 Convert TSSR to Excel** — turn TSSR PDFs into clean, topic-divided sheets
  (Antenna / Sector / Solution details + extraction log).

## Run locally

```bash
cd "Excel App"
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py          # or ./run.sh
```

Opens at <http://localhost:8501>.

## Deploy with Docker

```bash
docker build -t site-sheet-filler .
docker run -p 8501:8501 site-sheet-filler
```

The image contains only the app (`app.py`, `views/`, `pdf_excel_merger/`,
`.streamlit/`) — no samples, tests, or virtualenv. It can also be deployed to
Streamlit Community Cloud (point it at `app.py`) or any container host.

## How to use

1. **Upload** the master `.xlsx` and the site TSSR **PDFs** (and/or an
   already-converted Excel — both are combined best-of-both).
2. Confirm the **sheet**, **header row**, and the **key columns** (Site / Sector /
   Band / azimuth — auto-detected; under *Key columns & data source*).
3. Choose the **columns to fill** (TSSR columns are pre-selected) and the
   **document field** that feeds each (auto-mapped; editable).
4. Click **Fill master sheet**, review the report (incl. any **sector conflicts
   flagged for review**), and **download**.

## How it works

- **Extraction** (`pdf_excel_merger/tssr_extractor.py`): rebuilds each table cell
  from pdfplumber word geometry inside the `find_tables` grid — no truncation,
  robust to wrapped headers and faint borders.
- **Matching** (`pdf_excel_merger/tssr_fill.py`): resolves each master row's
  sector from azimuth (falling back to the Sector column with a 0/1-based
  offset), matches the Solution row by band, and reads the mapped fields.
  The band-token suffix (`_P`/`_1`/`_W`…) independently encodes the sector and is
  used to break ties and **flag conflicts** rather than guess silently.

## Testing & accuracy

```bash
source .venv/bin/activate
pip install -r requirements-dev.txt   # reportlab, for sample generation
python -m tests.test_tssr             # TSSR extract + fill, unit checks
python -m tests.test_pipeline         # generic key-based merge
python -m tests.accuracy_harness      # measured accuracy vs golden ground truth
```

The **accuracy harness** runs the real pipeline and compares written cells to a
hand-verified ground truth (`tests/golden/tssr_golden.csv`, covering a 1-indexed
and a 0-indexed site), reporting per-field accuracy / precision / recall and
gating at 90% (currently **100%**).

## Project layout

```
Excel App/
├── app.py                       # entry point (st.navigation)
├── views/
│   ├── fill.py                  # Fill Master Sheet page
│   └── convert.py               # Convert TSSR to Excel page
├── pdf_excel_merger/            # engine (UI-agnostic, tested)
│   ├── tssr_extractor.py        # PDF -> topic tables (geometry-based)
│   ├── tssr_fill.py             # match-by-key fill (configurable field map)
│   ├── excel_reader.py          # read master: sheets, header, columns
│   └── …                        # generic merge helpers + utils
├── tests/                       # test_tssr, test_pipeline, accuracy_harness, golden/
├── Dockerfile / .dockerignore   # deployment
├── requirements.txt             # runtime deps
└── requirements-dev.txt         # test/sample-generation deps
```

## Notes & limits

- Built for **text-based** TSSR PDFs (InfraMS template). Scanned/image PDFs would
  need OCR (not included).
- The band→token map and field map are editable in the UI for other layouts.
- Rows whose source data is ambiguous are **flagged for review**, never filled
  with a silent guess.
