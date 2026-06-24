#!/usr/bin/env bash
# Launch the PDF -> Excel Merger app. Creates the venv + installs deps on first run.
set -e
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  echo "Setting up virtual environment…"
  python3 -m venv .venv
  ./.venv/bin/python -m pip install --upgrade pip
  ./.venv/bin/python -m pip install -r requirements.txt
fi
exec ./.venv/bin/streamlit run app.py "$@"
