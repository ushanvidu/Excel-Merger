"""Generate sample files for trying out the app — network-sites scenario.

Creates, under ``sample_data/``:
  - ``sites_master.xlsx`` : the master sheet with one row per site (several
    fields left blank, to be filled from per-site documents).
  - per-site source documents, each describing ONE site:
      * key/value form PDFs       (site_COL-001.pdf, site_GAL-002.pdf)
      * a record-style table PDF  (site_KAN-003.pdf)
      * an already-converted Excel (site_MAT-004.xlsx)
      * an unmatched site PDF      (site_XYZ-999.pdf) to show unmatched handling

Requires the dev dependency reportlab:  pip install -r requirements-dev.txt
Run:  python make_samples.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

OUT = Path(__file__).parent / "sample_data"

MASTER_HEADERS = [
    "Site ID", "Site Name", "Region", "Power (kW)", "Battery Backup (h)", "Status"
]
# Power/Battery/Status start blank — those are what the per-site docs fill.
MASTER_ROWS = [
    ["COL-001", "Colombo Fort", "Western", "", "", ""],
    ["GAL-002", "Galle Face", "Western", "", "", ""],
    ["KAN-003", "Kandy Central", "Central", "", "", ""],
    ["MAT-004", "Matara South", "Southern", "", "", ""],
    ["JAF-005", "Jaffna North", "Northern", "", "", ""],  # no document -> stays blank
]


def _styled(data: list[list[str]]) -> Table:
    table = Table(data)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
    ]))
    return table


def _keyvalue_pdf(path: Path, fields: dict[str, str]) -> None:
    data = [["Parameter", "Value"]] + [[k, v] for k, v in fields.items()]
    SimpleDocTemplate(str(path), pagesize=A4).build([_styled(data)])


def _records_pdf(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    SimpleDocTemplate(str(path), pagesize=A4).build([_styled([headers, *rows])])


def _records_excel(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Site"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)


def _master(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sites"
    ws.append(MASTER_HEADERS)
    for row in MASTER_ROWS:
        ws.append(row)
    wb.save(path)


def main() -> None:
    OUT.mkdir(exist_ok=True)
    _master(OUT / "sites_master.xlsx")

    # One site per document, mixed formats:
    _keyvalue_pdf(OUT / "site_COL-001.pdf", {
        "Site ID": "COL-001", "Power (kW)": "5.5",
        "Battery Backup (h)": "8", "Status": "Active",
    })
    _keyvalue_pdf(OUT / "site_GAL-002.pdf", {
        "Site ID": "GAL-002", "Power (kW)": "6.0",
        "Battery Backup (h)": "10", "Status": "Active",
    })
    _records_pdf(
        OUT / "site_KAN-003.pdf",
        ["Site ID", "Power (kW)", "Battery Backup (h)", "Status"],
        [["KAN-003", "4.2", "6", "Maintenance"]],
    )
    _records_excel(
        OUT / "site_MAT-004.xlsx",
        ["Site ID", "Power (kW)", "Battery Backup (h)", "Status"],
        [["MAT-004", "3.8", "4", "Active"]],
    )
    _keyvalue_pdf(OUT / "site_XYZ-999.pdf", {
        "Site ID": "XYZ-999", "Power (kW)": "2.0",
        "Battery Backup (h)": "2", "Status": "Active",
    })

    print(f"Wrote samples to {OUT}/")
    print("  master : sites_master.xlsx (sheet 'Sites')")
    print("  docs   : site_COL-001.pdf, site_GAL-002.pdf, site_KAN-003.pdf,")
    print("           site_MAT-004.xlsx, site_XYZ-999.pdf (unmatched)")


if __name__ == "__main__":
    main()
