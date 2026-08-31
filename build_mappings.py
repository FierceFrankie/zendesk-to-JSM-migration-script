#!/usr/bin/env python3
"""build_mappings.py

Convert the Git-diff-friendly CSV mapping files in ``mappings/`` into a single
``mappings.xlsx`` workbook with one tab per mapping.

Why this exists
---------------
Binary xlsx files are painful to review and merge in Git: a one-cell change
rewrites the whole binary and produces no meaningful diff (and can corrupt on
bad merges). Keeping the *source of truth* as CSV means every change is a clean,
reviewable text diff. This script regenerates the workbook on demand.

The runtime loader in ``migrate.py`` can read either the CSVs directly or the
generated ``mappings.xlsx``, so this step is optional but recommended when you
want a single distributable workbook.

Workbook layout
---------------
Every tab uses: row 1 = header (ignored at load time), column A = Zendesk value,
column B = JSM value. The ``config`` tab is a simple key/value sheet.

Usage
-----
    python build_mappings.py                 # reads ./mappings, writes ./mappings.xlsx
    python build_mappings.py --src mappings --out mappings.xlsx
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from openpyxl import Workbook

# Maps each output worksheet (tab) name to its source CSV file. The order here
# is the order the tabs appear in the generated workbook.
SHEET_TO_CSV = {
    "priority": "priority.csv",
    "status": "status.csv",
    "tags": "tags.csv",
    "request_type": "request_type.csv",
    "config": "config.csv",
}


def _read_csv_rows(csv_path: Path) -> list[list[str]]:
    """Read a CSV file into a list of string rows.

    Missing files are tolerated (a warning is printed and an empty header-only
    sheet is produced) so that optional tabs like ``request_type`` do not break
    the build.
    """
    if not csv_path.exists():
        print(f"WARNING: {csv_path} not found; writing an empty sheet.", file=sys.stderr)
        return [["zendesk_value", "jsm_value"]]
    with csv_path.open(newline="", encoding="utf-8") as handle:
        return [row for row in csv.reader(handle)]


def build_workbook(src_dir: Path, out_path: Path) -> None:
    """Build ``out_path`` (an .xlsx) from the CSV files in ``src_dir``."""
    workbook = Workbook()
    # openpyxl creates one default sheet; remove it so we control all tabs.
    workbook.remove(workbook.active)

    for sheet_name, csv_name in SHEET_TO_CSV.items():
        worksheet = workbook.create_sheet(title=sheet_name)
        rows = _read_csv_rows(src_dir / csv_name)
        for row in rows:
            worksheet.append(row)
        print(f"  {sheet_name:<14} <- {csv_name} ({max(len(rows) - 1, 0)} data rows)")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(out_path)
    print(f"Wrote {out_path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build mappings.xlsx from CSV mapping files.")
    parser.add_argument("--src", default="mappings", help="Directory containing the mapping CSVs.")
    parser.add_argument("--out", default="mappings.xlsx", help="Output xlsx path.")
    args = parser.parse_args(argv)

    print("Building mappings workbook:")
    build_workbook(Path(args.src), Path(args.out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
