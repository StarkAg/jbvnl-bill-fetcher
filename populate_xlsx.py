#!/usr/bin/env python3
"""One-off: write the already-scraped amounts into Electricity.xlsx's 'Bill Amounts'
sheet, from the existing Bill Amounts.csv (no portal login needed).

Run AFTER closing Electricity.xlsx in Excel, so Excel doesn't overwrite the sheet.
"""
import csv
from pathlib import Path
import fetch_bills as fb

CSV = fb.BILLS_BASE / "Bill Amounts.csv"

rows = []
if CSV.exists():
    with open(CSV, newline="") as f:
        for r in csv.DictReader(f):
            # coerce numerics back
            for k in ("units_billed", "bill_amount", "arrears"):
                try:
                    r[k] = float(r[k]) if r[k] not in ("", None) else ""
                except ValueError:
                    pass
            rows.append(r)

if not rows:
    print("No Bill Amounts.csv found — run fetch_bills.py first.")
else:
    total = fb.store_amounts(rows)
    print(f"Wrote {len(rows)} rows into {fb.XLSX_PATH.name} → '{fb.SHEET_NAME}' ({total} total).")
