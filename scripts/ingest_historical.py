"""
One-time bulk load of the Commander's historical xlsx into Google Sheets.

Usage:
    PYTHONPATH=. python scripts/ingest_historical.py
    PYTHONPATH=. python scripts/ingest_historical.py --dry-run

Requirements:
    - .env file with GOOGLE_SERVICE_ACCOUNT_JSON and GOOGLE_SHEETS_ID (or env vars set)
    - Commander xlsx at path specified by VFI_HISTORICAL_XLSX (or default)

Behaviour:
    - Reads all 576 rows from the 'import list' tab.
    - Applies translation table from 'Names for Vlookup' tab.
    - Appends only new rows to the 'vfi_transactions' Sheets tab (idempotent).
    - Dedup key: (date, importer_ko, product_ko).
    - Prints: rows loaded / skipped / errors.
    - --dry-run: prints what would be written, makes no API calls.

Run this script only once. Subsequent MFDS loads use ingest_mfds.py.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Ensure src/ is importable when running with PYTHONPATH=.
from config import cfg
from src.parsers.historical_xlsx import get_expected_headers, parse
from src.sheets.client import batch_create_records, ensure_tab_exists

VFI_TAB = "vfi_transactions"


def main(dry_run: bool = False) -> int:
    """Run the historical ingest.

    Returns exit code (0 = success, 1 = error).
    """
    xlsx_path = cfg.historical_xlsx_path

    print(f"Velvet Food Imports — historical ingest")
    print(f"  Source: {xlsx_path}")
    print(f"  Target tab: {VFI_TAB}")
    print(f"  Dry-run: {dry_run}")
    print()

    # Validate source file exists
    if not Path(xlsx_path).exists():
        print(
            f"[error] Historical xlsx not found at: {xlsx_path}\n"
            f"  Set VFI_HISTORICAL_XLSX in your .env file to the correct path.",
            file=sys.stderr,
        )
        return 1

    # Parse all rows from the historical xlsx
    print("Parsing historical xlsx...")
    try:
        rows = parse(xlsx_path)
    except Exception as exc:
        print(f"[error] Parser failed: {exc}", file=sys.stderr)
        return 1

    print(f"  Parsed {len(rows)} rows from 'import list' tab.")

    if not rows:
        print("[warning] No rows parsed — check xlsx path and tab name.")
        return 0

    # Ensure the target Sheets tab exists with correct headers
    headers = get_expected_headers()

    if not dry_run:
        print(f"Ensuring tab '{VFI_TAB}' exists in Sheets...")
        try:
            ensure_tab_exists(VFI_TAB, headers)
        except Exception as exc:
            print(f"[error] Cannot open/create Sheets tab: {exc}", file=sys.stderr)
            return 1

    # Append new rows (dedup-safe)
    print(f"Writing to Sheets tab '{VFI_TAB}'...")
    try:
        result = batch_create_records(VFI_TAB, rows, dry_run=dry_run)
    except Exception as exc:
        print(f"[error] Sheets write failed: {exc}", file=sys.stderr)
        return 1

    # Summary report
    print()
    print("Summary:")
    print(f"  Loaded:  {result['written']}")
    print(f"  Skipped: {result['skipped']}  (already in Sheets)")
    print(f"  Errors:  {result['errors']}")
    print()

    if result["errors"] > 0:
        print("[warning] Some rows had errors — check output above.")
        return 1

    print("Ingest complete.")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="One-time bulk load of historical deer velvet import data into Google Sheets."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be written without making any API calls.",
    )
    args = parser.parse_args()
    sys.exit(main(dry_run=args.dry_run))
