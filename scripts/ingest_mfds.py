"""
MFDS weekly xlsx ingest script.

Triggered by Commander uploading a new 수입식품조회YYYYMMDD.xlsx to data/mfds/.
Also triggered by GitHub Actions on push to data/mfds/**.

Usage:
    PYTHONPATH=. python scripts/ingest_mfds.py
    PYTHONPATH=. python scripts/ingest_mfds.py --dry-run
    PYTHONPATH=. python scripts/ingest_mfds.py --file data/mfds/수입식품조회20260525.xlsx

Behaviour:
    - Scans VFI_MFDS_DIR (default: data/mfds/) for unprocessed xlsx files.
    - A file is considered processed if its name ends with '_done.xlsx'.
    - Parses each new file with src/parsers/mfds_xlsx.py.
    - Appends new rows to the 'vfi_transactions' Sheets tab (dedup-safe).
    - Renames each processed file from *.xlsx → *_done.xlsx.
    - Prints: files found / rows written / skipped / errors per file.
    - --dry-run: prints what would be written; does not rename files.
    - --file: process a single specific file (overrides directory scan).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from config import cfg
from src.parsers.historical_xlsx import get_expected_headers
from src.parsers.mfds_xlsx import get_expected_headers as mfds_headers, parse
from src.sheets.client import batch_create_records, ensure_tab_exists

VFI_TAB = "vfi_transactions"

# Combined header list: union of historical + MFDS fields.
# Historical rows will have empty strings for MFDS-only fields.
def _get_unified_headers() -> list[str]:
    hist = get_expected_headers()
    mfds = mfds_headers()
    seen = set(hist)
    extra = [h for h in mfds if h not in seen]
    return hist + extra


def _find_unprocessed_files(mfds_dir: Path) -> list[Path]:
    """Return all .xlsx files in mfds_dir that have not yet been processed.

    Processed files end with '_done.xlsx' and are excluded.
    """
    if not mfds_dir.exists():
        return []
    return sorted(
        p for p in mfds_dir.glob("*.xlsx")
        if not p.stem.endswith("_done")
    )


def _process_file(
    xlsx_path: Path,
    dry_run: bool,
    ensure_done: bool = True,
) -> dict[str, int]:
    """Parse one MFDS file and append its rows to Sheets.

    Args:
        xlsx_path:    Path to the MFDS xlsx file.
        dry_run:      If True, no API calls or file renames.
        ensure_done:  If True, rename file to *_done.xlsx after success.

    Returns:
        {'written': int, 'skipped': int, 'errors': int}
    """
    print(f"  Processing: {xlsx_path.name}")

    try:
        rows = parse(xlsx_path)
    except Exception as exc:
        print(f"    [error] Parse failed: {exc}", file=sys.stderr)
        return {"written": 0, "skipped": 0, "errors": 1}

    print(f"    Parsed {len(rows)} rows.")

    if not rows:
        print("    [warning] No rows in file — skipping.")
        return {"written": 0, "skipped": 0, "errors": 0}

    try:
        result = batch_create_records(VFI_TAB, rows, dry_run=dry_run)
    except Exception as exc:
        print(f"    [error] Sheets write failed: {exc}", file=sys.stderr)
        return {"written": 0, "skipped": 0, "errors": 1}

    print(
        f"    Written: {result['written']}  "
        f"Skipped: {result['skipped']}  "
        f"Errors: {result['errors']}"
    )

    # Rename to mark as processed
    if not dry_run and ensure_done and result["errors"] == 0:
        done_path = xlsx_path.with_stem(xlsx_path.stem + "_done")
        xlsx_path.rename(done_path)
        print(f"    Renamed to: {done_path.name}")

    return result


def main(dry_run: bool = False, single_file: Path | None = None) -> int:
    """Run the MFDS ingest.

    Returns exit code (0 = success, 1 = partial or full error).
    """
    print("Velvet Food Imports — MFDS ingest")
    print(f"  Dry-run: {dry_run}")
    print()

    # Determine which files to process
    if single_file is not None:
        if not single_file.exists():
            print(f"[error] File not found: {single_file}", file=sys.stderr)
            return 1
        files_to_process = [single_file]
    else:
        mfds_dir = cfg.mfds_dir
        files_to_process = _find_unprocessed_files(mfds_dir)
        print(f"  Scanning: {mfds_dir}")
        print(f"  Unprocessed files found: {len(files_to_process)}")

    if not files_to_process:
        print("  Nothing to process.")
        return 0

    # Ensure Sheets tab exists (combined headers)
    if not dry_run:
        headers = _get_unified_headers()
        try:
            ensure_tab_exists(VFI_TAB, headers)
        except Exception as exc:
            print(f"[error] Cannot open/create Sheets tab: {exc}", file=sys.stderr)
            return 1

    # Process each file
    total_written = 0
    total_skipped = 0
    total_errors = 0

    for xlsx_path in files_to_process:
        result = _process_file(xlsx_path, dry_run=dry_run)
        total_written += result["written"]
        total_skipped += result["skipped"]
        total_errors += result["errors"]

    print()
    print("Summary:")
    print(f"  Files processed: {len(files_to_process)}")
    print(f"  Rows written:    {total_written}")
    print(f"  Rows skipped:    {total_skipped}  (already in Sheets)")
    print(f"  Errors:          {total_errors}")

    return 1 if total_errors > 0 else 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Ingest MFDS weekly deer velvet import xlsx files into Google Sheets."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be written without making any API calls or renaming files.",
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=None,
        metavar="PATH",
        help="Process a single specific xlsx file instead of scanning data/mfds/.",
    )
    args = parser.parse_args()
    sys.exit(main(dry_run=args.dry_run, single_file=args.file))
