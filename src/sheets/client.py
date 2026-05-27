"""
Google Sheets API wrapper for Velvet Food Imports.

Design rules (fleet lessons — mandatory):
- Never call the Sheets API inside a row loop.
- Read the entire worksheet once with get_all_records(); cache in memory.
- Batch writes: up to 200 rows per call; 1.1 s sleep between batches.
- Dedup key for VFI: (처리일자, 수입업체, 제품명) — fields date, importer_ko, product_ko.
"""

import json
import time
from typing import Any

import gspread
from google.oauth2.service_account import Credentials

from config import cfg

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Dedup key field names (must match the column headers written to Sheets)
DEDUP_FIELDS: tuple[str, ...] = ("date", "importer_ko", "product_ko")

_BATCH_SIZE = 200
_BATCH_SLEEP = 1.1  # seconds between batch API calls


def _build_client() -> gspread.Client:
    """Authenticate and return a gspread client.

    Parses the service account JSON from the environment variable string.
    json.loads() is required because os.environ returns a str, not a dict.
    """
    sa_info = json.loads(cfg.service_account_json_str)
    creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return gspread.authorize(creds)


def open_sheet(tab_name: str) -> gspread.Worksheet:
    """Open a named worksheet in the VFI master Sheets file."""
    client = _build_client()
    spreadsheet = client.open_by_key(cfg.sheets_id)
    try:
        return spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError(
            f"Worksheet '{tab_name}' not found in Sheets ID {cfg.sheets_id}. "
            "Create the tab first or check the tab name."
        )


def get_all_records(tab_name: str) -> list[dict[str, Any]]:
    """Read all rows from a worksheet in a single API call.

    Returns a list of dicts keyed by header row values.
    Never call this inside a loop — one call per script run.
    """
    ws = open_sheet(tab_name)
    return ws.get_all_records()


def _make_dedup_key(row: dict[str, Any]) -> tuple:
    """Build the composite dedup key from a row dict."""
    return tuple(str(row.get(field, "")).strip() for field in DEDUP_FIELDS)


def batch_create_records(
    tab_name: str,
    rows: list[dict[str, Any]],
    *,
    dry_run: bool = False,
) -> dict[str, int]:
    """Append rows to a worksheet, skipping duplicates.

    Reads the entire worksheet once, builds a dedup set from existing rows,
    then appends only new rows in batches of up to _BATCH_SIZE.

    Args:
        tab_name: Name of the worksheet tab.
        rows: List of row dicts to insert. Keys must match existing headers.
        dry_run: If True, prints what would be written but makes no API calls.

    Returns:
        {'written': int, 'skipped': int, 'errors': int}
    """
    if not rows:
        return {"written": 0, "skipped": 0, "errors": 0}

    if dry_run:
        # In dry-run mode we do not open Sheets at all — just report what would be written
        print(f"  [dry-run] Would attempt to write {len(rows)} rows to tab '{tab_name}'.")
        print(f"  [dry-run] Dedup check skipped (no Sheets connection in dry-run mode).")
        return {"written": 0, "skipped": 0, "errors": 0}

    ws = open_sheet(tab_name)

    # --- Single bulk read — cache existing records in memory ---
    existing = ws.get_all_records()
    existing_keys: set[tuple] = {_make_dedup_key(r) for r in existing}

    # Determine column order from the header row
    if existing:
        headers = list(existing[0].keys())
    else:
        # Worksheet is empty — use keys from first incoming row
        headers = list(rows[0].keys())
        if not dry_run:
            ws.append_row(headers, value_input_option="RAW")

    # --- Filter new rows ---
    new_rows: list[list] = []
    skipped = 0
    errors = 0

    for row in rows:
        key = _make_dedup_key(row)
        if key in existing_keys:
            skipped += 1
            continue
        try:
            new_rows.append([row.get(h, "") for h in headers])
            existing_keys.add(key)  # prevent duplicates within this batch
        except Exception as exc:  # noqa: BLE001
            print(f"  [error] row skipped: {exc} — {row}")
            errors += 1

    if dry_run:
        print(f"  [dry-run] Would write {len(new_rows)} rows, skip {skipped}, errors {errors}")
        return {"written": 0, "skipped": skipped, "errors": errors}

    # --- Batch write ---
    written = 0
    for batch_start in range(0, len(new_rows), _BATCH_SIZE):
        batch = new_rows[batch_start : batch_start + _BATCH_SIZE]
        ws.append_rows(batch, value_input_option="USER_ENTERED")
        written += len(batch)
        if batch_start + _BATCH_SIZE < len(new_rows):
            time.sleep(_BATCH_SLEEP)

    return {"written": written, "skipped": skipped, "errors": errors}


def ensure_tab_exists(tab_name: str, headers: list[str]) -> gspread.Worksheet:
    """Create a worksheet tab with the given headers if it does not already exist.

    Returns the worksheet (existing or newly created).
    """
    client = _build_client()
    spreadsheet = client.open_by_key(cfg.sheets_id)
    try:
        ws = spreadsheet.worksheet(tab_name)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=tab_name, rows=2000, cols=len(headers))
        ws.append_row(headers, value_input_option="RAW")
        print(f"  Created tab '{tab_name}' with {len(headers)} columns.")
        return ws
