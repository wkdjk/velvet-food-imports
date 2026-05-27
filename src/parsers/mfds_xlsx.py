"""
Parser for MFDS weekly xlsx files (수입식품조회YYYYMMDD.xlsx).

Source: https://impfood.mfds.go.kr — 수입식품조회 download

Expected 14-column format (confirmed PRE-5, 2026-05-26):
  처리일자 | 수입업체 | 신고번호 | 품목제조번호 | 제품명(한글) | 제품명(영문) |
  품목류 | 수출국 | 제조국 | 제조사(영문) | 순중량(KG) | 수량(갯수) |
  유통기한 | 원산지

Internal field name mapping (KO → EN):
  처리일자       → date
  수입업체       → importer_ko
  신고번호       → report_no
  품목제조번호   → item_no
  제품명(한글)   → product_ko
  제품명(영문)   → product_en
  품목류         → product_type_ko
  수출국         → country_export_ko
  제조국         → country_origin_ko
  제조사(영문)   → exporter_en
  순중량(KG)    → weight_kg
  수량(갯수)    → quantity
  유통기한       → expiry_date
  원산지         → country_origin_raw

Notes:
- Some xlsx files have a title row before the header row; the parser auto-detects.
- product_type_en, country_export_en, country_origin_en, importer_en are
  resolved via the translation table (falls back to original string if unknown).
- type_frozen / type_dried / type_ambiguous flags added from product_type_ko.
- source is set to 'mfds'.
"""

from __future__ import annotations

import datetime
import re
from pathlib import Path
from typing import Any

import openpyxl

from src.parsers.translation_table import get_table

# Canonical 14 Korean column headers from MFDS portal
MFDS_COLUMNS: list[str] = [
    "처리일자",
    "수입업체",
    "신고번호",
    "품목제조번호",
    "제품명(한글)",
    "제품명(영문)",
    "품목류",
    "수출국",
    "제조국",
    "제조사(영문)",
    "순중량(KG)",
    "수량(갯수)",
    "유통기한",
    "원산지",
]

# Map: Korean column header → internal field name
COLUMN_MAP: dict[str, str] = {
    "처리일자": "date",
    "수입업체": "importer_ko",
    "신고번호": "report_no",
    "품목제조번호": "item_no",
    "제품명(한글)": "product_ko",
    "제품명(영문)": "product_en",
    "품목류": "product_type_ko",
    "수출국": "country_export_ko",
    "제조국": "country_origin_ko",
    "제조사(영문)": "exporter_en",
    "순중량(KG)": "weight_kg",
    "수량(갯수)": "quantity",
    "유통기한": "expiry_date",
    "원산지": "country_origin_raw",
}

_FROZEN_KEYWORDS = ("냉동", "생녹용")
_DRIED_KEYWORDS = ("건조", "건")


def _safe_str(value: Any) -> str:
    """Return stripped string or empty string for None."""
    if value is None:
        return ""
    return str(value).strip()


def _parse_date(value: Any) -> str:
    """Return ISO-format date string or empty string."""
    if value is None:
        return ""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    if not s or s == "-":
        return ""
    # Handle YYYYMMDD compact format from MFDS
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return s


def _classify_type(product_type_ko: str) -> dict[str, bool]:
    """Return frozen/dried/ambiguous flags from Korean product type."""
    kw = str(product_type_ko).strip()
    is_frozen = any(k in kw for k in _FROZEN_KEYWORDS)
    is_dried = "건조" in kw or ("건" in kw and "건강" not in kw and not is_frozen)
    if is_frozen and is_dried:
        return {"type_frozen": False, "type_dried": False, "type_ambiguous": True}
    if is_frozen:
        return {"type_frozen": True, "type_dried": False, "type_ambiguous": False}
    if is_dried:
        return {"type_frozen": False, "type_dried": True, "type_ambiguous": False}
    return {"type_frozen": False, "type_dried": False, "type_ambiguous": True}


def _find_header_row(
    all_rows: list[tuple], expected_columns: list[str]
) -> int | None:
    """Find the row index that contains the expected MFDS column headers.

    Some MFDS files have a title row before the header row; this handles both.
    Returns the 0-based index of the header row, or None if not found.
    """
    for i, row in enumerate(all_rows[:5]):  # header must be in first 5 rows
        row_strs = [_safe_str(c) for c in row]
        if all(col in row_strs for col in expected_columns[:7]):
            return i
    return None


def parse(xlsx_path: Path, xlsx_path_for_table: Path | None = None) -> list[dict[str, Any]]:
    """Parse an MFDS weekly xlsx file and return normalised row dicts.

    Args:
        xlsx_path: Path to the MFDS xlsx file.
        xlsx_path_for_table: Optional path to Commander xlsx for translation table.
                             Defaults to cfg.historical_xlsx_path via get_table().

    Returns:
        List of dicts with internal English field names.

    Raises:
        ValueError: If the expected column headers are not found in the file.
    """
    table = get_table(xlsx_path_for_table)

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    # MFDS files typically have one sheet
    ws = wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))

    header_idx = _find_header_row(all_rows, MFDS_COLUMNS)
    if header_idx is None:
        raise ValueError(
            f"Cannot find MFDS column headers in {xlsx_path}.\n"
            f"Expected at least: {MFDS_COLUMNS[:7]}\n"
            f"First row found: {list(all_rows[0]) if all_rows else 'empty'}"
        )

    raw_headers = [_safe_str(c) for c in all_rows[header_idx]]

    # Check all 14 columns are present
    missing = [col for col in MFDS_COLUMNS if col not in raw_headers]
    if missing:
        raise ValueError(
            f"MFDS file {xlsx_path} is missing columns: {missing}\n"
            f"Actual headers: {raw_headers}"
        )

    col_index: dict[str, int] = {h: i for i, h in enumerate(raw_headers) if h}

    results: list[dict[str, Any]] = []

    for raw_row in all_rows[header_idx + 1 :]:
        # Skip blank rows
        if all(v is None or str(v).strip() == "" for v in raw_row):
            continue

        row: dict[str, Any] = {}

        # Extract raw values
        for ko_col, internal in COLUMN_MAP.items():
            idx = col_index.get(ko_col)
            if idx is None or idx >= len(raw_row):
                row[internal] = ""
            else:
                row[internal] = raw_row[idx]

        # Normalise dates
        row["date"] = _parse_date(row["date"])
        row["expiry_date"] = _parse_date(row["expiry_date"])

        # Normalise string fields
        for f in ["importer_ko", "report_no", "item_no", "product_ko", "product_en",
                  "product_type_ko", "country_export_ko", "country_origin_ko",
                  "exporter_en", "country_origin_raw"]:
            row[f] = _safe_str(row.get(f, ""))

        # Numeric fields — keep as string to avoid Sheets type coercion issues
        row["weight_kg"] = _safe_str(row.get("weight_kg", ""))
        row["quantity"] = _safe_str(row.get("quantity", ""))

        # Resolve English translations via lookup table (unknown → original)
        row["importer_en"] = table.importer(row["importer_ko"])
        row["product_type_en"] = table.product_type(row["product_type_ko"])
        row["country_export_en"] = table.country_of_origin(row["country_export_ko"])
        row["country_origin_en"] = table.country_of_origin(row["country_origin_ko"])

        # Derive year/month/day from date string
        date_str = row["date"]
        if date_str and len(date_str) == 10:
            row["year"] = date_str[:4]
            row["month"] = date_str[5:7]
            row["day"] = date_str[8:10]
        else:
            row["year"] = ""
            row["month"] = ""
            row["day"] = ""

        # Classification flags
        row.update(_classify_type(row["product_type_ko"]))

        # Source marker
        row["source"] = "mfds"

        results.append(row)

    return results


def get_expected_headers() -> list[str]:
    """Return the canonical output field list for MFDS rows.

    This is a superset of the historical headers (includes MFDS-specific fields).
    """
    return [
        "year", "month", "day", "date",
        "importer_en", "importer_ko",
        "product_ko", "product_en",
        "product_type_ko", "product_type_en",
        "country_origin_en", "country_origin_ko",
        "country_export_en", "country_export_ko",
        "exporter_en", "expiry_date",
        "report_no", "item_no", "weight_kg", "quantity", "country_origin_raw",
        "type_frozen", "type_dried", "type_ambiguous",
        "source",
    ]
