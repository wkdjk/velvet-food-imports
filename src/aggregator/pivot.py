"""
Product Type × Country pivot aggregator — Velvet Food Imports.

Produces a cross-tabulation of product type (frozen / dried / ambiguous)
against country of export (NZ / Russia / China / Other).

Country bucketing logic:
  - 'NZ'     : country_export_en contains 'New Zealand' OR country_export_ko contains '뉴질랜드'
  - 'Russia'  : country_export_en contains 'Russia'     OR country_export_ko contains '러시아'
  - 'China'   : country_export_en contains 'China'      OR country_export_ko contains '중국'
  - 'Other'   : anything else, including empty strings

Optional year filter:  if `year` is provided only rows from that year are included.
                       If `year` is None, all years are aggregated.

Output:

    {
        "frozen":    {"NZ": 120, "Russia": 10, "China": 0, "Other": 5},
        "dried":     {"NZ":  90, "Russia":  5, "China": 2, "Other": 8},
        "ambiguous": {"NZ":  20, "Russia":  0, "China": 0, "Other": 4},
        "total":     {"NZ": 230, "Russia": 15, "China": 2, "Other": 17},
    }

Pure function — no Sheets API calls, no file I/O.
"""

from __future__ import annotations

from typing import Any

_COUNTRIES = ("NZ", "Russia", "China", "Other")


def _empty_country_counts() -> dict[str, int]:
    """Return a zeroed country count dict."""
    return {c: 0 for c in _COUNTRIES}


def _resolve_country(row: dict[str, Any]) -> str:
    """Map a row's export country to one of NZ / Russia / China / Other."""
    en = str(row.get("country_export_en", "") or "").strip()
    ko = str(row.get("country_export_ko", "") or "").strip()

    # NZ
    if "New Zealand" in en or "뉴질랜드" in ko:
        return "NZ"
    # Russia — 'Russia' covers 'Russian Federation' etc.
    if "Russia" in en or "러시아" in ko:
        return "Russia"
    # China
    if "China" in en or "중국" in ko:
        return "China"
    return "Other"


def _resolve_type(row: dict[str, Any]) -> str:
    """Return 'frozen', 'dried', or 'ambiguous' for a row."""
    if row.get("type_frozen"):
        return "frozen"
    if row.get("type_dried"):
        return "dried"
    return "ambiguous"


def aggregate(
    rows: list[dict[str, Any]],
    year: int | None = None,
) -> dict[str, dict[str, int]]:
    """Return the Product Type × Country pivot.

    Args:
        rows: Parsed row dicts.
        year: Optional integer year to filter on. If None, all years included.

    Returns:
        Dict with keys 'frozen', 'dried', 'ambiguous', 'total'.
        Each value is a dict of {"NZ": n, "Russia": n, "China": n, "Other": n}.
    """
    result: dict[str, dict[str, int]] = {
        "frozen":    _empty_country_counts(),
        "dried":     _empty_country_counts(),
        "ambiguous": _empty_country_counts(),
        "total":     _empty_country_counts(),
    }

    for row in rows:
        if year is not None:
            try:
                row_year = int(row.get("year", 0) or 0)
            except (ValueError, TypeError):
                continue
            if row_year != year:
                continue

        country = _resolve_country(row)
        ptype = _resolve_type(row)

        result[ptype][country] += 1
        result["total"][country] += 1

    return result
