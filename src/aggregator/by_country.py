"""
By-country aggregator — Velvet Food Imports.

Produces annual import counts per export country for 2016–present.

Country bucketing logic (identical to pivot.py):
  - 'NZ'     : country_export_en contains 'New Zealand' OR country_export_ko contains '뉴질랜드'
  - 'Russia'  : country_export_en contains 'Russia'     OR country_export_ko contains '러시아'
  - 'China'   : country_export_en contains 'China'      OR country_export_ko contains '중국'
  - 'Other'   : anything else (including empty strings)

Input:  list of parsed row dicts from historical_xlsx.parse() or mfds_xlsx.parse().

Output:

    {
        2016: {"NZ": 45, "Russia": 10, "China": 2, "Other": 3, "total": 60},
        2017: {...},
        ...
    }

Pure function — no Sheets API calls, no file I/O.
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any

_COUNTRIES = ("NZ", "Russia", "China", "Other")


def _resolve_country(row: dict[str, Any]) -> str:
    """Map a row's export country to one of NZ / Russia / China / Other.

    Uses the same bucketing rules as pivot.py for consistency.
    """
    en = str(row.get("country_export_en", "") or "").strip()
    ko = str(row.get("country_export_ko", "") or "").strip()

    if "New Zealand" in en or "뉴질랜드" in ko:
        return "NZ"
    if "Russia" in en or "러시아" in ko:
        return "Russia"
    if "China" in en or "중국" in ko:
        return "China"
    return "Other"


def aggregate(rows: list[dict[str, Any]]) -> dict[int, dict[str, int]]:
    """Return annual counts per export country.

    Args:
        rows: Parsed row dicts.

    Returns:
        Dict keyed by year (int).
        Each value is {"NZ": n, "Russia": n, "China": n, "Other": n, "total": n}.
        Years with no valid records are not included.
    """
    # accumulator: year → country → count
    counts: dict[int, dict[str, int]] = defaultdict(
        lambda: {c: 0 for c in (*_COUNTRIES, "total")}
    )

    for row in rows:
        try:
            year = int(row.get("year", 0) or 0)
        except (ValueError, TypeError):
            continue

        if year == 0:
            continue

        country = _resolve_country(row)
        counts[year][country] += 1
        counts[year]["total"] += 1

    # Return sorted by year ascending; convert defaultdict to plain dict.
    return {yr: counts[yr] for yr in sorted(counts)}
