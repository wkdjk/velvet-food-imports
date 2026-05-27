"""
Monthly aggregator — Velvet Food Imports.

Counts import records per calendar month, broken down by product type:
  frozen   → type_frozen is True
  dried    → type_dried is True
  ambiguous → type_ambiguous is True (or no flag set)

Input:  list of parsed row dicts as returned by src/parsers/historical_xlsx.parse()
        or src/parsers/mfds_xlsx.parse(). Works on any iterable of such dicts.

Output: dict keyed by (year, month) tuples → count breakdown dict.

        {
            (2016, 1): {"frozen": 3, "dried": 5, "ambiguous": 2, "total": 10},
            (2016, 2): {...},
            ...
        }

Pure function — no Sheets API calls, no file I/O.
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any


def aggregate(rows: list[dict[str, Any]]) -> dict[tuple[int, int], dict[str, int]]:
    """Return monthly counts broken down by product type.

    Args:
        rows: Parsed row dicts. Each dict is expected to contain:
              - 'year'          : str or int  (e.g. '2016' or 2016)
              - 'month'         : str or int  (e.g. '1' or 1)
              - 'type_frozen'   : bool
              - 'type_dried'    : bool
              - 'type_ambiguous': bool

    Returns:
        Dict keyed by (year: int, month: int) tuples.
        Months with no records are not present in the output.
        Values are {'frozen': n, 'dried': n, 'ambiguous': n, 'total': n}.
    """
    counts: dict[tuple[int, int], dict[str, int]] = defaultdict(
        lambda: {"frozen": 0, "dried": 0, "ambiguous": 0, "total": 0}
    )

    for row in rows:
        try:
            year = int(row.get("year", 0) or 0)
            month = int(row.get("month", 0) or 0)
        except (ValueError, TypeError):
            # Rows with unparseable year/month are silently skipped.
            continue

        if year == 0 or month == 0:
            continue

        key = (year, month)
        bucket = counts[key]
        bucket["total"] += 1

        if row.get("type_frozen"):
            bucket["frozen"] += 1
        elif row.get("type_dried"):
            bucket["dried"] += 1
        else:
            # Covers explicit type_ambiguous=True and rows with all flags False/missing.
            bucket["ambiguous"] += 1

    # Convert defaultdict to plain dict for a clean, predictable return type.
    return dict(counts)
