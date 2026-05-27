"""
By-importer aggregator — Velvet Food Imports.

Produces annual import counts per importer company.

Resolution order for the display name:
  1. English name from the translation table (importers dict), keyed on importer_ko.
  2. importer_en from the parsed row (already an English value for many rows).
  3. Fall back to importer_ko (Korean original).

Input:  list of parsed row dicts from historical_xlsx.parse() or mfds_xlsx.parse().

Output: list of dicts sorted by total volume descending:

        [
            {
                "importer_en": "Velvet Co Ltd",
                "importer_ko": "벨벳코",
                "year_counts":  {2016: 3, 2017: 5, ...},
                "total":        8,
            },
            ...
        ]

Pure function — no Sheets API calls, no file I/O.
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any


def aggregate(
    rows: list[dict[str, Any]],
    translation_importers: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """Return per-importer annual counts, sorted by total volume descending.

    Args:
        rows:                 Parsed row dicts.
        translation_importers: Optional mapping of Korean company name → English
                              company name (from TranslationTable.importers).
                              If None, importer_en from the row is used directly.

    Returns:
        List of importer dicts, sorted by 'total' descending.
        Each dict has: importer_en, importer_ko, year_counts, total.
    """
    if translation_importers is None:
        translation_importers = {}

    # accumulator: importer_ko → {year → count}
    ko_to_years: dict[str, dict[int, int]] = defaultdict(lambda: defaultdict(int))
    # importer_ko → best English name seen
    ko_to_en: dict[str, str] = {}

    for row in rows:
        ko = str(row.get("importer_ko", "") or "").strip()
        if not ko:
            ko = "__unknown__"

        try:
            year = int(row.get("year", 0) or 0)
        except (ValueError, TypeError):
            year = 0

        if year == 0:
            continue

        ko_to_years[ko][year] += 1

        # Resolve English name — translation table takes priority
        if ko not in ko_to_en:
            en_from_table = translation_importers.get(ko, "")
            en_from_row = str(row.get("importer_en", "") or "").strip()
            # Prefer table lookup, then row value, then fall back to Korean original
            ko_to_en[ko] = en_from_table or en_from_row or ko

    results: list[dict[str, Any]] = []

    for ko, year_dict in ko_to_years.items():
        total = sum(year_dict.values())
        # Convert inner defaultdict to plain dict with int keys sorted ascending.
        year_counts = {yr: year_dict[yr] for yr in sorted(year_dict)}
        results.append({
            "importer_en": ko_to_en.get(ko, ko),
            "importer_ko": ko if ko != "__unknown__" else "",
            "year_counts": year_counts,
            "total": total,
        })

    # Sort by total descending; secondary sort by importer_en for stable ordering.
    results.sort(key=lambda r: (-r["total"], r["importer_en"]))
    return results
