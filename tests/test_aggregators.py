"""
Unit tests for Group 2 aggregators: A-9, A-10, A-11, A-12.

Run:  PYTHONPATH=. python -m pytest tests/test_aggregators.py -v

All tests use synthetic row dicts — no file I/O, no Sheets API calls.
"""

from __future__ import annotations

import pytest

from src.aggregator import monthly, by_importer, pivot, by_country

# ---------------------------------------------------------------------------
# Helpers — row factory
# ---------------------------------------------------------------------------


def _row(
    year: int = 2020,
    month: int = 1,
    importer_ko: str = "수입사A",
    importer_en: str = "Importer A",
    country_export_en: str = "New Zealand",
    country_export_ko: str = "뉴질랜드",
    type_frozen: bool = False,
    type_dried: bool = True,
    type_ambiguous: bool = False,
) -> dict:
    return {
        "year": year,
        "month": month,
        "importer_ko": importer_ko,
        "importer_en": importer_en,
        "country_export_en": country_export_en,
        "country_export_ko": country_export_ko,
        "type_frozen": type_frozen,
        "type_dried": type_dried,
        "type_ambiguous": type_ambiguous,
    }


# ===========================================================================
# A-9: monthly.aggregate
# ===========================================================================


class TestMonthly:
    """Tests for src/aggregator/monthly.py — A-9."""

    def test_empty_input_returns_empty_dict(self):
        assert monthly.aggregate([]) == {}

    def test_single_row_dried(self):
        result = monthly.aggregate([_row(year=2020, month=3, type_frozen=False, type_dried=True, type_ambiguous=False)])
        assert result == {(2020, 3): {"frozen": 0, "dried": 1, "ambiguous": 0, "total": 1}}

    def test_single_row_frozen(self):
        result = monthly.aggregate([_row(year=2021, month=6, type_frozen=True, type_dried=False, type_ambiguous=False)])
        assert result == {(2021, 6): {"frozen": 1, "dried": 0, "ambiguous": 0, "total": 1}}

    def test_single_row_ambiguous(self):
        result = monthly.aggregate([_row(year=2019, month=12, type_frozen=False, type_dried=False, type_ambiguous=True)])
        assert result == {(2019, 12): {"frozen": 0, "dried": 0, "ambiguous": 1, "total": 1}}

    def test_all_flags_false_counted_as_ambiguous(self):
        """Rows with all flags False should be counted as ambiguous (defensive fallback)."""
        row = _row(type_frozen=False, type_dried=False, type_ambiguous=False)
        result = monthly.aggregate([row])
        assert result[(2020, 1)]["ambiguous"] == 1
        assert result[(2020, 1)]["total"] == 1

    def test_multi_row_single_month(self):
        rows = [
            _row(year=2020, month=5, type_frozen=True, type_dried=False, type_ambiguous=False),
            _row(year=2020, month=5, type_frozen=True, type_dried=False, type_ambiguous=False),
            _row(year=2020, month=5, type_frozen=False, type_dried=True, type_ambiguous=False),
        ]
        result = monthly.aggregate(rows)
        assert result[(2020, 5)] == {"frozen": 2, "dried": 1, "ambiguous": 0, "total": 3}

    def test_multi_year_separate_keys(self):
        rows = [
            _row(year=2018, month=1, type_frozen=True, type_dried=False, type_ambiguous=False),
            _row(year=2019, month=1, type_frozen=False, type_dried=True, type_ambiguous=False),
        ]
        result = monthly.aggregate(rows)
        assert (2018, 1) in result
        assert (2019, 1) in result
        assert result[(2018, 1)]["frozen"] == 1
        assert result[(2019, 1)]["dried"] == 1

    def test_year_zero_skipped(self):
        """Rows with year=0 (unparseable) must not appear in output."""
        rows = [_row(year=0, month=1)]
        assert monthly.aggregate(rows) == {}

    def test_month_zero_skipped(self):
        """Rows with month=0 must not appear in output."""
        rows = [_row(year=2020, month=0)]
        assert monthly.aggregate(rows) == {}

    def test_string_year_month_parsed_correctly(self):
        """Parser may return year/month as strings — aggregator must handle both."""
        row = _row(year=2022, month=8)
        row["year"] = "2022"
        row["month"] = "8"
        result = monthly.aggregate([row])
        assert (2022, 8) in result
        assert result[(2022, 8)]["total"] == 1

    def test_total_equals_sum_of_types(self):
        rows = [
            _row(year=2020, month=1, type_frozen=True, type_dried=False, type_ambiguous=False),
            _row(year=2020, month=1, type_frozen=False, type_dried=True, type_ambiguous=False),
            _row(year=2020, month=1, type_frozen=False, type_dried=False, type_ambiguous=True),
        ]
        result = monthly.aggregate(rows)
        b = result[(2020, 1)]
        assert b["total"] == b["frozen"] + b["dried"] + b["ambiguous"]


# ===========================================================================
# A-10: by_importer.aggregate
# ===========================================================================


class TestByImporter:
    """Tests for src/aggregator/by_importer.py — A-10."""

    def test_empty_input_returns_empty_list(self):
        assert by_importer.aggregate([]) == []

    def test_single_importer_single_year(self):
        rows = [_row(year=2020, importer_ko="수입사A", importer_en="Importer A")]
        result = by_importer.aggregate(rows)
        assert len(result) == 1
        assert result[0]["importer_ko"] == "수입사A"
        assert result[0]["year_counts"] == {2020: 1}
        assert result[0]["total"] == 1

    def test_translation_table_name_takes_priority(self):
        """When a translation dict is provided, its value should override the row's importer_en."""
        rows = [_row(year=2020, importer_ko="수입사A", importer_en="Old Name")]
        result = by_importer.aggregate(rows, translation_importers={"수입사A": "New Name from Table"})
        assert result[0]["importer_en"] == "New Name from Table"

    def test_fallback_to_row_importer_en(self):
        """Without a translation table entry, importer_en from the row is used."""
        rows = [_row(year=2020, importer_ko="수입사B", importer_en="Fallback English")]
        result = by_importer.aggregate(rows, translation_importers={})
        assert result[0]["importer_en"] == "Fallback English"

    def test_fallback_to_korean_when_no_english(self):
        """When no English name is available at all, Korean original is used."""
        row = _row(year=2020, importer_ko="수입사C", importer_en="")
        result = by_importer.aggregate([row], translation_importers={})
        assert result[0]["importer_en"] == "수입사C"

    def test_sorted_by_total_descending(self):
        rows = [
            _row(year=2020, importer_ko="소규모", importer_en="Small Co"),
            _row(year=2020, importer_ko="대규모", importer_en="Big Co"),
            _row(year=2021, importer_ko="대규모", importer_en="Big Co"),
            _row(year=2022, importer_ko="대규모", importer_en="Big Co"),
        ]
        result = by_importer.aggregate(rows)
        assert result[0]["importer_ko"] == "대규모"
        assert result[1]["importer_ko"] == "소규모"

    def test_multi_year_counts_accumulated(self):
        rows = [
            _row(year=2018, importer_ko="수입사A"),
            _row(year=2019, importer_ko="수입사A"),
            _row(year=2019, importer_ko="수입사A"),
            _row(year=2020, importer_ko="수입사A"),
        ]
        result = by_importer.aggregate(rows)
        assert result[0]["year_counts"] == {2018: 1, 2019: 2, 2020: 1}
        assert result[0]["total"] == 4

    def test_row_with_empty_importer_ko_skipped_if_no_year(self):
        """Row with year=0 should not appear in output."""
        rows = [_row(year=0, importer_ko="수입사A")]
        result = by_importer.aggregate(rows)
        assert result == []

    def test_two_importers_correct_totals(self):
        rows = [
            _row(year=2020, importer_ko="A사"),
            _row(year=2020, importer_ko="A사"),
            _row(year=2020, importer_ko="B사"),
        ]
        result = by_importer.aggregate(rows)
        totals = {r["importer_ko"]: r["total"] for r in result}
        assert totals["A사"] == 2
        assert totals["B사"] == 1

    def test_year_counts_keys_are_integers(self):
        rows = [_row(year=2020, importer_ko="수입사A")]
        result = by_importer.aggregate(rows)
        for key in result[0]["year_counts"]:
            assert isinstance(key, int)


# ===========================================================================
# A-11: pivot.aggregate
# ===========================================================================


class TestPivot:
    """Tests for src/aggregator/pivot.py — A-11."""

    def test_empty_input_returns_zeroed_structure(self):
        result = pivot.aggregate([])
        assert set(result.keys()) == {"frozen", "dried", "ambiguous", "total"}
        for ptype in result.values():
            for country_count in ptype.values():
                assert country_count == 0

    def test_nz_classification_by_english(self):
        row = _row(country_export_en="New Zealand", country_export_ko="", type_dried=True)
        result = pivot.aggregate([row])
        assert result["dried"]["NZ"] == 1
        assert result["total"]["NZ"] == 1

    def test_nz_classification_by_korean(self):
        row = _row(country_export_en="", country_export_ko="뉴질랜드", type_dried=True)
        result = pivot.aggregate([row])
        assert result["dried"]["NZ"] == 1

    def test_russia_classification(self):
        row = _row(country_export_en="Russia", country_export_ko="러시아", type_frozen=True, type_dried=False)
        result = pivot.aggregate([row])
        assert result["frozen"]["Russia"] == 1
        assert result["total"]["Russia"] == 1

    def test_china_classification(self):
        row = _row(country_export_en="China", country_export_ko="중국", type_frozen=True, type_dried=False)
        result = pivot.aggregate([row])
        assert result["frozen"]["China"] == 1

    def test_other_bucket_catches_unknowns(self):
        row = _row(country_export_en="Canada", country_export_ko="캐나다", type_dried=True)
        result = pivot.aggregate([row])
        assert result["dried"]["Other"] == 1
        assert result["total"]["Other"] == 1

    def test_other_bucket_for_empty_country(self):
        row = _row(country_export_en="", country_export_ko="")
        result = pivot.aggregate([row])
        assert result["total"]["Other"] == 1

    def test_year_filter_includes_matching_rows_only(self):
        rows = [
            _row(year=2020, country_export_en="New Zealand", type_dried=True),
            _row(year=2021, country_export_en="Russia", type_frozen=True, type_dried=False),
        ]
        result = pivot.aggregate(rows, year=2020)
        assert result["dried"]["NZ"] == 1
        assert result["frozen"]["Russia"] == 0

    def test_year_filter_none_aggregates_all_years(self):
        rows = [
            _row(year=2020, country_export_en="New Zealand", type_dried=True),
            _row(year=2021, country_export_en="New Zealand", type_dried=True),
        ]
        result = pivot.aggregate(rows, year=None)
        assert result["dried"]["NZ"] == 2

    def test_ambiguous_type_bucketed_correctly(self):
        row = _row(
            country_export_en="New Zealand",
            country_export_ko="뉴질랜드",
            type_frozen=False,
            type_dried=False,
            type_ambiguous=True,
        )
        result = pivot.aggregate([row])
        assert result["ambiguous"]["NZ"] == 1
        assert result["frozen"]["NZ"] == 0
        assert result["dried"]["NZ"] == 0

    def test_total_equals_sum_across_types(self):
        rows = [
            _row(country_export_en="New Zealand", country_export_ko="뉴질랜드", type_dried=True),
            _row(country_export_en="New Zealand", country_export_ko="뉴질랜드", type_frozen=True, type_dried=False),
            _row(country_export_en="Russia", country_export_ko="러시아", type_dried=True),
        ]
        result = pivot.aggregate(rows)
        assert result["total"]["NZ"] == result["frozen"]["NZ"] + result["dried"]["NZ"] + result["ambiguous"]["NZ"]
        assert result["total"]["Russia"] == 1


# ===========================================================================
# A-12: by_country.aggregate
# ===========================================================================


class TestByCountry:
    """Tests for src/aggregator/by_country.py — A-12."""

    def test_empty_input_returns_empty_dict(self):
        assert by_country.aggregate([]) == {}

    def test_single_row_nz(self):
        row = _row(year=2020, country_export_en="New Zealand")
        result = by_country.aggregate([row])
        assert result[2020]["NZ"] == 1
        assert result[2020]["total"] == 1

    def test_single_row_russia(self):
        row = _row(year=2019, country_export_en="Russia", country_export_ko="러시아")
        result = by_country.aggregate([row])
        assert result[2019]["Russia"] == 1

    def test_single_row_china(self):
        row = _row(year=2021, country_export_en="China", country_export_ko="중국")
        result = by_country.aggregate([row])
        assert result[2021]["China"] == 1

    def test_other_bucket(self):
        row = _row(year=2022, country_export_en="Australia", country_export_ko="호주")
        result = by_country.aggregate([row])
        assert result[2022]["Other"] == 1
        assert result[2022]["total"] == 1

    def test_multi_year_separate_entries(self):
        rows = [
            _row(year=2018, country_export_en="New Zealand", country_export_ko="뉴질랜드"),
            _row(year=2019, country_export_en="Russia", country_export_ko="러시아"),
        ]
        result = by_country.aggregate(rows)
        assert 2018 in result
        assert 2019 in result
        assert result[2018]["NZ"] == 1
        assert result[2019]["Russia"] == 1

    def test_year_zero_skipped(self):
        rows = [_row(year=0, country_export_en="New Zealand")]
        result = by_country.aggregate(rows)
        assert result == {}

    def test_total_is_sum_of_all_countries(self):
        rows = [
            _row(year=2020, country_export_en="New Zealand"),
            _row(year=2020, country_export_en="New Zealand"),
            _row(year=2020, country_export_en="Russia"),
            _row(year=2020, country_export_en="Canada"),
        ]
        result = by_country.aggregate(rows)
        y = result[2020]
        assert y["total"] == y["NZ"] + y["Russia"] + y["China"] + y["Other"]
        assert y["total"] == 4

    def test_output_sorted_by_year_ascending(self):
        rows = [
            _row(year=2022, country_export_en="New Zealand"),
            _row(year=2016, country_export_en="New Zealand"),
            _row(year=2019, country_export_en="New Zealand"),
        ]
        result = by_country.aggregate(rows)
        years = list(result.keys())
        assert years == sorted(years)

    def test_string_year_parsed_correctly(self):
        """year field may be a string from the parser — must be handled."""
        row = _row(year=2020, country_export_en="New Zealand")
        row["year"] = "2020"
        result = by_country.aggregate([row])
        assert 2020 in result
        assert result[2020]["NZ"] == 1

    def test_nz_detection_by_korean_field(self):
        """NZ must be detected from country_export_ko even when EN field is empty."""
        row = _row(year=2020, country_export_en="", country_export_ko="뉴질랜드")
        result = by_country.aggregate([row])
        assert result[2020]["NZ"] == 1
