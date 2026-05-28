"""
Unit tests for src/parsers/translation_table.py  (A-25).

All four lookup categories are exercised:
  - countries_of_origin
  - product_types
  - importers
  - exporters

The TranslationTable is constructed directly from a minimal synthetic xlsx
(built as a temp file in memory) so that no real Commander xlsx file, no
Google Sheets, and no credentials are needed.

Run:  PYTHONPATH=. python -m pytest tests/test_translation.py -v
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import openpyxl
import pytest

from src.parsers.translation_table import TranslationTable


# ---------------------------------------------------------------------------
# Fixture — minimal synthetic xlsx that mimics the 'Names for Vlookup' tab
# ---------------------------------------------------------------------------

def _make_vlookup_xlsx(tmp_path: Path) -> Path:
    """Write a minimal xlsx containing the 'Names for Vlookup' tab.

    Layout mirrors the real Commander file as documented in translation_table.py:
    - Column A is empty (None) — the parser only reads columns B and C.
    - The sentinel row for country_of_origin is (None, '뉴질랜드', 'New Zealand').
    - The sentinel row for product_type is (None, 'Product name (Korean)', 'Translation').
    - The sentinel row for importer   is (None, 'Company Korean name', 'Company English name').
    - The sentinel row for exporter   is (None, 'Company', 'Country').
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Names for Vlookup"

    # ---- Countries section ------------------------------------------------
    # Sentinel row doubles as the first data row for countries
    ws.append([None, "뉴질랜드", "New Zealand"])
    ws.append([None, "러시아",   "Russia"])
    ws.append([None, "중국",     "China"])

    # ---- Product types section --------------------------------------------
    ws.append([None, "Product name (Korean)", "Translation"])
    ws.append([None, "생녹용(냉동)", "Frozen velvet antler"])
    ws.append([None, "녹용(건조)",   "Dried velvet antler"])
    ws.append([None, "녹용",         "Velvet antler"])

    # ---- Importers section ------------------------------------------------
    ws.append([None, "Company Korean name", "Company English name"])
    ws.append([None, "한국수입사",   "Korea Importer Co"])
    ws.append([None, "서울벨벳",     "Seoul Velvet Ltd"])
    ws.append([None, "부산무역",     "Busan Trade Inc"])

    # ---- Exporters section ------------------------------------------------
    ws.append([None, "Company", "Country"])
    ws.append([None, "NZ Velvet Exports", "New Zealand"])
    ws.append([None, "Siberian Antler",   "Russia"])

    out = tmp_path / "test_vlookup.xlsx"
    wb.save(str(out))
    return out


@pytest.fixture(scope="module")
def vlookup_path(tmp_path_factory: pytest.TempPathFactory) -> Path:
    tmp = tmp_path_factory.mktemp("vlookup")
    return _make_vlookup_xlsx(tmp)


@pytest.fixture(scope="module")
def table(vlookup_path: Path) -> TranslationTable:
    return TranslationTable(xlsx_path=vlookup_path)


# ---------------------------------------------------------------------------
# Category 1 — countries_of_origin
# ---------------------------------------------------------------------------

class TestCountriesOfOrigin:
    def test_new_zealand_translates(self, table: TranslationTable) -> None:
        assert table.country_of_origin("뉴질랜드") == "New Zealand"

    def test_russia_translates(self, table: TranslationTable) -> None:
        assert table.country_of_origin("러시아") == "Russia"

    def test_china_translates(self, table: TranslationTable) -> None:
        assert table.country_of_origin("중국") == "China"

    def test_unknown_country_returns_original(self, table: TranslationTable) -> None:
        """Unknown key must return the original string, never raise."""
        assert table.country_of_origin("캐나다") == "캐나다"

    def test_empty_country_returns_empty(self, table: TranslationTable) -> None:
        assert table.country_of_origin("") == ""

    def test_whitespace_stripped_before_lookup(self, table: TranslationTable) -> None:
        """Surrounding whitespace in the key must not break the lookup."""
        assert table.country_of_origin("  뉴질랜드  ") == "New Zealand"


# ---------------------------------------------------------------------------
# Category 2 — product_types
# ---------------------------------------------------------------------------

class TestProductTypes:
    def test_frozen_velvet_translates(self, table: TranslationTable) -> None:
        assert table.product_type("생녹용(냉동)") == "Frozen velvet antler"

    def test_dried_velvet_translates(self, table: TranslationTable) -> None:
        assert table.product_type("녹용(건조)") == "Dried velvet antler"

    def test_plain_velvet_translates(self, table: TranslationTable) -> None:
        assert table.product_type("녹용") == "Velvet antler"

    def test_unknown_product_type_returns_original(self, table: TranslationTable) -> None:
        assert table.product_type("신제품유형") == "신제품유형"

    def test_empty_product_type_returns_empty(self, table: TranslationTable) -> None:
        assert table.product_type("") == ""


# ---------------------------------------------------------------------------
# Category 3 — importers
# ---------------------------------------------------------------------------

class TestImporters:
    def test_known_importer_translates(self, table: TranslationTable) -> None:
        assert table.importer("한국수입사") == "Korea Importer Co"

    def test_second_importer_translates(self, table: TranslationTable) -> None:
        assert table.importer("서울벨벳") == "Seoul Velvet Ltd"

    def test_third_importer_translates(self, table: TranslationTable) -> None:
        assert table.importer("부산무역") == "Busan Trade Inc"

    def test_unknown_importer_returns_original(self, table: TranslationTable) -> None:
        assert table.importer("미등록회사") == "미등록회사"

    def test_empty_importer_returns_empty(self, table: TranslationTable) -> None:
        assert table.importer("") == ""


# ---------------------------------------------------------------------------
# Category 4 — exporters
# ---------------------------------------------------------------------------

class TestExporters:
    def test_known_exporter_returns_country(self, table: TranslationTable) -> None:
        assert table.exporter_country("NZ Velvet Exports") == "New Zealand"

    def test_second_exporter_returns_country(self, table: TranslationTable) -> None:
        assert table.exporter_country("Siberian Antler") == "Russia"

    def test_unknown_exporter_returns_original(self, table: TranslationTable) -> None:
        assert table.exporter_country("Unknown Exporter Ltd") == "Unknown Exporter Ltd"

    def test_empty_exporter_returns_empty(self, table: TranslationTable) -> None:
        assert table.exporter_country("") == ""


# ---------------------------------------------------------------------------
# Cross-category — loading / singleton behaviour
# ---------------------------------------------------------------------------

class TestTableLoading:
    def test_all_four_dicts_populated(self, table: TranslationTable) -> None:
        """All four lookup dicts must have at least one entry after loading."""
        assert len(table.countries_of_origin) > 0
        assert len(table.product_types) > 0
        assert len(table.importers) > 0
        assert len(table.exporters) > 0

    def test_get_table_returns_new_instance_when_path_given(
        self, vlookup_path: Path
    ) -> None:
        """get_table(path) must bypass the singleton and return a fresh instance."""
        from src.parsers.translation_table import get_table
        t1 = get_table(vlookup_path)
        t2 = get_table(vlookup_path)
        # Each call with an explicit path returns a distinct object
        assert t1 is not t2

    def test_missing_tab_raises_value_error(self, tmp_path: Path) -> None:
        """If the xlsx has no 'Names for Vlookup' tab, ValueError is raised."""
        wb = openpyxl.Workbook()
        wb.active.title = "Sheet1"
        bad = tmp_path / "no_vlookup.xlsx"
        wb.save(str(bad))
        with pytest.raises(ValueError, match="Names for Vlookup"):
            TranslationTable(xlsx_path=bad)
