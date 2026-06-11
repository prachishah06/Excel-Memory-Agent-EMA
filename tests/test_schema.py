"""Smoke tests for the M3 workbook fixture generators.

These tests intentionally verify only that the generated workbooks exist,
open successfully, and expose the expected sheet/header layout. They do not
exercise any schema detection logic.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

import ema.config as config
from fixtures.make_fixtures import make_formatted


def _header_values(path: Path, *, sheet_name: str, row: int) -> list[object]:
    """Load a workbook and return the non-empty cell values for a header row."""

    workbook = load_workbook(path)
    try:
        worksheet = workbook[sheet_name]
        return [cell.value for cell in worksheet[row] if cell.value is not None]
    finally:
        workbook.close()


def test_ema_home_fixture_points_config_to_tmp_directory(ema_home: Path) -> None:
    """The `ema_home` fixture should redirect EMA runtime paths into tmp storage."""

    assert config.EMA_HOME == ema_home
    assert config.REGISTRY_PATH == ema_home / "registry.json"
    assert config.BACKUP_DIR == ema_home / "backups"
    assert config.LOG_PATH == ema_home / "ema.log"


def test_generated_fixture_workbooks_exist_and_open(
    plain_wb: Path, table_wb: Path, formula_wb: Path, offset_wb: Path
) -> None:
    """Core fixture workbooks should be created on disk and open successfully."""

    expected = [
        (plain_wb, "FoodLog"),
        (table_wb, "Expenses"),
        (formula_wb, "OrderLines"),
        (offset_wb, "Bloodwork"),
    ]

    for path, sheet_name in expected:
        assert path.exists()

        workbook = load_workbook(path)
        try:
            assert sheet_name in workbook.sheetnames
        finally:
            workbook.close()


def test_generated_fixture_workbooks_expose_expected_headers(
    plain_wb: Path, table_wb: Path, formula_wb: Path, offset_wb: Path
) -> None:
    """Each fixture workbook should contain the documented header row."""

    assert _header_values(plain_wb, sheet_name="FoodLog", row=1) == [
        "Date",
        "Meal",
        "Food",
        "Calories",
    ]
    assert _header_values(table_wb, sheet_name="Expenses", row=1) == [
        "Date",
        "Store",
        "Amount",
        "Category",
    ]
    assert _header_values(formula_wb, sheet_name="OrderLines", row=1) == [
        "Item",
        "Quantity",
        "Price",
        "Total",
    ]
    assert _header_values(offset_wb, sheet_name="Bloodwork", row=3) == [
        "Date",
        "Marker",
        "Value",
        "Unit",
    ]


def test_make_formatted_creates_a_styled_workbook_that_opens(tmp_path: Path) -> None:
    """The formatted workbook generator should produce an openable styled file."""

    path = make_formatted(tmp_path / "formatted.xlsx")

    assert path.exists()

    workbook = load_workbook(path)
    try:
        assert "Metrics" in workbook.sheetnames
        worksheet = workbook["Metrics"]
        headers = [cell.value for cell in worksheet[1] if cell.value is not None]
        assert headers == ["Week", "Weight", "Protein"]
        assert worksheet.freeze_panes == "A2"
        assert len(worksheet._charts) == 1
    finally:
        workbook.close()
