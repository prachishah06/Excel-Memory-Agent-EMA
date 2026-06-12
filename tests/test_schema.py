"""Strict TDD tests for M4 schema discovery and validation.

These tests define the expected behavior for `propose_schema`,
`validate_live_schema`, and `as_extraction_contract` before `ema/schema.py`
is implemented.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

import ema.config as config
import ema.schema as schema_module
from ema.errors import SchemaMismatchError
from ema.models import ColumnDef, ColumnType, WorkbookSchema


def _column_names(columns: list[ColumnDef]) -> list[str]:
    """Return column names in schema order."""

    return [column.name for column in columns]


def _schema_from_detected_columns(
    *,
    sheet: str,
    header_row: int,
    table_name: str | None,
    columns: list[ColumnDef],
    data_start_row: int,
) -> WorkbookSchema:
    """Build a persisted schema object from a detected schema proposal."""

    return WorkbookSchema(
        schema_version=config.SCHEMA_VERSION,
        sheet=sheet,
        header_row=header_row,
        table_name=table_name,
        columns=columns,
        data_start_row=data_start_row,
    )


def test_propose_schema_plain_workbook_detects_headers_and_types(
    plain_wb: Path,
) -> None:
    """Plain workbooks should produce a non-table schema from row 1 headers."""

    proposal = schema_module.propose_schema(plain_wb, "FoodLog")

    assert proposal.sheet == "FoodLog"
    assert proposal.header_row == 1
    assert proposal.table_name is None
    assert _column_names(proposal.columns) == ["Date", "Meal", "Food", "Calories"]
    assert [column.index for column in proposal.columns] == [0, 1, 2, 3]
    assert proposal.columns[3].type is ColumnType.NUMBER
    assert proposal.warnings == []


def test_propose_schema_table_workbook_detects_excel_table_and_headers(
    table_wb: Path,
) -> None:
    """Table workbooks should detect the Excel Table name and table headers."""

    proposal = schema_module.propose_schema(table_wb, "Expenses")

    assert proposal.sheet == "Expenses"
    assert proposal.header_row == 1
    assert proposal.table_name == "ExpensesTable"
    assert _column_names(proposal.columns) == ["Date", "Store", "Amount", "Category"]
    assert [column.index for column in proposal.columns] == [0, 1, 2, 3]
    assert proposal.columns[2].type is ColumnType.NUMBER
    assert proposal.warnings == []


def test_propose_schema_formula_workbook_flags_formula_column_and_warning(
    formula_wb: Path,
) -> None:
    """Formula workbooks should flag formula columns instead of treating them as normal inputs."""

    proposal = schema_module.propose_schema(formula_wb, "OrderLines")

    assert proposal.sheet == "OrderLines"
    assert proposal.header_row == 1
    assert proposal.table_name is None
    assert _column_names(proposal.columns) == ["Item", "Quantity", "Price", "Total"]
    assert proposal.columns[1].type is ColumnType.NUMBER
    assert proposal.columns[2].type is ColumnType.NUMBER
    assert proposal.columns[3].is_formula is True
    assert any("formula" in warning.lower() for warning in proposal.warnings)


def test_propose_schema_offset_header_workbook_uses_supplied_header_row(
    offset_wb: Path,
) -> None:
    """Offset-header workbooks should respect an explicit header row override."""

    proposal = schema_module.propose_schema(offset_wb, "Bloodwork", header_row=3)

    assert proposal.sheet == "Bloodwork"
    assert proposal.header_row == 3
    assert proposal.table_name is None
    assert _column_names(proposal.columns) == ["Date", "Marker", "Value", "Unit"]
    assert [column.index for column in proposal.columns] == [0, 1, 2, 3]
    assert proposal.columns[2].type is ColumnType.NUMBER


@pytest.mark.parametrize(
    ("fixture_name", "sheet_name", "header_row", "data_start_row"),
    [
        ("plain_wb", "FoodLog", 1, 2),
        ("table_wb", "Expenses", 1, 2),
        ("formula_wb", "OrderLines", 1, 2),
        ("offset_wb", "Bloodwork", 3, 4),
    ],
)
def test_validate_live_schema_accepts_unchanged_workbooks(
    request: pytest.FixtureRequest,
    fixture_name: str,
    sheet_name: str,
    header_row: int,
    data_start_row: int,
) -> None:
    """Live validation should pass when the workbook still matches the persisted schema."""

    workbook_path = request.getfixturevalue(fixture_name)
    proposal = schema_module.propose_schema(
        workbook_path,
        sheet_name,
        None if header_row == 1 else header_row,
    )
    schema = _schema_from_detected_columns(
        sheet=proposal.sheet,
        header_row=proposal.header_row,
        table_name=proposal.table_name,
        columns=proposal.columns,
        data_start_row=data_start_row,
    )

    schema_module.validate_live_schema(workbook_path, schema)


def test_validate_live_schema_raises_schema_mismatch_after_header_rename(
    plain_wb: Path,
) -> None:
    """Live validation should fail when workbook headers drift from the persisted schema."""

    proposal = schema_module.propose_schema(plain_wb, "FoodLog")
    schema = _schema_from_detected_columns(
        sheet=proposal.sheet,
        header_row=proposal.header_row,
        table_name=proposal.table_name,
        columns=proposal.columns,
        data_start_row=2,
    )

    workbook = load_workbook(plain_wb)
    try:
        worksheet = workbook["FoodLog"]
        worksheet["C1"] = "Dish"
        workbook.save(plain_wb)
    finally:
        workbook.close()

    with pytest.raises(SchemaMismatchError):
        schema_module.validate_live_schema(plain_wb, schema)


def test_as_extraction_contract_includes_every_column_type_required_and_description() -> None:
    """Extraction contracts should expose all columns with their schema metadata."""

    schema = WorkbookSchema(
        schema_version=config.SCHEMA_VERSION,
        sheet="FoodLog",
        header_row=1,
        table_name=None,
        columns=[
            ColumnDef(
                name="Date",
                index=0,
                type=ColumnType.DATE,
                required=True,
                description="Meal date in ISO format.",
            ),
            ColumnDef(
                name="Meal",
                index=1,
                type=ColumnType.TEXT,
                required=True,
                description="Meal name such as breakfast or dinner.",
            ),
            ColumnDef(
                name="Calories",
                index=2,
                type=ColumnType.NUMBER,
                required=False,
                description="Approximate calories for the meal.",
            ),
        ],
        data_start_row=2,
    )

    contract = schema_module.as_extraction_contract(schema)

    assert contract["type"] == "object"
    assert contract["additionalProperties"] is False
    assert set(contract["properties"]) == {"Date", "Meal", "Calories"}
    assert set(contract["required"]) == {"Date", "Meal"}
    assert contract["properties"]["Date"]["type"] == "string"
    assert contract["properties"]["Date"]["description"] == "Meal date in ISO format."
    assert contract["properties"]["Meal"]["type"] == "string"
    assert (
        contract["properties"]["Meal"]["description"]
        == "Meal name such as breakfast or dinner."
    )
    assert contract["properties"]["Calories"]["type"] == "number"
    assert (
        contract["properties"]["Calories"]["description"]
        == "Approximate calories for the meal."
    )
