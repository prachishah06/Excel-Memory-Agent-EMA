"""Tests for the M1 contract layer."""

from __future__ import annotations

import importlib
from datetime import datetime, timezone
from pathlib import Path

import pytest
from pydantic import ValidationError as PydanticValidationError

import ema.config as config
from ema import errors
from ema.models import (
    AppendRequest,
    ColumnDef,
    ColumnType,
    ExtractedField,
    ExtractionResult,
    FieldOrigin,
    Registry,
    WorkbookEntry,
    WorkbookSchema,
)


def _sample_schema() -> WorkbookSchema:
    """Build a representative workbook schema for contract tests."""

    return WorkbookSchema(
        schema_version=1,
        sheet="Food",
        header_row=1,
        table_name="FoodTable",
        columns=[
            ColumnDef(name="Date", index=0, type=ColumnType.DATE, required=True),
            ColumnDef(name="Food", index=1, type=ColumnType.TEXT, required=True),
            ColumnDef(name="Calories", index=2, type=ColumnType.NUMBER),
        ],
        data_start_row=2,
    )


def _sample_workbook_entry(tmp_path: Path) -> WorkbookEntry:
    """Build a representative workbook registry entry for round-trip tests."""

    timestamp = datetime(2026, 6, 11, 12, 0, tzinfo=timezone.utc)
    return WorkbookEntry(
        id="foodlog",
        name="Food Log",
        path=tmp_path / "FoodLog.xlsx",
        primary_sheet="Food",
        schema=_sample_schema(),
        created_at=timestamp,
        updated_at=timestamp,
        last_undo_token="backup-20260611T120000Z.xlsx",
    )


def _sample_extraction_result() -> ExtractionResult:
    """Build a representative extraction result for contract tests."""

    return ExtractionResult(
        workbook_id="foodlog",
        sheet="Food",
        fields=[
            ExtractedField(
                column="Date",
                value="2026-06-11",
                origin=FieldOrigin.DEFAULT,
                note="Filled from today's date.",
            ),
            ExtractedField(
                column="Food",
                value="Skyr",
                origin=FieldOrigin.EXTRACTED,
            ),
            ExtractedField(
                column="Calories",
                value=120,
                origin=FieldOrigin.GENERATED,
                note="Estimated from the description.",
            ),
        ],
        intent="Log a food entry.",
        warnings=["Calories were estimated."],
        model="qwen3:4b",
        needs_confirmation=True,
    )


def test_append_request_rejects_unknown_keys() -> None:
    """AppendRequest should fail loudly on unexpected input keys."""

    with pytest.raises(PydanticValidationError):
        AppendRequest.model_validate(
            {
                "workbook_id": "foodlog",
                "values": {"Food": "Skyr"},
                "dry_run": True,
                "unexpected": "boom",
            }
        )


def test_workbook_entry_round_trip_json(tmp_path: Path) -> None:
    """WorkbookEntry should serialize to and from JSON without data loss."""

    entry = _sample_workbook_entry(tmp_path)

    restored = WorkbookEntry.model_validate_json(entry.model_dump_json())

    assert restored == entry
    assert restored.path == entry.path
    assert restored.schema.columns[0].type is ColumnType.DATE


def test_registry_round_trip_json(tmp_path: Path) -> None:
    """Registry should round-trip cleanly through its JSON representation."""

    entry = _sample_workbook_entry(tmp_path)
    registry = Registry(workbooks={entry.id: entry})

    restored = Registry.model_validate_json(registry.model_dump_json())

    assert restored == registry
    assert restored.workbooks["foodlog"].name == "Food Log"


def test_extraction_result_round_trip_json() -> None:
    """ExtractionResult should preserve fields and origins through JSON."""

    result = _sample_extraction_result()

    restored = ExtractionResult.model_validate_json(result.model_dump_json())

    assert restored == result
    assert restored.fields[2].origin is FieldOrigin.GENERATED


def test_extraction_result_to_append_request_maps_fields_to_values() -> None:
    """ExtractionResult should convert field entries into AppendRequest values."""

    result = _sample_extraction_result()

    request = result.to_append_request(dry_run=True)

    assert request == AppendRequest(
        workbook_id="foodlog",
        sheet="Food",
        values={
            "Date": "2026-06-11",
            "Food": "Skyr",
            "Calories": 120,
        },
        dry_run=True,
    )


def test_config_constants_support_env_overrides(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Config constants should honor the documented environment overrides."""

    with monkeypatch.context() as env:
        env.setenv("EMA_HOME", str(tmp_path / "ema-home"))
        env.setenv("EMA_LLM_ENABLED", "true")
        env.setenv("EMA_LLM_MODEL", "custom-model")
        env.setenv("EMA_OLLAMA_HOST", "http://127.0.0.1:9999")

        importlib.reload(config)

        assert config.EMA_HOME == tmp_path / "ema-home"
        assert config.REGISTRY_PATH == config.EMA_HOME / "registry.json"
        assert config.BACKUP_DIR == config.EMA_HOME / "backups"
        assert config.LOG_PATH == config.EMA_HOME / "ema.log"
        assert config.LLM_ENABLED is True
        assert config.LLM_MODEL == "custom-model"
        assert config.OLLAMA_HOST == "http://127.0.0.1:9999"
        assert config.REQUIRE_CONFIRM_FOR_LLM is True
        assert config.EXPOSE_TEXT_TOOL is False

    importlib.reload(config)


def test_error_hierarchy_matches_architecture() -> None:
    """Every EMA error type should inherit from the documented base class."""

    assert issubclass(errors.WorkbookNotFoundError, errors.EmaError)
    assert issubclass(errors.WorkbookFileMissingError, errors.EmaError)
    assert issubclass(errors.WorkbookLockedError, errors.EmaError)
    assert issubclass(errors.SheetNotFoundError, errors.EmaError)
    assert issubclass(errors.SchemaMismatchError, errors.EmaError)
    assert issubclass(errors.ValidationError, errors.EmaError)
    assert issubclass(errors.WriteVerificationError, errors.EmaError)
    assert issubclass(errors.UndoError, errors.EmaError)
    assert issubclass(errors.RegistryCorruptError, errors.EmaError)
    assert issubclass(errors.LLMError, errors.EmaError)
    assert issubclass(errors.LLMUnavailableError, errors.LLMError)
    assert issubclass(errors.ExtractionError, errors.LLMError)
