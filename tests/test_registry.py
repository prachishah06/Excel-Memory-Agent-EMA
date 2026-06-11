"""Tests for the M2 registry persistence layer."""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path

import pytest

from ema.errors import RegistryCorruptError, WorkbookNotFoundError
from ema.models import ColumnDef, Registry, WorkbookEntry, WorkbookSchema
from ema.registry import RegistryStore


def _make_entry(tmp_path: Path, *, workbook_id: str = "foodlog") -> WorkbookEntry:
    """Build a representative workbook entry for registry tests."""

    timestamp = datetime(2026, 6, 11, 12, 0, tzinfo=timezone.utc)
    return WorkbookEntry(
        id=workbook_id,
        name="Food Log",
        path=tmp_path / "FoodLog.xlsx",
        primary_sheet="Food",
        schema=WorkbookSchema(
            schema_version=1,
            sheet="Food",
            header_row=1,
            table_name=None,
            columns=[
                ColumnDef(name="Date", index=0),
                ColumnDef(name="Food", index=1),
            ],
            data_start_row=2,
        ),
        created_at=timestamp,
        updated_at=timestamp,
        last_undo_token=None,
    )


def test_load_missing_file_returns_empty_registry(tmp_path: Path) -> None:
    """A missing registry file should behave like a fresh empty registry."""

    store = RegistryStore(tmp_path / "registry.json")

    loaded = store.load()

    assert loaded == Registry()


def test_save_and_load_round_trip_registry(tmp_path: Path) -> None:
    """Saved registries should round-trip through disk without data loss."""

    path = tmp_path / "registry.json"
    store = RegistryStore(path)
    entry = _make_entry(tmp_path)
    registry = Registry(workbooks={entry.id: entry})

    store.save(registry)
    loaded = store.load()

    assert loaded == registry
    assert path.exists()
    parsed = json.loads(path.read_text(encoding="utf-8"))
    assert parsed["workbooks"]["foodlog"]["name"] == "Food Log"


def test_atomic_save_leaves_no_partial_file_on_replace_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """A failed atomic replace should preserve the existing registry file."""

    path = tmp_path / "registry.json"
    store = RegistryStore(path)
    original_entry = _make_entry(tmp_path, workbook_id="original")
    updated_entry = _make_entry(tmp_path, workbook_id="updated")

    store.save(Registry(workbooks={original_entry.id: original_entry}))
    original_bytes = path.read_bytes()

    def fake_replace(src: os.PathLike[str] | str, dst: os.PathLike[str] | str) -> None:
        raise OSError("simulated replace failure")

    monkeypatch.setattr("ema.registry.os.replace", fake_replace)

    with pytest.raises(OSError, match="simulated replace failure"):
        store.save(Registry(workbooks={updated_entry.id: updated_entry}))

    assert path.read_bytes() == original_bytes
    leftover_temps = list(tmp_path.glob("registry.json.*.tmp"))
    assert leftover_temps == []


def test_get_unknown_id_raises_workbook_not_found(tmp_path: Path) -> None:
    """Lookup by unknown workbook ID should raise the typed not-found error."""

    store = RegistryStore(tmp_path / "registry.json")

    with pytest.raises(WorkbookNotFoundError, match="missing"):
        store.get("missing")


def test_load_corrupt_json_raises_registry_corrupt_error(tmp_path: Path) -> None:
    """Unreadable JSON should raise the typed registry corruption error."""

    path = tmp_path / "registry.json"
    path.write_text("{not valid json", encoding="utf-8")
    store = RegistryStore(path)

    with pytest.raises(RegistryCorruptError):
        store.load()


def test_crud_methods_persist_expected_registry_state(tmp_path: Path) -> None:
    """CRUD helpers should add, replace, list, and remove workbook entries."""

    store = RegistryStore(tmp_path / "registry.json")
    first = _make_entry(tmp_path, workbook_id="foodlog")
    replacement = _make_entry(tmp_path, workbook_id="foodlog")
    replacement.name = "Updated Food Log"
    second = _make_entry(tmp_path, workbook_id="expenses")
    second.name = "Expense Log"

    store.add(first)
    store.upsert(replacement)
    store.add(second)

    listed = store.list()
    assert [entry.id for entry in listed] == ["foodlog", "expenses"]
    assert store.get("foodlog").name == "Updated Food Log"

    store.remove("expenses")

    assert [entry.id for entry in store.list()] == ["foodlog"]
