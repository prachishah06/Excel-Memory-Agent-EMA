"""JSON registry persistence for Excel Memory Agent.

`RegistryStore` owns the `registry.json` file and provides atomic persistence
plus basic CRUD operations for registered workbooks. This module intentionally
contains no Excel-specific logic.
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from pydantic import ValidationError as PydanticValidationError

from ema import config
from ema.errors import RegistryCorruptError, WorkbookNotFoundError
from ema.models import Registry, WorkbookEntry


class RegistryStore:
    """Persist and retrieve the EMA workbook registry."""

    def __init__(self, path: Path = config.REGISTRY_PATH) -> None:
        """Initialize the store for a specific registry file path."""

        self._path = path

    @property
    def path(self) -> Path:
        """Return the registry file path used by this store."""

        return self._path

    def load(self) -> Registry:
        """Load the registry from disk.

        Returns an empty `Registry` when the file does not yet exist.
        Raises `RegistryCorruptError` when the file is unreadable or does not
        validate against the registry model.
        """

        if not self._path.exists():
            return Registry()

        try:
            raw_text = self._path.read_text(encoding="utf-8")
            raw_data = json.loads(raw_text)
            return Registry.model_validate(raw_data)
        except (OSError, json.JSONDecodeError, PydanticValidationError) as exc:
            raise RegistryCorruptError(
                f"Registry file is unreadable or invalid: {self._path}"
            ) from exc

    def save(self, reg: Registry) -> None:
        """Persist the registry atomically using a temp file and `os.replace`."""

        self._path.parent.mkdir(parents=True, exist_ok=True)

        payload = reg.model_dump_json(indent=2)
        temp_path: Path | None = None

        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                delete=False,
                dir=self._path.parent,
                prefix=f"{self._path.name}.",
                suffix=".tmp",
            ) as temp_file:
                temp_file.write(payload)
                temp_file.flush()
                os.fsync(temp_file.fileno())
                temp_path = Path(temp_file.name)

            os.replace(temp_path, self._path)
        except Exception:
            if temp_path is not None and temp_path.exists():
                temp_path.unlink()
            raise

    def add(self, entry: WorkbookEntry) -> None:
        """Add a new workbook entry to the registry.

        Raises `ValueError` when the workbook ID already exists; callers that
        want replace semantics should use `upsert`.
        """

        registry = self.load()
        if entry.id in registry.workbooks:
            raise ValueError(f"Workbook already exists in registry: {entry.id}")

        registry.workbooks[entry.id] = entry
        self.save(registry)

    def get(self, workbook_id: str) -> WorkbookEntry:
        """Return a workbook entry by ID or raise `WorkbookNotFoundError`."""

        registry = self.load()
        try:
            return registry.workbooks[workbook_id]
        except KeyError as exc:
            raise WorkbookNotFoundError(
                f"Workbook is not registered: {workbook_id}"
            ) from exc

    def list(self) -> list[WorkbookEntry]:
        """Return all registered workbooks in insertion order."""

        registry = self.load()
        return list(registry.workbooks.values())

    def remove(self, workbook_id: str) -> None:
        """Remove a workbook entry from the registry.

        Raises `WorkbookNotFoundError` when the workbook ID is unknown.
        """

        registry = self.load()
        if workbook_id not in registry.workbooks:
            raise WorkbookNotFoundError(f"Workbook is not registered: {workbook_id}")

        del registry.workbooks[workbook_id]
        self.save(registry)

    def upsert(self, entry: WorkbookEntry) -> None:
        """Insert or replace a workbook entry keyed by its workbook ID."""

        registry = self.load()
        registry.workbooks[entry.id] = entry
        self.save(registry)


__all__ = ["RegistryStore"]
