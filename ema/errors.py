"""Typed exception hierarchy for Excel Memory Agent.

The library raises these exceptions so that the CLI and MCP front-ends can
translate failures into clean, user-facing messages without swallowing
unexpected errors.
"""

from __future__ import annotations


class EmaError(Exception):
    """Base class for all EMA-specific errors."""


class WorkbookNotFoundError(EmaError):
    """Raised when a workbook ID is not present in the registry."""


class WorkbookFileMissingError(EmaError):
    """Raised when a registered workbook path no longer exists on disk."""


class WorkbookLockedError(EmaError):
    """Raised when Excel or a sync process appears to have a workbook locked."""


class SheetNotFoundError(EmaError):
    """Raised when the requested worksheet is not present in a workbook."""


class SchemaMismatchError(EmaError):
    """Raised when a live workbook no longer matches its persisted schema."""


class ValidationError(EmaError):
    """Raised when row values fail EMA's schema or writer validation."""


class WriteVerificationError(EmaError):
    """Raised when post-write verification fails after saving a workbook."""


class UndoError(EmaError):
    """Raised when an undo operation cannot be completed safely."""


class RegistryCorruptError(EmaError):
    """Raised when the registry JSON cannot be read or validated."""


class LLMError(EmaError):
    """Base class for all LLM-related failures."""


class LLMUnavailableError(LLMError):
    """Raised when the configured local model service is unavailable."""


class ExtractionError(LLMError):
    """Raised when text extraction cannot produce valid structured output."""


__all__ = [
    "EmaError",
    "ExtractionError",
    "LLMError",
    "LLMUnavailableError",
    "RegistryCorruptError",
    "SchemaMismatchError",
    "SheetNotFoundError",
    "UndoError",
    "ValidationError",
    "WorkbookFileMissingError",
    "WorkbookLockedError",
    "WorkbookNotFoundError",
    "WriteVerificationError",
]
