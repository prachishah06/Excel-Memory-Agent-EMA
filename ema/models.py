"""Pydantic data models for Excel Memory Agent.

All data crossing EMA module boundaries is represented as a Pydantic v2 model.
These contracts are the stable foundation for the registry, schema layer,
writer, service, and optional extraction flow.
"""

from __future__ import annotations

import warnings
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


ValueType = str | float | int | bool | None
"""Allowed scalar value types for structured row data."""


class ColumnType(str, Enum):
    """Supported logical column types used by EMA validation."""

    TEXT = "text"
    NUMBER = "number"
    DATE = "date"
    BOOL = "bool"


class ColumnDef(BaseModel):
    """Definition of a single workbook column."""

    name: str = Field(description="Exact header text as it appears in Excel.")
    index: int = Field(description="Zero-based offset within the table or range.")
    type: ColumnType = Field(
        default=ColumnType.TEXT,
        description="Logical type inferred or confirmed for the column.",
    )
    required: bool = Field(
        default=False,
        description="Whether a value is expected for this column during append.",
    )
    is_formula: bool = Field(
        default=False,
        description="Whether the destination column contains formulas and is not writable.",
    )
    description: str | None = Field(
        default=None,
        description="Optional hint used to guide extraction and validation.",
    )


class WorkbookSchema(BaseModel):
    """Persisted schema for one workbook sheet or table."""

    schema_version: int = Field(
        description="Persisted schema version; should match ema.config.SCHEMA_VERSION."
    )
    sheet: str = Field(description="Worksheet name the schema applies to.")
    header_row: int = Field(description="1-based Excel row index containing headers.")
    table_name: str | None = Field(
        default=None,
        description="Excel Table name when the data lives inside a ListObject.",
    )
    columns: list[ColumnDef] = Field(
        description="Ordered column definitions for the target range."
    )
    data_start_row: int = Field(
        description="1-based Excel row index where data rows begin."
    )


with warnings.catch_warnings():
    warnings.filterwarnings(
        "ignore",
        message=r'Field name "schema" in "WorkbookEntry" shadows an attribute in parent "BaseModel"',
        category=UserWarning,
    )

    class WorkbookEntry(BaseModel):
        """Registry entry describing one registered workbook."""

        id: str = Field(description="Stable slug identifier, for example 'foodlog'.")
        name: str = Field(description="Human-readable workbook display name.")
        path: Path = Field(description="Absolute path to the .xlsx workbook file.")
        primary_sheet: str = Field(description="Default worksheet used for operations.")
        schema: WorkbookSchema = Field(description="Persisted schema for the workbook.")
        created_at: datetime = Field(
            description="Timestamp when the workbook was registered."
        )
        updated_at: datetime = Field(
            description="Timestamp of the most recent metadata update."
        )
        last_undo_token: str | None = Field(
            default=None,
            description="Backup token that can restore the most recent write.",
        )


class Registry(BaseModel):
    """Top-level JSON registry document."""

    version: int = Field(default=1, description="Registry document version.")
    workbooks: dict[str, WorkbookEntry] = Field(
        default_factory=dict,
        description="Workbook entries keyed by workbook ID.",
    )


class SchemaProposal(BaseModel):
    """Suggested schema produced during workbook registration."""

    sheet: str = Field(description="Worksheet name the proposal targets.")
    header_row: int = Field(description="1-based Excel row index containing headers.")
    table_name: str | None = Field(
        default=None,
        description="Excel Table name when exactly one table was detected.",
    )
    columns: list[ColumnDef] = Field(
        description="Proposed ordered column definitions."
    )
    warnings: list[str] = Field(
        default_factory=list,
        description="Non-fatal issues discovered during schema detection.",
    )


class AppendRequest(BaseModel):
    """Structured request to append one row to a workbook."""

    model_config = ConfigDict(extra="forbid")

    workbook_id: str = Field(description="Registered workbook identifier.")
    sheet: str | None = Field(
        default=None,
        description="Optional worksheet override; defaults to the workbook primary sheet.",
    )
    values: dict[str, ValueType] = Field(
        description="Row values keyed by column name, not by position.",
    )
    dry_run: bool = Field(
        default=False,
        description="When true, return a preview without touching the workbook file.",
    )


class AppendResult(BaseModel):
    """Outcome of a structured or LLM-sourced append attempt."""

    ok: bool = Field(description="Whether the append operation succeeded.")
    workbook_id: str = Field(description="Registered workbook identifier.")
    sheet: str = Field(description="Worksheet that was targeted.")
    written_row: int | None = Field(
        description="1-based row index written to, or None for dry runs."
    )
    row_preview: dict[str, object] = Field(
        description="Preview of the row values that were or would be written."
    )
    backup_path: str | None = Field(
        description="Path to the backup used for recovery, if one was created."
    )
    undo_token: str | None = Field(
        description="Token that can restore the prior workbook state."
    )
    dry_run: bool = Field(description="Whether the operation was preview-only.")
    source: Literal["structured", "llm"] = Field(
        default="structured",
        description="Origin of the append request.",
    )
    generated_fields: list[str] = Field(
        default_factory=list,
        description="Fields generated or inferred by the LLM path.",
    )
    message: str = Field(description="Human-readable operation summary.")


class UndoResult(BaseModel):
    """Outcome of restoring a workbook from a backup."""

    ok: bool = Field(description="Whether the undo operation succeeded.")
    workbook_id: str = Field(description="Registered workbook identifier.")
    restored_from: str = Field(description="Backup path that was restored.")
    message: str = Field(description="Human-readable operation summary.")


class FieldOrigin(str, Enum):
    """Origin of an extracted field value."""

    EXTRACTED = "extracted"
    GENERATED = "generated"
    DEFAULT = "default"


class ExtractedField(BaseModel):
    """Single extracted field proposed by the LLM pipeline."""

    column: str = Field(description="Target column name.")
    value: ValueType = Field(description="Extracted, generated, or defaulted value.")
    origin: FieldOrigin = Field(description="Where the field value came from.")
    note: str | None = Field(
        default=None,
        description="Optional explanation attached to the extracted value.",
    )


class ExtractionResult(BaseModel):
    """Validated extraction proposal derived from natural-language input."""

    workbook_id: str = Field(description="Registered workbook identifier.")
    sheet: str = Field(description="Worksheet targeted by the extraction.")
    fields: list[ExtractedField] = Field(
        description="Ordered extracted fields to map into an append request."
    )
    intent: str = Field(description="Short description of the interpreted user intent.")
    warnings: list[str] = Field(
        default_factory=list,
        description="Non-fatal issues or caveats discovered during extraction.",
    )
    model: str = Field(description="Model name used to produce the extraction.")
    needs_confirmation: bool = Field(
        default=True,
        description="Whether the extraction must be confirmed before writing.",
    )

    def to_append_request(self, *, dry_run: bool) -> AppendRequest:
        """Convert the extracted fields into a structured append request."""

        return AppendRequest(
            workbook_id=self.workbook_id,
            sheet=self.sheet,
            values={field.column: field.value for field in self.fields},
            dry_run=dry_run,
        )


class TextAppendRequest(BaseModel):
    """Natural-language request that will later be converted into a row proposal."""

    workbook_id: str = Field(description="Registered workbook identifier.")
    text: str = Field(description="Natural-language user input to parse.")
    sheet: str | None = Field(
        default=None,
        description="Optional worksheet override; defaults to the workbook primary sheet.",
    )
    confirm: bool = Field(
        default=False,
        description="When false, preview only; when true, allow a write after validation.",
    )


__all__ = [
    "AppendRequest",
    "AppendResult",
    "ColumnDef",
    "ColumnType",
    "ExtractedField",
    "ExtractionResult",
    "FieldOrigin",
    "Registry",
    "SchemaProposal",
    "TextAppendRequest",
    "UndoResult",
    "ValueType",
    "WorkbookEntry",
    "WorkbookSchema",
]
