# Excel Memory Agent — Final Architecture

**Status:** Authoritative implementation specification. This is the single source of truth for development.
**Scope:** A safe Excel automation layer for AI agents — model-agnostic Excel infrastructure for automating *existing business workbooks* (expense trackers, bookkeeping, accounting, inventory, sales reports, tax workbooks). Delivered as a Python package (`ema/`), a JSON registry, a paranoid Excel writer, an Excel MCP server (thin FastMCP), and an optional local-LLM natural-language front-end (Ollama). EMA lets any agent — local or hosted — make safe updates to workbooks that contain formulas, tables, reports, and business logic.

---

## 1. Principles

1. **Safe Excel writing is the core.** `excel_io.py` is the largest, most-tested, most-careful module. Every write backs up first, writes atomically, and verifies.
2. **All logic lives in the library and is unit-testable without an LLM or MCP.** `EmaService` is the public API and the primary test surface.
3. **The LLM and the MCP host are two interchangeable front-ends.** Both produce the same validated `AppendRequest` and feed the same core. There is never more than one reasoning engine in a single path.
4. **The LLM proposes; the service decides; the writer executes.** LLM-sourced rows are previewed and require confirmation before any byte is written. The LLM never touches files.
5. **Local-first and opt-in.** The core installs and runs with no model and no cloud. The LLM is an optional extra, disabled by default.
6. **Simplicity over abstraction.** One file per concern. One provider Protocol with one real implementation. No plugin frameworks, no per-workbook prompt files.

**Scope guarantee (state in README):** EMA safely updates the *tabular regions* of business workbooks (plain ranges and Excel Tables). EMA does **not** guarantee loss-free round-tripping of pivots, charts, or VBA/macros. Backups and undo remain the safety mechanism: EMA backs up before every write, so nothing is unrecoverable.

---

## 2. Architecture

```
                 (NL text)                       (structured args)
        ┌──────────────────────┐          ┌──────────────────────────┐
        │  Local LLM front-end │          │   MCP host front-end      │
        │  (CLI / standalone)  │          │  (Claude/Cursor → tools)  │
        └──────────┬───────────┘          └────────────┬─────────────┘
                   │   both emit a validated AppendRequest             │
                   └───────────────┬───────────────────────────────────┘
                                   ▼
                          ┌─────────────────┐
                          │   EmaService    │   orchestration + safety policy
                          └────────┬────────┘
                 ┌─────────────────┼─────────────────┐
                 ▼                 ▼                 ▼
            registry.py        schema.py         excel_io.py
            (JSON state)   (truth + contract)   (backup→atomic→verify→undo)
```

### Run modes

| Mode | Extractor | Local LLM | Commit policy |
|------|-----------|-----------|---------------|
| **MCP (host)** — primary | Host LLM | No | May commit directly (still backs up + verifies). |
| **Standalone CLI** | Local Ollama | Yes | Preview by default; commit only on explicit confirm. |
| **Library** | Caller supplies structured `AppendRequest` | No | Caller's responsibility. |

---

## 3. Folder structure

```
excel-memory-agent/
├── ema/
│   ├── __init__.py          # version, public exports
│   ├── config.py            # paths & constants; LLM settings
│   ├── errors.py            # typed exception hierarchy
│   ├── models.py            # all Pydantic models (core + extraction DTOs)
│   ├── registry.py          # RegistryStore: load/save registry.json, CRUD
│   ├── schema.py            # detect headers/tables/types; extraction contract
│   ├── excel_io.py          # paranoid writer: backup, atomic save, append, verify, undo
│   ├── service.py           # EmaService: orchestration, safety gate (the API)
│   ├── llm.py               # LLMProvider Protocol, OllamaProvider, FakeProvider
│   ├── extract.py           # Extractor: prompt → provider → validate → proposal
│   ├── cli.py               # standalone natural-language entry point
│   └── server.py            # FastMCP tools — thin wrappers over EmaService
├── tests/
│   ├── conftest.py          # tmp registry, generated .xlsx, fake_provider fixtures
│   ├── fixtures/
│   │   └── make_fixtures.py # plain/table/formula/offset/formatted .xlsx generators
│   ├── test_models.py
│   ├── test_registry.py
│   ├── test_schema.py
│   ├── test_excel_io.py
│   ├── test_service.py
│   ├── test_extract.py
│   ├── test_cli.py
│   └── test_server.py
├── examples/
│   ├── Expenses.xlsx
│   ├── Inventory.xlsx
│   └── Bookkeeping.xlsx
├── .ema/                    # runtime state (kept OUTSIDE any cloud-synced folder)
├── pyproject.toml
├── README.md
└── architecture_final.md
```

Each concern is one file. `service.py` is the seam between the pure library and the front-ends.

**Engineering constraint:** `extract.py` must stay smaller than `excel_io.py`. The LLM code never imports `excel_io` — the call graph forbids the LLM from writing files.

---

## 4. Dependencies (`pyproject.toml`)

```toml
[project]
name = "excel-memory-agent"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
    "openpyxl>=3.1",
    "pydantic>=2.6",
    "mcp[cli]>=1.2",        # FastMCP lives in the official MCP SDK
]

[project.optional-dependencies]
llm = ["ollama>=0.3"]       # or httpx; optional, off by default
dev = ["pytest>=8.0", "pytest-cov"]

[project.scripts]
ema-server = "ema.server:main"
ema = "ema.cli:main"
```

No pandas. No PyYAML. The core installs without the `llm` extra.

---

## 5. Config (`ema/config.py`)

Constants only. Everything overridable by env var (tests point `EMA_HOME` at a `tmp_path`).

```python
from pathlib import Path
import os

EMA_HOME = Path(os.getenv("EMA_HOME", Path.home() / ".ema"))
REGISTRY_PATH = EMA_HOME / "registry.json"
BACKUP_DIR = EMA_HOME / "backups"
LOG_PATH = EMA_HOME / "ema.log"
LOCK_SUFFIX = "~$"                       # Excel's lock-file prefix marker
MAX_BACKUPS_PER_WORKBOOK = 20
SCHEMA_VERSION = 1                       # bump to invalidate persisted schemas

# LLM (opt-in)
LLM_ENABLED = os.getenv("EMA_LLM_ENABLED", "false").lower() == "true"
LLM_MODEL = os.getenv("EMA_LLM_MODEL", "qwen3:4b")
OLLAMA_HOST = os.getenv("EMA_OLLAMA_HOST", "http://localhost:11434")
REQUIRE_CONFIRM_FOR_LLM = True           # LLM-sourced rows always preview first
EXPOSE_TEXT_TOOL = False                 # gate the MCP append_from_text tool
```

**Logging:** stdlib `logging` to `LOG_PATH`. Never `print()` in the server — MCP stdio transport uses stdout for protocol traffic.

**Runtime state location:** keep `EMA_HOME` outside any cloud-synced folder (e.g. OneDrive) to avoid sync locks and conflict copies.

---

## 6. Error hierarchy (`ema/errors.py`)

Flat, typed hierarchy. The library raises these; front-ends translate to clean messages. Unexpected exceptions are not swallowed.

```python
class EmaError(Exception):
    """Base for all EMA errors."""

class WorkbookNotFoundError(EmaError):      # not in registry
class WorkbookFileMissingError(EmaError):   # registered but file deleted/moved
class WorkbookLockedError(EmaError):        # Excel/cloud sync has it open
class SheetNotFoundError(EmaError): ...
class SchemaMismatchError(EmaError):        # live file drifted from persisted schema
class ValidationError(EmaError):            # row values don't match column defs
class WriteVerificationError(EmaError):     # row not found after save (rollback)
class UndoError(EmaError):                  # nothing to undo / backup missing
class RegistryCorruptError(EmaError):       # registry.json unreadable

class LLMError(EmaError):                   # base for LLM failures
class LLMUnavailableError(LLMError):        # Ollama down / model missing
class ExtractionError(LLMError):            # no valid JSON after retry
```

---

## 7. Pydantic models (`ema/models.py`)

All data crossing module boundaries is a model. Pydantic v2.

```python
from __future__ import annotations
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Literal
from pydantic import BaseModel, Field, ConfigDict


# ---- Core ----

class ColumnType(str, Enum):
    TEXT = "text"
    NUMBER = "number"
    DATE = "date"
    BOOL = "bool"


class ColumnDef(BaseModel):
    name: str                       # header text, exact
    index: int                      # 0-based column offset within the table/range
    type: ColumnType = ColumnType.TEXT
    required: bool = False
    is_formula: bool = False        # column holds formulas -> writer refuses to write here
    description: str | None = None  # optional hint used by the extraction contract


class WorkbookSchema(BaseModel):
    schema_version: int             # == config.SCHEMA_VERSION when written
    sheet: str
    header_row: int                 # 1-based Excel row containing headers
    table_name: str | None = None   # ListObject name if the data is an Excel Table
    columns: list[ColumnDef]
    data_start_row: int             # 1-based row where data begins


class WorkbookEntry(BaseModel):
    id: str                         # slug, e.g. "expenses"
    name: str
    path: Path                      # absolute path to .xlsx
    primary_sheet: str
    schema: WorkbookSchema
    created_at: datetime
    updated_at: datetime
    last_undo_token: str | None = None


class Registry(BaseModel):
    version: int = 1
    workbooks: dict[str, WorkbookEntry] = Field(default_factory=dict)


# ---- Request / response DTOs ----

class SchemaProposal(BaseModel):
    sheet: str
    header_row: int
    table_name: str | None
    columns: list[ColumnDef]
    warnings: list[str] = Field(default_factory=list)


class AppendRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    workbook_id: str
    sheet: str | None = None                            # default: primary_sheet
    values: dict[str, str | float | int | bool | None]  # keyed by column NAME
    dry_run: bool = False


class AppendResult(BaseModel):
    ok: bool
    workbook_id: str
    sheet: str
    written_row: int | None
    row_preview: dict[str, object]
    backup_path: str | None
    undo_token: str | None
    dry_run: bool
    source: Literal["structured", "llm"] = "structured"
    generated_fields: list[str] = Field(default_factory=list)
    message: str


class UndoResult(BaseModel):
    ok: bool
    workbook_id: str
    restored_from: str
    message: str


# ---- Extraction (LLM) ----

class FieldOrigin(str, Enum):
    EXTRACTED = "extracted"   # value came from the user's text
    GENERATED = "generated"   # value invented/inferred by the model
    DEFAULT   = "default"     # value filled from a schema default (e.g. today's date)


class ExtractedField(BaseModel):
    column: str
    value: str | float | int | bool | None
    origin: FieldOrigin
    note: str | None = None


class ExtractionResult(BaseModel):
    workbook_id: str
    sheet: str
    fields: list[ExtractedField]
    intent: str
    warnings: list[str] = Field(default_factory=list)
    model: str
    needs_confirmation: bool = True

    def to_append_request(self, *, dry_run: bool) -> AppendRequest:
        return AppendRequest(
            workbook_id=self.workbook_id,
            sheet=self.sheet,
            values={f.column: f.value for f in self.fields},
            dry_run=dry_run,
        )


class TextAppendRequest(BaseModel):
    workbook_id: str
    text: str
    sheet: str | None = None
    confirm: bool = False     # False => preview only
```

**Design notes:**
- `values` keyed by **column name**, not position; the schema maps names to indices.
- `extra="forbid"` on `AppendRequest` fails loudly on key typos rather than silently dropping data.
- `is_formula` columns are first-class so the writer refuses them.
- `FieldOrigin` is the mechanism that prevents silent fabrication: generated values are visible in every preview and audited in `AppendResult.generated_fields`.

---

## 8. Registry (`ema/registry.py`)

Owns `registry.json`. Pure persistence + CRUD. No Excel knowledge. JSON serialized via `model_dump_json(indent=2)` (human-readable, diffable). `save` is atomic (temp + `os.replace`).

```python
class RegistryStore:
    def __init__(self, path: Path = config.REGISTRY_PATH): ...

    def load(self) -> Registry: ...        # missing -> empty Registry; unreadable -> RegistryCorruptError
    def save(self, reg: Registry) -> None: ...

    def add(self, entry: WorkbookEntry) -> None: ...
    def get(self, workbook_id: str) -> WorkbookEntry: ...   # raises WorkbookNotFoundError
    def list(self) -> list[WorkbookEntry]: ...
    def remove(self, workbook_id: str) -> None: ...
    def upsert(self, entry: WorkbookEntry) -> None: ...
```

---

## 9. Schema (`ema/schema.py`)

Single source of truth for columns/types/positions. Suggests at registration; the service persists the confirmed result; re-validates against the live file before every write.

```python
def propose_schema(path: Path, sheet: str, header_row: int | None = None) -> SchemaProposal:
    """
    1. Open workbook read-only.
    2. If the sheet has exactly one Excel Table (ListObject) -> use its range & header row.
    3. Else use header_row (default 1) and read that row as headers.
    4. For each column: infer type from the first N non-empty data cells.
    5. Flag columns whose data cells contain formulas -> is_formula=True + warning.
    6. Emit warnings: empty/duplicate headers, merged cells in range,
       multiple tables found, no data rows.
    """

def validate_live_schema(path: Path, schema: WorkbookSchema) -> None:
    """Confirm headers/positions still match the persisted schema.
       Raise SchemaMismatchError on drift. Called before every write."""

def as_extraction_contract(schema: WorkbookSchema) -> dict:
    """Produce the JSON shape + per-column type/required/description used by BOTH
       the extraction prompt and the post-extraction validator. One contract, no drift."""
```

Private helpers: `_read_header_row`, `_infer_column_type`, `_detect_tables`, `_cell_is_formula`.

**Type inference (intentionally simple):** sample up to 10 data cells; all int/float → `NUMBER`; all `datetime` → `DATE`; all bool → `BOOL`; else `TEXT`. Ambiguous → `TEXT`. Inference is a validation hint, never a hard gate that blocks a legitimate write.

---

## 10. Excel writer (`ema/excel_io.py`)

The core. Every public function treats the file as precious. Knows nothing about the LLM.

```python
def check_writable(path: Path) -> None:
    """Raise WorkbookFileMissingError or WorkbookLockedError.
       Lock detection: presence of Excel's '~$<name>' file, OR attempt to open
       for append and catch PermissionError."""

def backup(path: Path) -> Path:
    """Copy to BACKUP_DIR/<id-or-name>/<name>-<UTCstamp>.xlsx.
       Prune to MAX_BACKUPS_PER_WORKBOOK. Return backup path."""

def append_row(path: Path, schema: WorkbookSchema, row: dict[str, object],
               dry_run: bool = False) -> AppendResult:
    """
    1. check_writable(path)
    2. validate_live_schema(path, schema)
    3. map row{name->value} -> ordered cells using schema.columns
         - refuse if any target column is_formula -> ValidationError
         - coerce/validate types against ColumnDef -> ValidationError on mismatch
    4. compute target row:
         - Table -> extend tableXX.ref by one row, write within range
         - plain -> first empty row after last populated data row
    5. if dry_run: return preview (written_row=None); NO file touch
    6. backup(path) -> backup_path
    7. atomic write: load_workbook -> write cells -> wb.save(tmp)
                     -> _verify_opens(tmp) -> os.replace(tmp, path)
    8. verify_write(...) -> on mismatch restore backup, raise WriteVerificationError
    9. return AppendResult(undo_token=backup_token, ...)
    """

def verify_write(path: Path, schema: WorkbookSchema, expected: dict, row_idx: int) -> None:
    """Reopen, read row_idx, confirm values match expected. Raise on mismatch."""

def undo_last(path: Path, undo_token: str) -> UndoResult:
    """Resolve undo_token -> backup file; check_writable; copy backup over path
       (atomically). Raise UndoError if token unknown/backup missing."""
```

**Guarantees baked into the code:**
- Always `backup()` before save.
- `check_writable()` first; never retry into corruption.
- Refuse writes to `is_formula` columns.
- Table-aware insertion extends the `ListObject` range.
- Atomic save: temp file + `_verify_opens` + `os.replace`.
- Post-write `verify_write` with rollback to backup on failure.
- `keep_vba=False`; `.xlsm` out of scope.

**Undo token:** the backup filename (timestamped path) carried in `AppendResult`. The backup *is* the undo state; no separate undo store.

---

## 11. LLM provider (`ema/llm.py`)

One Protocol, one real implementation, one test fake. The only place that talks to a model. Optional dependency.

```python
from typing import Protocol

class LLMProvider(Protocol):
    def complete_json(self, system: str, user: str, json_schema: dict) -> dict:
        """Return a parsed JSON object conforming to json_schema, or raise LLMError."""

class OllamaProvider:
    def __init__(self, model: str = config.LLM_MODEL, host: str = config.OLLAMA_HOST): ...
    def complete_json(self, system, user, json_schema) -> dict:
        # POST /api/chat with format=json; parse; raise LLMUnavailableError if Ollama
        # is unreachable / model missing; raise LLMError on unparseable output.

class FakeProvider:   # tests only — deterministic
    def __init__(self, canned: dict): self._canned = canned
    def complete_json(self, *_): return self._canned
```

A future OpenAI/Anthropic provider is another small class implementing the same Protocol, added only when actually needed. No factory, no `providers/` package, no config-driven loader.

---

## 12. Extractor (`ema/extract.py`)

Converts natural language into a validated candidate row. Never writes files.

```python
class Extractor:
    def __init__(self, provider: LLMProvider): ...

    def extract(self, workbook: WorkbookEntry, text: str) -> ExtractionResult:
        """
        1. contract = schema.as_extraction_contract(workbook.schema)
        2. build a schema-driven system prompt (generic; no per-workbook prompts)
        3. data = provider.complete_json(system, user=text, json_schema=contract)
        4. validate values against ColumnDef types via the schema layer
        5. tag each field origin: EXTRACTED / GENERATED / DEFAULT; add notes
        6. on invalid/unparseable JSON: ONE retry echoing the error,
           else raise ExtractionError
        7. return ExtractionResult(needs_confirmation=True, model=...)
        """
```

One generic, schema-driven prompt builder. Workbook-specific guidance lives in `ColumnDef.description` (data), not in bespoke prompt files.

---

## 13. Service (`ema/service.py`)

The public API and primary test surface. Owns workbook resolution, default-sheet substitution, bookkeeping, and the LLM safety gate.

```python
class EmaService:
    def __init__(self, store: RegistryStore | None = None,
                 provider: LLMProvider | None = None): ...

    def register_workbook(self, path: str, name: str | None = None,
                          sheet: str | None = None, header_row: int | None = None,
                          confirm: bool = True) -> WorkbookEntry | SchemaProposal:
        """confirm=False -> return SchemaProposal (no persistence).
           confirm=True  -> persist WorkbookEntry built from the proposal."""

    def list_workbooks(self) -> list[WorkbookEntry]: ...

    def get_workbook_schema(self, workbook_id: str) -> WorkbookSchema: ...

    def append_row(self, req: AppendRequest) -> AppendResult:
        """Resolve workbook+sheet, delegate to excel_io.append_row,
           bump updated_at, persist last_undo_token."""

    def append_from_text(self, req: TextAppendRequest) -> AppendResult:
        """1. resolve workbook; 2. Extractor.extract(...) -> ExtractionResult;
           3. SAFETY GATE: if REQUIRE_CONFIRM_FOR_LLM and not req.confirm -> dry_run=True;
           4. build AppendRequest(source='llm'), carry generated_fields into AppendResult;
           5. delegate to append_row. Requires LLM_ENABLED and a provider."""

    def undo_last_append(self, workbook_id: str) -> UndoResult: ...
```

Safety policy lives here: LLM-sourced rows preview by default; structured rows follow the caller's `dry_run`.

---

## 14. MCP server (`ema/server.py`)

Thin. Each tool: validate input → call `EmaService` → return `model_dump(mode="json")`. Catch `EmaError`, return a structured error. Log to file, never stdout.

```python
from mcp.server.fastmcp import FastMCP
from ema.service import EmaService
from ema.errors import EmaError
from ema import config

mcp = FastMCP("excel-memory-agent")
svc = EmaService()

@mcp.tool()
def register_workbook(path: str, name: str | None = None, sheet: str | None = None,
                      header_row: int | None = None, confirm: bool = False) -> dict:
    """Register an .xlsx workbook. confirm=False returns a schema proposal to review;
       confirm=True persists it."""
    try:
        return svc.register_workbook(path, name, sheet, header_row, confirm).model_dump(mode="json")
    except EmaError as e:
        return {"ok": False, "error": type(e).__name__, "message": str(e)}

@mcp.tool()
def list_workbooks() -> dict: ...

@mcp.tool()
def get_workbook_schema(workbook_id: str) -> dict: ...

@mcp.tool()
def append_row(workbook_id: str, values: dict, sheet: str | None = None,
               dry_run: bool = False) -> dict:
    """Append one row of column-name->value pairs. Backs up first, writes atomically,
       verifies, returns an undo_token. Use dry_run=True to preview."""
    ...

@mcp.tool()
def undo_last_append(workbook_id: str) -> dict: ...

# Registered only when config.LLM_ENABLED and config.EXPOSE_TEXT_TOOL.
def append_from_text(workbook_id: str, text: str, confirm: bool = False) -> dict:
    """EMA-side extraction. Returns an ExtractionResult preview when confirm=False."""
    ...

def main() -> None:
    mcp.run()   # stdio transport
```

### Tool list

| Tool | Status | Purpose |
|------|--------|---------|
| `register_workbook` | Always | Register + propose/confirm schema (merged via `confirm`). |
| `list_workbooks` | Always | Enumerate registered workbooks. |
| `get_workbook_schema` | Always | Return persisted schema. |
| `append_row` | Always | Safe append (backup/atomic/verify/undo, `dry_run`). |
| `undo_last_append` | Always | Roll back the last write. |
| `append_from_text` | Gated (`LLM_ENABLED` + `EXPOSE_TEXT_TOOL`), off by default | EMA-side extraction for hosts that request it. |

Natural language's primary home is the CLI. Surfacing `append_from_text` to a capable host is an explicit opt-in.

---

## 15. CLI (`ema/cli.py`)

The home of the local-LLM path.

```
ema add "<text>" --workbook <id> [--sheet <name>] [--yes]
ema register <path> [--sheet <name>] [--header-row N] [--yes]
ema list
ema undo --workbook <id>
```

`ema add` default: extract → print the proposed row with each field's origin (generated fields marked) → ask to confirm. `--yes` commits directly. `main()` is the `ema` console entry point.

---

## 16. Extraction flow (concrete)

```
ema add "Save grocery expense 18.50 EUR at Aldi" --workbook expenses
  → EmaService.append_from_text(TextAppendRequest(confirm=False))
      1. entry = registry.get("expenses"); schema = entry.schema
      2. contract = schema.as_extraction_contract()   # {Date:date, Store:text, Amount:number, Category:text}
      3. result = Extractor.extract(entry, text)
            Amount=18.50 (extracted), Store="Aldi" (extracted),
            Date=2026-06-11 (generated: "today"), Category="Groceries" (generated)
            invalid JSON -> ONE retry echoing the error, else ExtractionError
      4. safety gate: confirm=False -> dry_run=True
      5. append_row(dry_run=True) -> preview, NO file write
  → CLI prints the proposed row, marks generated fields, asks: commit? [y/N]
  → on confirm: append_from_text(confirm=True) -> dry_run=False
                -> backup -> atomic write -> verify -> undo token
```

Every generated value is visible and labeled before any byte is written.

---

## 17. Test suite

Tests run with no real LLM and no MCP. They hit `EmaService` and the modules directly; the LLM path uses `FakeProvider`. Fixtures generate real `.xlsx` files.

### `tests/fixtures/make_fixtures.py`
- `make_plain(path)` — headers in row 1, a few data rows.
- `make_table(path)` — data inside a named Excel Table.
- `make_formula(path)` — one column is `=B2*C2`.
- `make_offset_headers(path)` — banner rows, headers in row 3.
- `make_formatted(path)` — frozen panes + a chart.

### `tests/conftest.py`
- `ema_home(tmp_path, monkeypatch)` — point `EMA_HOME` at `tmp_path`.
- `plain_wb`, `table_wb`, `formula_wb`, `offset_wb` — built into `tmp_path`.
- `service` — `EmaService` bound to the temp registry.
- `fake_provider` — `FakeProvider` returning canned extraction JSON.

### Cases per file

**`test_models.py`**
- `AppendRequest` rejects unknown keys.
- Round-trip serialize/deserialize for `WorkbookEntry`, `Registry`, `ExtractionResult`.
- `ExtractionResult.to_append_request` maps fields to values.

**`test_registry.py`**
- Missing file → empty `Registry`; round-trip save/load; atomic save (no partial file on simulated failure); unknown id → `WorkbookNotFoundError`; corrupt JSON → `RegistryCorruptError`.

**`test_schema.py`**
- `propose_schema` on plain/table/offset workbooks (correct headers, table name, header row).
- Formula column → `is_formula=True` + warning; duplicate/empty headers → warnings.
- `validate_live_schema` passes unchanged, raises `SchemaMismatchError` after a header rename.
- `as_extraction_contract` includes every column with type/required/description.

**`test_excel_io.py`** (most important)
- `check_writable` raises `WorkbookLockedError` on a `~$` lock file.
- `backup` creates a file and prunes beyond `MAX_BACKUPS_PER_WORKBOOK`.
- `append_row(plain)` writes the first empty row; correct values/index.
- `append_row(table)` extends the table; new row is inside the range.
- Formula column → `ValidationError`, file untouched.
- Type mismatch → `ValidationError`, file untouched.
- `dry_run=True` → preview, file byte-for-byte unchanged (mtime + hash).
- Simulated save failure leaves the original intact.
- `verify_write` mismatch → restores backup, raises `WriteVerificationError`.
- `undo_last` restores pre-append bytes.
- Formatted workbook: append succeeds, file still opens.

**`test_service.py`**
- `register_workbook(confirm=False/True)` behavior.
- `append_row` happy path; unknown `workbook_id`; default-sheet substitution; `updated_at` advances.
- `append_from_text` with `fake_provider`: `confirm=False` ⇒ `dry_run` preview, no write; `confirm=True` ⇒ writes; `generated_fields` populated; `source="llm"`.
- `undo_last_append` restores state.

**`test_extract.py`**
- Valid canned JSON → correct `ExtractedField` origins.
- Generated vs extracted flagging.
- Invalid JSON then valid on retry → succeeds; invalid twice → `ExtractionError`.
- Type mismatch from the model → rejected by validation.

**`test_cli.py`**
- `ema add` without `--yes` previews and does not write.
- `ema add --yes` commits.
- Generated fields are shown in the preview output.

**`test_server.py`**
- Each tool returns a JSON-serializable dict.
- `EmaError` → `{"ok": False, "error": ..., "message": ...}` (not raised).
- `append_from_text` is absent unless `LLM_ENABLED` + `EXPOSE_TEXT_TOOL`.

**Coverage target:** ≥90% on `excel_io.py` and `service.py`. Real-model checks live in a skipped `@pytest.mark.llm` smoke test. Heavy formatted-workbook test marked `@pytest.mark.slow`.

---

## 18. Build order (single developer)

**Core first. The LLM layer is built only after the core is green.**

1. `models.py` + `errors.py` + `config.py` — contracts. (~half day)
2. `registry.py` + `test_registry.py` — atomic JSON CRUD. (~half day)
3. `tests/fixtures/make_fixtures.py` — real files before the writer. (~half day)
4. `schema.py` + `test_schema.py` — `propose_schema`, `validate_live_schema`, `as_extraction_contract`. (~1 day)
5. `excel_io.py` + `test_excel_io.py` — the hard part. Sub-order: `check_writable` → `backup` → plain `append_row` → `verify_write` → atomic save → table support → `undo_last`. (~3–4 days)
6. `service.py` + `test_service.py` — structured path glue. (~half day)
7. `server.py` + `test_server.py` — thin wrappers; smoke test in an MCP host. (~half day)
8. `llm.py` — Protocol + `OllamaProvider` + `FakeProvider`. (~half day)
9. `extract.py` + `test_extract.py` — prompt, validate, origin flagging, one-retry (deterministic via `FakeProvider`). (~1–1.5 days)
10. `service.append_from_text` + safety-gate tests. (~half day)
11. `cli.py` + `test_cli.py` — preview/commit UX. (~half day)
12. Optional gated MCP `append_from_text`. (~half day)
13. Manual `@pytest.mark.llm` smoke test on real Qwen3; tune the system prompt once. (~half day)
14. `README.md` — quickstart, scope guarantee, MCP-vs-CLI guide, `ollama pull qwen3:4b`.

---

## 19. Definition of Done

- The structured MCP path works end-to-end on a real `Expenses.xlsx` (plain *and* Excel-Table variants), with backup + undo demonstrably recovering from a bad write.
- With Ollama running, `ema add "Record office supplies expense 42.90 EUR at Staples on 2026-06-10" --workbook expenses` prints a labeled preview and, on confirm, appends one verified row.
- Every test green; ≥90% coverage on `excel_io.py` and `service.py`; the core and the structured path remain entirely LLM-free.

---

## 20. Locked defaults

| Decision | Default |
|----------|---------|
| `EMA_HOME` | `~/.ema`, env-overridable; kept outside cloud-synced folders. |
| Undo depth | Last write only (backup-based). |
| `keep_vba` | `False`; `.xlsm` out of scope. |
| Date input | ISO `YYYY-MM-DD` → coerced to `datetime` for DATE columns. |
| Concurrency | Single-process; rely on `check_writable`. No locking daemon. |
| Logging | stdlib `logging` to `EMA_HOME/ema.log`; never stdout. |
| LLM | Off by default (`LLM_ENABLED=False`); optional `[llm]` extra; LLM-sourced rows always preview first. |
| Providers | One Protocol, `OllamaProvider` only; add others when needed. |
