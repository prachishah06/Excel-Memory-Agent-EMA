# Excel Memory Agent — Phase 1 MVP Design (Implementation-Ready)

**Scope:** The "Safe Append" core from the architecture review. One Python package (`ema/`), one JSON registry, a thin FastMCP server, and a real test suite.
**Audience:** A single developer who wants to open this file and start typing code top-to-bottom.
**Design law:** All logic lives in the library and is unit-testable *without* an LLM or MCP. The server is a dumb adapter.

---

## 1. What Phase 1 does (and does not)

**Does:**
- Register a workbook + sheet, confirm its schema, persist it to `registry.json`.
- List registered workbooks and return a schema.
- Append a row of *structured* values to a registered sheet — safely:
  - back up first → write atomically → respect Excel Tables → verify the write → return an **undo token**.
- Undo the most recent append.
- Preview (`dry_run`) what *would* be written.

**Does NOT (deferred):** intent parsing, nutrition math, multi-workbook routing, food-specific tools, SQLite, history beyond last-write undo, YAML config, pandas.

**Hard scope guarantee in the README:** EMA targets *simple tabular* and *Excel-Table* workbooks. Workbooks with charts/pivots/macros may lose those features on write; EMA backs up before every write so nothing is unrecoverable.

---

## 2. Folder structure

```
excel-memory-agent/
├── ema/
│   ├── __init__.py          # version, public exports
│   ├── config.py            # paths & constants (registry location, backup dir)
│   ├── errors.py            # typed exception hierarchy
│   ├── models.py            # all Pydantic models
│   ├── registry.py          # Registry: load/save registry.json, CRUD on workbooks
│   ├── schema.py            # detect headers/tables/types -> SchemaProposal
│   ├── excel_io.py          # THE paranoid writer: backup, atomic save, append, verify, undo
│   ├── service.py           # EmaService: orchestrates registry+schema+excel_io (the API)
│   └── server.py            # FastMCP tools — thin wrappers over EmaService
├── tests/
│   ├── conftest.py          # fixtures: tmp registry, generated .xlsx workbooks
│   ├── fixtures/
│   │   └── make_fixtures.py # generates plain/table/formula/formatted .xlsx on demand
│   ├── test_models.py
│   ├── test_registry.py
│   ├── test_schema.py
│   ├── test_excel_io.py
│   ├── test_service.py
│   └── test_server.py
├── examples/
│   └── FoodLog.xlsx
├── .ema/                    # runtime state dir (gitignored): registry.json + backups/
├── pyproject.toml
├── README.md
└── plan_review.md
```

**Why modules not packages:** each concern is one file. `service.py` is the seam between "pure library" and "MCP server" — test against `EmaService`, not against MCP.

---

## 3. Dependencies (`pyproject.toml`)

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
dev = ["pytest>=8.0", "pytest-cov"]

[project.scripts]
ema-server = "ema.server:main"
```

No pandas. No PyYAML. Python 3.11+ for `tomllib`/typing niceties and `datetime` UTC helpers.

---

## 4. Config (`ema/config.py`)

Constants only — no file format, no parser. Everything overridable by env var for tests.

```python
from pathlib import Path
import os

EMA_HOME = Path(os.getenv("EMA_HOME", Path.home() / ".ema"))
REGISTRY_PATH = EMA_HOME / "registry.json"
BACKUP_DIR = EMA_HOME / "backups"
LOCK_SUFFIX = "~$"          # Excel's lock-file prefix marker
MAX_BACKUPS_PER_WORKBOOK = 20
SCHEMA_VERSION = 1          # bump to invalidate persisted schemas
```

`EMA_HOME` env override is what lets tests point at a `tmp_path`.

---

## 5. Error hierarchy (`ema/errors.py`)

A flat, typed hierarchy so the MCP layer can translate to clean messages and tests can assert precisely.

```python
class EmaError(Exception):
    """Base for all EMA errors."""

class WorkbookNotFoundError(EmaError):      # not in registry
class WorkbookFileMissingError(EmaError):   # registered but file deleted/moved
class WorkbookLockedError(EmaError):        # Excel/OneDrive has it open
class SheetNotFoundError(EmaError):
class SchemaMismatchError(EmaError):        # live file drifted from persisted schema
class ValidationError(EmaError):            # row values don't match column defs
class WriteVerificationError(EmaError):     # row not found after save (rollback)
class UndoError(EmaError):                  # nothing to undo / backup missing
class RegistryCorruptError(EmaError):       # registry.json unreadable
```

Rule: the library raises these; `server.py` catches `EmaError` and returns a structured error result. Unexpected exceptions are *not* swallowed.

---

## 6. Pydantic models (`ema/models.py`)

All data crossing module boundaries is a model. Pydantic v2.

```python
from __future__ import annotations
from datetime import datetime
from enum import Enum
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict


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
    is_formula: bool = False        # column holds formulas -> EMA refuses to write here


class WorkbookSchema(BaseModel):
    schema_version: int             # == config.SCHEMA_VERSION when written
    sheet: str
    header_row: int                 # 1-based Excel row containing headers
    table_name: str | None = None   # ListObject name if the data is an Excel Table
    columns: list[ColumnDef]
    data_start_row: int             # 1-based row where data begins (header_row + 1)


class WorkbookEntry(BaseModel):
    id: str                         # slug, e.g. "foodlog"
    name: str                       # display name
    path: Path                      # absolute path to .xlsx
    primary_sheet: str
    schema: WorkbookSchema
    created_at: datetime
    updated_at: datetime


class Registry(BaseModel):
    version: int = 1
    workbooks: dict[str, WorkbookEntry] = Field(default_factory=dict)


# ---- Request / response DTOs (the service & MCP boundary) ----

class SchemaProposal(BaseModel):
    """What schema discovery suggests; user/LLM confirms before persistence."""
    sheet: str
    header_row: int
    table_name: str | None
    columns: list[ColumnDef]
    warnings: list[str] = Field(default_factory=list)   # e.g. "formula column detected"


class AppendRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    workbook_id: str
    sheet: str | None = None        # default: primary_sheet
    values: dict[str, str | float | int | bool | None]  # keyed by column NAME
    dry_run: bool = False


class AppendResult(BaseModel):
    ok: bool
    workbook_id: str
    sheet: str
    written_row: int | None         # 1-based row written (None if dry_run)
    row_preview: dict[str, object]  # final values mapped to columns
    backup_path: str | None
    undo_token: str | None          # opaque id to pass to undo_last_append
    dry_run: bool
    message: str


class UndoResult(BaseModel):
    ok: bool
    workbook_id: str
    restored_from: str              # backup path used
    message: str
```

**Key decisions:**
- `values` is keyed by **column name**, not position — the LLM speaks names, EMA maps to indices via the schema.
- `extra="forbid"` on `AppendRequest` makes the LLM's tool calls fail loudly on typos instead of silently dropping data.
- `is_formula` columns are first-class so the writer can refuse them (review risk R3).
- Schema carries `schema_version` so a bumped constant invalidates stale caches.

---

## 7. Registry (`ema/registry.py`)

Owns `registry.json`. Pure persistence + CRUD. No Excel knowledge.

```python
class RegistryStore:
    def __init__(self, path: Path = config.REGISTRY_PATH): ...

    def load(self) -> Registry: ...
        # missing file -> empty Registry(); unreadable -> RegistryCorruptError

    def save(self, reg: Registry) -> None: ...
        # atomic: write tmp then os.replace (same discipline as excel_io)

    # convenience CRUD operating on a loaded Registry
    def add(self, entry: WorkbookEntry) -> None: ...
    def get(self, workbook_id: str) -> WorkbookEntry: ...   # raises WorkbookNotFoundError
    def list(self) -> list[WorkbookEntry]: ...
    def remove(self, workbook_id: str) -> None: ...
    def upsert(self, entry: WorkbookEntry) -> None: ...
```

`save` uses the same temp-then-replace trick as the Excel writer. JSON is serialized via Pydantic's `model_dump_json(indent=2)` so the file is human-readable and diffable.

---

## 8. Schema discovery (`ema/schema.py`)

"Suggest, don't decide." Returns a `SchemaProposal`; the user/LLM confirms; the service persists it.

```python
def propose_schema(path: Path, sheet: str, header_row: int | None = None) -> SchemaProposal:
    """
    1. Open workbook read-only.
    2. If the sheet has exactly one Excel Table (ListObject) -> use its range & header row.
    3. Else use header_row (default 1) and read that row as headers.
    4. For each column: infer type from the first N non-empty data cells.
    5. Flag columns whose data cells contain formulas -> is_formula=True + warning.
    6. Emit warnings: empty header cells, duplicate headers, merged cells in range,
       multiple tables found, no data rows.
    """

def validate_live_schema(path: Path, schema: WorkbookSchema) -> None:
    """Re-read the file and confirm headers/positions still match the persisted schema.
       Raise SchemaMismatchError on drift. Called before every write."""
```

Helpers (private): `_read_header_row`, `_infer_column_type`, `_detect_tables`, `_cell_is_formula`.

**Type inference rule (intentionally dumb):** sample up to 10 data cells; if all parse as int/float → `NUMBER`; all `datetime` → `DATE`; all bool → `BOOL`; else `TEXT`. Ambiguous → `TEXT`. Inference is a *hint* for validation, never a hard gate that blocks a legitimate write.

---

## 9. The paranoid writer (`ema/excel_io.py`)

The heart of the project. Every public function assumes the file is precious.

```python
def check_writable(path: Path) -> None:
    """Raise WorkbookFileMissingError or WorkbookLockedError.
       Lock detection: presence of Excel's '~$<name>' file, OR attempt to open
       the file for append and catch PermissionError (Windows lock)."""

def backup(path: Path) -> Path:
    """Copy to config.BACKUP_DIR/<id-or-name>/<name>-<UTCstamp>.xlsx.
       Prune to MAX_BACKUPS_PER_WORKBOOK. Return backup path."""

def append_row(path: Path, schema: WorkbookSchema, row: dict[str, object],
               dry_run: bool = False) -> AppendResult:
    """
    Orchestrated, safe append:
      1. check_writable(path)
      2. validate_live_schema(path, schema)         # from schema.py
      3. map row{name->value} -> ordered cells using schema.columns
         - refuse if any target column is_formula -> ValidationError
         - coerce/validate types against ColumnDef -> ValidationError on mismatch
      4. compute target row:
         - Table  -> extend tableXX.ref by one row, write within range
         - plain  -> first empty row after last populated data row
      5. if dry_run: return AppendResult(written_row=None, dry_run=True, ...)  # NO file touch
      6. backup(path) -> backup_path
      7. atomic write:
            wb = load_workbook(path)         # keep_vba? no for MVP; keep formatting default
            ...write cells...
            wb.save(tmp); _verify_opens(tmp); os.replace(tmp, path)
      8. verify_write(path, schema, expected_row) -> WriteVerificationError -> restore backup
      9. return AppendResult(undo_token=backup_token, ...)
    """

def verify_write(path: Path, schema: WorkbookSchema, expected: dict, row_idx: int) -> None:
    """Reopen, read row_idx, confirm values match expected. Raise on mismatch."""

def undo_last(path: Path, undo_token: str) -> UndoResult:
    """Resolve undo_token -> backup file; check_writable; copy backup over path
       (itself atomic). Raise UndoError if token unknown/backup missing."""
```

**Append safety checklist baked into the code (maps to review risks):**
- R1 (corruption): always `backup()` before save; constrained scope documented.
- R2 (lock): `check_writable()` first; never retry into corruption.
- R3 (formulas): refuse writes to `is_formula` columns.
- R4 (tables): table-aware insertion path extends the `ListObject` range.
- R8 (partial write): temp file + `_verify_opens` + `os.replace` = atomic.
- Post-write `verify_write` with rollback to backup on failure.

**Undo token:** for MVP it's just the backup filename (an opaque, unguessable-enough timestamped path stored in `AppendResult`). No separate undo DB needed — the backup *is* the undo state.

---

## 10. Service / orchestration (`ema/service.py`)

The public API. MCP tools and any future CLI both call this. **This is the unit-test surface.**

```python
class EmaService:
    def __init__(self, store: RegistryStore | None = None):
        self.store = store or RegistryStore()

    def register_workbook(self, path: str, name: str | None = None,
                          sheet: str | None = None, header_row: int | None = None,
                          confirm: bool = True) -> WorkbookEntry | SchemaProposal:
        """If confirm=False -> return SchemaProposal for review (no persistence).
           If confirm=True  -> persist WorkbookEntry built from the proposal."""

    def list_workbooks(self) -> list[WorkbookEntry]: ...

    def get_workbook_schema(self, workbook_id: str) -> WorkbookSchema: ...

    def append_row(self, req: AppendRequest) -> AppendResult:
        """Resolve workbook+sheet, delegate to excel_io.append_row, bump updated_at,
           persist undo_token on the entry's last_undo slot (in registry)."""

    def undo_last_append(self, workbook_id: str) -> UndoResult: ...
```

The service is where `workbook_id` → `WorkbookEntry` resolution, default-sheet substitution, and `updated_at` bookkeeping live. `excel_io` stays file-focused; `registry` stays JSON-focused; `service` glues them.

---

## 11. MCP server (`ema/server.py`)

Thin. Each tool: build/validate a model → call `EmaService` → return `model_dump()`. Catch `EmaError`, return structured error. ~100 lines.

```python
from mcp.server.fastmcp import FastMCP
from ema.service import EmaService
from ema.models import AppendRequest
from ema.errors import EmaError

mcp = FastMCP("excel-memory-agent")
svc = EmaService()

@mcp.tool()
def register_workbook(path: str, name: str | None = None, sheet: str | None = None,
                      header_row: int | None = None, confirm: bool = False) -> dict:
    """Register an .xlsx workbook. confirm=False returns a schema proposal to review;
       confirm=True persists it. Always returns the (proposed/saved) schema."""
    try:
        result = svc.register_workbook(path, name, sheet, header_row, confirm)
        return result.model_dump(mode="json")
    except EmaError as e:
        return {"ok": False, "error": type(e).__name__, "message": str(e)}

@mcp.tool()
def list_workbooks() -> dict: ...

@mcp.tool()
def get_workbook_schema(workbook_id: str) -> dict: ...

@mcp.tool()
def append_row(workbook_id: str, values: dict, sheet: str | None = None,
               dry_run: bool = False) -> dict:
    """Append one row of column-name->value pairs to a registered workbook.
       Backs up first, writes atomically, verifies, returns an undo_token.
       Use dry_run=True to preview without writing."""
    ...

@mcp.tool()
def undo_last_append(workbook_id: str) -> dict: ...

def main() -> None:
    mcp.run()   # stdio transport
```

### MVP tool list (final)

| Tool | Purpose | Notes vs. original plan |
|------|---------|-------------------------|
| `register_workbook` | Register + propose/confirm schema | merges "register" + "get schema" via `confirm` flag |
| `list_workbooks` | Enumerate registered workbooks | kept |
| `get_workbook_schema` | Return persisted schema | kept |
| `append_row` | Safe append (backup/atomic/verify/undo, `dry_run`) | **generic only** |
| `undo_last_append` | Roll back last write | **added** (cheap, high-demo-value) |

**Cut from the original tool list:** `append_food_entry()` (premature specialization — Phase 2) and a standalone `save_workbook()` (saving is an atomic part of `append_row`, never a separate user step).

---

## 12. Test suite

Tests run with **no LLM and no MCP** — they hit `EmaService` and the modules directly. Fixtures generate real `.xlsx` files so we *prove* behavior on actual workbooks.

### `tests/fixtures/make_fixtures.py`
Functions that build workbooks with openpyxl and return paths:
- `make_plain(path)` — headers in row 1, a few data rows.
- `make_table(path)` — data inside a named Excel Table (`ListObject`).
- `make_formula(path)` — one column is `=B2*C2` (the "do not write here" case).
- `make_offset_headers(path)` — banner rows, headers in row 3.
- `make_formatted(path)` — frozen panes + a chart, to assert lossy-but-non-crashing behavior.

### `tests/conftest.py`
- `ema_home(tmp_path, monkeypatch)` — set `EMA_HOME` env to `tmp_path`, recreate config paths.
- `plain_wb`, `table_wb`, `formula_wb`, `offset_wb` — call the fixture builders into `tmp_path`.
- `service` — an `EmaService` bound to the temp registry.

### Test files and the cases each must cover

**`test_models.py`**
- `AppendRequest` rejects unknown keys (`extra="forbid"`).
- Round-trip serialize/deserialize for `WorkbookEntry`, `Registry`.
- `ColumnType` inference enum coercion.

**`test_registry.py`**
- Load missing file → empty `Registry`.
- `save` then `load` round-trips.
- `save` is atomic (no partial file if interrupted — simulate via monkeypatched failure).
- `get` unknown id → `WorkbookNotFoundError`.
- Corrupt JSON → `RegistryCorruptError`.

**`test_schema.py`**
- `propose_schema(plain_wb)` → correct headers, `header_row=1`, types inferred.
- `propose_schema(table_wb)` → detects `table_name`, correct range.
- `propose_schema(offset_wb, header_row=3)` → respects explicit header row.
- Formula column → `is_formula=True` + warning.
- Duplicate/empty headers → warnings populated.
- `validate_live_schema` passes on unchanged file, raises `SchemaMismatchError` after a header is renamed.

**`test_excel_io.py`** (the most important file)
- `check_writable` raises `WorkbookLockedError` when a `~$` lock file exists.
- `backup` creates a file and prunes beyond `MAX_BACKUPS_PER_WORKBOOK`.
- `append_row(plain)` writes to the first empty row; values & row index correct.
- `append_row(table)` extends the table range and the new row is *inside* the table.
- `append_row` into a formula column → `ValidationError`, file untouched.
- Type mismatch (text into NUMBER) → `ValidationError`, file untouched.
- `dry_run=True` → returns preview, **file byte-for-byte unchanged** (assert mtime+hash).
- Atomicity: simulated save failure leaves original intact (temp discarded).
- `verify_write` mismatch → restores backup, raises `WriteVerificationError`.
- `undo_last` restores the pre-append bytes.
- Formatted workbook: append succeeds, file still opens (document lossy parts as known).

**`test_service.py`**
- `register_workbook(confirm=False)` → `SchemaProposal`, registry unchanged.
- `register_workbook(confirm=True)` → entry persisted, retrievable via `list/get`.
- `append_row` happy path end-to-end → `AppendResult.ok`, `updated_at` advanced.
- `append_row` with unknown `workbook_id` → `WorkbookNotFoundError`.
- Default sheet substitution when `sheet=None`.
- `undo_last_append` after an append restores state.

**`test_server.py`**
- Each tool returns a JSON-serializable dict.
- `EmaError` is converted to `{"ok": False, "error": ..., "message": ...}` (not raised).
- `append_row` tool with bad `values` key surfaces a clean validation error.

**Coverage target:** ≥90% on `excel_io.py` and `service.py` specifically (the rest is incidental). Add a `pytest` marker `@pytest.mark.slow` for the formatted-workbook chart test if it's heavy.

---

## 13. Build order for a single developer (1–2 week part-time)

1. **`models.py` + `errors.py` + `config.py`** — define the contracts first. (~half day)
2. **`registry.py` + `test_registry.py`** — atomic JSON CRUD, fully green. (~half day)
3. **`tests/fixtures/make_fixtures.py`** — you need real files before writing the writer. (~half day)
4. **`schema.py` + `test_schema.py`** — propose + validate. (~1 day)
5. **`excel_io.py` + `test_excel_io.py`** — the hard part; budget the most time here. Build in this sub-order: `check_writable` → `backup` → plain `append_row` → `verify_write` → atomic save → table support → `undo_last`. (~3–4 days)
6. **`service.py` + `test_service.py`** — glue, now trivial. (~half day)
7. **`server.py` + `test_server.py`** — thin wrappers; manual smoke test in Claude/Cursor. (~half day)
8. **`README.md`** — quickstart, the scope guarantee, and a screenshot/gif of the breakfast demo.

**Definition of Done for Phase 1:** the breakfast demo works end-to-end through an MCP host on a real `FoodLog.xlsx` (plain *and* table variants), every test green, ≥90% coverage on the two core modules, and a backup + undo demonstrably recover from a bad write.

---

## 14. Open decisions to lock before coding (pick defaults, don't deliberate)

| Decision | Recommended default |
|----------|--------------------|
| `EMA_HOME` location | `~/.ema` (env-overridable). **Avoid placing state inside the OneDrive repo** to dodge sync locks. |
| Undo depth | Last write only (backup-based). Multi-step undo is Phase 2. |
| `keep_vba` on load | `False` for MVP; `.xlsm` is out of scope. |
| Date input format | ISO `YYYY-MM-DD` strings → coerced to `datetime` for DATE columns. |
| Concurrency | Single-process assumption; rely on `check_writable`. No locking daemon. |
| Logging | stdlib `logging` to `EMA_HOME/ema.log`; never to stdout (MCP stdio uses it). |

That last row matters: **MCP stdio transport uses stdout for protocol traffic — never `print()` in the server.** Log to a file.
