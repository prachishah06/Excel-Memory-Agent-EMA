# Excel Memory Agent — MVP Design v2 (Local LLM Layer)

**Supersedes:** `mvp_design.md` (v1). Read this for the LLM-related changes; the v1 *core* (`registry`, `schema`, `excel_io`, `service`) is unchanged and remains authoritative for those modules.
**Subject:** Adding an optional local LLM (Ollama / Qwen3 4B–8B) as a first-class extraction component.
**Design law (still in force):** Safe Excel writing is the project. The LLM is a *convenience front-end*, not the center of gravity. If this document ever makes the LLM bigger than `excel_io.py`, the design has failed.

---

## 0. The one insight that drives everything

> **A local LLM only earns its place when EMA is *not already sitting behind a host LLM*.**

When EMA runs as an MCP server inside Claude Desktop / Cursor, the host *is* a frontier LLM that already does intent + extraction far better than Qwen3 4B ever will. Bolting Ollama into that path means **two LLMs reasoning about the same sentence** — slower, dumber, redundant, and confusing.

So the local LLM is **not a new layer in the middle of the pipeline.** It is a **second, parallel front-end** that produces exactly the same structured object the MCP host produces: a validated `AppendRequest`. Both front-ends drop into the *same unchanged core*.

```
                 (NL text)                       (structured args)
        ┌──────────────────────┐          ┌──────────────────────────┐
        │  Local LLM front-end │          │   MCP host LLM front-end │
        │  (CLI / standalone)  │          │  (Claude/Cursor → tools) │
        └──────────┬───────────┘          └────────────┬─────────────┘
                   │   both emit a validated AppendRequest             │
                   └───────────────┬───────────────────────────────────┘
                                   ▼
                          ┌─────────────────┐
                          │   EmaService    │   ← unchanged orchestration
                          └────────┬────────┘
                 ┌─────────────────┼─────────────────┐
                 ▼                 ▼                 ▼
            registry.py        schema.py         excel_io.py   ← THE priority, untouched
            (JSON state)   (truth + contract)   (backup→atomic→verify→undo)
```

This framing is what keeps the LLM thin. Everything below is a consequence of it.

---

## 1. Architecture review: is adding a local LLM a good idea?

**Verdict: Yes — conditionally. Approve it as an *optional standalone front-end*; reject it as a mandatory pipeline layer.**

### Why it's worth doing
- **It unlocks the actual product promise without the cloud.** "Type a sentence in a terminal, the row lands in Excel, fully offline" is a genuinely compelling, *local-first* portfolio demo. The MCP path requires a frontier host; this path requires nothing but Ollama.
- **It's a clean, bounded addition** *if* you accept the insight in §0. It touches two new files and one CLI. The core stays frozen.
- **Portfolio signal:** "I integrated a local model with structured output, schema-constrained extraction, and a hard safety gate" is a stronger story than "I wrote an MCP tool wrapper."

### Why it's dangerous if done naively
- A 4B–8B local model **will** mis-extract, mis-map, and **fabricate** values. If those flow silently into a *tax* or *expense* log, you've built a tool that quietly corrupts financial records. That's worse than no tool.
- It's the classic scope-eater. "Generate missing fields," "support any workbook," "OpenAI/Anthropic too" — each is a reasonable sentence that, taken literally, turns EMA into an LLM framework with an Excel plugin. We must fence it.

### The non-negotiable mitigations (these define the feature)
1. **LLM output is a *proposal*, never a commit.** Any row originating from the LLM defaults to `dry_run` and requires explicit confirmation before `excel_io` touches the file.
2. **Schema-constrained extraction.** The model is asked to fill *exactly* the registered columns, returning JSON validated by Pydantic. Unknown keys rejected; types coerced/validated by the existing schema layer.
3. **"Generated" ≠ "extracted."** Every field the model *invented* (vs. lifted from the user's text) is flagged. Invented values in NUMBER/DATE columns of financial workbooks are surfaced loudly in the preview.
4. **The LLM cannot reach `excel_io`.** Enforced by call graph: extractor → `AppendRequest` → service → (confirm gate) → writer. No shortcut.

---

## 2. Risks & added complexity

| # | Risk introduced by the local LLM | Severity | Mitigation |
|---|----------------------------------|----------|------------|
| L1 | **Fabricated values written to real files** (esp. tax/expense) | **Critical** | Mandatory preview + confirm for LLM-sourced rows; flag generated fields; never auto-commit. |
| L2 | **Unreliable structured output** from a small model (invalid JSON, extra prose) | High | Request JSON via Ollama `format=json` / structured outputs; parse → validate → **one retry** with the error echoed back → else fail to manual. |
| L3 | **Wrong column mapping** ("18.50" into Date) | High | Schema layer validates types after extraction; mismatches reject the proposal, not the file. |
| L4 | **Latency / model not installed / Ollama down** | Medium | Treat LLM as optional dependency; clear error if unreachable; the structured path still works without it. |
| L5 | **Prompt sprawl / per-workbook prompt hacking** | Medium | One generic schema-driven prompt builder. No per-workbook prompt files. If a workbook needs special handling, that's data in the registry, not a bespoke prompt. |
| L6 | **Provider over-abstraction** | Medium | A *single* `LLMProvider` Protocol with one real impl (Ollama). No registry, no plugin loader, no config DSL. |
| L7 | **Double-LLM reasoning** when used inside an MCP host | Medium | Don't expose a "give me text and I'll reason" tool to hosts by default; the host already reasons. (Optional escape hatch documented, discouraged.) |
| L8 | **Non-determinism breaks tests** | Medium | Tests mock the provider (a `FakeProvider` returning canned JSON). The *real* model is exercised only in a manual/`@pytest.mark.llm` smoke test. |

**Net complexity added:** ~2 modules (`llm.py`, `extract.py`), 1 CLI (`cli.py`), ~3 models, 1 optional MCP tool, and one new test file. The core gains **zero** new responsibilities. That is the whole point.

---

## 3. How the architecture changes

**It barely does — by design.** The change is *additive at the edge*:

- **New:** an extraction front-end (`llm.py` + `extract.py`) and a standalone `cli.py`.
- **Changed:** `EmaService` gains one method (`append_from_text`) and one policy (the confirm gate for LLM-sourced rows). `schema.py` gains one helper (`as_extraction_contract`).
- **Unchanged:** `registry.py`, `excel_io.py`, `models.py` (core), `errors.py`, the entire safe-write pipeline, and all v1 MCP tools.

### Run modes (explicit)
| Mode | Who extracts | LLM used? | Commit policy |
|------|--------------|-----------|---------------|
| **MCP (host)** — primary | Host LLM (Claude/Cursor) | No local LLM | Host already reasoned + user saw the call → may commit directly (still backs up). |
| **Standalone CLI** — new | Local Ollama | Yes | **Preview by default**, commit only on `--yes`/confirm. |
| **Library** — for tests/automation | Caller supplies structured `AppendRequest` | No | Caller's responsibility. |

---

## 4. Responsibility split (authoritative)

### Local LLM (`llm.py` + `extract.py`)
- Understand intent from free text.
- Extract values for the **registered columns only**.
- Normalize obvious forms (date phrases → ISO, "EUR"/"€" → currency value, capitalize meal).
- Map extracted data to **column names** from the schema contract.
- Mark each field as `extracted` vs. `generated`, with a confidence note.
- Output a candidate `dict[column_name -> value]` + metadata. **Nothing else.**
- **Forbidden:** opening files, deciding to commit, inventing columns, computing domain facts it can't justify from input (it may *propose* `Protein=34` but must flag it `generated` so the human sees it).

### EMA Service (`service.py`)
- Resolve workbook/sheet; fetch persisted schema.
- For NL input: call the extractor, build an `AppendRequest`, **enforce the confirm gate** (LLM-sourced ⇒ `dry_run` unless explicitly confirmed).
- For structured input: behave exactly as v1.
- Delegate the actual write to `excel_io`. Own the bookkeeping (`updated_at`, undo token).
- **Owns the safety policy.** The LLM proposes; the service decides; the writer executes.

### Schema Layer (`schema.py`)
- Remains the single source of truth for columns/types/positions.
- **New:** `as_extraction_contract(schema) -> dict` — produces the JSON shape + per-column type/required/description that the prompt and the validator both use. One contract, two consumers (prompt + validation). No drift.
- Validates extracted values against `ColumnDef` types; rejects mismatches.

### Excel Writer (`excel_io.py`)
- **Completely unchanged.** It does not know the LLM exists. It receives ordered, validated cell values and performs backup → atomic write → verify → undo.
- This module stays the largest, most-tested, most-careful file in the repo. If `extract.py` ever has more lines than `excel_io.py`, stop and re-read §0.

---

## 5. Integration pattern recommendation

**Use a one-method Protocol with a single Ollama implementation. Defer everything else.**

```python
# ema/llm.py
from typing import Protocol

class LLMProvider(Protocol):
    def complete_json(self, system: str, user: str, json_schema: dict) -> dict:
        """Return a parsed JSON object conforming to json_schema, or raise LLMError."""

class OllamaProvider:
    def __init__(self, model: str = "qwen3:4b", host: str = "http://localhost:11434"): ...
    def complete_json(self, system, user, json_schema) -> dict:
        # POST /api/chat with format=json (or structured outputs); parse; raise LLMError on failure.

class FakeProvider:  # tests only
    def __init__(self, canned: dict): ...
    def complete_json(self, *_): return self._canned
```

- **Direct Ollama calls?** Yes, *inside* `OllamaProvider`. Don't scatter HTTP calls elsewhere.
- **Provider abstraction?** Yes, but the *minimum* viable one: a Protocol + one impl + a test fake. **No factory, no registry, no `providers/` package, no config-driven loader.**
- **Future OpenAI/Anthropic?** Trivially a second ~30-line class implementing the same Protocol, added *when actually needed*. Writing it now is speculative generality — don't. Document the seam; leave the socket empty.

Dependency: prefer the official `ollama` python client or plain `httpx`. Make it an **optional extra** (`pip install excel-memory-agent[llm]`) so the core installs without it.

---

## 6. Is MCP still appropriate?

**Yes, and it's still the *primary* integration.** Nothing about adding a local LLM weakens the MCP story; the two are orthogonal front-ends.

- Keep all v1 tools (`register_workbook`, `list_workbooks`, `get_workbook_schema`, `append_row`, `undo_last_append`) exactly as designed. In MCP mode the **host LLM is the extractor**, so these stay structured.
- **Do not** add a `append_from_text` MCP tool by default — it would invite double-LLM reasoning (L7). The natural-language path's home is the **CLI**, where there is no host LLM.
- *Optional, documented-as-discouraged escape hatch:* a single `append_from_text` tool guarded behind a config flag, for hosts that are weak or for users who explicitly want EMA-side extraction. Off by default.

---

## 7. Updated folder structure

```
excel-memory-agent/
├── ema/
│   ├── __init__.py
│   ├── config.py          # + OLLAMA_HOST, LLM_MODEL, LLM_ENABLED, REQUIRE_CONFIRM_FOR_LLM
│   ├── errors.py          # + LLMError, LLMUnavailableError, ExtractionError
│   ├── models.py          # core models (v1) + extraction DTOs (below)
│   ├── registry.py        # unchanged
│   ├── schema.py          # + as_extraction_contract()
│   ├── excel_io.py        # UNCHANGED (the priority)
│   ├── service.py         # + append_from_text(), confirm gate
│   ├── llm.py             # NEW: LLMProvider Protocol, OllamaProvider, FakeProvider
│   ├── extract.py         # NEW: Extractor — prompt build → provider → validate → proposal
│   ├── cli.py            # NEW: standalone NL entry point (the local-LLM home)
│   └── server.py          # v1 tools unchanged; optional gated append_from_text
├── tests/
│   ├── conftest.py        # + fake_provider fixture
│   ├── fixtures/make_fixtures.py
│   ├── test_models.py
│   ├── test_registry.py
│   ├── test_schema.py     # + test for as_extraction_contract
│   ├── test_excel_io.py
│   ├── test_service.py    # + append_from_text confirm-gate tests (with FakeProvider)
│   ├── test_extract.py    # NEW: extraction validation, generated-field flagging
│   ├── test_cli.py        # NEW: CLI preview vs commit
│   └── test_server.py
├── examples/
├── pyproject.toml         # + [llm] optional extra: ollama/httpx
├── README.md
├── plan_review.md
├── mvp_design.md          # v1 (core reference)
└── mvp_design_v2.md       # this file
```

New code footprint: **3 modules + 2 test files.** Core file count of the safe-write path: unchanged.

---

## 8. Updated module responsibilities (deltas only)

- **`config.py`** — add `LLM_ENABLED=False` (opt-in), `LLM_MODEL="qwen3:4b"`, `OLLAMA_HOST`, `REQUIRE_CONFIRM_FOR_LLM=True`. All env-overridable.
- **`errors.py`** — add `LLMError` (base), `LLMUnavailableError` (Ollama down / model missing), `ExtractionError` (couldn't produce valid JSON after retry).
- **`schema.py`** — add `as_extraction_contract(schema) -> dict`. Single source for both the prompt's column spec and post-extraction validation.
- **`service.py`** — add `append_from_text(...)`; implement the confirm gate; reuse v1 `append_row` underneath.
- **`llm.py`** — provider Protocol + Ollama + Fake. The *only* place that talks HTTP to a model.
- **`extract.py`** — `Extractor.extract(workbook, text) -> ExtractionResult`. Builds the prompt from the schema contract, calls the provider, validates, flags generated fields, retries once.
- **`cli.py`** — `ema add "<text>" --workbook foodlog [--yes]`. Default: print the proposed row + flags, ask to confirm. `--yes` commits.

---

## 9. Updated Pydantic models (additions to v1)

```python
from enum import Enum
from pydantic import BaseModel, Field

class FieldOrigin(str, Enum):
    EXTRACTED = "extracted"   # value came from the user's text
    GENERATED = "generated"   # value invented/inferred by the model
    DEFAULT   = "default"     # value filled from a schema default (e.g. today's date)

class ExtractedField(BaseModel):
    column: str
    value: str | float | int | bool | None
    origin: FieldOrigin
    note: str | None = None        # short justification, e.g. "parsed '18.50 EUR'"

class ExtractionResult(BaseModel):
    workbook_id: str
    sheet: str
    fields: list[ExtractedField]
    intent: str                    # free-text label, e.g. "append_expense" (informational)
    warnings: list[str] = Field(default_factory=list)
    model: str                     # which model produced this (provenance)
    needs_confirmation: bool = True

    def to_append_request(self, *, dry_run: bool) -> "AppendRequest":
        return AppendRequest(
            workbook_id=self.workbook_id,
            sheet=self.sheet,
            values={f.column: f.value for f in self.fields},
            dry_run=dry_run,
        )

class TextAppendRequest(BaseModel):       # the NL entry DTO
    workbook_id: str
    text: str
    sheet: str | None = None
    confirm: bool = False                 # False => preview only (safety default)
```

Note: `AppendResult` (v1) gains two optional fields for provenance: `source: Literal["structured","llm"]` and `generated_fields: list[str]`. These make the preview and the audit trail honest.

**Why `FieldOrigin` matters more than anything else here:** it is the mechanism that prevents silent fabrication (L1). The preview can render "⚠ Protein=34 (generated)" and the confirm gate can *require* extra acknowledgement when generated values land in NUMBER columns of a financial workbook.

---

## 10. Updated MCP tools

**Unchanged from v1 (primary path):** `register_workbook`, `list_workbooks`, `get_workbook_schema`, `append_row`, `undo_last_append`.

**New, optional, OFF by default:**

| Tool | Status | Purpose |
|------|--------|---------|
| `append_from_text(workbook_id, text, confirm=False)` | **Gated by `LLM_ENABLED` + `EXPOSE_TEXT_TOOL`** | EMA-side extraction for hosts that want it. Returns an `ExtractionResult` preview when `confirm=False`. Discouraged inside frontier hosts (double-LLM). |

The CLI — not MCP — is the intended home for natural language. Surfacing the text tool to a capable host is an explicit opt-in, documented with the L7 warning.

---

## 11. The extraction flow (concrete)

```
cli: ema add "Save grocery expense 18.50 EUR at Aldi" --workbook expenses
  → EmaService.append_from_text(TextAppendRequest(confirm=False))
      1. entry = registry.get("expenses"); schema = entry.schema
      2. contract = schema.as_extraction_contract()      # {Date:date, Store:text, Amount:number, Category:text}
      3. result = Extractor.extract(entry, text)
            - provider.complete_json(system, user=text, json_schema=contract)
            - parse → validate types via schema → flag origins
            - Amount=18.50 (extracted), Store="Aldi" (extracted),
              Date=2026-06-11 (generated: "today"), Category="Groceries" (generated)
            - on invalid JSON: ONE retry echoing the parse error, else ExtractionError
      4. confirm gate: confirm=False ⇒ req = result.to_append_request(dry_run=True)
      5. AppendResult = excel_io.append_row(..., dry_run=True)   # NO file write
  → CLI prints the proposed row, marks generated fields, asks: commit? [y/N]
  → on y: append_from_text(confirm=True) ⇒ dry_run=False ⇒ backup→atomic→verify→undo
```

Every dangerous value (`Date`, `Category` here) is visible and labeled *before* a single byte is written. That is the entire safety argument made operational.

---

## 12. Updated implementation plan (single developer)

Build the LLM layer **only after the v1 core is green.** It depends on a working `service` + `schema` + `excel_io`.

1. **(v1 first)** Complete `mvp_design.md` Phase 1. Non-negotiable prerequisite. (per v1 plan)
2. **`schema.as_extraction_contract()` + test** — pure function, trivial. (~half day)
3. **`llm.py`** — Protocol + `OllamaProvider` (httpx/ollama, `format=json`) + `FakeProvider`. (~half day)
4. **`extract.py` + `test_extract.py`** — prompt builder, validate-against-contract, origin flagging, one-retry. Test entirely with `FakeProvider` (deterministic). (~1–1.5 days)
5. **`service.append_from_text` + confirm gate + tests** — wire extractor → request → writer; assert `dry_run` default and that generated-field acknowledgement is enforced. (~half day)
6. **`cli.py` + `test_cli.py`** — `ema add "<text>" --workbook <id> [--yes]`; preview/commit UX. (~half day)
7. **Optional gated MCP `append_from_text`** — only if you want it; ship behind the flag. (~half day)
8. **Manual LLM smoke test** (`@pytest.mark.llm`, skipped in CI) — run real Qwen3 4B on the three example sentences; eyeball extraction quality; tune the system prompt once. (~half day)
9. **README** — add the local-LLM quickstart (`ollama pull qwen3:4b`), the safety model, and the "MCP host vs CLI" decision guide.

**Total LLM-layer effort:** ~3–4 part-time days *on top of* the v1 core. If it's ballooning past that, you're overbuilding the wrong half of the project.

**Definition of Done (v2):** with Ollama running, `ema add "Save dinner: 2 eggs, 200g skyr, 20g roasted edamame" --workbook foodlog` prints a labeled preview, and on confirm safely appends one verified row (with backup + undo) — while every v1 test and the structured MCP path remain green and entirely LLM-free.

---

## 13. Anti-overengineering guardrails (pin these to the wall)

1. **`extract.py` must stay smaller than `excel_io.py`.** Use it as a literal line-count tripwire.
2. **One prompt builder, schema-driven.** No per-workbook prompt files, no prompt template library.
3. **One provider Protocol, one real impl.** No `providers/` package until a *second real* provider ships.
4. **No LLM in the test-critical path.** Determinism via `FakeProvider`; real model only in a skipped smoke test.
5. **LLM never writes.** The call graph forbids `extract`/`llm` from importing `excel_io`.
6. **Opt-in everywhere.** `LLM_ENABLED=False` by default; core installs and runs without Ollama or the `[llm]` extra.
7. **No silent commits from the model.** `REQUIRE_CONFIRM_FOR_LLM=True`, always preview generated fields.

---

## 14. Final recommendation

**Adopt the local LLM — as a thin, optional, standalone front-end — and not one inch more.**

The feature is genuinely valuable: it delivers the "natural language → Excel, fully offline" promise that makes EMA a memorable local-first portfolio piece, and it does so by *adding to the edge* (`llm.py`, `extract.py`, `cli.py`) while leaving the safe-write core frozen. The architecture stays honest because the LLM and the MCP host are recognized as **two interchangeable front-ends emitting the same validated `AppendRequest`** — never a stack of two reasoning engines.

The single thing that makes this safe rather than reckless is the **proposal-then-confirm gate plus `FieldOrigin` flagging**: the model may suggest, but a human (or a deliberately structured host call) authorizes, and every fabricated value is visible before any byte is written. For a tool that edits real expense and tax records, that gate is not a feature — it's the price of admission.

**If you cannot commit to keeping it thin** — if you feel the pull toward provider plugins, prompt frameworks, multi-step agent loops, or auto-committing model output — then **don't add it yet.** Ship the v1 structured-MCP core, which is already portfolio-worthy and provably safe, and revisit the LLM once the boring, important half (never corrupting a workbook) is rock solid.

**Recommendation: Approve v2 as specified. Build it second, keep it small, and let `excel_io.py` remain the biggest file in the repo.**
