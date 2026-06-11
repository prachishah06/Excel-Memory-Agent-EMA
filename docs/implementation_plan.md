# Excel Memory Agent — Implementation Plan

**Derived from:** `architecture_final.md` (authoritative). This document does not alter the architecture, modules, models, or tools. It defines **how** and **in what order** to build them.

**Estimated effort:** ~9–12 part-time days for a single developer.

**Two principles govern sequencing:**
1. **Core before convenience.** The safe-write core (`excel_io`, `service`) is fully built and green before any LLM code exists.
2. **No untested module advances.** Each milestone ends green; later milestones depend on earlier ones being trustworthy.

---

## 1. Phased overview

| Phase | Milestones | Outcome |
|-------|-----------|---------|
| **A — Foundations** | M0–M3 | Contracts, persistence, and real `.xlsx` fixtures in place. |
| **B — Safe-write core** | M4–M5 | Schema discovery + the paranoid writer, fully tested. |
| **C — Structured product** | M6–M7 | `EmaService` + MCP server; the structured path works end-to-end. **First shippable build.** |
| **D — Local-LLM front-end** | M8–M11 | Provider, extractor, NL service method, CLI. |
| **E — Hardening & release** | M12–M14 | Optional gated tool, real-model smoke test, docs. |

**Two release gates:**
- **Gate 1 (end of Phase C):** structured, MCP-driven, LLM-free product is demonstrably safe. This is a complete, portfolio-worthy deliverable on its own.
- **Gate 2 (end of Phase E):** local natural-language path works with preview/confirm safety.

---

## 2. Milestones

Each milestone: **Goal → Build → Tests → Exit checkpoint → Deliverable.**

### Phase A — Foundations

#### M0 — Project scaffold
- **Goal:** A runnable, installable skeleton.
- **Build:** `pyproject.toml` (core deps only), `ema/__init__.py`, empty module files, `.ema/` gitignored, `pytest` configured (markers: `slow`, `llm`), CI-less local `pytest` run.
- **Tests:** `pytest` collects 0 tests without error; `pip install -e .[dev]` succeeds.
- **Checkpoint:** `import ema` works; `ema-server`/`ema` entry points resolve (even if stubbed).
- **Deliverable:** Installable empty package.

#### M1 — Contracts (`config.py`, `errors.py`, `models.py`)
- **Goal:** Freeze the data contracts every other module depends on.
- **Build:** All constants, the full exception hierarchy, all Pydantic models (core + extraction DTOs) exactly as specified.
- **Tests:** `test_models.py` — `AppendRequest` rejects unknown keys; round-trip serialize/deserialize for `WorkbookEntry`, `Registry`, `ExtractionResult`; `ExtractionResult.to_append_request` maps fields to values.
- **Checkpoint:** All models import and validate; no logic depends on undefined fields.
- **Deliverable:** Stable contracts; downstream modules can be typed against them.

#### M2 — Registry (`registry.py`)
- **Goal:** Atomic, human-readable JSON persistence.
- **Build:** `RegistryStore` with atomic `save` (temp + `os.replace`), `load`, and CRUD.
- **Tests:** `test_registry.py` — missing file → empty `Registry`; round-trip save/load; atomic save leaves no partial file on simulated failure; unknown id → `WorkbookNotFoundError`; corrupt JSON → `RegistryCorruptError`.
- **Checkpoint:** Registry survives interrupted writes; `registry.json` is diffable.
- **Deliverable:** Working persistence layer.

#### M3 — Test fixtures (`tests/fixtures/make_fixtures.py`, `conftest.py`)
- **Goal:** Real workbooks to build and test the writer against. **This precedes the writer deliberately.**
- **Build:** `make_plain`, `make_table`, `make_formula`, `make_offset_headers`, `make_formatted`; `conftest.py` fixtures `ema_home`, `plain_wb`, `table_wb`, `formula_wb`, `offset_wb`, `service`.
- **Tests:** A smoke test confirming each generator produces a file that opens in openpyxl.
- **Checkpoint:** All five workbook variants generate into `tmp_path`.
- **Deliverable:** Deterministic fixture suite usable by all later tests.

---

### Phase B — Safe-write core

#### M4 — Schema discovery (`schema.py`)
- **Goal:** Suggest, validate, and contract-ize schemas.
- **Build:** `propose_schema`, `validate_live_schema`, `as_extraction_contract`, and private helpers.
- **Tests:** `test_schema.py` — proposals on plain/table/offset workbooks; formula column flagged + warning; duplicate/empty headers warn; `validate_live_schema` passes unchanged and raises `SchemaMismatchError` on header rename; `as_extraction_contract` lists every column with type/required/description.
- **Checkpoint:** Schema correctly read from all fixtures; drift detection works.
- **Deliverable:** Schema layer ready for both writing and extraction.

#### M5 — Paranoid writer (`excel_io.py`) — **the critical milestone**
- **Goal:** Never corrupt a file; always recoverable.
- **Build (sub-order, each step tested before the next):**
  1. `check_writable`
  2. `backup` (with pruning)
  3. plain `append_row`
  4. `verify_write`
  5. atomic save (temp + `_verify_opens` + `os.replace`)
  6. table-aware `append_row`
  7. `undo_last`
- **Tests:** `test_excel_io.py` — lock detection; backup + pruning; plain append; table append within range; formula-column refusal (file untouched); type mismatch refusal (file untouched); `dry_run` leaves file byte-for-byte identical (mtime + hash); simulated save failure preserves original; `verify_write` mismatch restores backup; `undo_last` restores bytes; formatted workbook still opens after append.
- **Checkpoint:** **≥90% coverage on `excel_io.py`.** Backup + undo demonstrably recover from a forced bad write.
- **Deliverable:** The trustworthy heart of the project.

---

### Phase C — Structured product (Gate 1)

#### M6 — Service (`service.py`, structured path)
- **Goal:** The public API for the structured (non-LLM) path.
- **Build:** `EmaService.register_workbook`, `list_workbooks`, `get_workbook_schema`, `append_row`, `undo_last_append`. (Leave `append_from_text` for M10.)
- **Tests:** `test_service.py` — `register_workbook(confirm=False/True)`; `append_row` happy path; unknown `workbook_id`; default-sheet substitution; `updated_at` advances; `undo_last_append` restores state.
- **Checkpoint:** **≥90% coverage on `service.py`** (structured surface). Full register → append → undo cycle green.
- **Deliverable:** Importable library that safely appends rows.

#### M7 — MCP server (`server.py`)
- **Goal:** Expose the structured path to MCP hosts.
- **Build:** Five always-on tools as thin wrappers; `EmaError` → structured error dict; logging to file (never stdout); `main()` stdio entry.
- **Tests:** `test_server.py` — each tool returns JSON-serializable dict; `EmaError` is converted, not raised; `append_from_text` absent (flags off).
- **Checkpoint:** Manual smoke test in an MCP host (Claude Desktop / Cursor): register a real `FoodLog.xlsx`, append via the host, undo.
- **Deliverable & GATE 1:** Shippable, safe, structured MCP product. Tag `v0.1.0`.

---

### Phase D — Local-LLM front-end (opt-in)

#### M8 — LLM provider (`llm.py`)
- **Goal:** A single, swappable model interface.
- **Build:** `LLMProvider` Protocol, `OllamaProvider` (`format=json`, `LLMUnavailableError` on unreachable/missing model), `FakeProvider`. Add `[llm]` optional extra.
- **Tests:** `FakeProvider` returns canned JSON; `OllamaProvider` unreachable host → `LLMUnavailableError` (mock the HTTP call).
- **Checkpoint:** Core still installs and tests pass **without** the `[llm]` extra.
- **Deliverable:** Provider abstraction with a deterministic test fake.

#### M9 — Extractor (`extract.py`)
- **Goal:** NL text → validated candidate row, with origin flags.
- **Build:** `Extractor.extract` — schema-driven prompt, `complete_json`, type validation via schema, `FieldOrigin` tagging, one retry on bad JSON.
- **Tests:** `test_extract.py` (all via `FakeProvider`, deterministic) — correct origins; generated-vs-extracted flagging; invalid-then-valid retry succeeds; invalid twice → `ExtractionError`; model type mismatch rejected.
- **Checkpoint:** `extract.py` line count < `excel_io.py`; no import of `excel_io`.
- **Deliverable:** Deterministic, testable extraction.

#### M10 — NL service method (`service.append_from_text`)
- **Goal:** Wire extraction into the service behind the safety gate.
- **Build:** `append_from_text` — resolve workbook, extract, enforce `REQUIRE_CONFIRM_FOR_LLM` (preview unless `confirm=True`), build `AppendRequest(source="llm")`, carry `generated_fields`.
- **Tests:** extend `test_service.py` (with `fake_provider`) — `confirm=False` ⇒ `dry_run` preview, no write; `confirm=True` ⇒ writes; `generated_fields` populated; `source="llm"`.
- **Checkpoint:** No LLM-sourced row reaches the file without confirmation.
- **Deliverable:** Safe NL append at the service layer.

#### M11 — CLI (`cli.py`)
- **Goal:** The natural-language home for users.
- **Build:** `ema add/register/list/undo`; `add` previews with origin labels, `--yes` commits; `main()` entry point.
- **Tests:** `test_cli.py` — `add` without `--yes` previews and does not write; `add --yes` commits; generated fields shown in preview.
- **Checkpoint:** Full CLI flow works against fixtures with `FakeProvider`.
- **Deliverable:** Usable offline NL → Excel CLI.

---

### Phase E — Hardening & release (Gate 2)

#### M12 — Optional gated MCP text tool
- **Goal:** Allow EMA-side extraction for hosts that request it.
- **Build:** Register `append_from_text` only when `LLM_ENABLED` and `EXPOSE_TEXT_TOOL`.
- **Tests:** tool present only when both flags on; absent otherwise.
- **Checkpoint:** Default build still omits the tool.
- **Deliverable:** Opt-in host extraction path.

#### M13 — Real-model smoke test
- **Goal:** Validate extraction quality on a real local model.
- **Build:** `@pytest.mark.llm` test (skipped in normal runs) running Qwen3 on the three example sentences; tune the single system prompt once.
- **Tests:** manual/marked; assert extraction shape, not exact values.
- **Checkpoint:** The three reference sentences extract sensibly; prompt finalized.
- **Deliverable:** Confidence the NL path works with `ollama pull qwen3:4b`.

#### M14 — Documentation & release
- **Goal:** A reviewer can install and demo in minutes.
- **Build:** `README.md` — quickstart, scope guarantee, MCP-vs-CLI guide, Ollama setup, breakfast/expense demos (gif/screenshot).
- **Checkpoint:** Fresh-clone install + both demos reproduce from the README alone.
- **Deliverable & GATE 2:** Tag `v0.2.0` — full safe, local, AI-powered Excel append.

---

## 3. Testing strategy

- **Pyramid:** heavy unit coverage on `excel_io` and `service`; thin wrapper tests on `server`/`cli`; one marked real-model smoke test.
- **No LLM, no MCP in the default suite.** Every test runs offline and deterministically; the LLM path uses `FakeProvider`.
- **Prove on real files.** All write tests run against generated `.xlsx` fixtures, never mocks of openpyxl.
- **Safety assertions are first-class:** `dry_run` byte-equality (mtime + hash), backup-on-write existence, rollback-on-verify-failure, and confirmation-gate enforcement are explicit test cases, not afterthoughts.
- **Markers:** `@pytest.mark.slow` (formatted-workbook/chart), `@pytest.mark.llm` (real Ollama) — both excluded from the default fast run.
- **Coverage gates:** ≥90% on `excel_io.py` and `service.py`; remainder incidental.
- **Determinism:** seed nothing that varies; mock time/UTC stamps where assertions depend on them.

---

## 4. Checkpoints (must pass to advance)

| After | Hard checkpoint |
|-------|-----------------|
| M3 | All five workbook fixtures generate and open. |
| M5 | ≥90% coverage on `excel_io.py`; forced bad write is fully recovered via backup + undo. |
| M6 | ≥90% coverage on `service.py`; register → append → undo cycle green. |
| M7 (Gate 1) | Manual MCP-host demo on a real workbook succeeds; structured product is LLM-free and safe. |
| M9 | `extract.py` < `excel_io.py` in size; no `excel_io` import in LLM code. |
| M10 | No LLM-sourced row written without confirmation. |
| M14 (Gate 2) | Fresh-clone install + both demos reproduce from README. |

---

## 5. Deliverables summary

| Deliverable | Milestone | Tag |
|-------------|-----------|-----|
| Installable package skeleton | M0 | — |
| Persistence + fixtures | M2–M3 | — |
| Tested safe-write core | M5 | — |
| Structured library + MCP server | M6–M7 | **v0.1.0 (Gate 1)** |
| Provider + extractor + NL service | M8–M10 | — |
| Offline NL CLI | M11 | — |
| Hardened, documented release | M12–M14 | **v0.2.0 (Gate 2)** |

---

## 6. Dependency-ordered task graph

```
M0 ─ M1 ─ M2 ─┐
              ├─ M3 ─ M4 ─ M5 ─ M6 ─ M7 ──(Gate 1)── M8 ─ M9 ─ M10 ─ M11 ─ M12 ─ M13 ─ M14 ──(Gate 2)
              │                                  │
   (fixtures depend on M1 models)        (LLM phase depends on a green core)
```

Strictly sequential by default for a single developer. The only safe parallelization: writing test cases for milestone *N+1* while implementing milestone *N*, since the contracts (M1) are frozen early.
