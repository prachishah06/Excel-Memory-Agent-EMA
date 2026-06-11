# Excel Memory Agent — Architecture Review

**Reviewer role:** Principal Software Architect
**Subject:** EMA Architecture & Development Plan (Draft v1)
**Context:** Single-developer, portfolio-quality agentic AI project, local-first.
**Bias of this review:** Ruthless simplicity. Ship a working slice fast. Treat every "Engine," "Layer," and abstraction as guilty until proven necessary.

---

## TL;DR (read this if nothing else)

The **vision is good and marketable**. The **plan is over-architected for a one-person project** and, more importantly, **under-specifies the two things that will actually make or break it**: (1) safely writing to real-world `.xlsx` files without corrupting them, and (2) the LLM-to-row mapping (nutrition lookup, intent parsing). You've drawn seven boxes ("registry," "schema discovery engine," "memory layer," "intent parser," "row construction engine," "excel writer," "MCP server") when the MVP needs roughly **three modules and a JSON file**.

Biggest single mistake: **the plan treats Excel writing as a solved, trivial step ("Insert row, Preserve formatting, Save")** while listing "complex Excel formatting" as a *risk* at the bottom. That risk *is the project*. Everything else is comparatively easy.

Second biggest: **the "Memory Layer" and "Intent Parser" are presented as components you build, but in an MCP architecture the LLM client IS the intent parser and much of the memory.** You're planning to rebuild capabilities the agent already provides.

---

## Strengths

1. **Clear, concrete vision.** "Natural Language → Excel Updated Automatically" is a crisp, demoable value prop. The breakfast example is a perfect end-to-end test case.
2. **Local-first is the right call.** No cloud, no auth, no multi-tenant complexity. This eliminates 80% of the hard distributed-systems problems and is appropriate for a portfolio project.
3. **Phasing exists.** Having phases at all puts you ahead of most solo projects. The *instinct* to defer LLM memory past MVP is correct (even if the boundaries are wrong — see below).
4. **Sensible-looking stack.** Python + openpyxl + Pydantic + FastMCP is a reasonable, modern, "I know current tools" signal for a portfolio. Pytest included from the start is a green flag.
5. **Explicit Non-Goals.** Listing macros, cloud sync, and multi-user as out of scope shows judgment. Keep doing this.
6. **MCP is genuinely on-trend.** As a portfolio piece in 2026, "I built an MCP server" is a strong, current talking point. (Whether it's the *right* architecture is more nuanced — see next section.)

---

## Weaknesses

### W1. Over-decomposition: seven "components" for a CRUD-to-Excel tool
You have separate top-level components for Registry, Schema Discovery, Memory, Intent Parser, Row Construction, Excel Writer, and MCP Server. Several of these are 20–50 line modules masquerading as subsystems. "Schema Discovery Engine" with five responsibilities is, for MVP, "read row 1 of the sheet as headers." Naming a function an "Engine" doesn't make it one — it just creates folders, `__init__.py` files, and import overhead you'll fight.

### W2. The Intent Parser is redundant in an MCP design
This is the central conceptual flaw. In MCP, **the LLM (Claude/ChatGPT/Cursor) is the brain.** It already parses "Breakfast: 2 eggs, 200g skyr" into structured arguments and calls your tool. If you also build a Python `intent_parser` module, you're either:
- (a) duplicating the LLM's job with brittle regex/rules, or
- (b) calling a *second* LLM from inside your server — which defeats the "LLM-independent" principle and adds cost/latency/keys.

**Pick a lane.** For an MCP server, the right design is: tools expose *structured* parameters, and the host LLM fills them. Your server should mostly *validate and execute*, not *interpret*.

### W3. "Memory Layer" conflates three different things
"Memory" in the plan means at least three unrelated concepts:
- **Workbook metadata** (path, sheets, columns) → this is just the Registry. It's config, not memory.
- **Schema cache** (inferred headers/types) → a derived cache, invalidate on file mtime change.
- **Behavioral memory** ("common commands: save breakfast") → this is arguably the *LLM client's* job (system prompt / conversation), not your server's.

Treating these as one "Memory Layer" will produce a confused module. They have different lifetimes, owners, and invalidation rules.

### W4. Excel writing is dangerously underspecified
"Open workbook → Find insertion row → Insert row → Preserve formatting → Save." Each clause hides a landmine (detailed in Risks). openpyxl **silently drops or mangles**: charts, images, some conditional formatting, pivot tables, and (critically) it does **not recalculate formulas** — cached values can go stale and the file can show `0`/`#VALUE` until reopened in Excel. "Preserve formatting" with openpyxl is aspirational, not a checkbox.

### W5. The hard domain problem is invisible in the plan
"2 eggs, 200g skyr" → `Protein: 34`. Where does **34** come from? That requires a nutrition database or an LLM estimate. The "Row Construction Engine" example quietly performs a nutrition lookup that is nowhere in the architecture, has no data source, and is genuinely hard (portion parsing, unit conversion, food matching). This is the actual product value and it's a footnote.

### W6. "LLM-independent" + "MCP-compatible" are in tension with the NL features
MCP-compatible means you depend on an LLM host. "Natural Language Agent" (Phase 3) is *entirely* LLM-driven. So the system is not LLM-independent; it's **LLM-host-agnostic**, which is a different and weaker claim. Be precise in the README or it reads as confused.

### W7. No concurrency / file-lock story, but it's listed as a risk
Excel holds an exclusive lock on open `.xlsx` files on Windows. If the user has FoodLog.xlsx open and your agent writes, you'll get `PermissionError` on save (best case) or a corrupt/half-written file (worst case). The plan lists "concurrent workbook access" as a risk but proposes no mitigation. On Windows specifically, this WILL happen constantly.

### W8. No data-loss / safety strategy
You are mutating the user's real files. There is no mention of: backups before write, atomic writes (write-temp-then-rename), validation-after-write, or rollback. For a tool whose whole job is "modify my important spreadsheets automatically," the absence of a safety net is the most serious omission after W4.

### W9. Pandas is in the stack with no clear job
Pandas + openpyxl together usually means trouble: `pandas.to_excel` **destroys** formatting and rewrites the whole sheet. If you're appending single rows to formatted workbooks, **pandas is the wrong tool and an active hazard.** It's listed because it "feels like data work," not because it's needed. Drop it from MVP.

### W10. YAML config is premature
A solo, single-machine tool does not need YAML config files on day one. One JSON registry file + sensible defaults is enough. YAML adds a dependency and a parsing surface for zero MVP benefit.

---

## Risks (ranked by likelihood × damage)

| # | Risk | Likelihood | Damage | Notes / Mitigation |
|---|------|-----------|--------|--------------------|
| R1 | **openpyxl corrupts/strips a formatted workbook** (charts, images, pivot tables, some conditional formats) | High | High | Always back up before write; warn users that round-tripping complex workbooks via openpyxl is lossy; test against a *real* formatted file early. |
| R2 | **File locked because Excel has it open** (Windows) | Very High | Medium | Detect lock pre-write; fail loud with a clear message; never silently retry into corruption. |
| R3 | **Stale cached formula values** (openpyxl writes data, formulas don't recalc) | High | Medium–High | Don't write into formula columns; document that user must reopen in Excel for recalc; or restrict to value-only columns. |
| R4 | **Wrong insertion point** with Tables (`ListObject`), filters, frozen panes, or total rows | High | Medium | Use openpyxl Table objects properly (extend `tableXX.ref`); naive "append to first empty row" breaks tables and total rows. |
| R5 | **Schema misdetection** (headers not in row 1, multi-row headers, metadata banner rows) | Medium–High | Medium | Don't assume row 1 = headers. Make header row explicit at registration time; auto-detect as a *suggestion*, not truth. |
| R6 | **Nutrition/value derivation is wrong or impossible** | High | Medium | Decide explicitly: store raw text only (MVP) vs. LLM estimate vs. real DB. Don't silently guess macros into a health log. |
| R7 | **Ambiguous routing** across multiple workbooks (Phase 4) | Medium | Low–Medium | This is the LLM's job in MCP; give tools good descriptions and let the host disambiguate. Don't build a router. |
| R8 | **Partial write / crash mid-save** corrupts file | Low | High | Atomic save: write to temp file, validate it opens, then `os.replace`. |
| R9 | **OneDrive sync conflict** — note this repo lives under `OneDrive\Documents`. OneDrive may lock/sync files mid-write and create `-conflict` copies | Medium | Medium | Warn against operating on OneDrive-synced workbooks, or pause sync; at minimum document it. |

---

## Recommended Architecture

### Is MCP the right architecture?

**Mostly yes — with one strong caveat.** For a portfolio project in 2026, exposing this as an MCP server is a great, current choice and lets you skip building a chat UI entirely (Claude Desktop / Cursor *is* your UI).

**Caveat:** MCP should be a **thin adapter over a clean, independently-testable core library.** Do **not** put logic in the MCP layer. Build `ema/` as a normal importable Python package with a plain function API (`append_row(workbook_id, sheet, values)`), unit-test *that*, and make the MCP server a ~100-line file that wraps those functions as tools. This gives you:
- Tests without an LLM in the loop.
- A CLI fallback for free.
- The ability to demo without MCP if a reviewer doesn't have it set up.

**Reframe the components.** The LLM host owns intent parsing, conversation memory, and routing. Your server owns: registry, schema, **safe writing**, and validation. Internalize this and three of your seven boxes shrink to near-zero.

### Persistence layer recommendation

**Use a single JSON file for MVP. Migrate to SQLite only when you have a concrete reason.**

| Option | Verdict for this project |
|--------|--------------------------|
| **JSON** | ✅ **Winner for MVP.** Human-readable, zero dependency, trivial to inspect/debug, perfect for "a handful of registered workbooks + cached schemas." `registry.json`. |
| **SQLite** | ⏳ **Phase 2+.** Adopt when you add entry history, undo, duplicate detection, or daily summaries — i.e., when you need *queries* over many rows. Then it's clearly correct. |
| **YAML** | ❌ Don't. All of JSON's downsides (no queries) plus a parser dependency and whitespace footguns. Reserve YAML for hand-edited config only, and you don't need that yet. |
| **Per-workbook sidecar** | 🤔 Optional nicety: a `.ema.json` next to each workbook so config travels with the file. Adds complexity; skip for MVP. |

**Rule of thumb:** JSON until you need to *query history*, then SQLite. Never YAML for state.

### Schema discovery — recommended approach

The plan's auto-everything approach is too clever and will misfire on real spreadsheets. Use **"auto-suggest, human-confirm"**:

1. At `register_workbook`, read the sheet. Propose: header row index, column names, inferred types.
2. **Persist the confirmed schema** in `registry.json`. Treat it as truth thereafter.
3. Re-validate against the live file on each write (column count/names still match?); if drift detected, fail loud rather than write garbage.
4. **Prefer Excel Tables (`ListObject`) when present** — they give you unambiguous boundaries and append semantics for free. Detect them first; fall back to "headers in row N" only when there's no table.
5. Type inference: keep it minimal (text / number / date). Don't over-infer. Store as a hint for validation, not a hard contract.

### Excel writing — recommended strategy

This is where you should spend most of your engineering care.

1. **Back up first.** Copy to `.ema-backups/{name}-{timestamp}.xlsx` before any mutation. This single step de-risks the whole project.
2. **Atomic save.** Write to a temp file, reopen it to verify it's valid, then `os.replace()` over the original. Never write in place.
3. **Pre-flight lock check.** Detect if the file is open/locked (try opening for append, or check for the `~$` lock file Excel creates). Fail with a clear message; never fight the lock.
4. **Respect Tables.** If the target is a `ListObject`, append by extending the table range, not by blindly writing to the next empty row.
5. **Never write into formula columns.** Detect formula cells in the column; refuse or skip. Document that Excel must reopen for recalculation.
6. **Scope honesty.** Document loudly: "EMA is designed for *simple, tabular, openpyxl-safe* workbooks. Workbooks with charts, pivots, or macros may lose those on write." Better to constrain scope than corrupt a user's file.
7. **Validate after write.** Reopen, confirm the row landed with expected values, return that in the confirmation. This turns "I hope it worked" into "verified."

### Proposed project structure

Flatten it. Folders-per-concept is overhead at this size. One package, modules not packages:

```
excel-memory-agent/
├── ema/
│   ├── __init__.py
│   ├── registry.py        # load/save registry.json, register/list workbooks
│   ├── schema.py          # detect & confirm headers, tables, types
│   ├── excel_io.py        # the careful part: backup, atomic write, table-aware append
│   ├── models.py          # Pydantic: Workbook, ColumnDef, AppendRequest
│   └── server.py          # FastMCP: thin tool wrappers over the above
├── tests/
│   ├── fixtures/          # real .xlsx samples: plain, table, formula, formatted
│   ├── test_registry.py
│   ├── test_schema.py
│   └── test_excel_io.py
├── examples/
│   └── FoodLog.xlsx
├── registry.json          # gitignored or sample-only
├── pyproject.toml
├── README.md
└── plan_review.md
```

Note: **no `intent_parser/`, no `memory/`, no `row_construction/`, no `config/` (YAML)**. The LLM host handles intent; memory is `registry.json`; row construction is a few lines in `excel_io.py`/`server.py`; config defaults live in code.

---

## MVP Scope Changes

### Remove from Phase 1 MVP
- **Intent Parser module** — the MCP host LLM does this. Expose structured tool params instead.
- **Memory Layer (as a distinct subsystem)** — it's just `registry.json`. Don't build behavioral memory yet.
- **Row Construction "Engine"** — collapse into a small function; for MVP, map provided fields to columns. No nutrition math.
- **Pandas** — actively harmful for formatted-workbook appends. Cut it.
- **YAML config** — use code defaults + JSON registry.
- **`append_food_entry()` as a separate tool** — premature specialization. One generic `append_row` is enough for MVP; food-specific tools are Phase 2+ sugar.
- **Auto-everything schema discovery** — downgrade to "suggest at registration, confirm, persist."
- **Multi-workbook routing** — Phase 4 already; just don't let it leak into MVP thinking.

### Add to Phase 1 MVP (these are not optional for a file-mutating tool)
- **Automatic backup before every write.** Single highest-ROI feature.
- **Atomic write (temp + replace).**
- **File-lock / "Excel has it open" detection** with a clear error.
- **Post-write verification** (reopen, confirm row).
- **Explicit schema confirmation at registration** (header row, columns, types).
- **A real test fixture set**: a plain sheet, a Table-based sheet, a formula sheet, and a "formatted" sheet — so you *prove* what survives and what doesn't.
- **A `dry_run` / preview**: return the row that *would* be written before committing. Great for agent UX and demos.
- **An undo for the last write** (cheap, since you already have backups). Hugely improves the demo and your confidence.

### Revised Phase 1 success criteria
> User says "log breakfast: 2 eggs, 200g skyr." Claude/Cursor calls `append_row(FoodLog, "Daily Intake", {Date, Meal, Food})`. EMA backs up the file, appends the row (respecting the table), saves atomically, verifies the row, and returns a confirmation **including the backup path and an undo token**. The original formatting is intact for a simple/table workbook.

That is a tight, honest, *impressive* MVP.

---

## Future Enhancements (re-sequenced roadmap)

The original phases are roughly right but mis-prioritized — they defer *safety* and front-load *intelligence*. Reorder so each phase is independently demoable and safe.

- **Phase 1 — Safe Append (MVP).** Registry (JSON), schema confirm, table-aware atomic append with backup + verify + undo. Generic `append_row`. *This is the whole portfolio-worthy core.*
- **Phase 2 — Ergonomics & Persistence.** Food/expense-specific tool wrappers, schema caching keyed on file mtime, friendlier confirmations, optional per-workbook sidecar config. Introduce **SQLite** here *if* you add entry history.
- **Phase 3 — Conversational polish.** Lean fully into the MCP host for NL; ship a great system prompt / tool descriptions so "save lunch" just works. Add the nutrition/value-derivation feature *deliberately* (pick: store raw text / LLM estimate / real food DB) — this is a marquee feature, treat it as a project of its own.
- **Phase 4 — Multi-workbook.** Mostly "good tool descriptions + let the LLM route." Add duplicate detection (needs SQLite). 
- **Phase 5 — Advanced.** Row editing, daily summaries (SQLite queries), OCR receipts, voice. All genuinely optional flair. Don't anchor the project's identity here.

**MVP → Production gate:** "Production" for a local tool means: never corrupts a file, always backs up, handles the locked-file case gracefully, has tests against real `.xlsx` fixtures, and clearly documents its scope limits (no macros/pivots/charts guarantee). Hit those and it's genuinely usable, not just a demo.

---

## Implementation complexity & major technical challenges

**Overall complexity: Moderate, but lopsided.** ~80% of your effort and *all* of your bugs will live in `excel_io.py`. The rest is easy.

| Area | Complexity | Why |
|------|-----------|-----|
| MCP server wiring (FastMCP) | **Low** | Well-trodden; a few decorators. |
| Registry (JSON) | **Low** | Read/write a dict. |
| Schema detection (suggest + confirm) | **Low–Medium** | Easy if you confirm with the user; medium if you insist on full auto-detect (don't). |
| **Safe Excel writing** | **High** | Tables, formulas, formatting preservation, atomic saves, lock handling. *The* hard part. |
| Backup / atomic / undo | **Low–Medium** | Conceptually simple, easy to get *subtly* wrong (paths, OneDrive). |
| Nutrition/value derivation (if pursued) | **High** | Portion parsing + food matching + units. A standalone hard problem; defer and scope explicitly. |
| Multi-workbook routing | **Low** | Offload to the LLM host. |

**Top three challenges to plan around:**
1. **openpyxl's lossy round-trip on rich workbooks.** Mitigate by *constraining scope* and *backing up*, not by trying to perfectly preserve everything.
2. **Windows file locks + OneDrive.** Your repo is literally in a OneDrive folder — expect sync/lock surprises. Test there.
3. **Resisting your own scope creep.** Seven engines, voice input, OCR, calendar — the plan already shows the instinct to build everything. The discipline to ship the safe-append core *first* is the real challenge.

---

## Final Verdict

**Strong vision, good portfolio instincts, wrong center of gravity.**

The plan invests architecture in the *easy, LLM-handled* parts (intent, memory, routing) and treats the *genuinely hard, project-defining* part (safe, formatting-respecting, non-corrupting Excel writes) as a checkbox and a footnote-risk. Flip that.

**Do this:**
1. **Collapse seven components into one library (`ema/`) + a thin MCP server.** Delete the intent parser, the memory layer, and pandas from MVP.
2. **Persist with one `registry.json`.** SQLite later, only when you need history queries. Never YAML for state.
3. **Make `excel_io.py` excellent and paranoid:** backup → atomic write → table-aware append → verify → undo. This single module is your portfolio's "wow."
4. **Schema = suggest then confirm,** persisted as truth.
5. **Be honest about scope:** "simple/tabular/table-based workbooks; rich workbooks may lose chart/pivot/macro features on write."

Do that and you get a smaller, sharper, *safer* project that is **more** impressive than the sprawling v1 — because it actually works on real files and demonstrates the judgment to know what *not* to build. For a single developer, that judgment is the most valuable thing you can show.

**Recommendation: Approve the vision. Reject the v1 decomposition. Rebuild around a paranoid Excel-writer core behind a thin MCP adapter.**
