# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Excel Structure Inspector (`excel_inspector`): read-only inspection of `.xlsx` workbooks that detects every table in a workbook (including multiple stacked tables per sheet), produces a pandas-consumable `ReadPlan` per table, and can load everything into DataFrames with JSON/Markdown serialization and a CLI. Requires Python >= 3.14; pinned to `openpyxl==3.1.5`, `pandas==3.0.3`.

## Commands

macOS Homebrew Python is externally managed (PEP 668) — always use the project venv, never the system interpreter:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

```sh
# Tests (default suite; the slow 100k-row perf smoke is deselected via addopts)
.venv/bin/python -m pytest
.venv/bin/python -m pytest tests/test_header_locator.py            # one file
.venv/bin/python -m pytest tests/test_results.py -k json           # one test by keyword
.venv/bin/python -m pytest -m slow                                 # perf smoke only (~14s)
.venv/bin/python -m pytest -m "slow or not slow"                   # everything

# CLI
.venv/bin/python -m excel_inspector book.xlsx                      # Markdown (default, --max-rows caps rows)
.venv/bin/python -m excel_inspector book.xlsx --format json        # JSON schema v1.0, always all rows
```

There is no linter/formatter configured.

## Language

- **English** — content where AI alone is the audience: thinking/reasoning, code, identifiers, comments, and docstrings.
- **Korean** — anything the user reads, reviews, or decides on: chat responses and explanations, questions/options presented for a choice, result summaries and reports, plan/review documents shown to the user, and git commit messages.

## Workflow

- All code work happens in a dedicated git worktree — never edit the primary checkout directly. Create (or reuse) a worktree before touching code, do the work there, and merge back when done.

## Authoritative docs

`docs/excel-structure-inspector-spec.md` (spec, Korean), plus the v1 and v2 implementation plans in `docs/`. Code comments cite these constantly as `spec §N` and decision IDs `[D1]`–`[D6]` (defined in the spec's §0 revision-history table). When changing behavior, check the cited section first — the spec is the contract, and tests pin it.

## Architecture

### Two-layer public API (`excel_inspector/__init__.py`)

- **v1**: `inspect(path, options) -> WorkbookProfile` — metadata only (header row, boundaries, column types, merges, per-table `ReadPlan`s). Pair with `load_dataframe(path, read_plan)` (`adapters/pandas_loader.py`), which translates a `ReadPlan` into `pandas.read_excel` kwargs.
- **v2**: `extract(path, options) -> WorkbookResult` (`results.py`) — runs `inspect` and loads every detected table (one `TableResult` per table, `table_id` like `"매출!T1"`), with `to_json()` (deterministic schema v1.0) / `to_markdown()`.

### Analyzer pipeline (`pipeline.py`)

`inspect()` wires a fixed pipeline of `Analyzer` stages over one shared mutable `InspectionContext` (`context.py`):

```
Loader -> SheetEnumerator -> MergeScanner -> BlockSegmenter -> HeaderLocator
       -> BoundaryDetector -> TypeProfiler -> BlockAnalyzer -> MergeAnalyzer
       -> FormulaDetector -> PlanAggregator
```

Each analyzer reads fields earlier stages populated and fills its own. Robustness policy (spec §6): an analyzer exception is absorbed into `context.warnings` and the pipeline continues; only `InspectorError` subclasses (corrupt/encrypted workbook, from `exceptions.py`) propagate. `BlockSegmenter` splits each sheet into row bands; `BlockAnalyzer` re-runs header/boundary/type detection per band so stacked tables each become a `TableBlock` with their own read plan.

### Coordinate-system contract [D1] — the most important invariant

- `SheetProfile`, `TableBlock`, and all inspection-domain row/col positions are **openpyxl 1-based**.
- `ReadPlan` positions are **pandas 0-based**.
- `ColumnProfile.index` and `ReadPlan.dtype_map` keys are **0-based from the table top-left** (dtype_map keys are stringified ints, e.g. `"0"`) [D5].
- The single 1-based -> 0-based conversion happens **only in `PlanAggregator`** (`aggregator.py`). Never convert elsewhere; off-by-one alignment is pinned by golden tests.

### Loader mode policy [D3] (`loader.py`)

Three cached openpyxl handles per file, all owned by `Loader` (a context manager; every handle closed before `inspect()` returns):

- **structure** (`read_only=False, data_only=True`) — opened once; the only mode exposing `merged_cells`/trustworthy dimensions.
- **data** (`read_only=True, data_only=True`) — forward streaming over cached values for sampling.
- **formula** (`read_only=True, data_only=False`) — lazy by contract; opened only when formula markup is actually found.

Open failures are disambiguated by sniffing the OLE2 magic: encrypted container -> `EncryptedWorkbookError`, anything else -> `CorruptWorkbookError`.

### Override channel [D2] (`models.py`, `options.py`)

`InspectionOptions.sheet_overrides` carries per-sheet `SheetOverride`s. Every estimated field records `provenance` (`"heuristic"`/`"manual"`/`"default"`); an overridden field skips heuristic computation and gets `provenance="manual", confidence=1.0`. `SheetOverride.header_row` uses an `_UNSET` sentinel to distinguish three states: unspecified (defer to heuristic), `int` (forced row), and explicit `None` (declared headerless) — check via `header_row_set` / `options.py:has_header_override()`, never by the value alone.

### Heuristics [D4] (`heuristics.py`)

All scoring weights, sample sizes, and thresholds are fixed v1 constants calibrated against the fixture corpus (header scan scoring, `BLANK_RUN=2` block separation, density/keyword subtotal detection including Korean keywords 합계/소계/총계/계, type-inference sampling). Don't tune them without updating the fixture-pinned tests.

### Key guarantees (spec §8) that tests enforce

- **Read-only / idempotent**: file bytes (SHA-256) unchanged; repeated inspections deeply equal; no RNG.
- **Streaming + sampling**: large workbooks are never fully materialized; `inspect()` tracemalloc peak <= 200 MB on 100k rows (the `slow` smoke).
- **No silent loss**: skipped sheets, excluded subtotal rows, formula-cache gaps etc. all surface via `warnings`/`notes`.

## Tests

- Fixture corpus: every `tests/fixtures/*.xlsx` is generated deterministically by `tests/fixtures/generate.py`; the session-scoped `fixture_corpus` fixture regenerates it, so never hand-edit those files — edit the generator. Look up files via the `fixture_path` fixture.
- Analyzer unit tests use partial-context synthesis: `make_context()` / `make_sheet_profile()` from `tests/conftest.py` (also exposed as `context_factory` / `sheet_profile_factory` fixtures) build an `InspectionContext` with only the fields under test.
- The 100k-row perf workbook is built on demand into a pytest tmp dir (`perf_fixture_path`), deliberately outside the corpus.

## Conventions

- Spec/plan docs, fixture data, and some heuristic keywords are Korean; code/docstrings are English. Docstrings cite spec sections and decision IDs — keep doing this for new code.
- Multi-level headers flatten to `"상위 / 하위"`; duplicate column names get `.1`/`.2` suffixes; headerless tables get `col_0..col_n`; `numeric_text` columns stay strings (leading zeros survive). These serialization rules are a fixed contract (README "Serialization rules").
