# Excel Structure Inspector

Read-only inspection of `.xlsx` workbooks that turns every table of a workbook
into its own pandas DataFrame, with JSON / Markdown serialization and a CLI.
See `docs/excel-structure-inspector-spec.md` (spec, revision 1),
`docs/excel-structure-inspector-implementation-plan.md` (v1 plan), and
`docs/excel-structure-inspector-v2-plan.md` (v2 plan: result layer,
multi-table, merged/multi-level headers, formula detection, CLI).

## Environment

macOS / Homebrew Python is externally managed (PEP 668), so a virtualenv is
required.

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -U pip
.venv/bin/python -m pip install -r requirements.txt
```

## Quick start — `extract()` (v2 one-call API)

`extract()` inspects the workbook **read-only**, then loads every detected
table (one `TableResult` per table — a sheet with stacked tables yields
several) and returns a `WorkbookResult` ready for serialization.

```python
from excel_inspector import extract

result = extract("book.xlsx")              # -> WorkbookResult

for table in result.tables:                # every table in the workbook
    print(table.table_id)                  # "매출!T1", "매출!T2", ...
    print(table.dataframe.shape)           # a cleaned pandas DataFrame
    print(table.notes)                     # advisories (merges, formulas, ...)

print(result.to_json(indent=2))            # JSON document (schema v1.0)
print(result.to_markdown())                # human-readable Markdown tables
print(result.warnings)                     # workbook-level advisories
```

Overrides work exactly as with `inspect()`:

```python
from excel_inspector import InspectionOptions, SheetOverride, extract

options = InspectionOptions(
    sheet_overrides={
        "Sheet1": SheetOverride(header_row=4, dtype_force={"0": "string"}),
        "RawDump": SheetOverride(header_row=None),   # declare headerless
    }
)
result = extract("book.xlsx", options)
# A headerless table gets col_0..col_n column names, and its plan/notes carry
# "headerless sheet: dtype inference skipped" (dtype_force still applies).
```

## CLI

```sh
# Markdown tables (default; --max-rows bounds rows rendered per table)
.venv/bin/python -m excel_inspector book.xlsx
.venv/bin/python -m excel_inspector book.xlsx --max-rows 50

# Full JSON document (schema v1.0; always contains every row)
.venv/bin/python -m excel_inspector book.xlsx --format json
```

Exit codes: `0` on success; `1` for corrupt/encrypted/unreadable input, with
an explicit `error: ...` line on stderr and nothing on stdout.

## JSON schema (v1.0)

`to_json()` emits a deterministic document; the shape is stable (later phases
only add `tables` items):

```json
{
  "schema_version": "1.0",
  "file": "/abs/path/book.xlsx",
  "sheets": [
    {
      "name": "매출",
      "is_visible": true,
      "tables": [
        {
          "table_id": "매출!T1",
          "header_row": 4,
          "header_confidence": 0.88,
          "header_provenance": "heuristic",
          "columns": [
            {"index": 0, "name": "지역", "inferred_type": "text", "null_ratio": 0.0}
          ],
          "row_count": 6,
          "records": [
            {"지역": "서울", "제품코드": "00123", "출시일": "2026-01-05T00:00:00",
             "수량": 50, "매출액": 20000, "비고": null}
          ],
          "notes": []
        }
      ],
      "skipped": false,
      "skip_reason": null
    },
    {"name": "안내", "is_visible": true, "tables": [], "skipped": true,
     "skip_reason": "non-tabular"}
  ],
  "warnings": []
}
```

Serialization rules (fixed contract): dates -> ISO 8601 strings; missing
(`NaN`/`NA`/`NaT`) -> `null`; `numeric_text` stays a string (leading zeros
survive); numpy scalars -> native Python; headerless columns -> `col_0..col_n`;
duplicate column names -> `.1`/`.2` suffixes; multi-level headers flatten to
`"상위 / 하위"`.

`skip_reason` values: `"non-tabular"` (not a tabular candidate, spec §9) and
`"no-table-detected"` (tabular candidate, but every detected band was judged
non-table — issue #10; a `header_row` or `is_tabular` override forces
loading). A sheet whose header heuristic failed still loads via the v1
fallback (first row assumed as header) and carries that assumption in the
table's `notes` instead of being skipped.

## Lower-level API — `inspect()` + `load_dataframe()` (v1)

`inspect()` returns a `WorkbookProfile` whose per-sheet `read_plan` (a
`ReadPlan`, pandas 0-based) can be fed straight into the bundled adapter —
useful when you want the inspection metadata without loading, or your own
loading policy.

```python
from excel_inspector import inspect, load_dataframe

# Inspect: never mutates the file; degenerate sheets are explicit states
# (read_plan=None / needs_manual_header=...), only corrupt/encrypted raise.
profile = inspect("book.xlsx")

for sheet in profile.sheets:
    if sheet.read_plan is None:        # non-tabular (e.g. a README sheet)
        continue
    if sheet.needs_manual_header:      # header heuristic was not confident
        continue
    df = load_dataframe("book.xlsx", sheet.read_plan)
    print(sheet.name, df.shape)
    for block in sheet.blocks:         # per-table blocks (multi-table sheets)
        print(" ", block.block_index, block.read_plan)
```

Key guarantees (spec §8):

- **Read-only / idempotent** — the file's bytes (SHA-256) are unchanged by
  inspection, and repeated inspections produce a deeply equal
  `WorkbookProfile` (deterministic, RNG-free sampling). Every workbook handle
  the loader opens is closed before `inspect()` returns.
- **Streaming + sampling** — row data is read in `read_only` streaming mode;
  the header scan reads only the top rows and the type profiler samples a
  bounded number of rows, so large workbooks are not fully materialized. The
  structure workbook is opened once for merge/dimension metadata ([D3]); the
  formula-mode workbook opens lazily, only for files that actually contain
  formula markup.
- **No silent loss** — stacked multi-table sheets are extracted table by
  table; non-table bands (titles/footnotes), excluded subtotal rows, empty
  formula caches, and skipped dtype inference are all surfaced via
  `warnings`/`notes`.

## Test

```sh
.venv/bin/python -m pytest                      # default suite (fast)
.venv/bin/python -m pytest -m slow              # 100k-row perf smoke only
.venv/bin/python -m pytest -m "slow or not slow"  # everything
```

The fixture corpus (including a 5000-row `large_table` performance fixture) is
generated deterministically by `tests/fixtures/generate.py` and regenerated by
a session-scoped pytest fixture. The 100k-row perf workbook is built on demand
into a pytest tmp dir by the `slow`-marked smoke (deselected by default via
`addopts`) and asserts the spec §8 memory budget (`inspect()` tracemalloc peak
<= 200 MB).

## Status

v1 (Phases 0-8) and v2 (Phases 9-13) complete: data models, options, context,
analyzer/pipeline interfaces, domain exceptions, heuristics, the mode-aware
loader, all analyzers (sheet enumerator, merge scanner, block segmenter,
header locator, boundary detector, type profiler, block analyzer, merge
analyzer, formula detector), the plan aggregator, the pandas read-side
adapter, the result layer (`extract()` / `WorkbookResult` / JSON v1.0 /
Markdown), multi-table extraction, merged-header bridging, multi-level header
loading (MultiIndex + flattening), formula detection with `as_value` /
`as_formula` read hints, and the `python -m excel_inspector` CLI.
