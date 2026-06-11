"""Formula Detector tests (spec §4.7, plan v2 Phase 12) [D6].

Covers the four Phase 12 deliverables plus the plan's three review probes:

* golden ``formulas.xlsx``: the D column is flagged ``has_formula=True`` with
  ``read_hint="as_formula"``, a cache-empty warning surfaces, the aggregator
  skips its dtype inference and records an advisory note (Steps 2-4);
* golden ``formulas_cached.xlsx`` (probe 2): a hand-crafted file whose
  formula cells carry real cached ``<v>`` results — the only way to execute
  the ``as_value`` branch, since openpyxl never writes caches — loads the
  cached numbers and produces no warning/note;
* laziness (probe 3): an open-counter patch on ``Loader.formula_workbook``
  proves the formula-mode workbook is **never opened** for formula-free
  fixtures and is opened for the formula fixtures (probe 1's close-handle
  coverage lives in ``test_loader.py`` / ``test_idempotency.py``).

Fixture coordinates come from ``tests/fixtures/generate.py`` ``FIXTURES``
(1-based [D1]): both fixtures are header row 1 (``품목,수량,단가,금액``),
data rows 2-5, column D ``=B{r}*C{r}``; the cached variant injects
D2..D5 = 200/600/1200/2000.
"""

from __future__ import annotations

import json

import pytest

from excel_inspector import Loader, extract, inspect
from excel_inspector.analyzers.formula_detector import (
    _is_formula_text,
    _workbook_has_formula_markup,
)

#: The stable core of the cache-empty warning (plan v2 §6 Step 2).
CACHE_EMPTY = "formula cache empty (file never opened in Excel?)"


# ---------------------------------------------------------------------------
# Unit: formula-text predicate and the lazy zip-scan gate
# ---------------------------------------------------------------------------


def test_is_formula_text_rules() -> None:
    """Only a str starting with '=' is a formula (plan v2 §6 Step 2)."""

    assert _is_formula_text("=B2*C2") is True
    assert _is_formula_text("=SUM(A1:A9)") is True
    assert _is_formula_text("plain text") is False
    assert _is_formula_text("") is False
    assert _is_formula_text(None) is False
    assert _is_formula_text(42) is False


def test_formula_markup_scan_detects_formula_fixtures(fixture_path) -> None:
    """The raw zip scan finds the <f> markup in both formula fixtures."""

    assert _workbook_has_formula_markup(fixture_path("formulas")) is True
    assert _workbook_has_formula_markup(fixture_path("formulas_cached")) is True


def test_formula_markup_scan_clean_on_formula_free_fixture(
    fixture_path,
) -> None:
    """A formula-free workbook scans negative (the lazy gate stays shut)."""

    assert _workbook_has_formula_markup(fixture_path("header_simple")) is False
    assert _workbook_has_formula_markup(fixture_path("types_mixed")) is False


# ---------------------------------------------------------------------------
# Golden: formulas.xlsx — empty cache -> as_formula (plan v2 §6 Step 4)
# ---------------------------------------------------------------------------


def test_formulas_column_flagged_as_formula(fixture_path) -> None:
    """D column: has_formula=True + read_hint='as_formula' (empty cache)."""

    profile = inspect(fixture_path("formulas"))
    (sheet,) = profile.sheets
    flags = [(c.index, c.has_formula, c.read_hint) for c in sheet.columns]
    assert flags == [
        (0, False, "as_value"),
        (1, False, "as_value"),
        (2, False, "as_value"),
        (3, True, "as_formula"),
    ]


def test_formulas_block_columns_carry_the_same_flags(fixture_path) -> None:
    """Detection is block-coupled: the block's columns are the flagged ones.

    The flat sheet mirror shares the block's ColumnProfile instances (Phase
    10 mirror rule), so the per-block detection and the sheet view can never
    diverge.
    """

    profile = inspect(fixture_path("formulas"))
    (sheet,) = profile.sheets
    (block,) = sheet.blocks
    assert [c.has_formula for c in block.columns] == [
        False,
        False,
        False,
        True,
    ]
    assert block.columns[3].read_hint == "as_formula"
    assert all(a is b for a, b in zip(sheet.columns, block.columns))


def test_formulas_cache_empty_warning_recorded(fixture_path) -> None:
    """Exactly one cache-empty warning, naming the affected column."""

    profile = inspect(fixture_path("formulas"))
    hits = [w for w in profile.open_errors if CACHE_EMPTY in w]
    assert len(hits) == 1
    assert "column 3" in hits[0]


def test_formulas_dtype_inference_skipped_for_as_formula(fixture_path) -> None:
    """Aggregator: the as_formula column gets no dtype key (plan v2 §6 Step 3).

    Without Phase 12 the all-None cache makes the Type Profiler call column D
    'text' and the aggregator would emit ``'3': 'string'``; the as_formula
    skip must drop exactly that key while leaving the genuine text column
    ('품목' at index 0) typed.
    """

    profile = inspect(fixture_path("formulas"))
    (sheet,) = profile.sheets
    assert sheet.read_plan is not None
    assert "3" not in sheet.read_plan.dtype_map
    assert sheet.read_plan.dtype_map == {"0": "string"}


def test_formulas_plan_carries_as_formula_note(fixture_path) -> None:
    """The skip is visible: an advisory note lands on the plan (Step 3)."""

    profile = inspect(fixture_path("formulas"))
    (sheet,) = profile.sheets
    assert sheet.read_plan is not None
    notes = [n for n in sheet.read_plan.notes if "read_hint=as_formula" in n]
    assert len(notes) == 1
    assert "formula column 3" in notes[0]
    assert "dtype inference skipped" in notes[0]
    assert "data_only=False" in notes[0]


def test_formulas_block_plan_mirrors_flat_plan(fixture_path) -> None:
    """The mirror block's plan equals the flat plan (Phase 10 invariant)."""

    profile = inspect(fixture_path("formulas"))
    (sheet,) = profile.sheets
    (block,) = sheet.blocks
    assert block.read_plan == sheet.read_plan


def test_formulas_extract_records_and_notes(fixture_path) -> None:
    """extract(): empty-cache formula column is all null; note + warning ride.

    The result layer changes nothing about Phase 12 — the note flows through
    ``ReadPlan.notes`` into ``TableResult.notes`` and the warning into
    ``WorkbookResult.warnings`` — pinned here end to end.
    """

    result = extract(fixture_path("formulas"))
    (table,) = result.tables
    payload = table.to_dict()
    assert payload["row_count"] == 4
    assert [record["금액"] for record in payload["records"]] == [None] * 4
    assert [record["수량"] for record in payload["records"]] == [2, 3, 4, 5]
    assert any("read_hint=as_formula" in note for note in payload["notes"])
    assert any(CACHE_EMPTY in warning for warning in result.warnings)
    # The whole thing serializes (schema v1.0) without error.
    assert json.loads(result.to_json())["schema_version"] == "1.0"


# ---------------------------------------------------------------------------
# Golden: formulas_cached.xlsx — real caches -> as_value (probe 2)
# ---------------------------------------------------------------------------


def test_cached_fixture_really_has_caches_and_formulas(fixture_path) -> None:
    """The hand-crafted <v> injection is real in both read modes (probe 2).

    Measured directly through the loader: data mode reads the injected cached
    numbers (so the fixture is not silently cacheless, which would turn the
    as_value tests into dead code), while formula mode still reads the
    formula strings.
    """

    with Loader(fixture_path("formulas_cached")) as loader:
        data_ws = loader.data_workbook()["Sheet1"]
        cached = [
            row[3]
            for row in data_ws.iter_rows(min_row=2, max_row=5, values_only=True)
        ]
        assert cached == [200, 600, 1200, 2000]  # FIXTURES golden values
        formula_ws = loader.formula_workbook()["Sheet1"]
        formulas = [
            row[3]
            for row in formula_ws.iter_rows(
                min_row=2, max_row=5, values_only=True
            )
        ]
        assert formulas == ["=B2*C2", "=B3*C3", "=B4*C4", "=B5*C5"]


def test_cached_formula_column_reads_as_value(fixture_path) -> None:
    """With caches present the hint is as_value and nothing is warned/noted."""

    profile = inspect(fixture_path("formulas_cached"))
    (sheet,) = profile.sheets
    column = sheet.columns[3]
    assert column.has_formula is True
    assert column.read_hint == "as_value"
    assert not any(CACHE_EMPTY in w for w in profile.open_errors)
    assert sheet.read_plan is not None
    assert sheet.read_plan.notes == []
    # Column D profiled from the cached numbers -> number -> no dtype key.
    assert sheet.read_plan.dtype_map == {"0": "string"}


def test_cached_formula_extract_loads_cached_values(fixture_path) -> None:
    """extract() delivers the cached results through pandas (as_value path)."""

    result = extract(fixture_path("formulas_cached"))
    (table,) = result.tables
    assert [int(v) for v in table.dataframe["금액"]] == [200, 600, 1200, 2000]
    assert int(table.dataframe["금액"].sum()) == 4000
    assert not any(CACHE_EMPTY in w for w in result.warnings)


# ---------------------------------------------------------------------------
# Laziness (probe 3): the formula workbook is opened only when needed
# ---------------------------------------------------------------------------


@pytest.fixture
def formula_open_counter(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    """Count every ``Loader.formula_workbook`` call (delegating to the real one).

    A counting patch — not a mock that fakes success — so a counted call is a
    genuine formula-mode open request and a zero count is a meaningful "never
    opened" proof (plan v2 §6 review probe 3).
    """

    calls = {"count": 0}
    real = Loader.formula_workbook

    def _counting(self: Loader):  # noqa: ANN202 - test shim
        calls["count"] += 1
        return real(self)

    monkeypatch.setattr(Loader, "formula_workbook", _counting)
    return calls


#: Every openable fixture that contains no formulas (the whole pre-Phase-12
#: corpus plus the Phase 13 ``left_margin_with_subtotal`` variant). Kept
#: literal so a new formula fixture cannot silently join the "must not open"
#: list.
FORMULA_FREE_FIXTURES = [
    "header_simple",
    "header_offset",
    "offset_plus_subtotals",
    "merged_header",
    "multi_level_header",
    "multi_level_numeric_text",
    "types_mixed",
    "left_margin_cols",
    "left_margin_with_subtotal",
    "mixed_sheets",
    "hidden_sheet",
    "blank_run_terminates",
    "interior_blank",
    "empty_sheet",
    "header_only",
    "no_header",
    "large_table",
    "multi_table_stacked",
    "stacked_uneven_width",
    "title_between_tables",
]


@pytest.mark.parametrize("fixture_id", FORMULA_FREE_FIXTURES)
def test_formula_workbook_never_opened_without_formulas(
    fixture_path, formula_open_counter: dict[str, int], fixture_id: str
) -> None:
    """inspect() on a formula-free workbook never opens formula mode (lazy)."""

    inspect(fixture_path(fixture_id))
    assert formula_open_counter["count"] == 0


@pytest.mark.parametrize("fixture_id", ["formulas", "formulas_cached"])
def test_formula_workbook_opened_for_formula_fixtures(
    fixture_path, formula_open_counter: dict[str, int], fixture_id: str
) -> None:
    """The gate opens for real formula files (the counter actually counts)."""

    inspect(fixture_path(fixture_id))
    assert formula_open_counter["count"] >= 1


# ---------------------------------------------------------------------------
# Determinism (v2 test strategy §8)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("fixture_id", ["formulas", "formulas_cached"])
def test_formula_fixture_json_is_deterministic(
    fixture_path, fixture_id: str
) -> None:
    """Two extract() runs serialize identically (warnings/notes included)."""

    path = fixture_path(fixture_id)
    assert extract(path).to_json() == extract(path).to_json()
