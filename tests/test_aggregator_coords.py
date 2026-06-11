"""Aggregator coordinate-conversion + golden round-trip tests (Phase 4) [D1].

These tests close the [D1] coordinate contract loop end to end: they drive the
full ``inspect()`` pipeline (SheetEnumerator -> HeaderLocator ->
BoundaryDetector -> PlanAggregator) to produce a :class:`ReadPlan`, then feed
that plan *verbatim* into :func:`pandas.read_excel` and assert on the **loaded
DataFrame**. This is the regression net the spec demands (spec §4.8, §5.5,
implementation plan §5.2): the exact 1-based -> 0-based alignment is pinned by
an actual pandas load, not by remembering library internals.

Invariants fixed here (the inspector's core value — preventing alignment slip
and aggregation duplication):

* **No row slip**: leading title rows are absorbed into ``skiprows`` so the
  header binds to the right row and no title text leaks into the data.
* **No aggregation duplication**: interior subtotal rows and the trailing grand
  total are dropped, so re-aggregating the loaded rows does not double-count.
* **No last-row loss**: ``nrows`` spans the whole inclusive data region
  (interior skips are *not* subtracted), so the final data row is still read.
* **Column trimming**: a left filler column is excluded via ``usecols``.

Coordinate-helper unit tests live in ``test_aggregator.py``; this module is the
end-to-end golden layer. Assertions target the load-domain *meaning* (column
names, row counts, presence/absence of marker rows) rather than full-frame
equality, to keep the golden robust against incidental dtype churn (impl plan
§5.2 "golden fragility mitigation").
"""

from __future__ import annotations

import pandas as pd
import pytest

from excel_inspector import ReadPlan, SheetProfile, WorkbookProfile, inspect
from excel_inspector.adapters import load_dataframe


def _load(path, plan: ReadPlan) -> pd.DataFrame:
    """Load ``path`` through a :class:`ReadPlan` via the pandas adapter.

    Delegates to :func:`excel_inspector.adapters.load_dataframe` so the golden
    round-trip exercises the *real* read-side boundary (Phase 7) rather than an
    inline translation; the adapter owns the (0-based, pandas-domain) field ->
    kwarg mapping, the ``usecols`` "None == all columns" contract, and the [D5]
    dtype-key reduction.

    Args:
        path: The fixture workbook path.
        plan: The plan produced by ``inspect()`` for the target sheet.

    Returns:
        The loaded :class:`pandas.DataFrame`.
    """

    return load_dataframe(path, plan)


def _sheet(profile: WorkbookProfile, name: str) -> SheetProfile:
    """Return the named sheet profile, asserting it exists and has a plan."""

    matches = [s for s in profile.sheets if s.name == name]
    assert matches, f"sheet {name!r} not found in {[s.name for s in profile.sheets]}"
    sheet = matches[0]
    assert sheet.read_plan is not None, f"sheet {name!r} has no read plan"
    return sheet


# ---------------------------------------------------------------------------
# offset_plus_subtotals — the [D1] flagship: leading rows + interior subtotals
# + trailing grand total.
# ---------------------------------------------------------------------------


def test_offset_plus_subtotals_plan_coords(fixture_path) -> None:
    """The synthesized plan pins the exact [D1] coordinate conversion.

    Title rows 1-3 -> skiprows {0,1,2}; subtotal rows 8,12 and grand-total row
    13 -> {7,11,12}; header normalized to 0; nrows = full inclusive span
    (data 5..11 = 7) with interior skips NOT subtracted (spec §4.8 rule 4).
    """

    sheet = _sheet(inspect(fixture_path("offset_plus_subtotals")), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None

    # Inspection-domain boundaries (1-based) feeding the conversion.
    assert (sheet.data_start_row, sheet.data_end_row) == (5, 11)
    assert sheet.skip_rows == [8, 12, 13]

    # Loading-domain plan (0-based) [D1].
    assert plan.header == 0
    assert plan.skiprows == [0, 1, 2, 7, 11, 12]
    assert plan.usecols is None  # full-width A:D -> no column restriction
    assert plan.nrows == 7  # interior skips NOT subtracted


def test_offset_plus_subtotals_golden_roundtrip(fixture_path) -> None:
    """inspect() -> ReadPlan -> pandas yields exactly the 6 clean data rows.

    Subtotal/total leakage = 0, last data row preserved, columns intact, and
    the re-aggregated amount equals the file's own grand total (590), proving
    the interior/trailing aggregation rows were excluded (no double-count).
    """

    path = fixture_path("offset_plus_subtotals")
    sheet = _sheet(inspect(path), "Sheet1")
    df = _load(path, sheet.read_plan)

    # Exactly the data rows; no title, no subtotal, no grand total.
    assert len(df) == 6
    assert list(df.columns) == ["dept", "item", "month", "amount"]

    # No subtotal/total marker leaked into the body.
    assert not df["dept"].astype(str).str.contains("소계|합계").any()

    # Last real data row (관리/비품/40) survived -> nrows reached the end.
    last = df.iloc[-1]
    assert (last["dept"], last["item"], last["amount"]) == ("관리", "비품", 40)

    # Aggregation-duplication check: the 6 leaf rows sum to 590, equal to the
    # file's grand-total row (which was correctly excluded). Summing the loaded
    # rows therefore does NOT double-count the totals.
    assert df["amount"].sum() == 590
    assert list(df["amount"]) == [100, 200, 50, 80, 120, 40]


# ---------------------------------------------------------------------------
# left_margin_cols — usecols trims the left filler column.
# ---------------------------------------------------------------------------


def test_left_margin_cols_usecols_trims_filler(fixture_path) -> None:
    """A left description column is excluded by ``usecols='C:E'``.

    The real table lives in columns C-E; only those load, and the filler text
    in column A never appears as a column.
    """

    path = fixture_path("left_margin_cols")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None

    # Inspection-domain column boundaries (1-based) -> usecols.
    assert (sheet.data_left_col, sheet.data_right_col) == (3, 5)
    assert plan.usecols == "C:E"
    assert plan.nrows == 6  # data rows 2..7 inclusive

    df = _load(path, plan)
    assert list(df.columns) == ["sku", "qty", "price"]
    assert len(df) == 6
    # The filler column A header ("참고사항") is not present as a column.
    assert "참고사항" not in df.columns
    # First/last body rows are the table rows, not filler text.
    assert df.iloc[0]["sku"] == "A-1"
    assert df.iloc[-1]["sku"] == "A-6"


# ---------------------------------------------------------------------------
# header_offset — title rows do not bleed into the data.
# ---------------------------------------------------------------------------


def test_header_offset_title_not_in_data(fixture_path) -> None:
    """The 3-row title block is absorbed; no title text reaches the body."""

    path = fixture_path("header_offset")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None

    assert sheet.header_row == 4
    assert plan.header == 0
    assert plan.skiprows == [0, 1, 2]  # title rows 1-3 absorbed

    df = _load(path, plan)
    assert list(df.columns) == ["product", "region", "qty", "amount"]
    assert len(df) == 5

    # None of the title strings leaked into any cell of the loaded frame.
    flat = df.astype(str).to_numpy().ravel().tolist() + list(df.columns)
    for title in ("월간 판매 보고서", "작성일", "단위"):
        assert not any(title in cell for cell in flat), title

    # The first data row is the real first product, not a title line.
    assert df.iloc[0]["product"] == "Widget"


# ---------------------------------------------------------------------------
# header_simple — a plain row-1 header round-trips cleanly (no slip, no skips).
# ---------------------------------------------------------------------------


def test_header_simple_clean_roundtrip(fixture_path) -> None:
    """A row-1 header table loads with no skips and all 5 rows intact."""

    path = fixture_path("header_simple")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None

    assert plan.header == 0
    assert plan.skiprows == []
    assert plan.usecols is None
    assert plan.nrows == 5

    df = _load(path, plan)
    assert list(df.columns) == ["name", "age", "city", "score"]
    assert len(df) == 5
    assert df.iloc[0]["name"] == "Alice"
    assert df.iloc[-1]["name"] == "Eve"


# ---------------------------------------------------------------------------
# blank_run_terminates — the blank-run terminator stops the read before noise.
# ---------------------------------------------------------------------------


def test_blank_run_terminates_excludes_noise(fixture_path) -> None:
    """The 2-row blank run terminates the table; trailing noise is not loaded."""

    path = fixture_path("blank_run_terminates")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None

    assert (sheet.data_start_row, sheet.data_end_row) == (2, 5)
    assert plan.nrows == 4  # rows 2..5; noise rows 9-10 beyond the terminator

    df = _load(path, plan)
    assert list(df.columns) == ["name", "qty", "price"]
    assert len(df) == 4
    # The trailing noise row (Z-9) and stray label (기타 메모) are excluded.
    flat = df.astype(str).to_numpy().ravel().tolist()
    assert "Z-9" not in flat
    assert not any("기타 메모" in cell for cell in flat)
    assert df.iloc[-1]["name"] == "A-4"


# ---------------------------------------------------------------------------
# interior_blank — a single interior blank row is skipped, no all-NaN leak.
# ---------------------------------------------------------------------------


def test_interior_blank_no_nan_row_leaks(fixture_path) -> None:
    """A lone interior blank row is dropped; no all-NaN row reaches the frame.

    Regression for MEDIUM #4: a single blank row (below the BLANK_RUN of 2
    terminator threshold) between data rows must be recorded in ``skip_rows``
    and converted to ``skiprows`` so the loaded DataFrame has no all-NaN row and
    spans the whole data region (rows 2-6, the blank at row 4 excluded).
    """

    path = fixture_path("interior_blank")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None

    assert (sheet.data_start_row, sheet.data_end_row) == (2, 6)
    assert sheet.skip_rows == [4]
    # 0-based: interior blank row 4 -> skiprows index 3; nrows spans 2..6 = 5.
    assert plan.skiprows == [3]
    assert plan.nrows == 5

    df = _load(path, plan)
    assert list(df.columns) == ["name", "qty", "price"]
    # 4 real data rows; the blank row was dropped (not loaded as all-NaN).
    assert len(df) == 4
    assert not df.isna().all(axis=1).any()
    assert list(df["name"]) == ["A-1", "A-2", "A-3", "A-4"]


# ---------------------------------------------------------------------------
# Override golden: a header_row override drives the same coordinate conversion
# and the loaded frame matches the heuristic path (provenance=manual) [D2].
# ---------------------------------------------------------------------------


def test_header_override_golden_matches_heuristic(fixture_path) -> None:
    """A manual header_row override yields the identical loaded frame [D2].

    Forcing header_row=4 on header_offset must produce the same body as the
    heuristic detection, proving the override channel feeds the same [D1]
    conversion (rows 1-3 absorbed) and does not slip a row.
    """

    from excel_inspector import InspectionOptions, SheetOverride

    path = fixture_path("header_offset")
    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=4)}
    )
    sheet = _sheet(inspect(path, options), "Sheet1")
    assert sheet.header_provenance == "manual"
    assert sheet.header_confidence == 1.0
    plan = sheet.read_plan
    assert plan is not None
    assert plan.skiprows == [0, 1, 2]
    assert plan.header == 0

    df = _load(path, plan)
    assert list(df.columns) == ["product", "region", "qty", "amount"]
    assert len(df) == 5
    assert df.iloc[0]["product"] == "Widget"


# ---------------------------------------------------------------------------
# dtype_map golden [D5]: a forced dtype keyed by 0-based column position drives
# the loaded column's dtype through pandas.
# ---------------------------------------------------------------------------


def test_dtype_force_golden_applies_by_position(fixture_path) -> None:
    """``dtype_force`` keyed by 0-based position [D5] types the loaded column.

    Forcing position ``"1"`` (the ``age`` column) to ``string`` must surface as
    a string-dtype column in the loaded frame, confirming the key->position
    reduction the adapter performs.

    With the Type Profiler (Phase 5) wired in, the inferred text columns of
    ``header_simple`` (0=name, 2=city) also appear in the dtype_map; the forced
    position ``"1"`` (numeric ``age``, otherwise omitted) is added on top.
    """

    from excel_inspector import InspectionOptions, SheetOverride

    path = fixture_path("header_simple")
    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(dtype_force={"1": "string"})}
    )
    sheet = _sheet(inspect(path, options), "Sheet1")
    plan = sheet.read_plan
    assert plan is not None
    # Inferred text columns (0=name, 2=city) plus the forced numeric column 1.
    assert plan.dtype_map == {"0": "string", "1": "string", "2": "string"}

    df = _load(path, plan)
    # Column at 0-based position 1 is "age"; it must be string-typed.
    assert str(df.dtypes.iloc[1]) == "string"
    # Other columns are untouched by the force.
    assert str(df.dtypes.iloc[3]) == "float64"


# ---------------------------------------------------------------------------
# Cross-fixture parametrized guard: no plan ever reads a subtotal/total marker
# into the body for the aggregation-sensitive fixtures.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("fixture_id", "expected_rows"),
    [
        ("offset_plus_subtotals", 6),
        ("header_offset", 5),
        ("header_simple", 5),
        ("left_margin_cols", 6),
        ("blank_run_terminates", 4),
        ("interior_blank", 4),
    ],
)
def test_loaded_row_counts_match_plan_nrows(
    fixture_path, fixture_id: str, expected_rows: int
) -> None:
    """The loaded body length matches the clean data-row count for each fixture.

    Because interior subtotal rows are dropped from the output (but still spend
    the ``nrows`` budget), the loaded length can be < ``nrows``; it must equal
    the count of genuine data rows, never include skipped/aggregation rows.
    """

    path = fixture_path(fixture_id)
    sheet = _sheet(inspect(path), "Sheet1")
    df = _load(path, sheet.read_plan)
    assert len(df) == expected_rows
