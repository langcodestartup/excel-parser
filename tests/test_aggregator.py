"""Plan Aggregator v1 tests (spec §4.8, Phase 1) [D1][D5].

Covers the v1 minimal read-plan synthesis and the coordinate-conversion
skeleton helpers: 1-based -> 0-based row conversion, column-range -> usecols
translation, leading-row absorption, interior skip-row merging, nrows
computation, dtype_force application, and the non-tabular exclusion.
"""

from __future__ import annotations

import pytest

from excel_inspector import InspectionOptions, SheetOverride
from excel_inspector.aggregator import (
    PlanAggregator,
    build_read_plan,
    column_range_to_usecols,
    to_zero_based,
)

from conftest import make_context, make_sheet_profile  # type: ignore[import-not-found]


# -- coordinate-conversion helpers ----------------------------------------


def test_to_zero_based_converts() -> None:
    """A 1-based row converts to its 0-based equivalent [D1]."""

    assert to_zero_based(1) == 0
    assert to_zero_based(4) == 3


def test_to_zero_based_rejects_non_positive() -> None:
    """A non-1-based input is rejected (guards against double conversion)."""

    with pytest.raises(ValueError):
        to_zero_based(0)


def test_column_range_to_usecols() -> None:
    """1-based inclusive column span -> Excel-letter usecols range."""

    assert column_range_to_usecols(3, 5) == "C:E"
    assert column_range_to_usecols(1, 4) == "A:D"


def test_column_range_to_usecols_none_when_unbounded() -> None:
    """Missing either boundary means all columns (None)."""

    assert column_range_to_usecols(None, 5) is None
    assert column_range_to_usecols(3, None) is None
    assert column_range_to_usecols(None, None) is None


# -- v1 read plan synthesis -----------------------------------------------


def test_v1_simple_sheet_header_zero_engine_openpyxl() -> None:
    """No header/boundary info -> header=0, openpyxl engine, empty skiprows."""

    profile = make_sheet_profile(name="Sheet1", max_row=6, max_col=4)
    plan = build_read_plan(profile)

    assert plan.sheet_name == "Sheet1"
    assert plan.engine == "openpyxl"
    assert plan.header == 0
    assert plan.skiprows == []
    assert plan.usecols is None
    assert plan.nrows is None
    assert plan.dtype_map == {}


def test_leading_rows_absorbed_into_skiprows() -> None:
    """A header at 1-based row 4 absorbs rows 1-3 (0-based 0,1,2) [D1]."""

    profile = make_sheet_profile(name="Sheet1", header_row=4)
    plan = build_read_plan(profile)
    assert plan.skiprows == [0, 1, 2]
    assert plan.header == 0  # normalized to the post-skip frame top [D1]


def test_interior_skip_rows_converted_to_zero_based() -> None:
    """1-based subtotal rows merge into skiprows as 0-based absolute indices."""

    profile = make_sheet_profile(
        name="Sheet1",
        header_row=4,
        data_start_row=5,
        data_end_row=11,
        skip_rows=[8, 12, 13],
    )
    plan = build_read_plan(profile)
    # Leading 1-3 -> {0,1,2}; subtotals 8,12,13 -> {7,11,12}; merged & sorted.
    assert plan.skiprows == [0, 1, 2, 7, 11, 12]
    assert plan.header == 0


def test_nrows_is_full_inclusive_span() -> None:
    """nrows = full 1-based inclusive data span; interior skips NOT subtracted.

    pandas ``nrows`` counts original rows consumed after the header; interior
    ``skiprows`` are dropped from the output but still consume the budget, so
    subtracting them would drop the last data row (verified vs pandas 3.0.3).
    """

    profile = make_sheet_profile(
        name="Sheet1",
        data_start_row=5,
        data_end_row=11,
        skip_rows=[8],  # one interior subtotal within 5..11
    )
    plan = build_read_plan(profile)
    # span 5..11 = 7 rows, interior skip NOT subtracted.
    assert plan.nrows == 7


def test_usecols_from_column_boundaries() -> None:
    """left/right column boundaries become a usecols range."""

    profile = make_sheet_profile(
        name="Sheet1", data_left_col=3, data_right_col=5
    )
    plan = build_read_plan(profile)
    assert plan.usecols == "C:E"


def test_dtype_force_override_applied() -> None:
    """dtype_force [D5] is carried into the read plan's dtype_map."""

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(dtype_force={"1": "string"})}
    )
    profile = make_sheet_profile(name="Sheet1")
    plan = build_read_plan(profile, options)
    assert plan.dtype_map == {"1": "string"}


# -- provenance [D2] -------------------------------------------------------


def test_header_override_records_manual_provenance() -> None:
    """A header_row override sets provenance=manual + confidence=1.0 [D2]."""

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=4)}
    )
    profile = make_sheet_profile(name="Sheet1")
    plan = build_read_plan(profile, options)

    assert profile.header_provenance == "manual"
    assert profile.header_confidence == 1.0
    assert profile.header_row == 4
    assert profile.needs_manual_header is False
    # The override drives the coordinate conversion: rows 1-3 absorbed.
    assert plan.skiprows == [0, 1, 2]
    assert plan.header == 0


def test_v1_fallback_header_records_default_provenance() -> None:
    """No override and no heuristic -> provenance=default (honest) [D2]."""

    profile = make_sheet_profile(name="Sheet1", header_provenance="default")
    build_read_plan(profile)
    assert profile.header_provenance == "default"


def test_headerless_override_yields_plan_header_none() -> None:
    """An explicit headerless override -> plan.header is None (HIGH #3 / §9).

    A SheetOverride(header_row=None) declares the sheet has no header, so the
    plan must set ``header=None`` (pandas reads no header row); otherwise the
    first data row would be consumed as column names (§9 violation). Distinct
    from the detection fallback, which assumes header=0.
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    profile = make_sheet_profile(name="Sheet1", max_row=5, max_col=3)
    plan = build_read_plan(profile, options)

    assert plan.header is None
    assert profile.header_row is None
    assert profile.header_provenance == "manual"
    assert profile.header_confidence == 1.0
    # No leading rows to absorb (no header anchor), so no skiprows synthesized.
    assert plan.skiprows == []


def test_detection_fallback_header_is_zero_not_none() -> None:
    """No override + no detected header -> header=0 fallback, NOT None (HIGH #3).

    Contrast with the headerless override: when there is simply no detected
    header and no override, v1 still assumes the first row is the header.
    """

    profile = make_sheet_profile(name="Sheet1", max_row=5, max_col=3)
    plan = build_read_plan(profile)
    assert plan.header == 0


def test_headerless_override_plan_notes_dtype_skip() -> None:
    """L6 (plan v2 Phase 13 Step 2): a headerless plan says dtype was skipped.

    With ``header_row=None`` declared, no boundary/type analysis ran, so the
    columns stay unprofiled and the dtype_map is empty — previously a *silent*
    loss (W3 review, LOW). The plan must record the exact advisory note.
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    profile = make_sheet_profile(name="Sheet1", max_row=5, max_col=3)
    plan = build_read_plan(profile, options)

    assert plan.header is None
    assert plan.dtype_map == {}
    assert "headerless sheet: dtype inference skipped" in plan.notes


def test_detection_fallback_plan_has_no_headerless_note() -> None:
    """The header=0 detection *fallback* is not headerless — no L6 note.

    Only an explicit ``header_row=None`` declaration skips dtype inference by
    design; the fallback path still assumes a header, so the advisory would
    be wrong there.
    """

    profile = make_sheet_profile(name="Sheet1", max_row=5, max_col=3)
    plan = build_read_plan(profile)
    assert plan.header == 0
    assert "headerless sheet: dtype inference skipped" not in plan.notes


def test_manual_header_override_plan_has_no_headerless_note() -> None:
    """A manual header_row=N override is not headerless — no L6 note [D2]."""

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=1)}
    )
    profile = make_sheet_profile(name="Sheet1", max_row=5, max_col=3)
    plan = build_read_plan(profile, options)
    assert plan.header == 0
    assert "headerless sheet: dtype inference skipped" not in plan.notes


def test_heuristic_header_provenance_preserved() -> None:
    """A heuristic-detected header keeps provenance=heuristic (not clobbered)."""

    profile = make_sheet_profile(
        name="Sheet1", header_row=4, header_provenance="heuristic"
    )
    build_read_plan(profile)
    assert profile.header_provenance == "heuristic"


# -- skip-row sanitation (issue #9) ---------------------------------------


def test_stray_skip_row_at_or_above_header_is_ignored() -> None:
    """A skip_row at/above the header is discarded so the header still binds.

    Regression for issue #9: a stray skip_row of 2 (header is row 4) must not
    be folded into skiprows, which would otherwise shift the post-skip frame
    and break header normalization. A warning is recorded.
    """

    warnings: list[str] = []
    profile = make_sheet_profile(
        name="Sheet1",
        header_row=4,
        data_start_row=5,
        data_end_row=11,
        skip_rows=[2, 8],  # 2 is stray (at/above header); 8 is interior
    )
    plan = build_read_plan(profile, warnings=warnings)

    # Leading rows 1-3 absorbed; only the genuine interior skip (8 -> 7) added.
    assert plan.skiprows == [0, 1, 2, 7]
    assert plan.header == 0
    assert any("skip_row 2" in w for w in warnings)


def test_stray_skip_row_above_data_start_is_ignored() -> None:
    """A skip_row below the header but above data_start is discarded (issue #9)."""

    warnings: list[str] = []
    profile = make_sheet_profile(
        name="Sheet1",
        header_row=2,
        data_start_row=5,
        data_end_row=8,
        skip_rows=[3, 6],  # 3 is between header and data_start; 6 is interior
    )
    plan = build_read_plan(profile, warnings=warnings)

    # Leading row 1 absorbed (0); interior 6 -> 5. Stray 3 dropped.
    assert plan.skiprows == [0, 5]
    assert any("skip_row 3" in w for w in warnings)


# -- analyzer integration -------------------------------------------------


def test_aggregator_attaches_plan_to_tabular_sheet() -> None:
    """The analyzer attaches a plan to tabular sheets."""

    profile = make_sheet_profile(
        name="Sheet1", is_tabular_candidate=True, max_row=6, max_col=4
    )
    ctx = make_context(sheets=[profile])
    PlanAggregator().analyze(ctx)
    assert profile.read_plan is not None
    assert profile.read_plan.sheet_name == "Sheet1"


def test_aggregator_skips_non_tabular_sheet() -> None:
    """Non-tabular sheets are excluded from loading and get no plan (spec §9)."""

    profile = make_sheet_profile(name="README", is_tabular_candidate=False)
    ctx = make_context(sheets=[profile])
    PlanAggregator().analyze(ctx)
    assert profile.read_plan is None


def test_aggregator_name() -> None:
    """The analyzer reports a stable identifier."""

    assert PlanAggregator().name() == "plan_aggregator"
