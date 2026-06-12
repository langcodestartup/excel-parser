"""Block-level override channel [D7] (issue #9).

Covers the BlockOverride model 3-state sentinel contract, the
options.resolve_block_overrides anchor resolution rules, the BlockAnalyzer
specificity chain, and the declared-headerless block read plan.
"""

from __future__ import annotations

from excel_inspector.models import BlockOverride, InspectionOptions, SheetOverride


def test_block_override_three_states() -> None:
    """_UNSET sentinel keeps int / explicit-None / unspecified distinct (HIGH #2)."""

    unspecified = BlockOverride()
    assert unspecified.header_row_set is False
    assert unspecified.header_row is None  # sentinel collapsed for readers

    forced = BlockOverride(header_row=7)
    assert forced.header_row_set is True
    assert forced.header_row == 7

    headerless = BlockOverride(header_row=None)
    assert headerless.header_row_set is True
    assert headerless.header_row is None


def test_sheet_override_carries_block_overrides() -> None:
    """SheetOverride.block_overrides defaults empty; keys are anchor rows."""

    bare = SheetOverride()
    assert bare.block_overrides == {}

    override = SheetOverride(
        block_overrides={19: BlockOverride(header_row=None)}
    )
    assert override.block_overrides[19].header_row_set is True
    # The block channel alone is NOT a sheet-level header declaration.
    assert override.header_row_set is False


def test_block_override_is_publicly_exported() -> None:
    """BlockOverride is part of the public API surface."""

    import excel_inspector

    assert excel_inspector.BlockOverride is BlockOverride
    assert "BlockOverride" in excel_inspector.__all__


from excel_inspector.analyzers.block_segmenter import RowBand
from excel_inspector.options import resolve_block_overrides

#: Bands mirroring the multi_table_stacked fixture: [1..4] and [7..10].
_BANDS = [RowBand(1, 4), RowBand(7, 10)]


def _opts(block_overrides: dict[int, BlockOverride]) -> InspectionOptions:
    return InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(block_overrides=block_overrides)}
    )


def test_resolver_no_options_or_no_block_overrides() -> None:
    assert resolve_block_overrides(None, "Sheet1", _BANDS) == ({}, [])
    assert resolve_block_overrides(
        InspectionOptions(), "Sheet1", _BANDS
    ) == ({}, [])


def test_resolver_anchor_maps_any_row_inside_band() -> None:
    """Anchor 9 (mid-band) resolves to band start 7; keyed by band_start_row."""

    resolved, warnings = resolve_block_overrides(
        _opts({9: BlockOverride(header_row=None)}), "Sheet1", _BANDS
    )
    assert warnings == []
    assert set(resolved) == {7}
    assert resolved[7].header_row is None and resolved[7].header_row_set


def test_resolver_anchor_outside_all_bands_warns() -> None:
    """Row 5 is the blank separator -> no band -> warned and ignored."""

    resolved, warnings = resolve_block_overrides(
        _opts({5: BlockOverride(header_row=None)}), "Sheet1", _BANDS
    )
    assert resolved == {}
    assert len(warnings) == 1
    assert "anchor row 5" in warnings[0]
    assert "no detected table band" in warnings[0]


def test_resolver_duplicate_anchors_lowest_wins() -> None:
    resolved, warnings = resolve_block_overrides(
        _opts(
            {
                7: BlockOverride(header_row=7),
                9: BlockOverride(header_row=None),
            }
        ),
        "Sheet1",
        _BANDS,
    )
    assert set(resolved) == {7}
    assert resolved[7].header_row == 7  # lowest anchor's override won
    assert len(warnings) == 1
    assert "anchor row 9" in warnings[0] and "anchor row 7" in warnings[0]


def test_resolver_int_header_outside_anchored_band_warns() -> None:
    resolved, warnings = resolve_block_overrides(
        _opts({7: BlockOverride(header_row=2)}), "Sheet1", _BANDS
    )
    assert resolved == {}
    assert len(warnings) == 1
    assert "header_row 2" in warnings[0]
    assert "outside the anchored band" in warnings[0]


def test_resolver_empty_override_warns_and_does_not_claim() -> None:
    """An empty BlockOverride is a no-op; a later valid anchor still claims."""

    resolved, warnings = resolve_block_overrides(
        _opts(
            {
                7: BlockOverride(),
                9: BlockOverride(header_row=None),
            }
        ),
        "Sheet1",
        _BANDS,
    )
    assert set(resolved) == {7}
    assert resolved[7].header_row is None  # anchor 9's override claimed the band
    assert len(warnings) == 1
    assert "no override field specified" in warnings[0]


# ---------------------------------------------------------------------------
# Task 3 – PlanAggregator declared-headerless conversion path [D7]
# ---------------------------------------------------------------------------

from excel_inspector.aggregator import build_block_read_plan
from excel_inspector.models import TableBlock
from tests.conftest import make_sheet_profile


def _headerless_block() -> TableBlock:
    """A declared-headerless block on band rows 17-21 (design doc §6)."""

    return TableBlock(
        block_index=2,
        band_start_row=17,
        band_end_row=21,
        header_row=None,
        header_confidence=1.0,
        header_provenance="manual",
        data_start_row=17,
        data_end_row=21,
        data_left_col=None,
        data_right_col=None,
        skip_rows=[],
        columns=[],
        read_plan=None,
        subtotal_skip_labels={},
    )


def test_declared_headerless_block_plan() -> None:
    """header=None, rows 1-16 absorbed into skiprows, nrows = band length."""

    profile = make_sheet_profile(name="Sheet1", max_row=21, max_col=3)
    plan = build_block_read_plan(
        profile, _headerless_block(), None, None, band_scoped=True
    )
    assert plan.header is None
    assert plan.skiprows == list(range(16))  # 1-based rows 1-16 -> 0-based [D1]
    assert plan.nrows == 5
    assert plan.usecols is None
    assert plan.dtype_map == {}
    assert "headerless sheet: dtype inference skipped" in plan.notes


def test_sheet_level_headerless_plan_unchanged() -> None:
    """The sheet-level headerless path never sets data_start_row -> no
    skiprows absorption; the [D7] rule must not disturb it."""

    from excel_inspector.aggregator import build_read_plan

    profile = make_sheet_profile(name="Sheet1", max_row=10, max_col=3)
    opts = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    plan = build_read_plan(profile, opts, None)
    assert plan.header is None
    assert plan.skiprows == []
    assert plan.nrows is None


# ---------------------------------------------------------------------------
# Task 4 – BlockAnalyzer specificity chain + headerless block creation [D7]
# ---------------------------------------------------------------------------

from excel_inspector import inspect


def test_block_int_override_beats_sheet_int_for_same_band(fixture_path) -> None:
    """Sheet header_row=8 and a block override both target band [7..10];
    the block override (header_row=7) wins, with a conflict warning."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                header_row=8,
                block_overrides={9: BlockOverride(header_row=7)},
            )
        }
    )
    profile = inspect(fixture_path("multi_table_stacked"), opts)
    sheet = profile.sheets[0]
    b1, b2 = sheet.blocks
    assert b1.header_provenance == "heuristic"  # band 1 untouched
    assert b2.header_row == 7
    assert b2.header_provenance == "manual"
    assert any(
        "the block override wins" in w for w in profile.open_errors
    )


def test_block_headerless_override_creates_manual_band_block(fixture_path) -> None:
    """BlockOverride(header_row=None) on band [7..10]: the band becomes a
    manual headerless block spanning the whole band."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={8: BlockOverride(header_row=None)}
            )
        }
    )
    sheet = inspect(fixture_path("multi_table_stacked"), opts).sheets[0]
    assert len(sheet.blocks) == 2
    b2 = sheet.blocks[1]
    assert b2.header_row is None
    assert b2.header_provenance == "manual"
    assert b2.header_confidence == 1.0
    assert (b2.data_start_row, b2.data_end_row) == (7, 10)
    assert (b2.data_left_col, b2.data_right_col) == (None, None)
    assert b2.columns == []
    # [D7] declared-headerless plan shape (Task 3).
    assert b2.read_plan is not None
    assert b2.read_plan.header is None
    assert b2.read_plan.skiprows == list(range(6))  # rows 1-6 absorbed
    assert b2.read_plan.nrows == 4


def test_sheet_headerless_plus_block_overrides_block_channel_wins(
    fixture_path,
) -> None:
    """Sheet-wide header_row=None + block_overrides: contradiction warned,
    per-band analysis proceeds (band 1 heuristic, band 2 headerless)."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                header_row=None,
                block_overrides={7: BlockOverride(header_row=None)},
            )
        }
    )
    profile = inspect(fixture_path("multi_table_stacked"), opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 2
    assert sheet.blocks[0].header_provenance == "heuristic"
    assert sheet.blocks[1].header_row is None
    assert sheet.blocks[1].header_provenance == "manual"
    assert any("contradicts block_overrides" in w for w in profile.open_errors)


def test_sheet_headerless_without_block_overrides_unchanged(fixture_path) -> None:
    """Without block_overrides the sheet-wide headerless gate is intact
    (guard 4): no per-band analysis, no blocks."""

    opts = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    sheet = inspect(fixture_path("multi_table_stacked"), opts).sheets[0]
    assert sheet.blocks == []


def test_block_overrides_on_single_band_sheet_warn_and_ignore(
    fixture_path,
) -> None:
    """header_simple is single-band: block_overrides are ignored with a
    pointer to the sheet-level channel; the mirror block is untouched."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={1: BlockOverride(header_row=None)}
            )
        }
    )
    profile = inspect(fixture_path("header_simple"), opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 1
    assert sheet.blocks[0].header_row == 1  # heuristic mirror intact
    assert any(
        "single-band sheet" in w and "sheet-level SheetOverride" in w
        for w in profile.open_errors
    )


def test_resolver_warnings_surface_through_inspect(fixture_path) -> None:
    """An anchor in the blank separator (row 5) surfaces its warning."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={5: BlockOverride(header_row=None)}
            )
        }
    )
    profile = inspect(fixture_path("multi_table_stacked"), opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 2  # both bands fall back to the heuristic
    assert all(b.header_provenance == "heuristic" for b in sheet.blocks)
    assert any(
        "anchor row 5" in w and "no detected table band" in w
        for w in profile.open_errors
    )
