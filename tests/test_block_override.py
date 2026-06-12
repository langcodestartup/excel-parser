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
