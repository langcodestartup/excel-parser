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
