"""Block segmenter tests — Phase 10a (plan v2 §4 Task 10.1).

Covers three layers:

1.  ``split_row_bands`` unit contract (plan v2 §4 Task 10.1 test contract,
    1-based [D1]): bands split on ``BLANK_RUN`` (2) consecutive blank rows;
    a single blank row never splits.
2.  The :class:`BlockSegmenter` analyzer wiring: bands recorded on
    ``context.row_bands`` over the real fixture corpus.
3.  Pipeline warning behavior: ``inspect()``/``extract()`` on a multi-table
    sheet records a warning (existence only — the wording is NOT pinned, per
    the plan it changes once Task 10.2 extracts the blocks).
"""

from __future__ import annotations

from excel_inspector import InspectionOptions, Loader, extract, inspect
from excel_inspector.analyzers.block_segmenter import (
    BlockSegmenter,
    RowBand,
    split_row_bands,
)
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.context import InspectionContext
from excel_inspector.heuristics import BLANK_RUN
from excel_inspector.models import WorkbookProfile

# ---------------------------------------------------------------------------
# split_row_bands unit contract (plan v2 §4 Task 10.1)
# ---------------------------------------------------------------------------


def test_split_two_stacked_tables() -> None:
    rows = [("부서", "인원"), ("영업", 12), ("개발", 20), ("관리", 5),
            (None, None), (None, None),
            ("제품명", "단가"), ("키보드", 30000), ("마우스", 15000)]
    bands = split_row_bands(rows)
    assert bands == [RowBand(1, 4), RowBand(7, 9)]


def test_single_blank_does_not_split() -> None:
    rows = [("h",), ("a",), (None,), ("b",)]
    assert split_row_bands(rows) == [RowBand(1, 4)]


def test_default_blank_run_is_heuristic_constant() -> None:
    """The default splitting threshold is heuristics.BLANK_RUN (=2)."""

    assert BLANK_RUN == 2
    rows = [("a",), (None,), (None,), ("b",)]
    # Default (BLANK_RUN=2): splits.  blank_run=3: the 2-row gap stays inside.
    assert split_row_bands(rows) == [RowBand(1, 1), RowBand(4, 4)]
    assert split_row_bands(rows, blank_run=3) == [RowBand(1, 4)]


def test_empty_and_all_blank_inputs_yield_no_bands() -> None:
    assert split_row_bands([]) == []
    assert split_row_bands([(None, None), (None, None), (None, None)]) == []


def test_leading_and_trailing_blanks_are_trimmed() -> None:
    """Bands start and end on non-blank rows; edge blanks belong to no band."""

    rows = [(None,), ("a",), ("b",), (None,)]
    assert split_row_bands(rows) == [RowBand(2, 3)]
    # A trailing blank run beyond the threshold adds no extra band either.
    rows = [(None,), (None,), ("a",), (None,), (None,)]
    assert split_row_bands(rows) == [RowBand(3, 3)]


def test_empty_string_cells_count_as_blank() -> None:
    """'' is empty, matching the boundary detector's emptiness rule."""

    rows = [("a", 1), ("", ""), (None, ""), ("b", 2)]
    assert split_row_bands(rows) == [RowBand(1, 1), RowBand(4, 4)]


def test_rows_with_any_populated_cell_are_not_blank() -> None:
    """0 / False are populated cells (consistent with _is_empty semantics)."""

    rows = [(0,), (False,), (None,), (None,), ("x",)]
    assert split_row_bands(rows) == [RowBand(1, 2), RowBand(5, 5)]


def test_three_bands_top_down_order() -> None:
    rows = [("t1",), (None,), (None,),
            ("t2",), ("d",), (None,), (None,),
            ("t3",)]
    assert split_row_bands(rows) == [
        RowBand(1, 1), RowBand(4, 5), RowBand(8, 8),
    ]


def test_blank_run_below_one_is_rejected() -> None:
    import pytest

    with pytest.raises(ValueError):
        split_row_bands([("a",)], blank_run=0)


# ---------------------------------------------------------------------------
# BlockSegmenter analyzer over the real fixture corpus
# ---------------------------------------------------------------------------


def _run_segmenter(path) -> InspectionContext:
    """Run SheetEnumerator + BlockSegmenter over a real fixture."""

    context = InspectionContext(
        options=InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)
        BlockSegmenter().analyze(context)
    return context


def test_segmenter_records_bands_for_multi_table_stacked(fixture_path) -> None:
    """multi_table_stacked: bands [1..4] and [7..10] (FIXTURES coordinates)."""

    ctx = _run_segmenter(fixture_path("multi_table_stacked"))
    assert ctx.row_bands["Sheet1"] == [RowBand(1, 4), RowBand(7, 10)]
    # The second band carries a header candidate -> a warning was recorded.
    assert ctx.warnings


def test_segmenter_records_bands_for_stacked_uneven_width(fixture_path) -> None:
    """stacked_uneven_width: bands [1..4] and [7..10] (FIXTURES coordinates)."""

    ctx = _run_segmenter(fixture_path("stacked_uneven_width"))
    assert ctx.row_bands["Sheet1"] == [RowBand(1, 4), RowBand(7, 10)]
    assert ctx.warnings


def test_segmenter_single_band_sheet_warns_nothing(fixture_path) -> None:
    """header_simple: one band, no multi-block warning."""

    ctx = _run_segmenter(fixture_path("header_simple"))
    assert ctx.row_bands["Sheet1"] == [RowBand(1, 6)]
    assert ctx.warnings == []


def test_segmenter_skips_non_tabular_sheets(fixture_path) -> None:
    """mixed_sheets: the non-tabular README sheet is not segmented."""

    ctx = _run_segmenter(fixture_path("mixed_sheets"))
    assert "README" not in ctx.row_bands
    assert "Data" in ctx.row_bands


def test_segmenter_is_deterministic(fixture_path) -> None:
    """Two runs produce identical bands and warnings (warning-order contract)."""

    path = fixture_path("multi_table_stacked")
    first = _run_segmenter(path)
    second = _run_segmenter(path)
    assert first.row_bands == second.row_bands
    assert first.warnings == second.warnings


# ---------------------------------------------------------------------------
# Pipeline wiring: inspect()/extract() surface the multi-block warning
# ---------------------------------------------------------------------------


def test_inspect_multi_table_stacked_records_warning(fixture_path) -> None:
    """inspect() surfaces the multi-block suspicion via open_errors.

    Existence only — the wording is deliberately NOT pinned (plan v2 §4 Task
    10.1 Step 3: the message becomes an "extracted" notice after Task 10.2).
    """

    profile = inspect(fixture_path("multi_table_stacked"))
    assert profile.open_errors


def test_extract_multi_table_stacked_has_warning(fixture_path) -> None:
    """extract() on a stacked multi-table sheet yields non-empty warnings."""

    wr = extract(fixture_path("multi_table_stacked"))
    assert wr.warnings


def test_extract_stacked_uneven_width_has_warning(fixture_path) -> None:
    wr = extract(fixture_path("stacked_uneven_width"))
    assert wr.warnings


def test_extract_single_table_fixture_has_no_warning(fixture_path) -> None:
    """A plain single-table sheet gains no warnings from Phase 10a."""

    wr = extract(fixture_path("header_simple"))
    assert wr.warnings == []
