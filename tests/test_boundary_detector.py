"""Boundary Detector tests (spec §4.5, §7.2, Phase 3) [D2][D4][D6].

Two layers, mirroring the Header Locator suite:

1.  **Loader-backed fixture tests** run the wired
    ``SheetEnumerator -> HeaderLocator -> BoundaryDetector`` chain over the real
    data-mode workbook for each relevant corpus sample and assert the §7.2
    boundary fields (1-based ``data_start_row``/``data_end_row``,
    ``data_left_col``/``data_right_col``, ``skip_rows``).
2.  **Isolated partial-context tests** drive the analyzer through ``conftest``
    synthesis with a fake data-mode loader so the blank-run terminator, the
    keyword/low-density classification, the column-span derivation, and the
    [D2] ``skip_rows`` overrides can be exercised without touching disk.

Coordinates here are **openpyxl 1-based** (the inspection domain [D1]); the
single 1-based -> 0-based conversion is the aggregator's job, covered by
``test_aggregator.py``.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

from excel_inspector import (
    InspectionOptions,
    Loader,
    SheetOverride,
)
from excel_inspector.analyzers.boundary_detector import (
    BoundaryDetector,
    _header_column_span,
    _leading_label,
    _matches_keyword,
    _span_density,
)
from excel_inspector.analyzers.header_locator import HeaderLocator
from excel_inspector.analyzers.merge_analyzer import MergeScanner
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.context import InspectionContext
from excel_inspector.models import WorkbookProfile

from conftest import make_context, make_sheet_profile  # type: ignore[import-not-found]


def _run_on(
    path: Path, options: InspectionOptions | None = None
) -> InspectionContext:
    """Enumerate -> scan merges -> locate headers -> detect boundaries.

    Mirrors the real ``inspect()`` topology up to the Boundary Detector: the
    Merge Scanner (plan v2 Task 11.1) collects merge spans *before* boundary
    detection so the merged-header virtual fill is exercised exactly as in
    the pipeline. Sheets without merges get an empty span list — bitwise the
    pre-11a behavior.
    """

    context = InspectionContext(
        options=options or InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)
        MergeScanner().analyze(context)
        HeaderLocator().analyze(context)
        return BoundaryDetector().analyze(context)


# ---------------------------------------------------------------------------
# Identity
# ---------------------------------------------------------------------------


def test_name() -> None:
    """The analyzer reports a stable identifier."""

    assert BoundaryDetector().name() == "boundary_detector"


# ---------------------------------------------------------------------------
# Loader-backed fixture boundaries (spec §7.2)
# ---------------------------------------------------------------------------


def test_offset_plus_subtotals_boundaries(fixture_path) -> None:
    """offset_plus_subtotals: interior subtotals + trailing total (§7.2) [D1].

    Header row 4; data rows 5-7 and 9-11; subtotal rows 8 & 12 ('소계') and the
    grand-total row 13 ('합계') are ``skip_rows``. data_end_row is the last real
    data row (11), so the trailing totals at 12-13 sit beyond it. The table is
    full-width A-D, so column boundaries stay ``None`` (no usecols restriction).
    """

    sheet = _run_on(
        fixture_path("offset_plus_subtotals")
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 4
    assert sheet.data_start_row == 5
    assert sheet.data_end_row == 11
    assert sheet.skip_rows == [8, 12, 13]
    assert sheet.data_left_col is None
    assert sheet.data_right_col is None


def test_offset_plus_subtotals_records_skip_labels(fixture_path) -> None:
    """Issue #2: each excluded subtotal/total row's label is recorded (raw case).

    The detector keys ``subtotal_skip_labels`` by 1-based sheet row so the
    aggregator can name every dropped row in its "no silent loss" note. The
    labels keep their original characters (here Korean '소계'/'합계'), and only
    the non-blank skip rows are recorded.
    """

    sheet = _run_on(
        fixture_path("offset_plus_subtotals")
    ).workbook_profile.sheets[0]
    assert sheet.subtotal_skip_labels == {8: "소계", 12: "소계", 13: "합계"}


def test_blank_run_terminates_boundaries(fixture_path) -> None:
    """blank_run_terminates: a 2-row blank run ends the block (§7.2).

    Header row 1; data rows 2-5; the blank rows 6-7 terminate the table so the
    noise block at rows 9-10 is never reached. No interior subtotal -> empty
    skip_rows.
    """

    sheet = _run_on(
        fixture_path("blank_run_terminates")
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 5
    assert sheet.skip_rows == []


def test_left_margin_cols_column_boundaries(fixture_path) -> None:
    """left_margin_cols: a left filler column -> table at C-E (§7.2).

    The header row's longest contiguous run is columns C-E (the filler text in
    column A is broken from the table by the empty column B), so
    data_left_col=3, data_right_col=5 (usecols 'C:E'). Density uses the C-E span
    so the data rows 2-7 are all full data rows.
    """

    sheet = _run_on(
        fixture_path("left_margin_cols")
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_left_col == 3
    assert sheet.data_right_col == 5
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 7
    assert sheet.skip_rows == []


def test_left_margin_with_subtotal_keyword_anchors_at_left_col(
    fixture_path,
) -> None:
    """L7 (plan v2 Phase 13 Step 3): keyword scan starts at data_left_col.

    The variant fixture's left margin is NON-empty on the subtotal row (the
    A5 note '중간 점검용 행'), so a sheet-column-A keyword scan reads the
    margin note instead of the table's '소계' at C5 and silently leaks the
    subtotal into the data — exactly the trap the plain ``left_margin_cols``
    fixture (empty margin on body rows) cannot expose. The subtotal row's
    span density is 2/3 (above the low-density threshold), so the keyword
    rule is the only line of defense: ``skip_rows == [5]`` proves the scan
    anchored at the table's left column.
    """

    sheet = _run_on(
        fixture_path("left_margin_with_subtotal")
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_left_col == 3
    assert sheet.data_right_col == 5
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 7
    assert sheet.skip_rows == [5]  # the '소계' row, caught from column C
    # Issue #2: the recorded label is the table's own '소계' at C5, never the
    # left-margin note at A5 — the no-silent-loss note anchors at data_left_col.
    assert sheet.subtotal_skip_labels == {5: "소계"}


def test_header_only_has_no_data_region(fixture_path) -> None:
    """header_only: a header with zero data rows -> data_start_row None (§9)."""

    sheet = _run_on(fixture_path("header_only")).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_start_row is None
    assert sheet.data_end_row is None
    assert sheet.skip_rows == []


def test_header_simple_full_width_no_usecols(fixture_path) -> None:
    """header_simple: a left-anchored full-width table reports no usecols.

    The table spans A-D (the full used width), so per spec §4.5 the column
    boundaries are left ``None`` (read all columns) rather than 'A:D'.
    """

    sheet = _run_on(fixture_path("header_simple")).workbook_profile.sheets[0]
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 6
    assert sheet.data_left_col is None
    assert sheet.data_right_col is None


def test_wide_sparse_timeseries_sparse_date_rows_preserved(fixture_path) -> None:
    """issue #22: sparse date-axis observations are data, not skip rows."""

    sheet = _run_on(
        fixture_path("wide_sparse_timeseries")
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 4
    assert sheet.data_start_row == 5
    assert sheet.data_end_row == 16
    assert sheet.skip_rows == []


def test_no_header_sheet_left_untouched(fixture_path) -> None:
    """no_header: header estimation failed -> no anchor, boundaries stay None."""

    sheet = _run_on(fixture_path("no_header")).workbook_profile.sheets[0]
    assert sheet.header_row is None
    assert sheet.data_start_row is None
    assert sheet.data_end_row is None
    assert sheet.skip_rows == []


def test_merged_header_boundary_bridged(fixture_path) -> None:
    """merged_header: the A1:B1 merge no longer blocks the boundary (Task 11.1).

    Phase 11a characterization flip (plan v2 §5 Task 11.1 Step 3): the Merge
    Scanner's spans virtually fill B1, restoring the contiguous header span
    A-C, so the data region resolves to the FIXTURES-documented values —
    data_start_row=2, data_end_row=5 (1-based [D1]). The A6:A7 body-merge
    demo block (rows 6-7) is NOT part of the table body: a fully merge-grouped
    row group trailing a merge-free body is excluded, with a visible warning
    (fixture-contradiction guard: the answer is 5, never 7). The bridged table
    is full-width A-C, so the column boundaries stay ``None`` (no usecols
    restriction, spec §4.5), and the old 'discarded pending merge analysis'
    deferral warning is gone.
    """

    ctx = _run_on(fixture_path("merged_header"))
    sheet = ctx.workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 5          # NOT 7 — the demo block is excluded
    assert sheet.data_left_col is None      # full-width A-C -> no usecols
    assert sheet.data_right_col is None
    assert sheet.skip_rows == []
    # The old v1 deferral is resolved — no 'pending merge analysis' warning.
    assert not any("pending merge analysis" in w for w in ctx.warnings)
    # The demo-block exclusion is visible, never silent (rows 6-7 named).
    assert any(
        "trailing merged-row group (rows 6-7)" in w and "Sheet1" in w
        for w in ctx.warnings
    )


# ---------------------------------------------------------------------------
# Override path [D2]: skip_rows_add / skip_rows_remove
# ---------------------------------------------------------------------------


def test_skip_rows_remove_override(fixture_path) -> None:
    """skip_rows_remove drops a heuristic skip row (e.g. keep a subtotal) [D2].

    The override also pins ``header_row=4`` so the registered
    :class:`SheetOverride` does not get read as a headerless declaration (a
    SheetOverride with the default ``header_row=None`` would otherwise mark the
    sheet headerless and the boundary scan would be skipped — see
    ``options.has_header_override``).
    """

    options = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(header_row=4, skip_rows_remove=[8])
        }
    )
    sheet = _run_on(
        fixture_path("offset_plus_subtotals"), options
    ).workbook_profile.sheets[0]
    # Heuristic skip_rows were [8, 12, 13]; removing 8 leaves [12, 13].
    assert sheet.skip_rows == [12, 13]


def test_skip_rows_add_override(fixture_path) -> None:
    """skip_rows_add injects an extra skip row beyond the heuristic ones [D2].

    ``header_row=4`` is pinned for the same reason as
    :func:`test_skip_rows_remove_override`.
    """

    options = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(header_row=4, skip_rows_add=[6])
        }
    )
    sheet = _run_on(
        fixture_path("offset_plus_subtotals"), options
    ).workbook_profile.sheets[0]
    # 6 added to the heuristic [8, 12, 13]; result sorted & de-duplicated.
    assert sheet.skip_rows == [6, 8, 12, 13]


# ---------------------------------------------------------------------------
# Isolated partial-context tests with a fake data-mode loader
# ---------------------------------------------------------------------------


class _FakeWorksheet:
    """A read_only-style worksheet returning fixed 1-based-aligned rows."""

    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows

    def iter_rows(self, *, min_row, max_row, values_only):  # noqa: ANN001, ANN202
        end = len(self._rows) if max_row is None else min(max_row, len(self._rows))
        for r in range(min_row - 1, end):
            yield tuple(self._rows[r])


class _FakeWorkbook:
    """A workbook mapping sheet names to :class:`_FakeWorksheet`."""

    def __init__(self, sheets: dict[str, list[list[object]]]) -> None:
        self._sheets = {n: _FakeWorksheet(r) for n, r in sheets.items()}

    def __getitem__(self, name: str) -> _FakeWorksheet:
        return self._sheets[name]


class _FakeLoader:
    """A loader stub exposing only :meth:`data_workbook`."""

    def __init__(self, sheets: dict[str, list[list[object]]]) -> None:
        self._wb = _FakeWorkbook(sheets)

    def data_workbook(self) -> _FakeWorkbook:
        return self._wb


def test_isolated_blank_run_terminates_block() -> None:
    """Two consecutive blank rows end the block; rows beyond are not scanned."""

    rows = [
        ["name", "qty", "price"],  # row 1 header
        ["A", 1, 1.0],  # row 2
        ["B", 2, 2.0],  # row 3
        [None, None, None],  # row 4 blank
        [None, None, None],  # row 5 blank -> BLANK_RUN terminator
        ["noise", 9, 9.0],  # row 6 beyond the run (ignored)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 3
    assert profile.skip_rows == []


def test_isolated_blank_leading_header_key_column_kept() -> None:
    """Issue #16: a blank leading header cell can still be a key column.

    The header span is B:C because A1 is empty, but A2:A4 holds the time axis
    for the same data rows. The boundary detector must widen the table to
    A:C; since that is the full used width, it reports no explicit usecols
    restriction (``data_left_col``/``data_right_col`` stay ``None``).
    """

    rows = [
        [None, "USD", "JPY"],  # row 1 header; A1 is intentionally empty
        [_dt.datetime(1992, 1, 1), 2.2, 4.4],
        [_dt.datetime(1993, 1, 1), 3.3, 6.6],
        [_dt.datetime(1994, 1, 1), 4.4, 8.8],
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.data_left_col is None
    assert profile.data_right_col is None
    assert profile.skip_rows == []


def test_isolated_keyword_and_trailing_total() -> None:
    """A '소계' keyword row is skipped; a trailing total sits beyond data_end."""

    rows = [
        ["dept", "item", "amount"],  # row 1 header
        ["영업", "교통비", 100],  # row 2 data
        ["영업", "식대", 200],  # row 3 data
        ["소계", None, 300],  # row 4 subtotal (keyword)
        ["합계", None, 600],  # row 5 grand total (keyword, trailing)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 3  # last *real* data row
    assert profile.skip_rows == [4, 5]


def test_isolated_low_density_row_is_skipped() -> None:
    """A single-column (low-density) interior row is a skip candidate."""

    rows = [
        ["a", "b", "c", "d"],  # row 1 header (4 cols)
        ["w", 1, 2, 3],  # row 2 data (density 1.0)
        ["section", None, None, None],  # row 3 single col -> skip
        ["x", 4, 5, 6],  # row 4 data
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=4
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == [3]


def test_isolated_two_column_keyvalue_rows_preserved() -> None:
    """A 2-column key-value table's single-filled rows are NOT subtotals (MEDIUM #5).

    The 'non_empty == 1' subtotal rule applies only to tables >= 3 columns wide.
    Here a 2-column table whose every data row has only the value column filled
    (key column intentionally blank) must keep ALL rows as data, not skips.
    """

    rows = [
        ["key", "value"],  # row 1 header (2 cols)
        [None, 10],  # row 2 (single column filled, but width 2 -> data)
        [None, 20],  # row 3
        [None, 30],  # row 4
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=2
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == []


def test_isolated_one_column_rows_preserved() -> None:
    """A 1-column table's rows are all data (no single-column subtotal rule)."""

    rows = [
        ["value"],  # row 1 header
        ["a"],  # row 2
        ["b"],  # row 3
        ["c"],  # row 4
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=1
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == []


def test_isolated_three_column_single_filled_is_subtotal() -> None:
    """A >=3-column table keeps the single-filled subtotal rule (MEDIUM #5).

    The complement of the narrow-table guard: in a 3-column table a row with
    only one column filled is still a subtotal/separator candidate.
    """

    rows = [
        ["a", "b", "c"],  # row 1 header (3 cols)
        ["x", 1, 2],  # row 2 data
        ["section", None, None],  # row 3 single col -> skip (width 3)
        ["y", 3, 4],  # row 4 data
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == [3]


def test_isolated_wide_sparse_date_axis_rows_preserved() -> None:
    """Wide date-axis rows with few values are not low-density skips."""

    rows = [
        ["Period"] + [f"Q:TS:{i:03d}" for i in range(1, 12)],
        [_dt.datetime(2020, 3, 31), 1] + [None] * 10,
        [_dt.datetime(2020, 6, 30), None, 2] + [None] * 9,
        [_dt.datetime(2020, 9, 30), None, None, 3] + [None] * 8,
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=12
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == []


def test_isolated_single_interior_blank_recorded() -> None:
    """A single interior blank row is recorded in skip_rows (MEDIUM #4).

    One blank row between data rows (below the BLANK_RUN terminator threshold)
    must be captured as a skip so it never leaks into the loaded frame as an
    all-NaN row; the data region spans across it.
    """

    rows = [
        ["a", "b", "c"],  # row 1 header
        ["x", 1, 2],  # row 2 data
        [None, None, None],  # row 3 single interior blank
        ["y", 3, 4],  # row 4 data
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == [3]


def test_isolated_keyword_false_positive_data_rows_preserved() -> None:
    """Data labels embedding a keyword are not skipped (MEDIUM #6).

    '통계청' and '회계팀' contain '계' but must remain data rows; '소계'/'합계'
    leading-label rows are still skipped. This is the end-to-end boundary-level
    guard for the corrected leading-label keyword rule.
    """

    rows = [
        ["dept", "item", "amount"],  # row 1 header
        ["통계청", "교통비", 100],  # row 2 data (contains '계' but != '계')
        ["회계팀", "식대", 200],  # row 3 data (contains '계' but != '계')
        ["소계", None, 300],  # row 4 subtotal (leading '소계')
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 3  # '소계' at row 4 is the only skip
    assert profile.skip_rows == [4]


def test_isolated_custom_skip_keyword_override() -> None:
    """InspectionOptions.skip_keywords replaces the default keyword set [D2]."""

    rows = [
        ["region", "value"],  # row 1 header
        ["North", 10],  # row 2 data
        ["MARKER", 20],  # row 3 -> skipped only under the custom keyword
        ["South", 30],  # row 4 data
    ]
    options = InspectionOptions(skip_keywords=["marker"])
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=2
    )
    ctx = make_context(
        options=options, sheets=[profile], loader=_FakeLoader({"S": rows})
    )

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row == 2
    assert profile.data_end_row == 4
    assert profile.skip_rows == [3]


def test_isolated_non_tabular_sheet_is_skipped() -> None:
    """Non-tabular sheets are never scanned for boundaries (spec §9)."""

    profile = make_sheet_profile(
        name="README", is_tabular_candidate=False, header_row=1, max_row=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"README": []}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row is None
    assert profile.data_end_row is None


def test_isolated_no_header_sheet_is_skipped() -> None:
    """A sheet with header_row=None has no anchor and is left untouched."""

    profile = make_sheet_profile(name="S", header_row=None, max_row=5, max_col=3)
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": []}))

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row is None
    assert profile.data_end_row is None


def test_no_loader_warns_and_skips() -> None:
    """With no loader the analyzer warns and leaves boundary fields untouched."""

    profile = make_sheet_profile(name="S", header_row=1, max_row=5, max_col=3)
    ctx = make_context(sheets=[profile], loader=None)

    BoundaryDetector().analyze(ctx)
    assert profile.data_start_row is None
    assert any("boundary_detector" in w for w in ctx.warnings)


# ---------------------------------------------------------------------------
# Helper unit tests (§7.2 components)
# ---------------------------------------------------------------------------


def test_header_column_span_longest_contiguous_run() -> None:
    """The longest contiguous non-empty run fixes the 1-based span (§7.2)."""

    # A filled, B empty, C-E filled -> longest run is C:E (3 > 1).
    assert _header_column_span(["x", None, "a", "b", "c"], 5) == (3, 5)
    # Full contiguous run -> whole width.
    assert _header_column_span(["a", "b", "c"], 3) == (1, 3)
    # Empty header -> no span.
    assert _header_column_span([None, None], 2) == (None, None)


def test_span_density_over_table_columns() -> None:
    """Density is computed over the table column span only (§7.2)."""

    # Span C:E (1-based 3..5); only those columns count.
    row = ["filler", None, "a", 1, 2]
    density, non_empty = _span_density(row, 3, 5)
    assert non_empty == 3
    assert density == 1.0
    # A single populated cell in the span -> low density.
    row2 = ["filler", None, "a", None, None]
    density2, non_empty2 = _span_density(row2, 3, 5)
    assert non_empty2 == 1
    assert abs(density2 - 1 / 3) < 1e-9


def test_matches_keyword_leading_label_startswith() -> None:
    """Keyword matching anchors to the leading label via startswith (§7.2).

    Multi-character keywords match the row's leading (first non-empty) label by
    case-insensitive ``startswith``; the single-character ``"계"`` matches only
    on exact equality. This is the corrected MEDIUM #6 contract: no arbitrary
    substring matching, so data labels do not false-match.
    """

    # Leading label exactly equals / starts with a multi-char keyword.
    assert _matches_keyword(["소계", None, 300], ["합계", "소계"]) is True
    assert _matches_keyword(["소계 합산", None, 300], ["소계"]) is True
    # Case-insensitive startswith on the leading label.
    assert _matches_keyword(["Total", 1], ["total"]) is True
    assert _matches_keyword(["Grand Total", 1], ["grand total"]) is True
    # A normal data label is not a keyword.
    assert _matches_keyword(["영업", "교통비", 100], ["합계", "소계"]) is False
    # No leading label (empty / numeric-first row) never matches.
    assert _matches_keyword([None, None], ["소계"]) is False
    assert _matches_keyword([100, "소계"], ["소계"]) is False


def test_matches_keyword_no_substring_false_positives() -> None:
    """Data labels that merely *contain* a keyword do not match (MEDIUM #6).

    ``통계청`` / ``회계팀`` contain ``계``; ``Total Wine`` contains ``Total``.
    Under the corrected leading-label rules none of these are skip rows: the
    single-char ``계`` needs exact equality, and ``Total`` only matches a label
    that *starts with* ``total`` — ``Total Wine`` does (a label genuinely
    leading with the keyword), but ``통계청`` / ``회계팀`` never do.
    """

    keywords = ["합계", "소계", "총계", "계", "Total", "Subtotal", "Grand Total"]
    # Single-char '계' must be exact: these contain it but do not equal it.
    assert _matches_keyword(["통계청", 1, 2], keywords) is False
    assert _matches_keyword(["회계팀", 1, 2], keywords) is False
    # Exact single-char '계' still matches.
    assert _matches_keyword(["계", None, 99], keywords) is True
    # 'Total Wine' starts with 'Total' -> a leading-label match is expected;
    # the false-positive guard is for labels that merely *embed* the keyword.
    assert _matches_keyword(["Wine Total", 1], keywords) is False


def test_leading_label_starts_at_left_col() -> None:
    """L7: the leading-label scan is anchored at the table's left_col [D1]."""

    row = [None, "여백 메모", "소계", 350]
    # Default (left_col=1, the pre-L7 / full-width behavior): the first
    # populated cell anywhere in the row wins.
    assert _leading_label(row) == "여백 메모"
    # Anchored at the table's left column (C=3): the margin is invisible.
    assert _leading_label(row, left_col=3) == "소계"
    # Anchored past the last populated cell -> no label at all.
    assert _leading_label(row, left_col=5) is None
    # A numeric first cell inside the span is data, not a label.
    assert _leading_label(row, left_col=4) is None


def test_matches_keyword_scans_from_table_left_col() -> None:
    """L7 (plan v2 Phase 13 Step 3): keywords anchor at data_left_col.

    Two failure modes are pinned: (a) a non-keyword margin note must not
    *shadow* the table's own '소계' label, and (b) a keyword-looking margin
    note must not *fake* a subtotal for a normal data row.
    """

    # (a) margin note shadows the real subtotal label for an A-column scan.
    subtotal_row = ["중간 점검용 행", None, "소계", None, 7.5]
    assert _matches_keyword(subtotal_row, ["소계"]) is False  # pre-L7 miss
    assert _matches_keyword(subtotal_row, ["소계"], left_col=3) is True

    # (b) a keyword in the margin must not flag a genuine data row.
    data_row = ["소계", None, "A-1", 10, 1.5]
    assert _matches_keyword(data_row, ["소계"], left_col=3) is False


# ---------------------------------------------------------------------------
# Phase 10b-1: block-local _detect_block core (plan v2 Task 10.2 Step 1)
# ---------------------------------------------------------------------------


def test_detect_block_is_profile_local() -> None:
    """_detect_block accumulates on BlockBoundary only — guard 3 (plan v2).

    The shared profile's boundary fields stay untouched and no warning reaches
    the context; the returned block-local result carries the outcome (here the
    whole-sheet window, so the values match what the sheet-level applier would
    write).
    """

    rows = [
        ["dept", "item", "amount"],  # row 1 header
        ["영업", "교통비", 100],  # row 2 data
        ["소계", None, 100],  # row 3 subtotal (keyword)
        ["관리", "식대", 200],  # row 4 data
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert result.data_start_row == 2
    assert result.data_end_row == 4
    assert result.skip_rows == [3]
    assert result.data_left_col is None and result.data_right_col is None
    # Guard 3: nothing leaked onto the shared profile or the context.
    assert profile.data_start_row is None
    assert profile.data_end_row is None
    assert profile.skip_rows == []
    assert ctx.warnings == []


def test_detect_block_merge_deferral_warning_is_block_local() -> None:
    """The unreliable-span deferral warning lives on the result (guard 3).

    A merge-narrowed header (interior gap) discards the column boundaries; the
    deferral warning must be accumulated block-locally and forwarded to the
    context only by the sheet-level applier, never by the core itself.
    """

    rows = [
        ["이름", None, "점수"],  # row 1 merged-style header (interior gap)
        ["김", None, 90],  # row 2
        ["이", None, 85],  # row 3
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert result.data_start_row is None and result.data_end_row is None
    assert result.data_left_col is None and result.data_right_col is None
    assert any("merge analysis" in w for w in result.warnings)
    assert ctx.warnings == []  # forwarded only by the applier (_detect)


def test_detect_block_row_window_bounds_the_scan() -> None:
    """A row_window stops the body scan at min(window_end, max_row).

    Without a window the scan crosses the single interior blank (row 4) and
    reaches the rows below; with ``row_window=(1, 3)`` the scan ends at row 3,
    so neither the blank nor the lower rows are ever visited.
    """

    rows = [
        ["a", "b", "c"],  # row 1 header
        ["x", 1, 2],  # row 2 data
        ["y", 3, 4],  # row 3 data
        [None, None, None],  # row 4 single interior blank (below BLANK_RUN)
        ["z", 5, 6],  # row 5 data
        ["w", 7, 8],  # row 6 data
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))
    detector = BoundaryDetector()

    whole = detector._detect_block(ctx, profile, 1)
    assert (whole.data_start_row, whole.data_end_row) == (2, 6)
    assert whole.skip_rows == [4]

    windowed = detector._detect_block(ctx, profile, 1, row_window=(1, 3))
    assert (windowed.data_start_row, windowed.data_end_row) == (2, 3)
    assert windowed.skip_rows == []
