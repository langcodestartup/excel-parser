"""Header Locator tests (spec §4.3, §7.1, Phase 2) [D2][D4].

Two layers:

1.  **Loader-backed fixture tests** run the analyzer over the real data-mode
    workbook for each corpus sample and assert the §7.1 header estimate
    (header_row 1-based, confidence trend, ``needs_manual_header``).
2.  **Isolated partial-context tests** drive the analyzer (and its scoring
    helpers) through ``conftest`` synthesis with a fake loader so the override
    path, the threshold knob, and the non-tabular skip can be exercised without
    touching disk.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

from excel_inspector import (
    InspectionOptions,
    Loader,
    SheetOverride,
)
from excel_inspector.analyzers.header_locator import (
    HeaderLocator,
    _distinctness,
    _non_empty_string_ratio,
    _score_row,
    _type_consistency,
)
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.context import InspectionContext
from excel_inspector.models import WorkbookProfile

from conftest import make_context, make_sheet_profile  # type: ignore[import-not-found]


def _run_on(
    path: Path, options: InspectionOptions | None = None
) -> InspectionContext:
    """Enumerate then locate headers over a fixture; return the context."""

    context = InspectionContext(
        options=options or InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)
        return HeaderLocator().analyze(context)


# ---------------------------------------------------------------------------
# Identity
# ---------------------------------------------------------------------------


def test_name() -> None:
    """The analyzer reports a stable identifier."""

    assert HeaderLocator().name() == "header_locator"


# ---------------------------------------------------------------------------
# Loader-backed fixture estimates (spec §7.1)
# ---------------------------------------------------------------------------


def test_header_simple_row_one(fixture_path) -> None:
    """header_simple: an all-string row 1 above mixed data -> header_row=1."""

    sheet = _run_on(fixture_path("header_simple")).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.needs_manual_header is False
    assert sheet.header_provenance == "heuristic"
    assert sheet.is_multi_level_header is False
    assert sheet.header_confidence >= 0.5


def test_header_offset_row_four(fixture_path) -> None:
    """header_offset: a 3-row title block precedes the header -> header_row=4."""

    sheet = _run_on(fixture_path("header_offset")).workbook_profile.sheets[0]
    assert sheet.header_row == 4
    assert sheet.needs_manual_header is False
    assert sheet.header_provenance == "heuristic"
    assert sheet.header_confidence >= 0.5


def test_offset_plus_subtotals_row_four(fixture_path) -> None:
    """offset_plus_subtotals: title block + interior subtotals -> header_row=4."""

    sheet = _run_on(
        fixture_path("offset_plus_subtotals")
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 4
    assert sheet.needs_manual_header is False
    assert sheet.header_confidence >= 0.5


def test_no_header_needs_manual(fixture_path) -> None:
    """no_header: homogeneous numeric data, no string header -> manual needed."""

    ctx = _run_on(fixture_path("no_header"))
    sheet = ctx.workbook_profile.sheets[0]
    assert sheet.header_row is None
    assert sheet.header_confidence == 0.0
    assert sheet.needs_manual_header is True
    assert sheet.header_provenance == "heuristic"
    assert any("header_locator" in w for w in ctx.warnings)


def test_types_mixed_row_one(fixture_path) -> None:
    """types_mixed: string header over number/numeric_text/date/mixed columns."""

    sheet = _run_on(fixture_path("types_mixed")).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.needs_manual_header is False
    assert sheet.header_confidence >= 0.5


def test_offset_outscores_simple_confidence(fixture_path) -> None:
    """A clearly-delimited header (title + data) is at least as confident.

    Confidence *trend* check: the offset header sitting above well-typed data
    should score no lower than the simple header, both comfortably above the
    threshold.
    """

    simple = _run_on(fixture_path("header_simple")).workbook_profile.sheets[0]
    offset = _run_on(fixture_path("header_offset")).workbook_profile.sheets[0]
    assert simple.header_confidence >= 0.5
    assert offset.header_confidence >= simple.header_confidence


def test_non_tabular_sheet_is_skipped(fixture_path) -> None:
    """mixed_sheets: README (non-tabular) is left untouched by the locator."""

    ctx = _run_on(fixture_path("mixed_sheets"))
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    readme = sheets["README"]
    # The locator never ran on README: header fields stay at enumerator defaults.
    assert readme.header_row is None
    assert readme.needs_manual_header is False
    assert readme.header_provenance == "default"
    # The tabular Data sheet got a heuristic header.
    assert sheets["Data"].header_row == 1
    assert sheets["Data"].header_provenance == "heuristic"


# ---------------------------------------------------------------------------
# Override path [D2]
# ---------------------------------------------------------------------------


def test_override_skips_scoring_and_records_manual(fixture_path) -> None:
    """A header_row override [D2] bypasses scoring: manual + confidence 1.0.

    no_header would normally fail the heuristic; forcing header_row=1 must
    yield a confident manual estimate instead.
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=1)}
    )
    sheet = _run_on(fixture_path("no_header"), options).workbook_profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.header_confidence == 1.0
    assert sheet.header_provenance == "manual"
    assert sheet.needs_manual_header is False


def test_override_none_declares_no_header() -> None:
    """An explicit header_row=None override declares 'no header' (manual) [D2].

    Isolated: a fake loader proves scoring is skipped (it would raise if used).
    """

    class _ExplodingLoader:
        def data_workbook(self):  # noqa: ANN202 - test double
            raise AssertionError("scoring must be skipped under override")

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    profile = make_sheet_profile(name="Sheet1", max_row=5, max_col=3)
    ctx = make_context(
        options=options, sheets=[profile], loader=_ExplodingLoader()
    )

    HeaderLocator().analyze(ctx)
    assert profile.header_row is None
    assert profile.header_confidence == 1.0
    assert profile.header_provenance == "manual"
    assert profile.needs_manual_header is False


def test_dtype_only_override_preserves_heuristic_header(fixture_path) -> None:
    """A dtype_force-only SheetOverride must NOT suppress header detection.

    Regression for HIGH #2: before the SheetOverride sentinel, *any* registered
    SheetOverride read as a header override (header_row defaulted to None), so a
    user who set only ``dtype_force`` would have the sheet mis-declared
    headerless and the heuristic header locator skipped. With the sentinel, the
    override does not touch the header channel and the heuristic still runs.
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(dtype_force={"1": "string"})}
    )
    sheet = _run_on(
        fixture_path("header_simple"), options
    ).workbook_profile.sheets[0]
    # The heuristic header was preserved, not overridden away.
    assert sheet.header_row == 1
    assert sheet.header_provenance == "heuristic"
    assert sheet.needs_manual_header is False


def test_skip_rows_only_override_preserves_heuristic_header(fixture_path) -> None:
    """A skip_rows-only SheetOverride also preserves heuristic header (HIGH #2)."""

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(skip_rows_add=[6])}
    )
    sheet = _run_on(
        fixture_path("header_offset"), options
    ).workbook_profile.sheets[0]
    assert sheet.header_row == 4
    assert sheet.header_provenance == "heuristic"


# ---------------------------------------------------------------------------
# Isolated partial-context tests with a fake data-mode loader
# ---------------------------------------------------------------------------


class _FakeWorksheet:
    """A read_only-style worksheet returning fixed rows for iter_rows."""

    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows

    def iter_rows(self, *, min_row, max_row, values_only):  # noqa: ANN001, ANN202
        end = min(max_row, len(self._rows))
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


def test_isolated_offset_header_detected() -> None:
    """Synthetic title-then-header table -> header_row=4 via partial context."""

    rows = [
        ["보고서", None, None, None],
        ["작성일", None, None, None],
        ["단위", None, None, None],
        ["name", "qty", "price", "city"],
        ["A", 1, 1.0, "Seoul"],
        ["B", 2, 2.0, "Busan"],
        ["C", 3, 3.0, "Daegu"],
        ["D", 4, 4.0, "Incheon"],
        ["E", 5, 5.0, "Gwangju"],
    ]
    profile = make_sheet_profile(name="S", max_row=len(rows), max_col=4)
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    HeaderLocator().analyze(ctx)
    assert profile.header_row == 4
    assert profile.header_provenance == "heuristic"
    assert profile.header_confidence >= 0.5


def test_isolated_custom_threshold_forces_manual() -> None:
    """A high confidence threshold pushes an otherwise-good header to manual."""

    rows = [
        ["name", "qty", "price"],
        ["A", 1, 1.0],
        ["B", 2, 2.0],
        ["C", 3, 3.0],
    ]
    profile = make_sheet_profile(name="S", max_row=len(rows), max_col=3)
    options = InspectionOptions(header_confidence_threshold=0.99)
    ctx = make_context(
        options=options, sheets=[profile], loader=_FakeLoader({"S": rows})
    )

    HeaderLocator().analyze(ctx)
    assert profile.header_row is None
    assert profile.needs_manual_header is True
    assert profile.header_confidence == 0.0


def test_isolated_empty_sheet_needs_manual() -> None:
    """A sheet that yields no rows cannot be scored -> manual + warning."""

    profile = make_sheet_profile(name="S", max_row=0, max_col=0)
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": []}))

    HeaderLocator().analyze(ctx)
    assert profile.header_row is None
    assert profile.needs_manual_header is True
    assert any("header_locator" in w for w in ctx.warnings)


def test_no_loader_warns_and_skips() -> None:
    """With no loader the analyzer warns and leaves header fields untouched."""

    profile = make_sheet_profile(name="S", max_row=5, max_col=3)
    ctx = make_context(sheets=[profile], loader=None)

    HeaderLocator().analyze(ctx)
    assert profile.header_row is None
    assert any("header_locator" in w for w in ctx.warnings)


# ---------------------------------------------------------------------------
# Scoring-helper unit tests (§7.1 components)
# ---------------------------------------------------------------------------


def test_non_empty_string_ratio() -> None:
    """Only non-empty string cells count toward the ratio (§7.1)."""

    assert _non_empty_string_ratio(["a", "b", "c", "d"], 4) == 1.0
    assert _non_empty_string_ratio(["a", None, "", 3], 4) == 0.25
    assert _non_empty_string_ratio([1, 2, 3], 3) == 0.0
    assert _non_empty_string_ratio([], 0) == 0.0


def test_type_consistency_perfect_and_mixed() -> None:
    """Per-column type consistency averages the dominant-category fraction."""

    below = [[1, "x"], [2, "y"], [3, "z"]]
    assert _type_consistency(below, 2) == 1.0
    # Column 0 stays numeric (1.0); column 1 is 2/3 string -> mean 0.833...
    mixed = [[1, "x"], [2, "y"], [3, 4]]
    assert abs(_type_consistency(mixed, 2) - (1.0 + 2 / 3) / 2) < 1e-9
    assert _type_consistency([], 2) == 0.0


def test_distinctness_header_vs_data() -> None:
    """A string header over numeric data is highly distinct (type differs)."""

    candidate = ["name", "qty", "price"]
    below = [[1, 2, 3.0], [4, 5, 6.0]]
    # Every column's candidate category (string) differs from the numeric
    # dominant below, so each type_diff is 1.0; distinctness is high.
    assert _distinctness(candidate, below, 3) >= 0.5
    # A numeric "candidate" identical in shape to the data is not distinct.
    assert _distinctness([1, 2, 3.0], below, 3) < 0.5


def test_score_row_combines_components() -> None:
    """The composite score equals the §7.1 weighted sum of its parts.

    The consistency term carries the lookahead-evidence factor
    ``n_below / HEADER_LOOKAHEAD_ROWS`` (issue #8): two observed below rows
    earn 2/5 of the full consistency weight.
    """

    rows = [
        ["name", "qty", "price"],
        ["A", 1, 1.0],
        ["B", 2, 2.0],
    ]
    s = _score_row(0, rows, 3)
    expected = (
        0.5 * _non_empty_string_ratio(rows[0], 3)
        + 0.3 * _type_consistency(rows[1:3], 3) * (2 / 5)
        + 0.2 * _distinctness(rows[0], rows[1:3], 3)
    )
    assert abs(s - expected) < 1e-12


def test_date_cells_categorized_as_date() -> None:
    """datetime cells are treated as a 'date' category, not a string."""

    below = [
        [_dt.datetime(2026, 1, 1), 1],
        [_dt.datetime(2026, 1, 2), 2],
    ]
    # Column 0 is all dates -> fully consistent.
    assert _type_consistency(below, 2) == 1.0


# ---------------------------------------------------------------------------
# Phase 10b-1: row_window scoping of the _locate core (plan v2 Task 10.2 Step 1)
# ---------------------------------------------------------------------------


def test_locate_window_returns_absolute_header_row() -> None:
    """_locate maps the window-local best index to the absolute sheet row.

    Guard 2 (plan v2 Task 10.2): ``header_row = window_start + best_index``,
    so the header of a lower stacked band is located at its absolute 1-based
    position (7), not at a sheet-top-relative 1. Degenerate windows (beyond
    the content, or inverted) yield ``(None, 0.0)``.
    """

    rows = [
        ["부서", "인원", "예산"],  # row 1: table 1 header
        ["영업", 12, 100],  # rows 2-4: table 1 data
        ["개발", 20, 200],
        ["관리", 5, 300],
        [None, None, None],  # rows 5-6: blank separator
        [None, None, None],
        ["제품명", "단가", "재고"],  # row 7: table 2 header
        ["키보드", 30000, 10],  # rows 8-10: table 2 data
        ["마우스", 15000, 20],
        ["모니터", 210000, 5],
    ]
    profile = make_sheet_profile(name="S", max_row=len(rows), max_col=3)
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))
    locator = HeaderLocator()

    band_row, band_score = locator._locate(ctx, profile, row_window=(7, 10))
    assert band_row == 7  # absolute sheet row, window_start + best_index
    assert band_score >= 0.5

    # A window beyond the content yields no sampleable rows.
    assert locator._locate(ctx, profile, row_window=(11, 12)) == (None, 0.0)
    # An inverted window is an empty sample, not a crash.
    assert locator._locate(ctx, profile, row_window=(8, 7)) == (None, 0.0)

    # The core never mutates the profile (the applier's job).
    assert profile.header_row is None
    assert profile.header_confidence == 0.0


def test_sample_rows_scan_clamped_to_window() -> None:
    """The scan end is min(window_start + HEADER_SCAN_ROWS - 1, window_end).

    Guard 2's clamp rule (plan v2 Task 10.2): a short band clamps the read to
    its own end; a tall window is clamped by ``HEADER_SCAN_ROWS`` from the
    window's start. The whole-sheet default still requests exactly rows
    ``1 .. HEADER_SCAN_ROWS`` (the v1 read, pinned by test_performance too).
    """

    from excel_inspector.heuristics import HEADER_SCAN_ROWS

    requests: list[tuple[int, int]] = []

    class _RecordingWorksheet:
        def iter_rows(self, *, min_row, max_row, values_only):  # noqa: ANN001, ANN202
            requests.append((min_row, max_row))
            return iter(())

    class _RecordingLoader:
        def data_workbook(self):  # noqa: ANN202 - test double
            return {"S": _RecordingWorksheet()}

    profile = make_sheet_profile(name="S", max_row=100, max_col=3)
    ctx = make_context(sheets=[profile], loader=_RecordingLoader())
    locator = HeaderLocator()

    locator._sample_rows(ctx, profile)  # whole-sheet default
    locator._sample_rows(ctx, profile, 7, 10)  # clamped by window_end
    locator._sample_rows(ctx, profile, 7, 200)  # clamped by HEADER_SCAN_ROWS

    assert requests == [
        (1, HEADER_SCAN_ROWS),
        (7, 10),
        (7, 7 + HEADER_SCAN_ROWS - 1),
    ]


def test_col_count_window_scoped_uses_band_local_width() -> None:
    """Window-scoped scoring ignores the sheet-global max_col (guard 1).

    A narrow (3-column) band on a sheet whose global ``max_col`` is 8 must be
    scored against its own used width, or its header score would be diluted
    into a "not a table" misjudgment (plan v2 Task 10.2 guard 1). The
    whole-sheet default keeps preferring ``max_col`` (v1 behavior).
    """

    rows = [["코드", "명칭", "수량"], ["A-1", "키보드", 1]]
    profile = make_sheet_profile(name="S", max_row=len(rows), max_col=8)

    assert HeaderLocator._col_count(profile, rows) == 8  # v1 whole-sheet
    assert HeaderLocator._col_count(profile, rows, window_scoped=True) == 3


# ---------------------------------------------------------------------------
# Issue #8: lookahead-evidence weighting of type_consistency
# ---------------------------------------------------------------------------


def test_small_mixed_table_prefers_true_header_over_bottom_row() -> None:
    """Issue #8 regression: a bottom all-string data row must not win.

    Pre-fix, the §7.1 scoring handed row 4 ('거래처수'/'5개사'/'100%') a
    trivially-perfect type_consistency — its lookahead window held a single
    row, and one row is always self-consistent — so it outscored the true
    header at row 1 (whose consistency the genuinely mixed '값' column drags
    down) and rows 1-3 vanished from the loaded frame. The consistency term
    now scales with the observed lookahead evidence, so the true header wins.
    """

    rows = [
        ["항목", "값", "달성률"],
        ["총매출", 48300, "92%"],
        ["총수량", 318, "104%"],
        ["거래처수", "5개사", "100%"],
        ["반품건수", 1, "-"],
    ]
    scores = [_score_row(i, rows, 3) for i in range(len(rows))]
    assert max(range(len(scores)), key=scores.__getitem__) == 0
    assert scores[0] >= 0.5  # still above the manual-header threshold


def test_type_consistency_scales_with_lookahead_evidence() -> None:
    """The consistency weight is earned by evidence, not granted by default.

    Identical candidate and below-row content, differing only in how many
    below rows exist: a single-row window may claim only 1/5 of the full
    consistency term, so the score gap is exactly
    ``0.3 * 1.0 * (1 - 1/5) = 0.24``. Repeating the same below row keeps
    str_ratio and distinctness identical across both variants, so the gap
    isolates the evidence factor.
    """

    candidate = ["name", "qty"]
    below_row = ["widget", 10]
    full = [candidate] + [list(below_row) for _ in range(5)]
    single = [candidate, list(below_row)]

    gap = _score_row(0, full, 2) - _score_row(0, single, 2)
    assert abs(gap - 0.3 * (1.0 - 1.0 / 5.0)) < 1e-9


def test_small_legitimate_table_header_stays_above_threshold() -> None:
    """Evidence weighting must not sink a legitimate small table's header.

    A 3-row table (header + 2 typed data rows) keeps its header at row 1,
    comfortably above the 0.5 threshold, even though only 2/5 of the
    lookahead window exists.
    """

    rows = [["name", "qty"], ["widget", 10], ["gadget", 5]]
    scores = [_score_row(i, rows, 2) for i in range(len(rows))]
    assert max(range(len(scores)), key=scores.__getitem__) == 0
    assert scores[0] >= 0.5
