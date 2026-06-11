"""Merge Analyzer tests (spec §4.4, Phase 6) [D3].

Three layers, mirroring the Boundary Detector / Type Profiler suites:

1.  **Loader-backed fixture tests** run the wired
    ``SheetEnumerator -> HeaderLocator -> BoundaryDetector -> TypeProfiler ->
    MergeAnalyzer`` chain over the real **structure-mode** workbook and assert
    the §4.4 merge classification (``header`` vs ``body``), the
    ``is_multi_level_header`` flag, and (via the aggregator) the body-merge
    forward-fill recommendation recorded on ``ReadPlan.notes``.
2.  **Isolated partial-context tests** drive the analyzer through ``conftest``
    synthesis with a fake **structure-mode** loader so the classification rules,
    the multi-level flag, and the headerless conservative default can be
    exercised without touching disk. A structural assertion proves the analyzer
    uses ``structure_workbook()`` and never ``data_workbook()`` [D3].
3.  **Helper unit test** pins the ``_classify_kind`` rule directly.

merged_cells ordering is not guaranteed by openpyxl, so every set of regions is
compared as a ``set`` of ``(range, kind)`` pairs and notes as a ``set`` (spec
note in the task; aggregator docstring).

Coordinates here are **openpyxl 1-based** (the inspection domain [D1]).
"""

from __future__ import annotations

from pathlib import Path

from excel_inspector import InspectionOptions, Loader, inspect
from excel_inspector.aggregator import PlanAggregator
from excel_inspector.analyzers.boundary_detector import BoundaryDetector
from excel_inspector.analyzers.header_locator import HeaderLocator
from excel_inspector.analyzers.merge_analyzer import (
    MergeAnalyzer,
    _classify_kind,
)
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.analyzers.type_profiler import TypeProfiler
from excel_inspector.context import InspectionContext
from excel_inspector.models import WorkbookProfile

from conftest import make_context, make_sheet_profile  # type: ignore[import-not-found]


def _run_on(
    path: Path, options: InspectionOptions | None = None
) -> InspectionContext:
    """Enumerate -> headers -> boundaries -> types -> merges; return context."""

    context = InspectionContext(
        options=options or InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)
        HeaderLocator().analyze(context)
        BoundaryDetector().analyze(context)
        TypeProfiler().analyze(context)
        return MergeAnalyzer().analyze(context)


def _merge_set(profile) -> set[tuple[str, str]]:  # noqa: ANN001
    """Return the sheet's merges as an order-insensitive ``(range, kind)`` set."""

    return {(m.range, m.kind) for m in profile.merges}


# ---------------------------------------------------------------------------
# Identity
# ---------------------------------------------------------------------------


def test_name() -> None:
    """The analyzer reports a stable identifier."""

    assert MergeAnalyzer().name() == "merge_analyzer"


# ---------------------------------------------------------------------------
# Loader-backed fixture classification (spec §4.4)
# ---------------------------------------------------------------------------


def test_merged_header_classifies_header_and_body(fixture_path) -> None:
    """merged_header: A1:B1 is a header merge, A6:A7 is a body merge (§4.4).

    The header merge overlaps header_row 1; the body merge (rows 6-7) sits below
    it. is_multi_level_header is False (no header band strictly above row 1).
    """

    sheet = _run_on(fixture_path("merged_header")).workbook_profile.sheets[0]

    assert sheet.header_row == 1
    assert _merge_set(sheet) == {("A1:B1", "header"), ("A6:A7", "body")}
    assert sheet.is_multi_level_header is False


def test_multi_level_header_sets_flag(fixture_path) -> None:
    """multi_level_header: row-1 group merges over a row-2 leaf header (§4.4).

    Both group merges (A1:B1, C1:D1) sit on row 1, strictly above the resolved
    header_row 2, so is_multi_level_header is True. v1 only judges the flag; it
    does NOT branch the load path [D6]. Both merges classify as ``header``.
    """

    sheet = _run_on(
        fixture_path("multi_level_header")
    ).workbook_profile.sheets[0]

    assert sheet.header_row == 2
    assert sheet.is_multi_level_header is True
    assert _merge_set(sheet) == {("A1:B1", "header"), ("C1:D1", "header")}


def test_no_merges_sheet_is_empty(fixture_path) -> None:
    """A plain table has no merges and is not a multi-level header (§4.4)."""

    sheet = _run_on(fixture_path("header_simple")).workbook_profile.sheets[0]

    assert sheet.merges == []
    assert sheet.is_multi_level_header is False


# ---------------------------------------------------------------------------
# Body-merge forward-fill note flows into ReadPlan.notes (spec §4.4)
# ---------------------------------------------------------------------------


def test_body_merge_fill_note_in_read_plan(fixture_path) -> None:
    """The body merge yields a forward-fill recommendation on ReadPlan.notes.

    Driven through the full ``inspect()`` pipeline (MergeAnalyzer before the
    aggregator). The note names the body merge's A1 range and recommends
    forward-fill; the actual fill is the loader's job (spec §4.4).
    """

    profile = inspect(fixture_path("merged_header"))
    sheet = profile.sheets[0]
    plan = sheet.read_plan
    assert plan is not None

    # Order is not guaranteed; compare as a set. Exactly one body merge (A6:A7).
    assert set(plan.notes) == {
        "body merge A6:A7 -> forward-fill top-left value (spec §4.4)"
    }


def test_header_only_merge_yields_no_fill_note(fixture_path) -> None:
    """A sheet whose only merges are header merges records no fill note (§4.4).

    multi_level_header has only ``header``-kind merges, so the aggregator emits
    no forward-fill note onto the read plan.
    """

    profile = inspect(fixture_path("multi_level_header"))
    plan = profile.sheets[0].read_plan
    assert plan is not None
    assert plan.notes == []


# ---------------------------------------------------------------------------
# Structure-mode usage [D3]: a fake loader proving read_only is never used
# ---------------------------------------------------------------------------


class _FakeRange:
    """A merged-cell range stub exposing CellRange-style bounds and an A1 str.

    The Phase 11a collection step (``_collect_spans``) reads all four bounds
    (``min_row``/``min_col``/``max_row``/``max_col``) like openpyxl's real
    ``CellRange``; classification itself still depends on ``min_row`` only,
    so the column/extent defaults are immaterial to these tests.
    """

    def __init__(
        self,
        a1: str,
        min_row: int,
        min_col: int = 1,
        max_row: int | None = None,
        max_col: int | None = None,
    ) -> None:
        self._a1 = a1
        self.min_row = min_row
        self.min_col = min_col
        self.max_row = max_row if max_row is not None else min_row
        self.max_col = max_col if max_col is not None else min_col

    def __str__(self) -> str:
        return self._a1


class _FakeMergedCells:
    """A ``merged_cells`` stub exposing a ``ranges`` iterable."""

    def __init__(self, ranges: list[_FakeRange]) -> None:
        self.ranges = ranges


class _FakeWorksheet:
    """A structure-mode worksheet stub exposing ``merged_cells``."""

    def __init__(self, ranges: list[_FakeRange]) -> None:
        self.merged_cells = _FakeMergedCells(ranges)


class _FakeWorkbook:
    """A workbook mapping sheet names to :class:`_FakeWorksheet`."""

    def __init__(self, sheets: dict[str, list[_FakeRange]]) -> None:
        self._sheets = {n: _FakeWorksheet(r) for n, r in sheets.items()}

    def __getitem__(self, name: str) -> _FakeWorksheet:
        return self._sheets[name]


class _StructureOnlyLoader:
    """A loader stub exposing only :meth:`structure_workbook` [D3].

    Calling :meth:`data_workbook` raises so any accidental use of the read_only
    streaming mode (which lacks ``merged_cells``) fails the test loudly.
    """

    def __init__(self, sheets: dict[str, list[_FakeRange]]) -> None:
        self._wb = _FakeWorkbook(sheets)
        self.structure_calls = 0

    def structure_workbook(self) -> _FakeWorkbook:
        self.structure_calls += 1
        return self._wb

    def data_workbook(self):  # noqa: ANN202
        raise AssertionError(
            "MergeAnalyzer must use structure mode only; read_only worksheets "
            "lack merged_cells [D3]"
        )


def test_uses_structure_mode_only() -> None:
    """The analyzer reads structure mode and never the data (read_only) mode [D3]."""

    profile = make_sheet_profile(name="S", header_row=1, max_row=5, max_col=3)
    loader = _StructureOnlyLoader({"S": [_FakeRange("A1:B1", 1)]})
    ctx = make_context(sheets=[profile], loader=loader)

    MergeAnalyzer().analyze(ctx)

    assert loader.structure_calls == 1
    assert _merge_set(profile) == {("A1:B1", "header")}


def test_isolated_classifies_header_above_and_body_below() -> None:
    """A merge above the header is ``header``; one below is ``body`` (§4.4).

    With header_row 3: a merge on row 1 (above) and a merge overlapping row 3
    are both ``header``; a merge starting at row 4 (below) is ``body``.
    is_multi_level_header is True because the row-1 merge is strictly above the
    header.
    """

    profile = make_sheet_profile(name="S", header_row=3, max_row=8, max_col=4)
    loader = _StructureOnlyLoader(
        {
            "S": [
                _FakeRange("A1:D1", 1),  # above header -> header band
                _FakeRange("A3:B3", 3),  # overlaps header -> header
                _FakeRange("A4:A6", 4),  # below header -> body
            ]
        }
    )
    ctx = make_context(sheets=[profile], loader=loader)

    MergeAnalyzer().analyze(ctx)

    assert _merge_set(profile) == {
        ("A1:D1", "header"),
        ("A3:B3", "header"),
        ("A4:A6", "body"),
    }
    assert profile.is_multi_level_header is True


def test_isolated_no_band_above_header_not_multi_level() -> None:
    """Header merges only on the header row itself are not multi-level (§4.4)."""

    profile = make_sheet_profile(name="S", header_row=2, max_row=6, max_col=4)
    loader = _StructureOnlyLoader(
        {"S": [_FakeRange("A2:B2", 2), _FakeRange("A3:A4", 3)]}
    )
    ctx = make_context(sheets=[profile], loader=loader)

    MergeAnalyzer().analyze(ctx)

    assert _merge_set(profile) == {
        ("A2:B2", "header"),
        ("A3:A4", "body"),
    }
    assert profile.is_multi_level_header is False


def test_isolated_headerless_sheet_defaults_to_body() -> None:
    """With no header anchor, every merge is conservatively ``body`` (§4.4).

    A sheet whose ``header_row`` is ``None`` (estimation failed / declared
    headerless) has no anchor to classify against; merges default to ``body``
    (the safe default — body only adds a fill note) and the multi-level flag
    stays ``False``.
    """

    profile = make_sheet_profile(name="S", header_row=None, max_row=6, max_col=3)
    loader = _StructureOnlyLoader(
        {"S": [_FakeRange("A1:B1", 1), _FakeRange("A3:A4", 3)]}
    )
    ctx = make_context(sheets=[profile], loader=loader)

    MergeAnalyzer().analyze(ctx)

    assert _merge_set(profile) == {
        ("A1:B1", "body"),
        ("A3:A4", "body"),
    }
    assert profile.is_multi_level_header is False


def test_isolated_non_tabular_sheet_skipped() -> None:
    """Non-tabular sheets are skipped (excluded from loading, spec §9)."""

    profile = make_sheet_profile(
        name="README", is_tabular_candidate=False, header_row=1, max_col=1
    )
    loader = _StructureOnlyLoader({"README": [_FakeRange("A1:B1", 1)]})
    ctx = make_context(sheets=[profile], loader=loader)

    MergeAnalyzer().analyze(ctx)

    # Untouched: the analyzer never read its merges. Since Phase 11a the
    # no-scanner fallback opens the structure workbook lazily, so a workbook
    # with no tabular sheet is never fetched at all.
    assert profile.merges == []
    assert loader.structure_calls == 0


# ---------------------------------------------------------------------------
# Helper unit test
# ---------------------------------------------------------------------------


def test_classify_kind_rule() -> None:
    """``_classify_kind`` pins the §4.4 boundary rule directly."""

    # header_row = 3
    assert _classify_kind(1, 3) == "header"  # above
    assert _classify_kind(3, 3) == "header"  # overlapping (equal)
    assert _classify_kind(4, 3) == "body"  # below
    # Unknown header -> conservative body default.
    assert _classify_kind(1, None) == "body"
    assert _classify_kind(99, None) == "body"
