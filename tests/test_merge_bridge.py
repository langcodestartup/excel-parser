"""Merged-header boundary bridge tests — Phase 11a (plan v2 §5 Task 11.1).

Layers covered:

1.  **Characterization flip golden** — ``merged_header`` resolves to the
    FIXTURES-documented boundaries (``data_start_row=2``, ``data_end_row=5``;
    1-based [D1]) instead of the v1 ``None``/``None`` deferral, and the A6:A7
    body-merge demo block (rows 6-7) is excluded from the table body with a
    visible warning (fixture-contradiction guard: the answer is 5, never 7).
2.  **Step 1 — collection/classification split** — the :class:`MergeScanner`
    collects *unclassified* sorted spans once per sheet before the Boundary
    Detector; the :class:`MergeAnalyzer` classifies the collected spans later
    (without re-reading the workbook) once the header is final.
3.  **Step 2 — virtual fill** — empty header cells covered by a merge
    intersecting the header row count as filled for the column-span
    derivation, including inside a band-scoped row window.
4.  **Trailing merged-group rule** — only a fully merge-grouped row group
    *trailing* a merge-free body is excluded; group-label tables (merges
    interior to / covering the body) are never clamped.
5.  **Block attribution** — a band-scoped block's plan carries only the
    forward-fill notes of merges whose rows intersect its band.
6.  **Determinism** — openpyxl guarantees no ``merged_cells.ranges`` order, so
    notes are asserted as sets and the serialized JSON is asserted stable
    across repeated ``extract()`` calls.

Coordinates in assertions are openpyxl 1-based for ``SheetProfile`` /
``TableBlock`` / ``MergeSpan`` and pandas 0-based for ``ReadPlan`` fields
[D1]; fixture layouts are documented in ``tests/fixtures/generate.py``
(FIXTURES, the single source).
"""

from __future__ import annotations

from excel_inspector import (
    MergeAnalyzer,
    MergeScanner,
    MergeSpan,
    extract,
    inspect,
)
from excel_inspector.aggregator import build_block_read_plan
from excel_inspector.analyzers.boundary_detector import BoundaryDetector
from excel_inspector.models import MergeRegion, TableBlock

from conftest import make_context, make_sheet_profile  # type: ignore[import-not-found]

#: The body-merge forward-fill note pinned by the existing Phase 6 golden.
_FILL_NOTE = "body merge A6:A7 -> forward-fill top-left value (spec §4.4)"


# ---------------------------------------------------------------------------
# Characterization flip goldens — merged_header end-to-end (Step 3)
# ---------------------------------------------------------------------------


def test_merged_header_inspect_boundaries_resolved(fixture_path) -> None:
    """The FIXTURES-documented boundaries: data 2-5, demo block excluded.

    Plan v2 Task 11.1 Step 3 + fixture-contradiction guard: ``data_end_row``
    is 5 — the A6:A7 body-merge demo rows 6-7 are NOT table body. The bridged
    header spans the full width A-C, so the column boundaries stay ``None``
    (no usecols restriction, spec §4.5).
    """

    sheet = inspect(fixture_path("merged_header")).sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 5          # never 7 (plan guard)
    assert (sheet.data_left_col, sheet.data_right_col) == (None, None)
    assert sheet.skip_rows == []


def test_merged_header_read_plan_pinned(fixture_path) -> None:
    """The bridged sheet's ReadPlan (0-based loading domain [D1]).

    The data region now resolves, so the plan reads exactly the 4 data rows
    (nrows = 5 - 2 + 1) and the profiled text columns drive the dtype_map
    [D5] (column 2, ``점수``, is ``number`` -> omitted).
    """

    plan = inspect(fixture_path("merged_header")).sheets[0].read_plan
    assert plan is not None
    assert plan.header == 0
    assert plan.skiprows == []
    assert plan.nrows == 4
    assert plan.usecols is None
    assert plan.dtype_map == {"0": "string", "1": "string"}


def test_merged_header_columns_profiled(fixture_path) -> None:
    """The resolved data region is type-profiled (was empty pre-11a).

    Column B has no header label (the A1:B1 merge leaves B1 empty in data
    mode), so its name is ``None`` while its body still types as text.
    """

    sheet = inspect(fixture_path("merged_header")).sheets[0]
    assert [(c.index, c.name, c.inferred_type) for c in sheet.columns] == [
        (0, "이름", "text"),
        (1, None, "text"),
        (2, "점수", "number"),
    ]


def test_merged_header_extract_golden(fixture_path) -> None:
    """extract() loads exactly the 4 table rows; the demo block never leaks."""

    wr = extract(fixture_path("merged_header"))
    (table,) = wr.tables
    assert table.table_id == "Sheet1!T1"
    assert table.header_row == 1
    df = table.dataframe
    assert df.shape == (4, 3)
    # pandas labels the merged-away empty header cell 'Unnamed: 1'.
    assert list(df.columns) == ["이름", "Unnamed: 1", "점수"]
    assert list(df["이름"]) == ["Kim", "Lee", "Park", "Choi"]
    assert int(df["점수"].sum()) == 305
    # Zero leakage of the excluded A6:A7 demo block (rows 6-7).
    flat = df.astype(str).to_numpy().ravel().tolist()
    for demo_value in ("그룹", "정대만", "송태섭", "50", "55"):
        assert all(demo_value not in cell for cell in flat), demo_value


def test_merged_header_fill_note_in_extract_notes(fixture_path) -> None:
    """Task 11.1 Step 3: the forward-fill note reaches TableResult.notes.

    ``merged_cells.ranges`` ordering is not guaranteed by openpyxl, so the
    notes are compared as a set (plan v2 Task 11.1 determinism note).
    """

    wr = extract(fixture_path("merged_header"))
    (table,) = wr.tables
    assert set(table.notes) == {_FILL_NOTE}


def test_merged_header_exclusion_warning_visible(fixture_path) -> None:
    """The demo-block exclusion is warned, and the old deferral is gone."""

    wr = extract(fixture_path("merged_header"))
    assert any(
        "trailing merged-row group (rows 6-7)" in w and "clamped to 5" in w
        for w in wr.warnings
    )
    assert not any("pending merge analysis" in w for w in wr.warnings)


def test_merged_header_extract_is_deterministic(fixture_path) -> None:
    """Sorted span collection makes notes/warnings/JSON order-stable."""

    p = fixture_path("merged_header")
    first = extract(p)
    second = extract(p)
    assert first.warnings == second.warnings
    assert first.to_json() == second.to_json()


def test_multi_level_header_classified_from_collected_spans(
    fixture_path,
) -> None:
    """The pipeline (spans-primary) classification path keeps the §4.4 golden.

    multi_level_header's row-1 group merges sit strictly above the row-2 leaf
    header: both classify ``header`` and the multi-level flag stays True when
    the Merge Analyzer consumes the scanner's spans instead of re-reading the
    workbook.
    """

    sheet = inspect(fixture_path("multi_level_header")).sheets[0]
    assert sheet.header_row == 2
    assert sheet.is_multi_level_header is True
    assert {(m.range, m.kind) for m in sheet.merges} == {
        ("A1:B1", "header"),
        ("C1:D1", "header"),
    }
    # Header-row merges intersect row... 1 only, not the row-2 leaf header, so
    # the virtual fill is a no-op and the boundaries keep their v1 values.
    assert sheet.data_start_row == 3
    assert sheet.data_end_row == 6


# ---------------------------------------------------------------------------
# Step 1 — MergeScanner: collection without classification
# ---------------------------------------------------------------------------


class _FakeRange:
    """A CellRange-style stub carrying the four bounds and an A1 ``str``."""

    def __init__(
        self, a1: str, min_row: int, min_col: int, max_row: int, max_col: int
    ) -> None:
        self._a1 = a1
        self.min_row = min_row
        self.min_col = min_col
        self.max_row = max_row
        self.max_col = max_col

    def __str__(self) -> str:
        return self._a1


class _FakeMergedCells:
    def __init__(self, ranges: list[_FakeRange]) -> None:
        self.ranges = ranges


class _FakeStructureWorksheet:
    def __init__(self, ranges: list[_FakeRange]) -> None:
        self.merged_cells = _FakeMergedCells(ranges)


class _FakeStructureWorkbook:
    def __init__(self, sheets: dict[str, list[_FakeRange]]) -> None:
        self._sheets = {
            name: _FakeStructureWorksheet(r) for name, r in sheets.items()
        }

    def __getitem__(self, name: str) -> _FakeStructureWorksheet:
        return self._sheets[name]


class _StructureLoader:
    """A loader stub exposing only :meth:`structure_workbook` [D3]."""

    def __init__(self, sheets: dict[str, list[_FakeRange]]) -> None:
        self._wb = _FakeStructureWorkbook(sheets)
        self.structure_calls = 0

    def structure_workbook(self) -> _FakeStructureWorkbook:
        self.structure_calls += 1
        return self._wb

    def data_workbook(self):  # noqa: ANN202
        raise AssertionError(
            "merge collection must use structure mode only [D3]"
        )


def test_scanner_collects_sorted_unclassified_spans() -> None:
    """Spans land on context.merge_spans, sorted, with no classification.

    The fake ranges are deliberately supplied out of order (openpyxl does not
    guarantee any), and the scanner must not touch ``profile.merges`` — the
    classification belongs to the Merge Analyzer after the header is final
    (plan v2 Task 11.1 Step 1).
    """

    profile = make_sheet_profile(name="S", max_row=8, max_col=4)
    loader = _StructureLoader(
        {
            "S": [
                _FakeRange("A6:A7", 6, 1, 7, 1),
                _FakeRange("A1:B1", 1, 1, 1, 2),
            ]
        }
    )
    ctx = make_context(sheets=[profile], loader=loader)

    MergeScanner().analyze(ctx)

    assert ctx.merge_spans["S"] == [
        MergeSpan(range="A1:B1", min_row=1, min_col=1, max_row=1, max_col=2),
        MergeSpan(range="A6:A7", min_row=6, min_col=1, max_row=7, max_col=1),
    ]
    assert profile.merges == []          # collection classifies nothing
    assert loader.structure_calls == 1   # once per workbook


def test_scanner_records_empty_list_for_merge_free_sheet() -> None:
    """A scanned sheet without merges gets an (empty) entry — 'scanned' state.

    The empty list is meaningful: it tells the Merge Analyzer the sheet was
    already scanned, so the no-scanner fallback never re-reads the workbook.
    """

    profile = make_sheet_profile(name="S", max_row=3, max_col=2)
    ctx = make_context(sheets=[profile], loader=_StructureLoader({"S": []}))

    MergeScanner().analyze(ctx)

    assert ctx.merge_spans == {"S": []}


def test_scanner_skips_non_tabular_sheets() -> None:
    """Non-tabular sheets are excluded from loading (spec §9) — not scanned."""

    profile = make_sheet_profile(
        name="README", is_tabular_candidate=False, max_row=3, max_col=1
    )
    loader = _StructureLoader({"README": [_FakeRange("A1:B1", 1, 1, 1, 2)]})
    ctx = make_context(sheets=[profile], loader=loader)

    MergeScanner().analyze(ctx)

    assert "README" not in ctx.merge_spans


def test_scanner_without_loader_warns() -> None:
    """No loader -> a visible warning, no crash (spec §6 robustness)."""

    profile = make_sheet_profile(name="S", max_row=3, max_col=2)
    ctx = make_context(sheets=[profile], loader=None)

    MergeScanner().analyze(ctx)

    assert ctx.merge_spans == {}
    assert any("merge_scanner" in w for w in ctx.warnings)


def test_analyzer_classifies_collected_spans_without_workbook() -> None:
    """Pre-collected spans are classified with NO workbook re-read (Step 1).

    The loader stub raises on every access, proving the analyzer's primary
    path runs entirely off ``context.merge_spans``. With header_row=1 the
    A1:B1 merge classifies ``header`` and A6:A7 ``body``.
    """

    class _ExplodingLoader:
        def structure_workbook(self):  # noqa: ANN202
            raise AssertionError(
                "MergeAnalyzer must not re-read the workbook when spans "
                "were already collected (Task 11.1 Step 1)"
            )

        data_workbook = structure_workbook

    profile = make_sheet_profile(name="S", header_row=1, max_row=7, max_col=3)
    ctx = make_context(sheets=[profile], loader=_ExplodingLoader())
    ctx.merge_spans["S"] = [
        MergeSpan(range="A1:B1", min_row=1, min_col=1, max_row=1, max_col=2),
        MergeSpan(range="A6:A7", min_row=6, min_col=1, max_row=7, max_col=1),
    ]

    MergeAnalyzer().analyze(ctx)

    assert {(m.range, m.kind) for m in profile.merges} == {
        ("A1:B1", "header"),
        ("A6:A7", "body"),
    }
    assert profile.is_multi_level_header is False


# ---------------------------------------------------------------------------
# Step 2 — virtual fill in the boundary core (isolated, no disk)
# ---------------------------------------------------------------------------


class _FakeDataWorksheet:
    """A read_only-style worksheet returning fixed 1-based-aligned rows."""

    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows

    def iter_rows(self, *, min_row, max_row, values_only):  # noqa: ANN001, ANN202
        end = len(self._rows) if max_row is None else min(max_row, len(self._rows))
        for r in range(min_row - 1, end):
            yield tuple(self._rows[r])


class _FakeDataWorkbook:
    def __init__(self, sheets: dict[str, list[list[object]]]) -> None:
        self._sheets = {n: _FakeDataWorksheet(r) for n, r in sheets.items()}

    def __getitem__(self, name: str) -> _FakeDataWorksheet:
        return self._sheets[name]


class _FakeDataLoader:
    """A loader stub exposing only :meth:`data_workbook`."""

    def __init__(self, sheets: dict[str, list[list[object]]]) -> None:
        self._wb = _FakeDataWorkbook(sheets)

    def data_workbook(self) -> _FakeDataWorkbook:
        return self._wb


def test_virtual_fill_bridges_merged_header_span() -> None:
    """A header-row merge restores the collapsed span (Task 11.1 Step 2).

    Without spans the interior gap defers everything (pinned by the existing
    ``test_detect_block_merge_deferral_warning_is_block_local``); with the
    collected A1:B1 span the same rows resolve to data 2-3 over the bridged
    full-width span.
    """

    rows = [
        ["이름", None, "점수"],  # row 1 merged header (B1 empty via A1:B1)
        ["김", "김철수", 90],  # row 2
        ["이", "이영희", 85],  # row 3
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A1:B1", min_row=1, min_col=1, max_row=1, max_col=2)
    ]

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert (result.data_start_row, result.data_end_row) == (2, 3)
    # Bridged span A-C == full width -> no usecols restriction (spec §4.5).
    assert (result.data_left_col, result.data_right_col) == (None, None)
    assert result.skip_rows == []
    assert result.warnings == []         # no deferral warning anymore


def test_virtual_fill_only_for_merges_touching_the_header_row() -> None:
    """A merge that does NOT intersect the header row fills nothing.

    The body-side A2:B2 merge must not bridge the genuine header gap, so the
    v1 deferral behavior is preserved (the gap is real, not merge-made).
    """

    rows = [
        ["이름", None, "점수"],  # row 1: genuinely gapped header
        ["병합", None, 90],  # row 2 (A2:B2 merged, B2 empty)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A2:B2", min_row=2, min_col=1, max_row=2, max_col=2)
    ]

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert (result.data_start_row, result.data_end_row) == (None, None)
    assert any("pending merge analysis" in w for w in result.warnings)


def test_virtual_fill_inside_band_window() -> None:
    """The bridge works for a band-scoped (windowed) lower block too.

    A merged-header table sits in the band rows 5-8; the detection core is
    invoked exactly like the Block Analyzer would (header 5, window (5, 8))
    and must resolve the band's data rows 6-8.
    """

    rows = [
        ["상단표", "x", "y"],  # row 1 (another band)
        ["a", 1, 2],  # row 2
        [None, None, None],  # row 3 blank
        [None, None, None],  # row 4 blank (band separator)
        ["이름", None, "점수"],  # row 5 merged header (A5:B5)
        ["김", "김철수", 90],  # row 6
        ["이", "이영희", 85],  # row 7
        ["박", "박민수", 70],  # row 8
    ]
    profile = make_sheet_profile(
        name="S", header_row=None, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A5:B5", min_row=5, min_col=1, max_row=5, max_col=2)
    ]

    result = BoundaryDetector()._detect_block(
        ctx, profile, 5, row_window=(5, 8)
    )

    assert (result.data_start_row, result.data_end_row) == (6, 8)
    assert result.skip_rows == []
    assert result.warnings == []


# ---------------------------------------------------------------------------
# Trailing merged-group exclusion — clamp vs group-label tables (no disk)
# ---------------------------------------------------------------------------


def test_trailing_merged_group_is_excluded_with_warning() -> None:
    """A merge-grouped block trailing a merge-free body is clamped out.

    The merged_header shape in miniature: flat data rows 2-3, then rows 4-5
    grouped by a vertical A4:A5 merge. data_end_row clamps to 3 and the
    exclusion is named in the block-local warnings (never silent).
    """

    rows = [
        ["이름", "별명", "점수"],  # row 1 header
        ["김", "a", 90],  # row 2 flat data
        ["이", "b", 85],  # row 3 flat data
        ["그룹", "c", 50],  # row 4 (A4:A5 merge anchor)
        [None, "d", 55],  # row 5 (merge continuation)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A4:A5", min_row=4, min_col=1, max_row=5, max_col=1)
    ]

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert (result.data_start_row, result.data_end_row) == (2, 3)
    assert any(
        "trailing merged-row group (rows 4-5)" in w for w in result.warnings
    )


def test_interior_merged_group_table_is_never_clamped() -> None:
    """Group-label tables keep every row (spec §4.4 forward-fill intent).

    A merged group *followed by a flat data row* is interior table structure,
    so the trailing-exclusion rule must stay off — even though the last rows
    are also merge-grouped (the table's own style includes merges).
    """

    rows = [
        ["그룹", "이름", "점수"],  # row 1 header
        ["A", "김철수", 90],  # row 2 (A2:A3 merge anchor)
        [None, "이영희", 85],  # row 3 (merge continuation)
        ["B", "박민수", 70],  # row 4 flat data
        ["C", "최지우", 60],  # row 5 (A5:A6 merge anchor)
        [None, "정대만", 50],  # row 6 (merge continuation)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A2:A3", min_row=2, min_col=1, max_row=3, max_col=1),
        MergeSpan(range="A5:A6", min_row=5, min_col=1, max_row=6, max_col=1),
    ]

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert (result.data_start_row, result.data_end_row) == (2, 6)
    assert result.warnings == []


def test_fully_merged_body_is_never_clamped() -> None:
    """A body covered end-to-end by merged groups has no flat anchor -> kept."""

    rows = [
        ["그룹", "이름", "점수"],  # row 1 header
        ["A", "김철수", 90],  # row 2 (A2:A3 merge anchor)
        [None, "이영희", 85],  # row 3 (merge continuation)
        ["B", "박민수", 70],  # row 4 (A4:A5 merge anchor)
        [None, "최지우", 60],  # row 5 (merge continuation)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A2:A3", min_row=2, min_col=1, max_row=3, max_col=1),
        MergeSpan(range="A4:A5", min_row=4, min_col=1, max_row=5, max_col=1),
    ]

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert (result.data_start_row, result.data_end_row) == (2, 5)
    assert result.warnings == []


def test_single_row_body_merge_groups_nothing() -> None:
    """A horizontal (single-row) body merge never triggers the exclusion."""

    rows = [
        ["이름", "별명", "점수"],  # row 1 header
        ["김", "a", 90],  # row 2 flat data
        ["병합", None, 55],  # row 3 (A3:B3 horizontal merge anchor)
    ]
    profile = make_sheet_profile(
        name="S", header_row=1, max_row=len(rows), max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeDataLoader({"S": rows}))
    ctx.merge_spans["S"] = [
        MergeSpan(range="A3:B3", min_row=3, min_col=1, max_row=3, max_col=2)
    ]

    result = BoundaryDetector()._detect_block(ctx, profile, 1)

    assert (result.data_start_row, result.data_end_row) == (2, 3)
    assert result.warnings == []


# ---------------------------------------------------------------------------
# Block attribution — per-block plans own only their band's merges (Step 1)
# ---------------------------------------------------------------------------


def _block(band_start: int, band_end: int, header: int, index: int) -> TableBlock:
    """A minimal table-judged block for plan synthesis (1-based [D1])."""

    return TableBlock(
        block_index=index,
        band_start_row=band_start,
        band_end_row=band_end,
        header_row=header,
        header_confidence=0.9,
        header_provenance="heuristic",
        data_start_row=header + 1,
        data_end_row=band_end,
        data_left_col=None,
        data_right_col=None,
        skip_rows=[],
        columns=[],
        read_plan=None,
    )


def test_band_scoped_plan_owns_only_intersecting_merges() -> None:
    """Row intersection attributes each body merge to exactly one block.

    With body merges A2:A3 (band 1) and A8:A9 (band 2), each band-scoped plan
    must carry only its own forward-fill note — never the sibling's (plan v2
    Task 11.1 Step 1, block attribution).
    """

    profile = make_sheet_profile(
        name="S",
        max_row=10,
        max_col=3,
        merges=[
            MergeRegion(range="A2:A3", kind="body"),
            MergeRegion(range="A8:A9", kind="body"),
        ],
    )
    top = _block(1, 4, header=1, index=0)
    bottom = _block(7, 10, header=7, index=1)

    top_plan = build_block_read_plan(profile, top, band_scoped=True)
    bottom_plan = build_block_read_plan(profile, bottom, band_scoped=True)

    assert set(top_plan.notes) == {
        "body merge A2:A3 -> forward-fill top-left value (spec §4.4)"
    }
    assert set(bottom_plan.notes) == {
        "body merge A8:A9 -> forward-fill top-left value (spec §4.4)"
    }


def test_mirror_plan_keeps_all_sheet_merges() -> None:
    """The single-band mirror path keeps every sheet merge (compat invariant).

    The flat ``build_read_plan`` path sees all sheet merges, so the mirror
    block's independently-computed plan must too — otherwise the pinned
    mirror-plan == flat-plan equality would break.
    """

    profile = make_sheet_profile(
        name="S",
        max_row=10,
        max_col=3,
        merges=[
            MergeRegion(range="A2:A3", kind="body"),
            MergeRegion(range="A8:A9", kind="body"),
        ],
    )
    mirror = _block(1, 10, header=1, index=0)

    plan = build_block_read_plan(profile, mirror, band_scoped=False)

    assert set(plan.notes) == {
        "body merge A2:A3 -> forward-fill top-left value (spec §4.4)",
        "body merge A8:A9 -> forward-fill top-left value (spec §4.4)",
    }
