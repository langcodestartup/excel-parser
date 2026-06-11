"""Sheet Enumerator tests (spec §4.2, Phase 1).

Exercises the analyzer against the real structure-mode loader on fixtures
(visibility, used range, max row/col, dimension trust, tabular-candidate
classification) and against synthetic loaders for the override path and the
no-loader degradation path.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_inspector import (
    InspectionOptions,
    Loader,
    SheetOverride,
)
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.context import InspectionContext
from excel_inspector.exceptions import CorruptWorkbookError
from excel_inspector.models import WorkbookProfile

from conftest import make_context  # type: ignore[import-not-found]


def _run_on(path: Path, options: InspectionOptions | None = None) -> InspectionContext:
    """Run the enumerator over a fixture and return the enriched context."""

    context = InspectionContext(
        options=options or InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        return SheetEnumerator().analyze(context)


def test_name() -> None:
    """The analyzer reports a stable identifier."""

    assert SheetEnumerator().name() == "sheet_enumerator"


def test_header_simple_metadata(fixture_path) -> None:
    """header_simple: one visible tabular sheet, A1:D6, trusted dims."""

    ctx = _run_on(fixture_path("header_simple"))
    sheets = ctx.workbook_profile.sheets
    assert [s.name for s in sheets] == ["Sheet1"]

    s = sheets[0]
    assert s.is_visible is True
    assert s.is_tabular_candidate is True
    assert s.used_range == "A1:D6"
    assert (s.max_row, s.max_col) == (6, 4)
    assert s.used_range_trusted is True


def test_left_margin_cols_dimensions(fixture_path) -> None:
    """left_margin_cols: used range spans the filler column too (A1:E7)."""

    ctx = _run_on(fixture_path("left_margin_cols"))
    s = ctx.workbook_profile.sheets[0]
    # Sheet enumerator reports the raw used range; column boundary detection
    # (C-E) is a later phase. max_col includes the table's rightmost column E.
    assert (s.max_row, s.max_col) == (7, 5)
    assert s.used_range == "A1:E7"


def test_mixed_sheets_visibility_and_tabular(fixture_path) -> None:
    """mixed_sheets: README is non-tabular; Data is tabular; order preserved."""

    ctx = _run_on(fixture_path("mixed_sheets"))
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    assert [s.name for s in ctx.workbook_profile.sheets] == ["README", "Data"]

    readme = sheets["README"]
    assert readme.is_visible is True
    assert readme.is_tabular_candidate is False  # single sparse text column
    assert readme.max_col == 1

    data = sheets["Data"]
    assert data.is_tabular_candidate is True
    assert (data.max_row, data.max_col) == (5, 3)
    assert data.used_range == "A1:C5"


def test_empty_sheet_is_non_tabular(fixture_path) -> None:
    """empty_sheet: openpyxl reports a 1x1 range; flagged non-tabular."""

    ctx = _run_on(fixture_path("empty_sheet"))
    s = ctx.workbook_profile.sheets[0]
    assert (s.max_row, s.max_col) == (1, 1)
    assert s.used_range == "A1:A1"
    assert s.is_tabular_candidate is False


def test_hidden_and_very_hidden_sheets_report_not_visible(fixture_path) -> None:
    """hidden_sheet: hidden/veryHidden sheets are is_visible False (issue #6)."""

    ctx = _run_on(fixture_path("hidden_sheet"))
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    assert [s.name for s in ctx.workbook_profile.sheets] == [
        "Visible",
        "Hidden",
        "VeryHidden",
    ]
    assert sheets["Visible"].is_visible is True
    assert sheets["Hidden"].is_visible is False
    assert sheets["VeryHidden"].is_visible is False


def test_is_tabular_override_forces_true(fixture_path) -> None:
    """is_tabular override [D2] overrides the heuristic for the README sheet."""

    options = InspectionOptions(
        sheet_overrides={"README": SheetOverride(is_tabular=True)}
    )
    ctx = _run_on(fixture_path("mixed_sheets"), options)
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    assert sheets["README"].is_tabular_candidate is True
    # Override is recorded with provenance=manual [D2].
    assert sheets["README"].is_tabular_provenance == "manual"
    # The non-overridden sheet keeps heuristic provenance.
    assert sheets["Data"].is_tabular_provenance == "heuristic"


def test_is_tabular_override_forces_false(fixture_path) -> None:
    """is_tabular override [D2] can demote an otherwise-tabular sheet."""

    options = InspectionOptions(
        sheet_overrides={"Data": SheetOverride(is_tabular=False)}
    )
    ctx = _run_on(fixture_path("mixed_sheets"), options)
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    assert sheets["Data"].is_tabular_candidate is False


def test_no_loader_warns_and_skips() -> None:
    """With no loader the analyzer records a warning and skips (spec §9)."""

    context = make_context(loader=None)
    result = SheetEnumerator().analyze(context)
    assert result.workbook_profile.sheets == []
    assert any("sheet_enumerator" in w for w in result.warnings)


# ---------------------------------------------------------------------------
# Content-aware tabular gate (issue #3): populated columns + density
# ---------------------------------------------------------------------------


class _FakeWorksheet:
    """Minimal read_only-style worksheet yielding canned rows for sampling."""

    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows

    def iter_rows(self, min_row=1, max_row=None, values_only=False):
        end = len(self._rows) if max_row is None else max_row
        for row in self._rows[min_row - 1 : end]:
            yield tuple(row)


class _FakeWorkbook:
    def __init__(self, sheets: dict[str, _FakeWorksheet]) -> None:
        self._sheets = sheets

    def __getitem__(self, name: str) -> _FakeWorksheet:
        return self._sheets[name]


class _FakeLoader:
    def __init__(self, sheets: dict[str, _FakeWorksheet]) -> None:
        self._wb = _FakeWorkbook(sheets)

    def data_workbook(self) -> _FakeWorkbook:
        return self._wb


class _BoomLoader:
    def data_workbook(self):
        raise RuntimeError("boom")


class _CorruptLoader:
    def data_workbook(self):
        raise CorruptWorkbookError("boom")


def test_offset_cover_is_non_tabular(fixture_path) -> None:
    """cover_offset: text in column B is one populated column -> non-tabular."""

    ctx = _run_on(fixture_path("cover_offset"))
    s = ctx.workbook_profile.sheets[0]
    assert s.max_col == 2  # the offset that fooled the legacy max_col>1 gate
    assert s.is_tabular_candidate is False
    assert s.is_tabular_provenance == "heuristic"


def test_sparse_cover_is_non_tabular(fixture_path) -> None:
    """cover_sparse: 3 scattered columns, density 0.333 < 0.5 -> non-tabular."""

    ctx = _run_on(fixture_path("cover_sparse"))
    s = ctx.workbook_profile.sheets[0]
    assert s.max_col == 5
    assert s.is_tabular_candidate is False


@pytest.mark.parametrize(
    "fixture_id,sheet,expected",
    [
        ("stacked_uneven_width", "Sheet1", True),  # corpus-floor density (0.688)
        ("sparse_real_table", "Sheet1", True),  # sparse real table density 0.583 > 0.5
        ("no_header", "Sheet1", True),  # all-numeric real table (pop_cols=3)
        ("header_only", "Sheet1", True),  # header-only real table (pop_cols=3)
        ("hidden_sheet", "Hidden", True),  # 2-column table, high density
        ("mixed_sheets", "README", False),  # single sparse text column
        ("empty_sheet", "Sheet1", False),  # empty sheet
        ("cover_offset", "표지", False),  # issue #3: B-offset single column
        ("cover_sparse", "표지", False),  # multi-column but sparse (density 0.333)
    ],
)
def test_tabular_classification(fixture_path, fixture_id, sheet, expected) -> None:
    """Whole-sheet tabular gate stays correct across the corpus (regression)."""

    ctx = _run_on(fixture_path(fixture_id))
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    assert sheets[sheet].is_tabular_candidate is expected


def test_empty_sample_falls_back_to_dimensions() -> None:
    """No content in the sample -> defer to the legacy max_col>1 dims rule."""

    loader = _FakeLoader({"S": _FakeWorksheet([])})
    ctx = make_context(loader=loader)
    result, prov = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=10, max_col=3
    )
    assert result is True and prov == "heuristic"


def test_sampling_failure_falls_back_with_warning() -> None:
    """A sampling exception must not break enumeration (spec §6): fall back."""

    ctx = make_context(loader=_BoomLoader())
    result, prov = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=10, max_col=1
    )
    assert result is False and prov == "heuristic"  # dims fallback: 1 not > 1
    assert any("sheet_enumerator" in w and "S" in w for w in ctx.warnings)


def test_inspector_error_propagates_not_swallowed() -> None:
    """Loader domain errors (corrupt/encrypted) must propagate, not be absorbed
    into a warning + dims fallback (spec §6/§9; consistent with pipeline.py)."""

    ctx = make_context(loader=_CorruptLoader())
    with pytest.raises(CorruptWorkbookError):
        SheetEnumerator()._is_tabular_candidate(ctx, "S", max_row=10, max_col=2)
    assert ctx.warnings == []  # not downgraded to a warning


def test_density_rule_counts_only_content_rows() -> None:
    """density = filled/(populated_cols*populated_rows); a blank middle row is
    NOT counted in populated_rows. Here pc=2, content_rows=2, filled=3 ->
    0.75 >= 0.5 -> tabular (pins the populated_rows denominator term)."""

    rows = [["a", "b"], [None, None], ["c", None]]
    loader = _FakeLoader({"S": _FakeWorksheet(rows)})
    ctx = make_context(loader=loader)
    result, prov = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=3, max_col=2
    )
    assert result is True and prov == "heuristic"


def test_density_rule_low_density_is_non_tabular() -> None:
    """pc=3, content_rows=3, filled=3 -> 3/9=0.333 < 0.5 -> non-tabular
    (exercises the density branch directly via the fake loader)."""

    rows = [["a", None, None], [None, None, "b"], [None, "c", None]]
    loader = _FakeLoader({"S": _FakeWorksheet(rows)})
    ctx = make_context(loader=loader)
    result, _ = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=3, max_col=3
    )
    assert result is False
