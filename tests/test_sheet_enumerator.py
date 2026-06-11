"""Sheet Enumerator tests (spec §4.2, Phase 1).

Exercises the analyzer against the real structure-mode loader on fixtures
(visibility, used range, max row/col, dimension trust, tabular-candidate
classification) and against synthetic loaders for the override path and the
no-loader degradation path.
"""

from __future__ import annotations

from pathlib import Path

from excel_inspector import (
    InspectionOptions,
    Loader,
    SheetOverride,
)
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.context import InspectionContext
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
