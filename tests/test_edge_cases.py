"""Edge-case robustness tests (spec §8, §9, Phase 8).

The inspector must never raise on a *structurally* awkward but openable
workbook; degenerate sheets are surfaced as explicit states (``read_plan`` /
``data_start_row`` / ``needs_manual_header`` / ``is_tabular_candidate``) plus
accumulated ``warnings``/``open_errors``, not exceptions (spec §6). Only the
loader's corrupt/encrypted domain errors abort (spec §9).

This file pins the end-to-end behavior for each degenerate fixture:

* ``empty_sheet``    — no usable area: non-tabular, no read plan.
* ``header_only``    — header but zero data rows: tabular, no resolved data
  region, header-only-style plan.
* ``no_header``      — pure data, header heuristic fails: ``needs_manual_header``.
* ``mixed_sheets``   — a README (non-tabular) alongside a real table.
* ``hidden_sheet``   — visibility is reported but never blocks inspection.
"""

from __future__ import annotations

import pytest

from excel_inspector import (
    CorruptWorkbookError,
    EncryptedWorkbookError,
    inspect,
)


def test_empty_sheet_explicit_state(fixture_path) -> None:
    """An empty sheet is non-tabular with no read plan and no exception."""

    profile = inspect(fixture_path("empty_sheet"))
    assert profile.open_errors == [] or isinstance(profile.open_errors, list)

    sheet = profile.sheets[0]
    assert sheet.is_tabular_candidate is False
    assert sheet.read_plan is None
    assert sheet.data_start_row is None
    assert sheet.data_end_row is None


def test_header_only_explicit_state(fixture_path) -> None:
    """A header-only sheet resolves no data region but still gets a plan.

    The sheet is tabular (3 columns), the header is detectable, but there are
    zero data rows below it, so ``data_start_row`` / ``data_end_row`` stay
    ``None`` and the read plan carries ``nrows=None`` (a header-only plan). No
    exception is raised.
    """

    profile = inspect(fixture_path("header_only"))
    sheet = profile.sheets[0]

    assert sheet.is_tabular_candidate is True
    assert sheet.data_start_row is None
    assert sheet.data_end_row is None
    assert sheet.read_plan is not None
    assert sheet.read_plan.nrows is None
    # No columns profiled (no data region to sample) — the "unknown" sentinel
    # means the aggregator emitted no dtype keys.
    assert sheet.columns == []
    assert sheet.read_plan.dtype_map == {}


def test_no_header_explicit_state(fixture_path) -> None:
    """Pure data with no header fails the heuristic -> needs_manual_header.

    The sheet is still tabular and inspection completes without raising; the
    failure is expressed as ``header_row=None`` + ``needs_manual_header=True``
    + a recorded warning, per spec §9 (not an exception).
    """

    profile = inspect(fixture_path("no_header"))
    sheet = profile.sheets[0]

    assert sheet.is_tabular_candidate is True
    assert sheet.header_row is None
    assert sheet.needs_manual_header is True
    # The low-confidence header is recorded as a warning, surfaced on the
    # workbook profile's open_errors (spec §6).
    assert any("header" in w.lower() for w in profile.open_errors)
    # No header anchor -> boundary detection is skipped -> no resolved region.
    assert sheet.data_start_row is None


def test_mixed_sheets_explicit_state(fixture_path) -> None:
    """A README (non-tabular) and a real table coexist without error."""

    profile = inspect(fixture_path("mixed_sheets"))
    sheets = {s.name: s for s in profile.sheets}

    readme = sheets["README"]
    assert readme.is_tabular_candidate is False
    assert readme.read_plan is None

    data = sheets["Data"]
    assert data.is_tabular_candidate is True
    assert data.read_plan is not None
    assert data.header_row == 1
    assert data.data_start_row == 2


def test_hidden_sheets_inspected_with_visibility(fixture_path) -> None:
    """Hidden / veryHidden sheets are inspected; visibility is just reported.

    Visibility never blocks inspection (spec §4.2): all three sheets appear in
    order, the hidden two report ``is_visible=False``, and every sheet still
    gets a profile (tabular sheets get a read plan).
    """

    profile = inspect(fixture_path("hidden_sheet"))
    by_name = {s.name: s for s in profile.sheets}

    assert [s.name for s in profile.sheets] == [
        "Visible",
        "Hidden",
        "VeryHidden",
    ]
    assert by_name["Visible"].is_visible is True
    assert by_name["Hidden"].is_visible is False
    assert by_name["VeryHidden"].is_visible is False

    # Visibility does not exclude a sheet from loading; each remains tabular and
    # planned (they all have a 2+ column data table).
    for sheet in profile.sheets:
        assert sheet.is_tabular_candidate is True
        assert sheet.read_plan is not None


def test_every_openable_fixture_never_raises(openable_fixture) -> None:
    """No openable fixture makes inspect() raise (spec §6 robustness).

    Degenerate states must be values, not exceptions; the only exceptions are
    the loader's corrupt/encrypted domain errors (covered separately).
    """

    profile = inspect(openable_fixture)
    assert isinstance(profile.open_errors, list)
    assert len(profile.sheets) >= 1


@pytest.mark.parametrize(
    ("fixture_id", "error"),
    [
        ("corrupt", CorruptWorkbookError),
        ("encrypted", EncryptedWorkbookError),
    ],
)
def test_loader_domain_errors_are_the_only_abort(
    fixture_path, fixture_id: str, error: type[Exception]
) -> None:
    """Corrupt/encrypted files are the *only* abort path (spec §9)."""

    with pytest.raises(error):
        inspect(fixture_path(fixture_id))
