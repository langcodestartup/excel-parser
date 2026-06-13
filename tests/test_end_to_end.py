"""End-to-end inspect() tests (spec §3, Phase 1 exit criteria).

Drives the full v1 pipeline (Loader -> SheetEnumerator -> PlanAggregator) via
the public :func:`inspect` entry point and asserts the produced
:class:`WorkbookProfile`: sheet listing, used ranges, basic v1 read plans, and
the corrupt/encrypted domain-exception paths.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_inspector import (
    CorruptWorkbookError,
    EncryptedWorkbookError,
    InspectionOptions,
    SheetOverride,
    extract,
    inspect,
)

_REAL_XLSX_DIR = Path(__file__).resolve().parent / "real_xlsx"


def test_inspect_header_simple(fixture_path) -> None:
    """A normal table yields one tabular sheet with a v1 read plan."""

    path = fixture_path("header_simple")
    profile = inspect(path)

    assert profile.file_path == str(path)
    assert [s.name for s in profile.sheets] == ["Sheet1"]

    sheet = profile.sheets[0]
    assert sheet.used_range == "A1:D6"
    assert (sheet.max_row, sheet.max_col) == (6, 4)

    plan = sheet.read_plan
    assert plan is not None
    assert plan.sheet_name == "Sheet1"
    assert plan.engine == "openpyxl"
    assert plan.header == 0  # v1: first row assumed header (0-based)
    assert plan.skiprows == []
    assert plan.usecols is None


def test_inspect_mixed_sheets_splits_tabular(fixture_path) -> None:
    """README (non-tabular) gets no plan; Data (tabular) gets one."""

    profile = inspect(fixture_path("mixed_sheets"))
    sheets = {s.name: s for s in profile.sheets}

    assert sheets["README"].is_tabular_candidate is False
    assert sheets["README"].read_plan is None

    assert sheets["Data"].is_tabular_candidate is True
    assert sheets["Data"].read_plan is not None
    assert sheets["Data"].read_plan.header == 0


def test_inspect_empty_sheet_has_no_plan(fixture_path) -> None:
    """An empty sheet is non-tabular and carries no read plan."""

    profile = inspect(fixture_path("empty_sheet"))
    sheet = profile.sheets[0]
    assert sheet.is_tabular_candidate is False
    assert sheet.read_plan is None


def test_inspect_with_dtype_force_override(fixture_path) -> None:
    """dtype_force [D5] flows through inspect() into the read plan.

    With the Type Profiler (Phase 5) now wired in, the dtype_map also carries
    inferred dtypes for the text columns of ``header_simple`` (``name`` at
    position ``0`` and ``city`` at position ``2`` are both text -> ``"string"``;
    the numeric ``age``/``score`` columns are omitted). The ``dtype_force``
    override is applied on top and wins per key, so position ``"0"`` is the
    forced ``"string"``.
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(dtype_force={"0": "string"})}
    )
    profile = inspect(fixture_path("header_simple"), options)
    plan = profile.sheets[0].read_plan
    assert plan is not None
    # Inferred text columns (0=name, 2=city) plus the forced override on 0.
    assert plan.dtype_map == {"0": "string", "2": "string"}


def test_inspect_header_override_records_manual_provenance(
    fixture_path,
) -> None:
    """A header_row override flows through inspect() as provenance=manual [D2].

    header_offset has a real header on 1-based row 4; forcing it via override
    must mark the sheet ``header_provenance='manual'`` with confidence 1.0 and
    drive the skiprows conversion (rows 1-3 absorbed).
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=4)}
    )
    profile = inspect(fixture_path("header_offset"), options)
    sheet = profile.sheets[0]

    assert sheet.header_provenance == "manual"
    assert sheet.header_confidence == 1.0
    assert sheet.header_row == 4
    assert sheet.read_plan is not None
    assert sheet.read_plan.skiprows == [0, 1, 2]
    assert sheet.read_plan.header == 0


def test_inspect_heuristic_header_provenance(fixture_path) -> None:
    """The wired HeaderLocator (Phase 2) detects header_simple's row-1 header.

    With the header analyzer now in the pipeline, an estimated header carries
    ``provenance="heuristic"`` (no override). The honest ``"default"`` fallback
    only applies when no header is detected and none is overridden.
    """

    profile = inspect(fixture_path("header_simple"))
    sheet = profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.header_provenance == "heuristic"
    assert sheet.needs_manual_header is False


def test_inspect_headerless_override_plan_header_none(fixture_path) -> None:
    """no_header + SheetOverride(header_row=None) -> plan.header is None (HIGH #3).

    The no_header fixture is pure data (no header row); declaring it headerless
    via override must produce a plan with ``header=None`` so the first data row
    is preserved (not consumed as column names), per spec §9.
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    profile = inspect(fixture_path("no_header"), options)
    sheet = profile.sheets[0]

    assert sheet.header_row is None
    assert sheet.header_provenance == "manual"
    assert sheet.read_plan is not None
    assert sheet.read_plan.header is None


def test_inspect_headerless_override_loads_first_row_as_data(
    fixture_path,
) -> None:
    """The headerless plan loads the first data row as data, not a header.

    End-to-end: feeding the headerless plan into pandas must yield as many rows
    as there are data rows in the fixture (5), with the first row preserved.
    """

    from excel_inspector.adapters import load_dataframe

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    path = fixture_path("no_header")
    sheet = inspect(path, options).sheets[0]
    plan = sheet.read_plan
    assert plan is not None
    assert plan.header is None

    df = load_dataframe(path, plan)
    # All 5 data rows present; the first row (1, 100, 1.1) was NOT eaten as a
    # header, so the body still starts with the integer 1.
    assert len(df) == 5
    assert df.iloc[0, 0] == 1


def test_inspect_corrupt_raises(fixture_path) -> None:
    """A corrupt file aborts inspection with CorruptWorkbookError (spec §9)."""

    with pytest.raises(CorruptWorkbookError):
        inspect(fixture_path("corrupt"))


def test_inspect_encrypted_raises(fixture_path) -> None:
    """An encrypted file aborts inspection with EncryptedWorkbookError."""

    with pytest.raises(EncryptedWorkbookError):
        inspect(fixture_path("encrypted"))


def test_inspect_every_openable_fixture(openable_fixture) -> None:
    """inspect() runs without error over every openable fixture."""

    profile = inspect(openable_fixture)
    assert profile.file_path == str(openable_fixture)
    assert isinstance(profile.sheets, list)
    assert len(profile.sheets) >= 1


@pytest.mark.parametrize(
    "filename,expected_rows,expected_cols",
    [
        ("bis_pp_selected.xlsx", 387, 249),
        ("bis_totcredit.xlsx", 333, 1134),
    ],
)
def test_bis_quarterly_series_wide_sparse_extracted(
    filename: str, expected_rows: int, expected_cols: int
) -> None:
    """issue #22: 실전 BIS Quarterly Series는 non-tabular로 누락되지 않는다."""

    wr = extract(_REAL_XLSX_DIR / filename)
    entry = {sheet.name: sheet for sheet in wr.sheets}["Quarterly Series"]

    assert entry.skipped is False
    assert entry.skip_reason is None
    (table,) = entry.tables
    assert table.header_row == 4
    assert table.dataframe.shape == (expected_rows, expected_cols)
    assert table.dataframe.columns[0] == "Period"
