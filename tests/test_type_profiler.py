"""Type Profiler tests (spec §4.6, §7.3, Phase 5) [D4][D5].

Three layers, mirroring the Boundary Detector suite:

1.  **Loader-backed fixture tests** run the wired
    ``SheetEnumerator -> HeaderLocator -> BoundaryDetector -> TypeProfiler``
    chain over the real data-mode workbook and assert the §7.3 per-column
    classification (``number`` / ``numeric_text`` / ``date`` / ``mixed`` /
    ``text``), the 0-based ``index`` from the table top-left [D5], the column
    ``name`` from the header row, and ``null_ratio``.
2.  **Isolated partial-context tests** drive the analyzer through ``conftest``
    synthesis with a fake data-mode loader so the missing-cell handling, the
    null_ratio denominator, the skip-row exclusion, the deterministic even
    sampling, and the headerless / no-data-region branches can be exercised
    without touching disk.
3.  **Helper unit tests** pin the classification order and the even-sample
    index rule directly.

End-to-end the produced ``dtype_map`` is asserted on the assembled
:class:`ReadPlan`.

Coordinates here are **openpyxl 1-based** (the inspection domain [D1]); the
single 1-based -> 0-based conversion is the aggregator's job.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

from excel_inspector import InspectionOptions, Loader, SheetOverride, inspect
from excel_inspector.analyzers.boundary_detector import BoundaryDetector
from excel_inspector.analyzers.header_locator import HeaderLocator
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.analyzers.type_profiler import (
    TypeProfiler,
    _classify_column,
    _even_sample_indices,
)
from excel_inspector.context import InspectionContext
from excel_inspector.models import WorkbookProfile

from conftest import make_context, make_sheet_profile  # type: ignore[import-not-found]


def _run_on(
    path: Path, options: InspectionOptions | None = None
) -> InspectionContext:
    """Enumerate -> headers -> boundaries -> types; return the context."""

    context = InspectionContext(
        options=options or InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)
        HeaderLocator().analyze(context)
        BoundaryDetector().analyze(context)
        return TypeProfiler().analyze(context)


# ---------------------------------------------------------------------------
# Identity
# ---------------------------------------------------------------------------


def test_name() -> None:
    """The analyzer reports a stable identifier."""

    assert TypeProfiler().name() == "type_profiler"


# ---------------------------------------------------------------------------
# Loader-backed fixture classification (spec §7.3)
# ---------------------------------------------------------------------------


def test_types_mixed_column_classification(fixture_path) -> None:
    """types_mixed: id=number, code=numeric_text, date=date, mixed=mixed (§7.3).

    Column B holds digit strings stored as Excel text ("007", ...): every value
    parses as a number but is stored as a string -> ``numeric_text``. Column C
    holds native datetimes -> ``date``. Column D interleaves ints and arbitrary
    strings -> ``mixed``. Column A is native ints -> ``number``. ``index`` is
    0-based from the table top-left and ``name`` is the header label [D5].
    """

    sheet = _run_on(fixture_path("types_mixed")).workbook_profile.sheets[0]
    by_index = {c.index: c for c in sheet.columns}

    assert [c.index for c in sheet.columns] == [0, 1, 2, 3]
    assert by_index[0].name == "id"
    assert by_index[0].inferred_type == "number"
    assert by_index[1].name == "code"
    assert by_index[1].inferred_type == "numeric_text"
    assert by_index[2].name == "date"
    assert by_index[2].inferred_type == "date"
    assert by_index[3].name == "mixed"
    assert by_index[3].inferred_type == "mixed"


def test_types_mixed_null_ratio_all_populated(fixture_path) -> None:
    """types_mixed has no missing cells in the data region -> null_ratio 0.0."""

    sheet = _run_on(fixture_path("types_mixed")).workbook_profile.sheets[0]
    assert all(c.null_ratio == 0.0 for c in sheet.columns)


def test_header_simple_text_and_number_columns(fixture_path) -> None:
    """header_simple: name/city are text, age/score are number (§7.3)."""

    sheet = _run_on(fixture_path("header_simple")).workbook_profile.sheets[0]
    types = {c.name: c.inferred_type for c in sheet.columns}
    assert types == {
        "name": "text",
        "age": "number",
        "city": "text",
        "score": "number",
    }


def test_left_margin_cols_index_is_table_relative(fixture_path) -> None:
    """left_margin_cols: ColumnProfile.index is 0-based from the table top-left.

    The table occupies sheet columns C-E (usecols 'C:E'); the profiled columns
    must be indexed 0,1,2 (NOT 2,3,4) so the dtype_map key equals the
    usecols-selected frame position [D5]. Names come from the C1:E1 header.
    """

    sheet = _run_on(fixture_path("left_margin_cols")).workbook_profile.sheets[0]
    assert [(c.index, c.name) for c in sheet.columns] == [
        (0, "sku"),
        (1, "qty"),
        (2, "price"),
    ]
    assert sheet.columns[0].inferred_type == "text"  # sku
    assert sheet.columns[1].inferred_type == "number"  # qty
    assert sheet.columns[2].inferred_type == "number"  # price


def test_offset_plus_subtotals_skips_excluded_from_sample(fixture_path) -> None:
    """Subtotal/total rows are excluded before sampling -> clean dept column.

    offset_plus_subtotals has subtotal rows 8/12 ('소계') and grand-total row 13
    ('합계') folded into ``skip_rows``. They must be removed before sampling, so
    column A ('dept') sees only the real data labels ('영업'/'관리') and stays a
    clean ``text`` column (the subtotal '소계'/'합계' labels never enter the
    sample, and column A has no missing cells in the data region).
    """

    sheet = _run_on(
        fixture_path("offset_plus_subtotals")
    ).workbook_profile.sheets[0]
    dept = sheet.columns[0]
    assert dept.name == "dept"
    assert dept.inferred_type == "text"
    assert dept.null_ratio == 0.0


def test_no_data_region_sheet_has_no_columns(fixture_path) -> None:
    """header_only: no resolved data region -> no profiled columns (§5.3)."""

    sheet = _run_on(fixture_path("header_only")).workbook_profile.sheets[0]
    assert sheet.data_start_row is None
    assert sheet.columns == []


# ---------------------------------------------------------------------------
# End-to-end dtype_map [D5]
# ---------------------------------------------------------------------------


def test_end_to_end_dtype_map_from_inferred_types(fixture_path) -> None:
    """inspect() maps inferred types to the plan's dtype_map [D5] (spec §4.8).

    number -> omitted; numeric_text/text -> 'string'; date -> 'datetime64[ns]';
    mixed -> 'object'. Keys are the 0-based table-relative positions.
    """

    sheet = inspect(fixture_path("types_mixed")).sheets[0]
    plan = sheet.read_plan
    assert plan is not None
    assert plan.dtype_map == {
        "1": "string",  # code (numeric_text)
        "2": "datetime64[ns]",  # date
        "3": "object",  # mixed
    }
    # id (number, position 0) is intentionally absent.
    assert "0" not in plan.dtype_map


def test_end_to_end_dtype_force_wins_over_inferred(fixture_path) -> None:
    """A dtype_force override wins per key over the inferred dtype [D5]."""

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(dtype_force={"1": "Int64"})}
    )
    sheet = inspect(fixture_path("types_mixed"), options).sheets[0]
    plan = sheet.read_plan
    assert plan is not None
    # Forced "1" overrides the inferred "string"; the rest stay inferred.
    assert plan.dtype_map["1"] == "Int64"
    assert plan.dtype_map["2"] == "datetime64[ns]"
    assert plan.dtype_map["3"] == "object"


def test_end_to_end_dtype_map_loads_in_pandas(fixture_path) -> None:
    """The inferred dtype_map is loadable: numeric_text keeps leading zeros."""

    from excel_inspector.adapters import load_dataframe

    path = fixture_path("types_mixed")
    plan = inspect(path).sheets[0].read_plan
    assert plan is not None

    df = load_dataframe(path, plan)
    # numeric_text column "code" keeps its leading zeros as strings.
    assert df["code"].tolist() == ["007", "012", "034", "056", "078", "090"]
    # date column is a datetime dtype.
    assert str(df["date"].dtype).startswith("datetime64")


# ---------------------------------------------------------------------------
# Isolated partial-context tests with a fake data-mode loader
# ---------------------------------------------------------------------------


class _FakeWorksheet:
    """A read_only-style worksheet returning fixed 1-based-aligned rows."""

    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows

    def iter_rows(self, *, min_row, max_row, values_only):  # noqa: ANN001, ANN202
        end = (
            len(self._rows)
            if max_row is None
            else min(max_row, len(self._rows))
        )
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


def test_isolated_numeric_text_vs_number() -> None:
    """A digit-string text column -> numeric_text; native ints -> number (§7.3)."""

    rows = [
        ["id", "code"],  # row 1 header
        [1, "007"],  # row 2
        [2, "012"],  # row 3
        [3, "034"],  # row 4
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=1,
        data_start_row=2,
        data_end_row=4,
        max_row=len(rows),
        max_col=2,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    assert profile.columns[0].inferred_type == "number"
    assert profile.columns[1].inferred_type == "numeric_text"


def test_isolated_date_column() -> None:
    """A column of native datetimes -> date (§7.3)."""

    rows = [
        ["when"],  # row 1 header
        [_dt.datetime(2026, 1, 1)],  # row 2
        [_dt.date(2026, 1, 2)],  # row 3
        [_dt.datetime(2026, 1, 3)],  # row 4
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=1,
        data_start_row=2,
        data_end_row=4,
        max_row=len(rows),
        max_col=1,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    assert profile.columns[0].inferred_type == "date"


def test_isolated_mixed_column() -> None:
    """Interleaved ints and arbitrary strings -> mixed (no type >= 0.95) (§7.3)."""

    rows = [
        ["m"],  # row 1 header
        [100],  # row 2
        ["N/A"],  # row 3
        [300],  # row 4
        ["pending"],  # row 5
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=1,
        data_start_row=2,
        data_end_row=5,
        max_row=len(rows),
        max_col=1,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    assert profile.columns[0].inferred_type == "mixed"


def test_isolated_null_ratio_denominator_is_sample_row_count() -> None:
    """null_ratio denominator = number of sampled data rows (§5.3 / §7.3).

    Two of four sampled data rows are missing in the value column -> 0.5; the
    type is still inferred from the present cells (here both numbers -> number).
    """

    rows = [
        ["k", "v"],  # row 1 header
        ["a", 10],  # row 2
        ["b", None],  # row 3 missing
        ["c", 30],  # row 4
        ["d", None],  # row 5 missing
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=1,
        data_start_row=2,
        data_end_row=5,
        max_row=len(rows),
        max_col=2,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    value_col = profile.columns[1]
    assert value_col.null_ratio == 0.5
    assert value_col.inferred_type == "number"


def test_isolated_skip_rows_excluded_from_sample() -> None:
    """Interior skip rows are removed before sampling and from the denominator.

    Row 3 is an interior skip (subtotal). It must not enter the sample, so the
    'amount' column sees only the two real data rows (100, 200) -> number, and
    the null_ratio denominator is 2 (not 3).
    """

    rows = [
        ["dept", "amount"],  # row 1 header
        ["영업", 100],  # row 2 data
        ["소계", 300],  # row 3 interior skip
        ["관리", 200],  # row 4 data
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=1,
        data_start_row=2,
        data_end_row=4,
        skip_rows=[3],
        max_row=len(rows),
        max_col=2,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    amount = profile.columns[1]
    assert amount.inferred_type == "number"
    # Denominator excludes the skipped row -> 2 sampled rows, both present.
    assert amount.null_ratio == 0.0
    dept = profile.columns[0]
    assert dept.inferred_type == "text"  # '소계' never sampled


def test_isolated_headerless_sheet_names_are_none() -> None:
    """A header_row=None sheet profiles columns with name=None (no name source)."""

    rows = [
        [1, "a"],  # row 1 data
        [2, "b"],  # row 2 data
        [3, "c"],  # row 3 data
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=None,
        data_start_row=1,
        data_end_row=3,
        max_row=len(rows),
        max_col=2,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    assert [c.name for c in profile.columns] == [None, None]
    assert profile.columns[0].inferred_type == "number"
    assert profile.columns[1].inferred_type == "text"


def test_isolated_no_data_region_is_skipped() -> None:
    """A sheet without a resolved data region profiles no columns (§5.3)."""

    profile = make_sheet_profile(
        name="S", header_row=1, data_start_row=None, data_end_row=None, max_col=3
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": []}))

    TypeProfiler().analyze(ctx)
    assert profile.columns == []


def test_isolated_non_tabular_sheet_is_skipped() -> None:
    """Non-tabular sheets are never profiled (spec §9)."""

    profile = make_sheet_profile(
        name="README",
        is_tabular_candidate=False,
        header_row=1,
        data_start_row=2,
        data_end_row=4,
        max_col=3,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"README": []}))

    TypeProfiler().analyze(ctx)
    assert profile.columns == []


def test_no_loader_warns_and_skips() -> None:
    """With no loader the analyzer warns and profiles nothing."""

    profile = make_sheet_profile(
        name="S", header_row=1, data_start_row=2, data_end_row=4, max_col=3
    )
    ctx = make_context(sheets=[profile], loader=None)

    TypeProfiler().analyze(ctx)
    assert profile.columns == []
    assert any("type_profiler" in w for w in ctx.warnings)


def test_isolated_partial_span_index_offset() -> None:
    """A partial column span yields table-relative indices and offset names.

    The table occupies sheet columns C-E (left=3, right=5); the profiled
    columns must be index 0,1,2 with names from the C..E header cells, NOT the
    sheet-absolute positions [D5].
    """

    rows = [
        ["filler", None, "sku", "qty", "price"],  # row 1 header (C-E table)
        ["x", None, "A-1", 10, 1.5],  # row 2
        ["y", None, "A-2", 20, 2.5],  # row 3
    ]
    profile = make_sheet_profile(
        name="S",
        header_row=1,
        data_start_row=2,
        data_end_row=3,
        data_left_col=3,
        data_right_col=5,
        max_row=len(rows),
        max_col=5,
    )
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    TypeProfiler().analyze(ctx)
    assert [(c.index, c.name, c.inferred_type) for c in profile.columns] == [
        (0, "sku", "text"),
        (1, "qty", "number"),
        (2, "price", "number"),
    ]


# ---------------------------------------------------------------------------
# Helper unit tests (§7.3 components)
# ---------------------------------------------------------------------------


def test_classify_column_order() -> None:
    """The §7.3 judgment order classifies each canonical column shape."""

    # All native ints -> number.
    assert _classify_column([1, 2, 3], 3) == "number"
    # All digit strings stored as text -> numeric_text.
    assert _classify_column(["1", "2", "3"], 3) == "numeric_text"
    # All datetimes -> date.
    assert (
        _classify_column(
            [_dt.datetime(2026, 1, 1), _dt.date(2026, 1, 2)], 2
        )
        == "date"
    )
    # All non-numeric strings -> text.
    assert _classify_column(["a", "b", "c"], 3) == "text"
    # Ints + arbitrary strings -> mixed.
    assert _classify_column([1, "x", 2, "y"], 4) == "mixed"


def test_classify_column_missing_excluded() -> None:
    """Missing cells are dropped before classification (denominator = present)."""

    # One missing among native ints -> still number (missing excluded).
    assert _classify_column([1, None, 3], 3) == "number"
    # All-missing column commits to text (no numeric/date evidence).
    assert _classify_column([None, "", None], 3) == "text"


def test_classify_column_numeric_text_requires_all_string_storage() -> None:
    """numeric_text only when EVERY non-missing value is stored as a string.

    A column mixing a native int with digit strings still parses fully as
    numbers, but the native int storage makes it a plain ``number`` (not
    ``numeric_text``).
    """

    assert _classify_column(["1", 2, "3"], 3) == "number"
    assert _classify_column(["1", "2", "3"], 3) == "numeric_text"


def test_even_sample_indices_deterministic_and_unique() -> None:
    """Even sampling is deterministic, ascending, unique, and capped (§7.3)."""

    # sample_size >= count -> the whole population.
    assert _even_sample_indices(3, 10) == [0, 1, 2]
    # Even spread; ascending and unique.
    picks = _even_sample_indices(10, 5)
    assert picks == [0, 2, 4, 6, 8]
    assert picks == sorted(set(picks))
    # Degenerate inputs.
    assert _even_sample_indices(0, 5) == []
    assert _even_sample_indices(5, 0) == []


# ---------------------------------------------------------------------------
# Phase 10b-1: block-local _profile_block core (plan v2 Task 10.2 Step 1)
# ---------------------------------------------------------------------------


def test_profile_block_explicit_boundaries_and_no_mutation() -> None:
    """_profile_block profiles from block-local parameters and mutates nothing.

    The boundaries are explicit arguments (not profile reads) and the produced
    columns are returned, not applied — the shared profile's ``columns`` stays
    untouched (the sheet-level applier's job).
    """

    rows = [
        ["제품명", "단가"],  # row 1: block header
        ["키보드", 30000],  # row 2
        ["마우스", 15000],  # row 3
    ]
    profile = make_sheet_profile(name="S", max_row=len(rows), max_col=2)
    ctx = make_context(sheets=[profile], loader=_FakeLoader({"S": rows}))

    columns = TypeProfiler()._profile_block(
        ctx,
        profile,
        header_row=1,
        data_start_row=2,
        data_end_row=3,
        skip_rows=[],
        data_left_col=None,
        data_right_col=None,
    )

    assert columns is not None
    assert [(c.index, c.name, c.inferred_type) for c in columns] == [
        (0, "제품명", "text"),
        (1, "단가", "number"),
    ]
    # The core never applies to the shared profile.
    assert profile.columns == []


def test_sampled_row_numbers_row_window_clamp() -> None:
    """row_window clamps the eligible span; an empty intersection samples nothing.

    Clamp rule (plan v2 Task 10.2 Step 1): eligible rows are
    ``max(data_start_row, window_start) .. min(data_end_row, window_end)``
    minus the skip rows; ``row_window=None`` keeps the v1 behavior.
    """

    full = TypeProfiler._sampled_row_numbers(2, 10, [5])
    assert full == [2, 3, 4, 6, 7, 8, 9, 10]

    clamped = TypeProfiler._sampled_row_numbers(2, 10, [5], row_window=(4, 8))
    assert clamped == [4, 6, 7, 8]

    assert TypeProfiler._sampled_row_numbers(2, 10, [], row_window=(20, 30)) == []
