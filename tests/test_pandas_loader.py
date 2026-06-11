"""Tests for the ``ReadPlan`` -> pandas adapter (Phase 7) [D1][D5].

Two layers:

1.  **``read_plan_to_kwargs`` unit tests** — the pure field -> kwarg mapping,
    the ``usecols`` "None == all columns" omission, the headerless ``header=None``
    pass-through, and the [D5] dtype-key reduction (string position keys ->
    positional ints).
2.  **``load_dataframe`` golden round-trips** — drive the full ``inspect()``
    pipeline to a real :class:`ReadPlan`, then load each tabular fixture through
    the adapter and assert the load-domain meaning: clean row counts, no
    subtotal/title leakage, ``usecols`` trimming, applied dtypes, and the
    headerless first-row-as-data contract. Aggregation sums (``amount``) are
    pinned so a coordinate slip or a doubled total would fail loudly.

Assertions target meaning (column names, row counts, sums, dtypes) rather than
full-frame equality, per the impl plan §5.2 golden-fragility mitigation.
"""

from __future__ import annotations

import pandas as pd
import pytest

from excel_inspector import (
    InspectionOptions,
    ReadPlan,
    SheetOverride,
    SheetProfile,
    WorkbookProfile,
    inspect,
)
from excel_inspector.adapters import load_dataframe, read_plan_to_kwargs


def _sheet(profile: WorkbookProfile, name: str) -> SheetProfile:
    """Return the named sheet profile, asserting it exists and has a plan."""

    matches = [s for s in profile.sheets if s.name == name]
    assert matches, (
        f"sheet {name!r} not found in {[s.name for s in profile.sheets]}"
    )
    sheet = matches[0]
    assert sheet.read_plan is not None, f"sheet {name!r} has no read plan"
    return sheet


# ---------------------------------------------------------------------------
# read_plan_to_kwargs — the pure field -> kwarg mapping.
# ---------------------------------------------------------------------------


def test_kwargs_full_field_mapping() -> None:
    """Every plan field maps onto the expected read_excel kwarg [D1]."""

    plan = ReadPlan(
        sheet_name="Sheet1",
        engine="openpyxl",
        header=0,
        usecols="C:E",
        skiprows=[0, 1, 2],
        nrows=7,
        dtype_map={"0": "string", "2": "object"},
    )
    kwargs = read_plan_to_kwargs(plan)

    assert kwargs["sheet_name"] == "Sheet1"
    assert kwargs["engine"] == "openpyxl"
    assert kwargs["header"] == 0
    assert kwargs["skiprows"] == [0, 1, 2]
    assert kwargs["nrows"] == 7
    assert kwargs["usecols"] == "C:E"
    # [D5] dtype keys reduced from 0-based position strings to positional ints.
    assert kwargs["dtype"] == {0: "string", 2: "object"}


def test_kwargs_omits_usecols_when_none() -> None:
    """A ``usecols=None`` plan omits the kwarg (None == all columns)."""

    plan = ReadPlan(sheet_name="S", header=0, usecols=None)
    kwargs = read_plan_to_kwargs(plan)
    assert "usecols" not in kwargs


def test_kwargs_omits_dtype_when_empty() -> None:
    """An empty ``dtype_map`` produces no ``dtype`` kwarg."""

    plan = ReadPlan(sheet_name="S", header=0, dtype_map={})
    kwargs = read_plan_to_kwargs(plan)
    assert "dtype" not in kwargs


def test_kwargs_headerless_passes_header_none() -> None:
    """A headerless plan keeps ``header=None`` (spec §9, HIGH #3)."""

    plan = ReadPlan(sheet_name="S", header=None)
    kwargs = read_plan_to_kwargs(plan)
    # header must be present and explicitly None, not dropped.
    assert "header" in kwargs
    assert kwargs["header"] is None


def test_kwargs_always_emits_skiprows_and_nrows() -> None:
    """``skiprows`` (possibly empty) and ``nrows`` (possibly None) always emit."""

    plan = ReadPlan(sheet_name="S", header=0, skiprows=[], nrows=None)
    kwargs = read_plan_to_kwargs(plan)
    assert kwargs["skiprows"] == []
    assert kwargs["nrows"] is None


# ---------------------------------------------------------------------------
# load_dataframe golden round-trips over the tabular corpus.
# ---------------------------------------------------------------------------


def test_load_offset_plus_subtotals_six_rows_no_leak(fixture_path) -> None:
    """offset_plus_subtotals -> 6 clean rows, zero subtotal/total leakage.

    The flagship [D1] case: leading title rows absorbed, interior subtotals and
    the trailing grand total dropped. The re-aggregated ``amount`` equals the
    file's own grand total (590), proving no aggregation double-count.
    """

    path = fixture_path("offset_plus_subtotals")
    sheet = _sheet(inspect(path), "Sheet1")
    df = load_dataframe(path, sheet.read_plan)

    assert len(df) == 6
    assert list(df.columns) == ["dept", "item", "month", "amount"]
    # No subtotal/total marker leaked into the body.
    assert not df["dept"].astype(str).str.contains("소계|합계").any()
    # Aggregation sum matches the (excluded) grand-total row -> no double-count.
    assert df["amount"].sum() == 590
    assert list(df["amount"]) == [100, 200, 50, 80, 120, 40]


def test_load_left_margin_cols_applies_usecols(fixture_path) -> None:
    """left_margin_cols -> usecols='C:E' trims the left filler column."""

    path = fixture_path("left_margin_cols")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan.usecols == "C:E"

    df = load_dataframe(path, plan)
    assert list(df.columns) == ["sku", "qty", "price"]
    assert len(df) == 6
    # The filler column A header is never present.
    assert "참고사항" not in df.columns
    assert df.iloc[0]["sku"] == "A-1"
    assert df.iloc[-1]["sku"] == "A-6"


def test_load_header_offset_title_not_mixed_in(fixture_path) -> None:
    """header_offset -> the 3-row title block never bleeds into the body."""

    path = fixture_path("header_offset")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    assert plan.header == 0
    assert plan.skiprows == [0, 1, 2]

    df = load_dataframe(path, plan)
    assert list(df.columns) == ["product", "region", "qty", "amount"]
    assert len(df) == 5

    flat = df.astype(str).to_numpy().ravel().tolist() + list(df.columns)
    for title in ("월간 판매 보고서", "작성일", "단위"):
        assert not any(title in cell for cell in flat), title
    assert df.iloc[0]["product"] == "Widget"


def test_load_types_mixed_applies_dtype(fixture_path) -> None:
    """types_mixed -> the inferred dtype_map applies through the adapter [D5].

    numeric_text keeps leading zeros (string dtype) and the date column loads as
    a datetime dtype — both via the adapter's dtype-key reduction.
    """

    path = fixture_path("types_mixed")
    sheet = _sheet(inspect(path), "Sheet1")
    plan = sheet.read_plan
    # 0=id (number, omitted), 1=code (numeric_text), 2=date, 3=mixed.
    assert plan.dtype_map == {"1": "string", "2": "datetime64[ns]", "3": "object"}

    df = load_dataframe(path, plan)
    # numeric_text 'code' column keeps its leading zeros as strings.
    assert df["code"].tolist() == ["007", "012", "034", "056", "078", "090"]
    assert str(df["code"].dtype) == "string"
    # date column is a datetime dtype.
    assert str(df["date"].dtype).startswith("datetime64")


def test_load_no_header_headerless_override(fixture_path) -> None:
    """no_header + SheetOverride(header_row=None) -> loaded with no header.

    The headerless plan (``header=None``) must load all 5 data rows with the
    first row preserved as data, not consumed as column names (spec §9).
    """

    options = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    path = fixture_path("no_header")
    sheet = _sheet(inspect(path, options), "Sheet1")
    plan = sheet.read_plan
    assert plan.header is None

    df = load_dataframe(path, plan)
    # All 5 data rows present; the first row (1, 100, 1.1) was NOT eaten as a
    # header, so the body still starts with the integer 1.
    assert len(df) == 5
    assert df.iloc[0, 0] == 1


@pytest.mark.parametrize(
    ("fixture_id", "expected_rows", "amount_sum"),
    [
        ("offset_plus_subtotals", 6, 590),
        ("header_offset", 5, None),
        ("header_simple", 5, None),
        ("left_margin_cols", 6, None),
        ("blank_run_terminates", 4, None),
        ("interior_blank", 4, None),
    ],
)
def test_load_row_counts_match_clean_data(
    fixture_path, fixture_id: str, expected_rows: int, amount_sum: int | None
) -> None:
    """Loaded body length matches the clean data-row count for each fixture.

    Where an ``amount`` column exists with a known total, the sum is asserted to
    pin against aggregation duplication (a leaked subtotal/total would inflate
    it).
    """

    path = fixture_path(fixture_id)
    sheet = _sheet(inspect(path), "Sheet1")
    df = load_dataframe(path, sheet.read_plan)
    assert len(df) == expected_rows
    if amount_sum is not None:
        assert df["amount"].sum() == amount_sum
