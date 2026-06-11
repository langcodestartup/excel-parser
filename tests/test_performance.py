"""Performance / sampling smoke tests (spec §8, Phase 8).

The inspector must scale by *sampling and streaming*, never by a full scan or a
full in-memory materialization of the row data (spec §8) [D3]. These tests use
the ``large_table`` fixture (5000 data rows + a trailing total) to assert:

1.  **Correctness at scale** — boundaries, the trailing-total skip, the sampled
    column types, and the read plan are exactly what the small fixtures' rules
    predict, just at 5000 rows.
2.  **Bounded sampling** — the Header Locator reads only the top
    ``HEADER_SCAN_ROWS`` rows, and the Type Profiler retains at most
    ``TYPE_SAMPLE_ROWS`` data rows (+ the header) regardless of table height, so
    neither materializes the whole sheet.
3.  **Bounded memory (best-effort)** — a ``tracemalloc`` peak sanity bound that a
    full 5000-row materialization in the analyzers would exceed. This is a
    coarse upper-bound check, not a precise measurement (spec §8: Phase 8
    best-effort).

The row-data analyzers (Header Locator, Boundary Detector, Type Profiler) read
in read_only streaming mode; the structure workbook (loaded once for merges /
dimensions, [D3]) is the only full in-memory open and is accepted by the spec's
200 MB residency budget.

Plan v2 Phase 13 Step 4 adds the ``@pytest.mark.slow`` 100k-row smoke
(``test_inspect_100k_rows_peak_memory_within_spec_budget``) pinning the spec
§8 quantitative target itself. It is deselected by the default ``addopts``
(``-m "not slow"``); run it with ``pytest -m slow`` (or everything with
``pytest -m "slow or not slow"``).
"""

from __future__ import annotations

import time
import tracemalloc
from pathlib import Path

import pytest

from excel_inspector import Loader, inspect
from excel_inspector.analyzers.boundary_detector import BoundaryDetector
from excel_inspector.analyzers.header_locator import HeaderLocator
from excel_inspector.analyzers.sheet_enumerator import SheetEnumerator
from excel_inspector.analyzers.type_profiler import TypeProfiler
from excel_inspector.context import InspectionContext
from excel_inspector.heuristics import HEADER_SCAN_ROWS, TYPE_SAMPLE_ROWS
from excel_inspector.models import InspectionOptions, WorkbookProfile

#: Generous wall-clock budget for inspecting the 5000-row fixture. A full,
#: doubly-materialized scan still runs in well under this; the budget exists to
#: catch a pathological regression (e.g. an accidental O(n^2) scan), not to
#: benchmark. CI machines are slow, so the margin is wide.
_TIME_BUDGET_S = 10.0

#: tracemalloc peak ceiling (Python allocations only). Comfortably above the
#: observed ~10 MB but far below what fully materializing 5000 rows twice in the
#: analyzers would add. Best-effort sanity bound (spec §8).
_PEAK_MB_CEILING = 80.0


def test_large_table_boundaries_and_plan(fixture_path) -> None:
    """At 5000 rows the boundaries / total-skip / plan are still exact (spec §8)."""

    profile = inspect(fixture_path("large_table"))
    sheet = profile.sheets[0]

    assert sheet.header_row == 1
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == 5001  # last real data row (total excluded)
    assert sheet.skip_rows == [5002]  # the trailing 합계 grand total
    assert (sheet.max_row, sheet.max_col) == (5002, 4)

    plan = sheet.read_plan
    assert plan is not None
    # nrows is the whole inclusive data span (interior skips not subtracted, [D1]).
    assert plan.nrows == 5000
    # 0-based skiprows: the total at 1-based 5002 -> 5001.
    assert plan.skiprows == [5001]
    # Only the text 'name' column (position 1) gets a dtype key; numbers omitted.
    assert plan.dtype_map == {"1": "string"}


def test_large_table_column_types(fixture_path) -> None:
    """The deterministic sample classifies every column correctly at scale."""

    sheet = inspect(fixture_path("large_table")).sheets[0]
    types = {c.name: c.inferred_type for c in sheet.columns}
    assert types == {
        "id": "number",
        "name": "text",
        "amount": "number",
        "flag": "number",
    }
    assert all(c.null_ratio == 0.0 for c in sheet.columns)


def test_type_profiler_samples_are_bounded(fixture_path) -> None:
    """The Type Profiler retains <= TYPE_SAMPLE_ROWS rows, not all 5000 (spec §8).

    Drive the pipeline up to the Type Profiler, then ask it directly how many
    rows its deterministic sample selects: it must be capped at
    ``TYPE_SAMPLE_ROWS`` even though the table has 5000 eligible data rows. This
    proves the analyzer samples rather than scanning the whole sheet.
    """

    path = fixture_path("large_table")
    context = InspectionContext(
        options=InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)
        HeaderLocator().analyze(context)
        BoundaryDetector().analyze(context)

        sheet = context.workbook_profile.sheets[0]
        # 5000 eligible data rows (inclusive 2..5001; the total at 5002 is a skip).
        assert sheet.data_start_row == 2
        assert sheet.data_end_row == 5001
        assert sheet.data_end_row - sheet.data_start_row + 1 == 5000
        assert sheet.skip_rows == [5002]

        # Call-site updated for the Phase 10b-1 block-local signature (explicit
        # boundaries instead of a profile read); the assertions are unchanged.
        sample_rows = TypeProfiler()._sampled_row_numbers(  # noqa: SLF001
            sheet.data_start_row, sheet.data_end_row, sheet.skip_rows
        )
        assert len(sample_rows) == TYPE_SAMPLE_ROWS
        assert len(sample_rows) < 5000
        # Sampled rows lie strictly inside the data region and never include the
        # skipped total row.
        assert all(2 <= r <= 5001 for r in sample_rows)
        assert 5002 not in sample_rows


def test_header_locator_reads_only_top_rows(fixture_path) -> None:
    """The Header Locator scans at most HEADER_SCAN_ROWS rows (spec §7.1).

    A counting spy over the data-mode worksheet's ``iter_rows`` confirms the
    header scan never streams past the top ``HEADER_SCAN_ROWS`` of the 5002-row
    sheet — i.e. it samples the top, not the whole sheet.
    """

    path = fixture_path("large_table")
    context = InspectionContext(
        options=InspectionOptions(),
        workbook_profile=WorkbookProfile(file_path=str(path)),
    )
    with Loader(path) as loader:
        context.loader = loader
        SheetEnumerator().analyze(context)

        original_data_workbook = loader.data_workbook
        max_rows_requested: list[int | None] = []

        class _CountingWorksheet:
            def __init__(self, ws: object) -> None:
                self._ws = ws

            def iter_rows(self, *, min_row, max_row, values_only):  # noqa: ANN001, ANN202
                max_rows_requested.append(max_row)
                yield from self._ws.iter_rows(
                    min_row=min_row, max_row=max_row, values_only=values_only
                )

        class _CountingWorkbook:
            def __init__(self, wb: object) -> None:
                self._wb = wb

            def __getitem__(self, name: str) -> _CountingWorksheet:
                return _CountingWorksheet(self._wb[name])

        def _spy_data_workbook() -> _CountingWorkbook:
            return _CountingWorkbook(original_data_workbook())

        loader.data_workbook = _spy_data_workbook  # type: ignore[method-assign]
        HeaderLocator().analyze(context)

    assert max_rows_requested, "header locator did not read any rows"
    # The header scan caps its read at HEADER_SCAN_ROWS — never the full sheet.
    assert all(req == HEADER_SCAN_ROWS for req in max_rows_requested)


def test_inspect_large_table_is_fast_and_bounded(fixture_path) -> None:
    """inspect() on 5000 rows is fast and stays under the memory sanity bound.

    Best-effort (spec §8): a wall-clock budget catches pathological scans and a
    ``tracemalloc`` peak ceiling catches an accidental full materialization of
    the row data in the analyzers.
    """

    path: Path = fixture_path("large_table")

    tracemalloc.start()
    start = time.perf_counter()
    profile = inspect(path)
    elapsed = time.perf_counter() - start
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    # Correctness guard so a fast-but-wrong run cannot pass.
    assert profile.sheets[0].data_end_row == 5001

    assert elapsed < _TIME_BUDGET_S, f"inspection too slow: {elapsed:.2f}s"
    peak_mb = peak / 1e6
    assert peak_mb < _PEAK_MB_CEILING, f"peak allocation too high: {peak_mb:.1f}MB"


#: Spec §8 quantitative target: at a 100k-row file, inspection's memory
#: increase stays <= 200 MB (sampling-based). Asserted on the tracemalloc
#: peak (Python allocations) per plan v2 Phase 13 Step 4.
_SPEC_PEAK_MB_BUDGET = 200.0


@pytest.mark.slow
def test_inspect_100k_rows_peak_memory_within_spec_budget(
    perf_fixture_path: Path, perf_table_data_rows: int
) -> None:
    """inspect() on a 100k-row workbook peaks <= 200 MB (spec §8; slow smoke).

    The tracemalloc peak over a full ``inspect()`` of the on-demand
    ``build_perf_100k`` workbook (100k data rows x 3 columns, plan v2 Phase
    13 Step 4) must stay within the spec §8 budget. Measured baseline
    (openpyxl 3.1.5 / Python 3.14.5): **171.4 MB**, dominated by the
    structure-mode (``read_only=False``, [D3]) cell tree at ~390-540 B/cell —
    an accidental full materialization of row data in the analyzers would
    add far more than the ~30 MB of remaining headroom, which is exactly the
    regression this smoke guards.

    Known boundary (recorded in the Phase 13 handoff): at 100k x **4**
    columns the same measurement is 216.8 MB — the [D3] full structure load
    alone breaches the §8 budget at roughly >= 370k cells. That is an
    architecture-level finding, not an analyzer regression; do not "fix" a
    failure of this test by widening the budget or shrinking the fixture.
    """

    tracemalloc.start()
    profile = inspect(perf_fixture_path)
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    # Correctness guard so a fast-but-wrong (or short-read) run cannot pass.
    sheet = profile.sheets[0]
    assert sheet.header_row == 1
    assert sheet.data_start_row == 2
    assert sheet.data_end_row == perf_table_data_rows + 1
    assert sheet.skip_rows == []
    assert sheet.read_plan is not None
    assert sheet.read_plan.nrows == perf_table_data_rows

    peak_mb = peak / 1e6
    assert peak_mb <= _SPEC_PEAK_MB_BUDGET, (
        f"inspect() peak allocation {peak_mb:.1f} MB exceeds the spec §8 "
        f"budget of {_SPEC_PEAK_MB_BUDGET:.0f} MB at "
        f"{perf_table_data_rows} rows"
    )
