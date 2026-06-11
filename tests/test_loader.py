"""Loader tests (spec §4.1, Phase 1; formula mode plan v2 Phase 12) [D3].

Covers the loader's responsibilities: opening the three modes (structure /
data / formula), mapping corrupt/encrypted files to domain exceptions,
lazy/cached handles, and handle cleanup (no leaks) via both ``close()`` and
the context manager protocol. Phase 12 activated the formula mode
(``read_only=True, data_only=False``); the close/weakref coverage explicitly
includes the ``_formula_wb`` handle (plan v2 §6 review probe 1 — forgetting
it in ``close()`` would leak while every pre-Phase-12 close test stayed
green).
"""

from __future__ import annotations

import gc
import weakref
from pathlib import Path

import pytest

from excel_inspector import (
    CorruptWorkbookError,
    EncryptedWorkbookError,
    Loader,
)


def test_structure_workbook_opens_and_exposes_merges(fixture_path) -> None:
    """Structure mode opens and exposes merged_cells (read_only lacks them)."""

    loader = Loader(fixture_path("merged_header"))
    try:
        wb = loader.structure_workbook()
        ws = wb["Sheet1"]
        ranges = {str(m) for m in ws.merged_cells.ranges}
        assert ranges == {"A1:B1", "A6:A7"}
    finally:
        loader.close()


def test_structure_workbook_is_cached(fixture_path) -> None:
    """Repeated structure_workbook() calls return the same cached handle."""

    with Loader(fixture_path("header_simple")) as loader:
        first = loader.structure_workbook()
        second = loader.structure_workbook()
        assert first is second


def test_data_workbook_streams_rows(fixture_path) -> None:
    """Data mode opens a read_only streaming workbook over cached values."""

    with Loader(fixture_path("header_simple")) as loader:
        wb = loader.data_workbook()
        ws = wb["Sheet1"]
        first_row = next(ws.iter_rows(values_only=True))
        assert first_row == ("name", "age", "city", "score")


def test_data_workbook_is_cached(fixture_path) -> None:
    """Repeated data_workbook() calls return the same cached handle."""

    with Loader(fixture_path("header_simple")) as loader:
        assert loader.data_workbook() is loader.data_workbook()


def test_structure_and_data_are_distinct_handles(fixture_path) -> None:
    """Structure and data modes are separate workbook instances [D3]."""

    with Loader(fixture_path("header_simple")) as loader:
        assert loader.structure_workbook() is not loader.data_workbook()


def test_formula_workbook_streams_formula_strings(fixture_path) -> None:
    """Formula mode (data_only=False) reads stored formula strings (Phase 12).

    Replaces the v1 ``NotImplementedError`` stub assertion — the activation
    is the explicit Phase 12 Step 1 deliverable (plan v2 §6).
    """

    with Loader(fixture_path("formulas")) as loader:
        wb = loader.formula_workbook()
        ws = wb["Sheet1"]
        rows = list(ws.iter_rows(min_row=2, max_row=5, values_only=True))
        assert [row[3] for row in rows] == [
            "=B2*C2",
            "=B3*C3",
            "=B4*C4",
            "=B5*C5",
        ]


def test_formula_workbook_is_cached_and_distinct(fixture_path) -> None:
    """The formula handle is cached and separate from the other modes [D3]."""

    with Loader(fixture_path("formulas")) as loader:
        first = loader.formula_workbook()
        assert first is loader.formula_workbook()
        assert first is not loader.data_workbook()
        assert first is not loader.structure_workbook()


def test_corrupt_file_raises_corrupt_error(fixture_path) -> None:
    """A truncated zip maps to CorruptWorkbookError (spec §4.1, §9)."""

    with Loader(fixture_path("corrupt")) as loader:
        with pytest.raises(CorruptWorkbookError):
            loader.structure_workbook()


def test_encrypted_file_raises_encrypted_error(fixture_path) -> None:
    """An OLE2-wrapped password-protected file maps to EncryptedWorkbookError."""

    with Loader(fixture_path("encrypted")) as loader:
        with pytest.raises(EncryptedWorkbookError):
            loader.structure_workbook()


def test_encrypted_data_mode_also_raises_encrypted(fixture_path) -> None:
    """Data mode disambiguates encryption the same way structure mode does."""

    with Loader(fixture_path("encrypted")) as loader:
        with pytest.raises(EncryptedWorkbookError):
            loader.data_workbook()


def test_encrypted_formula_mode_also_raises_encrypted(fixture_path) -> None:
    """Formula mode shares the open-error translation (Phase 12)."""

    with Loader(fixture_path("encrypted")) as loader:
        with pytest.raises(EncryptedWorkbookError):
            loader.formula_workbook()


def test_missing_file_raises_corrupt_error(tmp_path: Path) -> None:
    """A nonexistent file is surfaced as a corrupt (unreadable) workbook."""

    with Loader(tmp_path / "does_not_exist.xlsx") as loader:
        with pytest.raises(CorruptWorkbookError):
            loader.structure_workbook()


def test_close_is_idempotent_and_releases_handles(fixture_path) -> None:
    """close() releases handles, is idempotent, and blocks reopening."""

    loader = Loader(fixture_path("header_simple"))
    loader.structure_workbook()
    loader.data_workbook()

    loader.close()
    loader.close()  # idempotent: must not raise

    with pytest.raises(RuntimeError):
        loader.structure_workbook()


def test_context_manager_closes_on_exit(fixture_path) -> None:
    """Exiting the context manager closes handles even after use."""

    with Loader(fixture_path("header_simple")) as loader:
        loader.structure_workbook()
    # After exit the loader is closed and refuses to reopen.
    with pytest.raises(RuntimeError):
        loader.data_workbook()


def test_no_handle_leak_after_close(fixture_path) -> None:
    """Workbook handles are garbage-collected after close (no leak) [D3]."""

    loader = Loader(fixture_path("header_simple"))
    wb_ref = weakref.ref(loader.structure_workbook())
    loader.close()
    gc.collect()
    assert wb_ref() is None


def test_no_data_handle_leak_after_close(fixture_path) -> None:
    """The data-mode handle is also garbage-collected after close (issue #4)."""

    loader = Loader(fixture_path("header_simple"))
    wb_ref = weakref.ref(loader.data_workbook())
    loader.close()
    gc.collect()
    assert wb_ref() is None


def test_both_modes_reclaimed_after_close(fixture_path) -> None:
    """Both handles open simultaneously are both reclaimed on close (issue #4)."""

    loader = Loader(fixture_path("header_simple"))
    structure_ref = weakref.ref(loader.structure_workbook())
    data_ref = weakref.ref(loader.data_workbook())
    loader.close()
    gc.collect()
    assert structure_ref() is None
    assert data_ref() is None


def test_no_formula_handle_leak_after_close(fixture_path) -> None:
    """The formula-mode handle is garbage-collected after close (probe 1).

    Plan v2 §6 review probe 1: omitting ``_formula_wb`` from the ``close()``
    handle tuple keeps every pre-Phase-12 close test green while silently
    leaking the third workbook — this weakref extension is the trap-spring.
    """

    loader = Loader(fixture_path("formulas"))
    wb_ref = weakref.ref(loader.formula_workbook())
    loader.close()
    gc.collect()
    assert wb_ref() is None


def test_all_three_modes_reclaimed_after_close(fixture_path) -> None:
    """Structure + data + formula handles are all reclaimed together (probe 1)."""

    loader = Loader(fixture_path("formulas"))
    structure_ref = weakref.ref(loader.structure_workbook())
    data_ref = weakref.ref(loader.data_workbook())
    formula_ref = weakref.ref(loader.formula_workbook())
    loader.close()
    gc.collect()
    assert structure_ref() is None
    assert data_ref() is None
    assert formula_ref() is None


def test_closed_loader_refuses_formula_mode(fixture_path) -> None:
    """A closed loader refuses to (re)open the formula mode like the others."""

    loader = Loader(fixture_path("formulas"))
    loader.close()
    with pytest.raises(RuntimeError):
        loader.formula_workbook()


def test_close_isolates_per_handle_failures(fixture_path) -> None:
    """A failure closing the first handle still closes the second + marks closed.

    Regression for issue #3: the structure handle's ``close()`` is patched to
    raise; the data handle must still be closed, ``_closed`` must be set, and
    the collected error surfaces as an ExceptionGroup after cleanup.
    """

    loader = Loader(fixture_path("header_simple"))
    structure = loader.structure_workbook()
    data = loader.data_workbook()

    data_closed: dict[str, bool] = {"called": False}
    real_data_close = data.close

    def _flag_data_close() -> None:
        data_closed["called"] = True
        real_data_close()

    def _boom() -> None:
        raise RuntimeError("structure close boom")

    structure.close = _boom  # type: ignore[method-assign]
    data.close = _flag_data_close  # type: ignore[method-assign]

    with pytest.raises(ExceptionGroup) as excinfo:
        loader.close()

    # Second handle was closed despite the first failing.
    assert data_closed["called"] is True
    # Loader is marked closed unconditionally and refuses to reopen.
    assert loader._closed is True
    with pytest.raises(RuntimeError):
        loader.structure_workbook()
    # The collected close error is surfaced.
    assert any(isinstance(e, RuntimeError) for e in excinfo.value.exceptions)
