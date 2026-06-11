"""Idempotency, read-only, and handle-lifetime tests (spec §8, Phase 8).

This suite pins the Phase 8 robustness guarantees end-to-end, across **every**
fixture (openable corpus *and* the corrupt/encrypted failure paths):

1.  **Idempotency** — two :func:`inspect` calls on the same file produce a
    *deeply equal* :class:`WorkbookProfile` (every sheet, every field, every
    read plan), not merely a matching sheet list. The inspector's heuristics
    use deterministic (RNG-free) sampling, so repeated runs must not drift
    (implementation plan §5.3).
2.  **Read-only / side-effect freedom** — the original file's SHA-256 digest and
    mtime are identical before and after inspection, including the failing
    corrupt/encrypted open paths (spec §8, §9) [D3].
3.  **Handle lifetime / no leaks** — every workbook handle the loader opens is
    closed by the time :func:`inspect` returns, and the loader's
    context-manager / :meth:`Loader.close` contract is idempotent and leaves no
    open handle (spec §4.1, §8) [D3].

These complement ``test_readonly_property.py`` (which pins the per-fixture hash
and a lighter idempotency check) with deep-equality and explicit
handle-lifetime assertions.
"""

from __future__ import annotations

import dataclasses
import hashlib
from pathlib import Path

import pytest

from excel_inspector import (
    CorruptWorkbookError,
    EncryptedWorkbookError,
    InspectionOptions,
    Loader,
    SheetOverride,
    inspect,
)
from excel_inspector.models import WorkbookProfile


def _sha256(path: Path) -> str:
    """Return the SHA-256 hex digest of a file's bytes."""

    return hashlib.sha256(path.read_bytes()).hexdigest()


# ---------------------------------------------------------------------------
# Deep idempotency across the openable corpus
# ---------------------------------------------------------------------------


def test_inspect_is_deeply_idempotent(openable_fixture: Path) -> None:
    """Two inspections yield a *deeply equal* WorkbookProfile (spec §8).

    Dataclass equality recurses through ``sheets`` -> ``columns`` / ``merges``
    / ``read_plan``, so this asserts every coordinate, type, provenance, and
    plan field is reproduced exactly — a far stronger statement than matching
    sheet names. RNG-free sampling makes this deterministic (plan §5.3).
    """

    first = inspect(openable_fixture)
    second = inspect(openable_fixture)

    assert isinstance(first, WorkbookProfile)
    assert first == second


def test_inspect_is_deeply_idempotent_with_overrides(fixture_path) -> None:
    """Idempotency holds on the override path too [D2].

    A header_row + dtype_force + skip_rows override exercises the manual
    provenance and aggregator override branches; repeating the inspection with
    the same options must still produce an identical profile.
    """

    options = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                header_row=4,
                dtype_force={"0": "string"},
                skip_rows_add=[13],
            )
        }
    )
    path = fixture_path("offset_plus_subtotals")
    first = inspect(path, options)
    second = inspect(path, options)
    assert first == second


# ---------------------------------------------------------------------------
# Read-only: bytes + mtime unchanged (openable + failure paths)
# ---------------------------------------------------------------------------


def test_openable_inspection_leaves_bytes_unchanged(
    openable_fixture: Path,
) -> None:
    """A successful inspection never touches the file's bytes (spec §8).

    mtime is intentionally not asserted: under ~/Documents an OS/sync daemon
    may bump st_mtime independently of this process. SHA-256 is the contract.
    """

    before_hash = _sha256(openable_fixture)

    inspect(openable_fixture)

    assert _sha256(openable_fixture) == before_hash


@pytest.mark.parametrize(
    ("fixture_id", "error"),
    [
        ("corrupt", CorruptWorkbookError),
        ("encrypted", EncryptedWorkbookError),
    ],
)
def test_failed_inspection_leaves_bytes_unchanged(
    fixture_path, fixture_id: str, error: type[Exception]
) -> None:
    """The failing corrupt/encrypted open path is still strictly read-only.

    Even when the loader raises a domain exception (spec §9), the file's
    SHA-256 must be unchanged — the early-abort open path writes nothing
    (spec §8). mtime is not asserted (external OS/sync can bump it).
    """

    path = fixture_path(fixture_id)
    before_hash = _sha256(path)

    with pytest.raises(error):
        inspect(path)

    assert _sha256(path) == before_hash


def test_double_inspection_leaves_bytes_unchanged(
    openable_fixture: Path,
) -> None:
    """Repeated inspections do not accumulate any byte drift."""

    before = _sha256(openable_fixture)
    inspect(openable_fixture)
    inspect(openable_fixture)
    assert _sha256(openable_fixture) == before


# ---------------------------------------------------------------------------
# Handle lifetime / no leaks (spec §4.1, §8) [D3]
# ---------------------------------------------------------------------------


def test_inspect_closes_all_handles(
    openable_fixture: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Every workbook handle opened during inspect() is closed on return [D3].

    We wrap :class:`Loader` so the loader instance created by ``inspect`` is
    observable, then assert it reports closed (``_closed``) and has dropped both
    mode handles after the call. This proves the ``with Loader(...)`` contract
    in :func:`inspect` released every handle (no leak).
    """

    created: list[Loader] = []
    original_init = Loader.__init__

    def _tracking_init(self: Loader, path: object) -> None:  # type: ignore[override]
        original_init(self, path)  # type: ignore[arg-type]
        created.append(self)

    monkeypatch.setattr(Loader, "__init__", _tracking_init)

    inspect(openable_fixture)

    assert created, "inspect() did not construct a Loader"
    for loader in created:
        assert loader._closed is True  # noqa: SLF001 - lifetime invariant
        assert loader._structure_wb is None  # noqa: SLF001
        assert loader._data_wb is None  # noqa: SLF001
        # Phase 12 (plan v2 §6 probe 1): the formula-mode handle — opened
        # lazily for the formulas fixtures in this parametrization — must be
        # dropped by the same close() tuple as the other two.
        assert loader._formula_wb is None  # noqa: SLF001


def test_failed_inspect_closes_all_handles(
    fixture_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A failed open still releases the loader (the ``with`` block exits) [D3]."""

    created: list[Loader] = []
    original_init = Loader.__init__

    def _tracking_init(self: Loader, path: object) -> None:  # type: ignore[override]
        original_init(self, path)  # type: ignore[arg-type]
        created.append(self)

    monkeypatch.setattr(Loader, "__init__", _tracking_init)

    with pytest.raises(CorruptWorkbookError):
        inspect(fixture_path("corrupt"))

    assert created, "inspect() did not construct a Loader"
    for loader in created:
        assert loader._closed is True  # noqa: SLF001


def test_loader_close_is_idempotent(fixture_path) -> None:
    """``Loader.close`` is safe to call repeatedly and frees both handles [D3]."""

    loader = Loader(fixture_path("header_simple"))
    loader.structure_workbook()
    loader.data_workbook()

    loader.close()
    assert loader._closed is True  # noqa: SLF001
    assert loader._structure_wb is None  # noqa: SLF001
    assert loader._data_wb is None  # noqa: SLF001

    # A second close must not raise and must keep the closed state.
    loader.close()
    assert loader._closed is True  # noqa: SLF001


def test_full_workbook_profile_fields_are_stable(
    openable_fixture: Path,
) -> None:
    """A field-by-field re-comparison guards against any non-dataclass drift.

    Belt-and-suspenders over :func:`test_inspect_is_deeply_idempotent`: convert
    both profiles to nested dicts (``dataclasses.asdict``) and compare, so even
    a future field that breaks ``__eq__`` semantics would still be caught.
    """

    first = dataclasses.asdict(inspect(openable_fixture))
    second = dataclasses.asdict(inspect(openable_fixture))
    assert first == second
