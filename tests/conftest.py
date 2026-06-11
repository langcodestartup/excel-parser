"""Shared pytest fixtures and partial-context synthesis helpers (spec §6).

This module provides three test-support facilities:

1.  **Partial-context synthesis** (spec §6): :func:`make_context` and the
    :func:`make_sheet_profile` :class:`SheetProfile` factory let analyzer unit
    tests build an :class:`InspectionContext` populated with *only* the fields
    under test, for isolation. Both are exposed as fixtures
    (``context_factory``, ``sheet_profile_factory``).
2.  **Corpus generation** (implementation plan §5.1): the session-scoped
    ``fixture_corpus`` fixture (re)generates the entire ``tests/fixtures/*.xlsx``
    corpus via ``tests/fixtures/generate.py`` before any fixture-consuming test
    runs, so the corpus is always present and deterministic.
3.  **Fixture path access**: ``fixture_path`` (a lookup function fixture),
    ``fixtures`` (the ``FIXTURES`` metadata table), and a parametrized
    ``openable_fixture`` fixture over every openable sample.
"""

from __future__ import annotations

import importlib.util
import sys
from collections.abc import Callable, Iterable
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from excel_inspector.context import InspectionContext
from excel_inspector.models import (
    ColumnProfile,
    InspectionOptions,
    MergeRegion,
    SheetProfile,
    WorkbookProfile,
)

_FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"


def _load_generate_module() -> ModuleType:
    """Import ``tests/fixtures/generate.py`` directly.

    ``tests`` is not an installed package, so we load the generator by file
    path rather than relying on package-qualified imports.
    """

    if "fixtures_generate" in sys.modules:
        return sys.modules["fixtures_generate"]

    spec = importlib.util.spec_from_file_location(
        "fixtures_generate", _FIXTURES_DIR / "generate.py"
    )
    if spec is None or spec.loader is None:  # pragma: no cover - defensive
        raise RuntimeError("could not load tests/fixtures/generate.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules["fixtures_generate"] = module
    spec.loader.exec_module(module)
    return module


# ---------------------------------------------------------------------------
# Partial-context synthesis helpers (spec §6)
# ---------------------------------------------------------------------------


def make_sheet_profile(
    name: str = "Sheet1",
    *,
    is_visible: bool = True,
    is_tabular_candidate: bool = True,
    is_tabular_provenance: str = "heuristic",
    used_range: str = "",
    used_range_trusted: bool = True,
    max_row: int = 0,
    max_col: int = 0,
    header_row: int | None = None,
    header_confidence: float = 0.0,
    header_provenance: str = "default",
    needs_manual_header: bool = False,
    is_multi_level_header: bool = False,
    merges: Iterable[MergeRegion] | None = None,
    data_start_row: int | None = None,
    data_end_row: int | None = None,
    data_left_col: int | None = None,
    data_right_col: int | None = None,
    skip_rows: Iterable[int] | None = None,
    columns: Iterable[ColumnProfile] | None = None,
    read_plan: Any | None = None,
) -> SheetProfile:
    """Build a :class:`SheetProfile` with only the fields a test cares about.

    All row/column coordinates are **openpyxl 1-based** (the inspection domain
    [D1]); ``ColumnProfile.index`` inside ``columns`` is 0-based from the table
    top-left [D5]. Every argument has a neutral default so analyzer tests can
    supply just the inputs the analyzer reads and assert on the fields it fills.

    Returns:
        A :class:`SheetProfile` populated from the given arguments.
    """

    return SheetProfile(
        name=name,
        is_visible=is_visible,
        is_tabular_candidate=is_tabular_candidate,
        is_tabular_provenance=is_tabular_provenance,
        used_range=used_range,
        used_range_trusted=used_range_trusted,
        max_row=max_row,
        max_col=max_col,
        header_row=header_row,
        header_confidence=header_confidence,
        header_provenance=header_provenance,
        needs_manual_header=needs_manual_header,
        is_multi_level_header=is_multi_level_header,
        merges=list(merges or []),
        data_start_row=data_start_row,
        data_end_row=data_end_row,
        data_left_col=data_left_col,
        data_right_col=data_right_col,
        skip_rows=list(skip_rows or []),
        columns=list(columns or []),
        read_plan=read_plan,
    )


def make_context(
    *,
    options: InspectionOptions | None = None,
    workbook_profile: WorkbookProfile | None = None,
    sheets: list[SheetProfile] | None = None,
    loader: object | None = None,
    warnings: list[str] | None = None,
) -> InspectionContext:
    """Synthesize a partial :class:`InspectionContext` for unit tests (spec §6).

    Only the fields a test needs must be supplied; everything else defaults to
    an empty/clean value so analyzers can be exercised in isolation.

    Args:
        options: Inspection options; defaults to a fresh ``InspectionOptions``.
        workbook_profile: A pre-built workbook profile. Mutually informative
            with ``sheets``: if omitted, one is built (optionally seeded with
            ``sheets``).
        sheets: Sheets to seed a freshly-built workbook profile with. Ignored
            when ``workbook_profile`` is provided.
        loader: A loader (real or fake) to inject; defaults to ``None``.
        warnings: Initial warning list; defaults to empty.

    Returns:
        A partially-populated :class:`InspectionContext`.
    """

    if workbook_profile is None:
        workbook_profile = WorkbookProfile(sheets=list(sheets or []))

    return InspectionContext(
        options=options or InspectionOptions(),
        loader=loader,
        workbook_profile=workbook_profile,
        warnings=list(warnings or []),
    )


@pytest.fixture
def context_factory() -> Callable[..., InspectionContext]:
    """Return the :func:`make_context` helper as a fixture for convenience."""

    return make_context


@pytest.fixture
def sheet_profile_factory() -> Callable[..., SheetProfile]:
    """Return the :func:`make_sheet_profile` factory as a fixture."""

    return make_sheet_profile


# ---------------------------------------------------------------------------
# Fixture corpus generation + path access (implementation plan §5.1)
# ---------------------------------------------------------------------------


@pytest.fixture(scope="session")
def fixture_corpus() -> dict[str, Path]:
    """Regenerate the full ``tests/fixtures/*.xlsx`` corpus once per session.

    Returns:
        Mapping of fixture id -> generated file path. Deterministic and
        idempotent; safe to depend on from any test.
    """

    generate = _load_generate_module()
    return generate.generate_all(_FIXTURES_DIR)


@pytest.fixture(scope="session")
def fixtures() -> dict[str, Any]:
    """Return the ``FIXTURES`` metadata table (id -> ``FixtureSpec``)."""

    return _load_generate_module().FIXTURES


@pytest.fixture
def fixture_path(
    fixture_corpus: dict[str, Path],
) -> Callable[[str], Path]:
    """Return a ``lookup(fixture_id) -> Path`` for the generated corpus.

    Depends on ``fixture_corpus`` so the file is guaranteed to exist before a
    test resolves its path.
    """

    def _lookup(fixture_id: str) -> Path:
        if fixture_id not in fixture_corpus:
            raise KeyError(
                f"unknown fixture id {fixture_id!r}; "
                f"known: {sorted(fixture_corpus)}"
            )
        return fixture_corpus[fixture_id]

    return _lookup


def _openable_fixture_ids() -> list[str]:
    """Collect ids of fixtures openpyxl can open (excludes corrupt/encrypted)."""

    fixtures = _load_generate_module().FIXTURES
    return [fid for fid, spec in fixtures.items() if spec.openable]


@pytest.fixture(scope="session")
def perf_fixture_path(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Materialize the 100k-row perf-smoke workbook (plan v2 Phase 13 Step 4).

    Deliberately *outside* ``fixture_corpus``: ``build_perf_100k`` is not in
    ``BUILDERS``, so default corpus generation (and every corpus-parametrized
    test) never pays the build cost. Only the ``@pytest.mark.slow`` memory
    smoke requests this fixture; the file lands in a session tmp dir so the
    repository's fixtures directory stays free of multi-megabyte artifacts.
    """

    generate = _load_generate_module()
    path = tmp_path_factory.mktemp("perf") / "perf_100k.xlsx"
    path.write_bytes(generate.build_perf_100k())
    return path


@pytest.fixture(scope="session")
def perf_table_data_rows() -> int:
    """The 100k builder's data-row count (single source of truth)."""

    return int(_load_generate_module().PERF_TABLE_DATA_ROWS)


@pytest.fixture(params=_openable_fixture_ids())
def openable_fixture(
    request: pytest.FixtureRequest, fixture_corpus: dict[str, Path]
) -> Path:
    """Parametrized path over every openable (non-negative) fixture."""

    return fixture_corpus[request.param]
