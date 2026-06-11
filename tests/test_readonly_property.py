"""Read-only / idempotency property tests (spec §8, Phase 1 exit criteria).

Inspection must be side-effect free: the original file's bytes must be identical
before and after :func:`inspect`, and repeated inspections must produce identical
results [D3]. Verified across every openable fixture.

The spec §8 contract is byte-level (SHA-256) invariance, NOT mtime. The corpus
lives under ~/Documents, where OS/sync daemons (Spotlight, iCloud) may bump
st_mtime independently of this process; asserting mtime equality is therefore
flaky and outside what a read-only inspector can guarantee. We assert the hash.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from excel_inspector import (
    CorruptWorkbookError,
    EncryptedWorkbookError,
    inspect,
)


def _sha256(path: Path) -> str:
    """Return the SHA-256 hex digest of a file's bytes."""

    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_inspect_does_not_mutate_file_bytes(openable_fixture: Path) -> None:
    """File bytes (SHA-256) are unchanged by inspection (spec §8)."""

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
def test_inspect_does_not_mutate_negative_fixtures(
    fixture_path, fixture_id: str, error: type[Exception]
) -> None:
    """Corrupt/encrypted files are byte-stable across a failed open.

    The loader raises a domain exception for these (issue #12); the file's
    SHA-256 must be identical before and after, proving the failing open path
    is still strictly read-only (spec §8, §9).
    """

    path = fixture_path(fixture_id)
    before_hash = _sha256(path)

    with pytest.raises(error):
        inspect(path)

    assert _sha256(path) == before_hash


def test_inspect_is_idempotent(openable_fixture: Path) -> None:
    """Two inspections of the same file produce the same structural result."""

    first = inspect(openable_fixture)
    second = inspect(openable_fixture)

    assert [s.name for s in first.sheets] == [s.name for s in second.sheets]
    for s1, s2 in zip(first.sheets, second.sheets, strict=True):
        assert s1.used_range == s2.used_range
        assert (s1.max_row, s1.max_col) == (s2.max_row, s2.max_col)
        assert s1.is_tabular_candidate == s2.is_tabular_candidate
        assert (s1.read_plan is None) == (s2.read_plan is None)
        if s1.read_plan is not None and s2.read_plan is not None:
            assert s1.read_plan == s2.read_plan
