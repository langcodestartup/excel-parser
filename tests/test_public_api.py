"""Public API surface tests (spec §3/§5.5, Phase 8).

Pins the v1 public surface of :mod:`excel_inspector`: the :func:`inspect`
entry point ``inspect(path, options=None) -> WorkbookProfile`` and the read-side
adapter :func:`load_dataframe` (plus :func:`read_plan_to_kwargs`) must be
importable directly from the top-level package, so a consumer never has to reach
into submodules. Guards against an accidental removal/rename in ``__init__``.
"""

from __future__ import annotations

import inspect as _inspect

import excel_inspector as ei


def test_inspect_is_exported_with_v1_signature() -> None:
    """``inspect(path, options=None) -> WorkbookProfile`` is the v1 entry point."""

    assert "inspect" in ei.__all__
    assert callable(ei.inspect)

    sig = _inspect.signature(ei.inspect)
    params = list(sig.parameters)
    assert params == ["path", "options"]
    assert sig.parameters["options"].default is None


def test_load_dataframe_is_exported() -> None:
    """``load_dataframe`` is reachable from the top-level package (Phase 8)."""

    assert "load_dataframe" in ei.__all__
    assert callable(ei.load_dataframe)
    # Same object as the adapters submodule export (single definition).
    from excel_inspector.adapters import load_dataframe as adapter_load

    assert ei.load_dataframe is adapter_load


def test_read_plan_to_kwargs_is_exported() -> None:
    """The plan->kwargs translator is also part of the public surface."""

    assert "read_plan_to_kwargs" in ei.__all__
    assert callable(ei.read_plan_to_kwargs)


def test_public_round_trip_from_top_level(fixture_path) -> None:
    """inspect() -> load_dataframe() works using only top-level imports."""

    path = fixture_path("header_simple")
    profile = ei.inspect(path)
    sheet = profile.sheets[0]
    assert sheet.read_plan is not None

    df = ei.load_dataframe(path, sheet.read_plan)
    assert list(df.columns) == ["name", "age", "city", "score"]
    assert len(df) == 5
