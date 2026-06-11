"""Loading adapters bridging :class:`ReadPlan` to external libraries.

``pandas_loader.py`` (the ``ReadPlan`` -> :func:`pandas.read_excel` translation,
Phase 7) is the single read-side boundary [D1]. This package is established in
Phase 0 to fix the module layout (implementation plan §2).
"""

from __future__ import annotations

from .pandas_loader import load_dataframe, read_plan_to_kwargs

__all__ = [
    "load_dataframe",
    "read_plan_to_kwargs",
]
