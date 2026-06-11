"""``ReadPlan`` -> :func:`pandas.read_excel` adapter (spec §4.8/§5.5) [D1][D5].

This is the *single* place where the inspector's pandas-domain (0-based)
:class:`~excel_inspector.models.ReadPlan` is translated into a concrete
:func:`pandas.read_excel` call. Keeping the translation here (rather than inline
in every test or caller) means the read-side contract is defined once and pinned
by golden round-trip tests (spec §5.5 "계약 주의"; implementation plan Phase 7).

The plan fields map onto ``read_excel`` keyword arguments verbatim — they are
already in the pandas 0-based domain (the 1-based -> 0-based conversion lives
solely in ``aggregator.py`` [D1]), so this adapter performs **no coordinate
math**. It only:

* selects the kwargs pandas actually needs (``sheet_name``/``engine``/
  ``header``/``usecols``/``skiprows``/``nrows``/``dtype``);
* reduces the :attr:`ReadPlan.dtype_map` keys — 0-based column-position strings
  [D5] — to the *integer positional* keys pandas expects for a ``dtype`` dict;
* preserves the ``header=None`` (headerless) contract (spec §9): the first data
  row is loaded as data, not consumed as column names.

Kwarg mapping (``read_plan_to_kwargs``):

==================  =========================  ============================
ReadPlan field      read_excel kwarg           Notes
==================  =========================  ============================
sheet_name          sheet_name                 always
engine              engine                     always (fixed "openpyxl")
header              header                     always; ``None`` => headerless
usecols             usecols                    only when not ``None``
skiprows            skiprows                   always (may be empty list)
nrows               nrows                      always (may be ``None``)
dtype_map           dtype                      only when non-empty; keys
                                               ``str`` -> ``int`` (positional)
==================  =========================  ============================

dtype key reduction [D5]: ``dtype_map`` keys are 0-based column-position
*strings* relative to the usecols-selected frame. pandas accepts a ``dtype``
dict keyed by **positional integers**, so each ``"0"``/``"1"`` key is cast to
the int ``0``/``1``. Because ``usecols`` selects exactly the profiled table
span, position 0 of the selected frame is the table's first column — the same
basis :attr:`ColumnProfile.index` uses, so no offset is applied.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

import pandas as pd

if TYPE_CHECKING:  # pragma: no cover - typing only
    from ..models import ReadPlan


def _dtype_map_to_pandas(dtype_map: dict[str, str]) -> dict[int, str]:
    """Reduce 0-based position-string keys to positional integer keys [D5].

    :attr:`ReadPlan.dtype_map` is keyed by the 0-based column position expressed
    as a string (``"0"``, ``"1"``, ...). pandas' ``read_excel(dtype=...)``
    accepts a dict keyed by integer column positions, so each key is cast to its
    integer value while the dtype string value is left untouched.

    Args:
        dtype_map: The plan's ``{position_string: pandas_dtype}`` map.

    Returns:
        A new ``{position_int: pandas_dtype}`` map suitable for ``read_excel``.
    """

    return {int(key): dtype for key, dtype in dtype_map.items()}


def read_plan_to_kwargs(plan: "ReadPlan") -> dict[str, Any]:
    """Translate a :class:`ReadPlan` into :func:`pandas.read_excel` kwargs [D1].

    No coordinate conversion happens here — the plan is already 0-based (pandas
    domain). ``sheet_name``/``engine``/``header``/``skiprows``/``nrows`` are
    always emitted; ``usecols`` is emitted only when the plan restricts columns
    (the ``None`` == "all columns" contract); ``dtype`` is emitted only when the
    plan carries a non-empty ``dtype_map`` (its string keys reduced to positional
    ints [D5]).

    The ``header=None`` (headerless) case is preserved verbatim so pandas reads
    no header row and the first data row stays as data (spec §9, HIGH #3).

    Args:
        plan: The read plan produced by the aggregator for one sheet.

    Returns:
        A kwargs dict ready to splat into :func:`pandas.read_excel`.
    """

    kwargs: dict[str, Any] = {
        "sheet_name": plan.sheet_name,
        "engine": plan.engine,
        # header may be an int, a list[int] (v1+ multi-level), or None
        # (headerless); all three are valid read_excel values and are passed
        # through unchanged.
        "header": plan.header,
        "skiprows": plan.skiprows,
        "nrows": plan.nrows,
    }
    if plan.usecols is not None:
        kwargs["usecols"] = plan.usecols
    if plan.dtype_map:
        kwargs["dtype"] = _dtype_map_to_pandas(plan.dtype_map)
    return kwargs


def load_dataframe(file_path: str | Path, plan: "ReadPlan") -> pd.DataFrame:
    """Load ``file_path`` into a DataFrame following ``plan`` (spec §4.8) [D1].

    Thin wrapper over :func:`pandas.read_excel` driven entirely by
    :func:`read_plan_to_kwargs`. This is the inspector's read-side boundary: the
    only point at which a :class:`ReadPlan` becomes actual loaded data.

    Args:
        file_path: Path to the ``.xlsx`` workbook to load.
        plan: The read plan for the target sheet.

    Returns:
        The loaded :class:`pandas.DataFrame`, aligned per the plan (no row slip,
        skipped subtotal/blank rows excluded, columns trimmed to ``usecols``).
    """

    kwargs = read_plan_to_kwargs(plan)
    return pd.read_excel(file_path, **kwargs)
