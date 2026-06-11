"""Data models for the Excel Structure Inspector.

This module defines the explicit data structures produced by inspection, per
spec §5. The top-level :class:`WorkbookProfile` holds several
:class:`SheetProfile` instances, and each sheet carries its own
:class:`ReadPlan`.

Coordinate-system contract [D1]:
    * :class:`SheetProfile` and :class:`ColumnProfile` row positions are
      **openpyxl 1-based** (inspection domain).
    * :class:`ReadPlan` row positions are **pandas 0-based** (loading domain).
    * :attr:`ColumnProfile.index` is **0-based relative to the table top-left**.

The single coordinate conversion 1-based -> 0-based happens only in the
Plan Aggregator (``aggregator.py``); see [D1].
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Final

# ---------------------------------------------------------------------------
# Input override contract (spec §5.0) [D2]
# ---------------------------------------------------------------------------


class _Unset:
    """Sentinel type for "this override field was not specified" [D2].

    Distinct from ``None`` (an explicit *headerless* declaration) so the header
    channel can tell "defer to the heuristic" apart from "this sheet has no
    header row" (HIGH #2). A dedicated class gives a readable ``repr`` and a
    stable singleton identity.
    """

    _instance: "_Unset | None" = None

    def __new__(cls) -> "_Unset":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:  # pragma: no cover - cosmetic
        return "<UNSET>"


#: Singleton sentinel marking an unspecified :attr:`SheetOverride.header_row`.
_UNSET: Final[_Unset] = _Unset()


@dataclass
class SheetOverride:
    """Per-sheet manual overrides applied during inspection [D2].

    Fields left at their defaults mean "no override; defer to heuristics".

    The header channel uses a sentinel default (:data:`_UNSET`) so that a
    :class:`SheetOverride` carrying *only* ``dtype_force`` / ``is_tabular`` /
    ``skip_rows_*`` does **not** read as a headerless declaration (HIGH #2).
    Three distinct header states are therefore representable:

    * ``header_row`` left at :data:`_UNSET` — defer to the heuristic locator.
    * ``header_row=<int>`` — force the header to that 1-based row.
    * ``header_row=None`` — declare the sheet has *no* header row.

    Attributes:
        header_row: Forced header row, **1-based** (openpyxl domain); ``None``
            to declare the sheet has no header row; or left at the
            :data:`_UNSET` sentinel to defer to the heuristic. Whether the
            field was actually specified is exposed via
            :attr:`header_row_set` (computed in ``__post_init__``).
        skip_rows_add: Extra rows (1-based) to add to ``skip_rows``.
        skip_rows_remove: Rows (1-based) to remove from heuristic
            ``skip_rows``.
        dtype_force: Forced dtypes keyed by 0-based column position string
            (matching :attr:`ReadPlan.dtype_map` key convention) [D5].
        is_tabular: Forced tabular-candidate flag, or ``None`` to defer.
    """

    header_row: int | None | _Unset = _UNSET
    skip_rows_add: list[int] = field(default_factory=list)
    skip_rows_remove: list[int] = field(default_factory=list)
    dtype_force: dict[str, str] = field(default_factory=dict)
    is_tabular: bool | None = None
    #: Whether ``header_row`` was explicitly specified (set in ``__post_init__``;
    #: not a constructor argument). ``True`` for an int or an explicit ``None``.
    header_row_set: bool = field(init=False, default=False)

    def __post_init__(self) -> None:
        """Record whether ``header_row`` was actually specified [D2]/HIGH #2.

        ``header_row_set`` is ``True`` iff the caller passed a value other than
        the :data:`_UNSET` sentinel (an int *or* an explicit ``None``). When it
        is still the sentinel we collapse it back to ``None`` so the public
        ``header_row`` type stays ``int | None`` for any downstream reader that
        ignores the sentinel.
        """

        self.header_row_set = self.header_row is not _UNSET
        if self.header_row is _UNSET:
            # Collapse the sentinel to None for downstream type simplicity;
            # header_row_set=False already records that it was not specified.
            self.header_row = None


@dataclass
class InspectionOptions:
    """Top-level inspection overrides injected at the entry point [D2].

    Attributes:
        sheet_overrides: Per-sheet overrides keyed by sheet name.
        header_confidence_threshold: Header confidence threshold (default
            ``0.5``; see spec §7.1).
        skip_keywords: Replacement/additional boundary keywords, or ``None``
            to use the v1 default list (see ``heuristics.SKIP_KEYWORDS``).
    """

    sheet_overrides: dict[str, SheetOverride] = field(default_factory=dict)
    header_confidence_threshold: float = 0.5
    skip_keywords: list[str] | None = None


# ---------------------------------------------------------------------------
# Inspection result models (spec §5.1 - §5.5)
# ---------------------------------------------------------------------------


@dataclass
class MergeRegion:
    """A merged-cell region classified as header or body (spec §5.4).

    Attributes:
        range: Merge range in A1 notation (e.g. ``"A1:C1"``).
        kind: ``"header"`` or ``"body"``.
    """

    range: str
    kind: str


@dataclass
class ColumnProfile:
    """Per-column profile (spec §5.3).

    The :attr:`index` is **0-based relative to the table top-left** [D5].

    Attributes:
        index: Column position, 0-based from the table top-left.
        name: The raw header cell value seen at inspection time, or ``None``
            when unavailable. For a column whose leaf header cell is an empty
            continuation of a **vertical** header merge, this is the merge
            anchor's value (issue #1) — the real header text one or more rows
            up; a horizontal group label, by contrast, stays out of ``name``
            and surfaces only in the flattened :attr:`resolved_name`.
        inferred_type: One of ``"number"``, ``"text"``, ``"numeric_text"``,
            ``"date"``, ``"mixed"``. Defaults to the ``"unknown"`` sentinel
            until the Type Profiler (Phase 5) overwrites it; the aggregator
            skips dtype inference for ``"unknown"`` columns so an unprofiled
            column is never silently typed as text.
        null_ratio: Missing ratio (denominator = number of sampled data rows).
        has_formula: Whether any sampled data cell of the column stores a
            formula — set by the Formula Detector (plan v2 Phase 12) from a
            formula-mode (``data_only=False``) sample; defaults ``False``.
        read_hint: ``"as_value"`` (default — cached results are present and
            trustworthy) or ``"as_formula"`` (Phase 12: the column's formula
            cells have **no** cached results, so a value-mode load yields only
            nulls; re-read with ``data_only=False`` to get the formula
            strings). The aggregator skips dtype inference for
            ``"as_formula"`` columns and records an advisory note on the plan.
        resolved_name: The **post-load** column name at this position in the
            loaded frame — flattened (multi-level ``"상위 / 하위"``),
            deduplicated (``.N`` suffixes), and stringified, i.e. the exact
            ``records`` key (adversarial review MEDIUM #2). ``None`` during
            inspection; populated positionally (via :attr:`index`, the
            0-based position within the usecols-selected frame [D5]) when a
            ``TableResult`` is built. :attr:`name` keeps the raw header cell
            value seen at inspection time.
    """

    index: int
    name: str | None = None
    inferred_type: str = "unknown"
    null_ratio: float = 0.0
    has_formula: bool = False
    read_hint: str = "as_value"
    resolved_name: str | None = None


@dataclass
class ReadPlan:
    """The single inspection<->loading contract (spec §5.5).

    Row positions are **pandas 0-based**.  Fields map directly onto
    :func:`pandas.read_excel` parameters; their exact alignment is pinned by
    golden tests [D1].

    Attributes:
        sheet_name: Target sheet.
        engine: pandas engine (fixed ``"openpyxl"``).
        header: Header row (0-based, post-skip normalized). An ``int`` for a
            single header row; a contiguous ``list[int]`` of 0-based absolute
            row indices for a multi-level header band (plan v2 Task 11.2)
            [D6] — absolute equals post-skip there because no row at/above
            the band is ever skipped; ``None`` for headerless.
        usecols: Excel column-letter range (e.g. ``"B:H"``) or ``None`` for
            all columns.
        skiprows: Rows to skip (0-based absolute indices).
        nrows: Number of rows to read, or ``None``.
        dtype_map: Keys = 0-based column-position strings, values = pandas
            dtype strings [D5].
        notes: Loader hints (e.g. body-merge forward-fill recommendation).
    """

    sheet_name: str
    engine: str = "openpyxl"
    header: int | list[int] | None = 0
    usecols: str | None = None
    skiprows: list[int] = field(default_factory=list)
    nrows: int | None = None
    dtype_map: dict[str, str] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)


@dataclass
class TableBlock:
    """One detected table block inside a sheet (1-based inspection coordinates).

    Plan v2 Phase 10 (§4.0): a sheet stacking several tables vertically is
    split into row bands, and every band judged to be a table yields one
    :class:`TableBlock` carrying its own header/boundary/type results and —
    once the aggregator has run — its own :class:`ReadPlan`.

    Coordinate contract [D1]: every ``*_row`` / ``*_col`` field is **openpyxl
    1-based** (inspection domain); only :attr:`read_plan` speaks pandas
    0-based, converted solely by the aggregator. :attr:`ColumnProfile.index`
    inside :attr:`columns` is 0-based from the *block's* table top-left [D5].

    Mirror rule (plan v2 §4.0): ``SheetProfile``'s flat fields (``header_row``
    etc.) mirror the **top-most** block (``blocks[0]``); only bands judged to
    be tables enter ``blocks``, so the mirror is always the top-most *table*.

    Attributes:
        block_index: 0-based position within ``SheetProfile.blocks`` (top-down).
        band_start_row: First row of the enclosing band (1-based, inclusive).
        band_end_row: Last row of the enclosing band (1-based, inclusive).
        header_row: The block's header row (1-based), or ``None``.
        header_confidence: Header estimation confidence (0~1).
        header_provenance: ``"heuristic"`` / ``"manual"`` / ``"default"``.
        data_start_row: First real data row (1-based), or ``None``.
        data_end_row: Last real data row (1-based), or ``None``.
        data_left_col: Left data column boundary (1-based), or ``None``.
        data_right_col: Right data column boundary (1-based), or ``None``.
        skip_rows: Excluded rows (1-based), [D2] overrides already folded for
            rows falling inside this block's band.
        columns: The block's column profiles ([D5] block-local 0-based index).
        read_plan: The block's read plan, or ``None`` until the aggregator runs.
        subtotal_skip_labels: Labels of the excluded non-blank skip rows
            (subtotal/total/low-density), keyed by 1-based sheet row, for the
            aggregator's "no silent loss" note (issue #2); value ``None`` when
            the excluded row has no leading string label.
    """

    block_index: int
    band_start_row: int
    band_end_row: int
    header_row: int | None
    header_confidence: float
    header_provenance: str
    data_start_row: int | None
    data_end_row: int | None
    data_left_col: int | None
    data_right_col: int | None
    skip_rows: list[int]
    columns: list[ColumnProfile]
    read_plan: ReadPlan | None
    #: Labels of the excluded *non-blank* skip rows (subtotal/total/low-density),
    #: keyed by 1-based sheet row, for the aggregator's "no silent loss" note
    #: (issue #2). Value is the row's raw leading label, or ``None`` when the
    #: excluded row has no leading string label. Default empty.
    subtotal_skip_labels: dict[int, str | None] = field(default_factory=dict)


@dataclass
class SheetProfile:
    """Per-sheet profile (spec §5.2).

    Row positions are **openpyxl 1-based**.

    Attributes:
        name: Sheet name.
        is_visible: Visibility.
        is_tabular_candidate: Estimated tabular-form flag.
        is_tabular_provenance: ``"heuristic"`` / ``"manual"`` — records whether
            ``is_tabular_candidate`` came from the heuristic or an
            ``InspectionOptions`` override [D2].
        used_range: Used range (e.g. ``"A1:H120"``).
        used_range_trusted: Whether dimensions are trusted (read_only
            correction flag) [D3].
        max_row: Maximum row (collected in structure mode).
        max_col: Maximum column (collected in structure mode).
        header_row: Estimated header row (1-based), or ``None``.
        header_confidence: Header estimation confidence (0~1).
        header_provenance: ``"heuristic"`` / ``"manual"`` / ``"default"``.
        needs_manual_header: Header estimation failed -> manual spec required.
        is_multi_level_header: Multi-level header flag (v1 always ``False``).
        merges: Merge regions.
        data_start_row: Data start row (1-based), or ``None``.
        data_end_row: Data end row (1-based), or ``None``.
        data_left_col: Left data column boundary (1-based), or ``None``.
        data_right_col: Right data column boundary (1-based), or ``None``.
        skip_rows: Excluded rows such as subtotals/blanks (1-based).
        columns: Column profiles.
        read_plan: Final read plan, or ``None`` until the aggregator runs.
        blocks: Detected table blocks, top-down (plan v2 Phase 10). Empty
            when the sheet contributes no table block (non-tabular / empty /
            headerless fallback) — the flat fields then keep their v1
            semantics. When non-empty, the flat header/boundary/column fields
            mirror ``blocks[0]`` (the top-most table) and ``read_plan`` equals
            ``blocks[0].read_plan``.
        subtotal_skip_labels: Labels of the excluded non-blank skip rows
            (subtotal/total/low-density), keyed by 1-based sheet row, for the
            aggregator's "no silent loss" note (issue #2). Mirrors ``blocks[0]``
            for a multi-band sheet.
    """

    name: str
    is_visible: bool = True
    is_tabular_candidate: bool = True
    is_tabular_provenance: str = "heuristic"
    used_range: str = ""
    used_range_trusted: bool = True
    max_row: int = 0
    max_col: int = 0
    header_row: int | None = None
    header_confidence: float = 0.0
    header_provenance: str = "default"
    needs_manual_header: bool = False
    is_multi_level_header: bool = False
    merges: list[MergeRegion] = field(default_factory=list)
    data_start_row: int | None = None
    data_end_row: int | None = None
    data_left_col: int | None = None
    data_right_col: int | None = None
    skip_rows: list[int] = field(default_factory=list)
    columns: list[ColumnProfile] = field(default_factory=list)
    read_plan: ReadPlan | None = None
    blocks: list[TableBlock] = field(default_factory=list)
    #: Labels of the excluded *non-blank* skip rows (subtotal/total/low-density),
    #: keyed by 1-based sheet row, for the aggregator's "no silent loss" note
    #: (issue #2). Mirrors ``blocks[0]`` for a multi-band sheet. Default empty.
    subtotal_skip_labels: dict[int, str | None] = field(default_factory=dict)


@dataclass
class WorkbookProfile:
    """Top-level workbook profile (spec §5.1).

    Attributes:
        file_path: Original file path.
        sheets: Per-sheet profiles.
        open_errors: Warnings/errors from the open stage.
    """

    file_path: str = ""
    sheets: list[SheetProfile] = field(default_factory=list)
    open_errors: list[str] = field(default_factory=list)
