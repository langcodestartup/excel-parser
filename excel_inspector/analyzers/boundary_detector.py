"""Boundary Detector analyzer (spec §4.5, §7.2) [D4][D6].

Identifies the data region of each tabular sheet *below its header*, so the
loader never double-counts subtotals/totals and never reads beyond the real
data. This is the v1 priority analyzer because preventing aggregation
duplication is the inspector's core value [D6].

Given the ``header_row`` populated by the Header Locator (spec §4.3), this
analyzer reads the rows below it in **data mode** (read_only streaming) [D3]
and applies the §7.2 rules:

* **Column boundaries** — the longest contiguous run of non-empty cells in the
  header row fixes ``data_left_col``/``data_right_col`` (1-based). A blank
  header cell immediately to the left of that run is folded in when the body
  beneath it is a consistently populated key column (issue #16). The width of
  the final span (``data_col_count``) is the density denominator, so a left
  filler column does not drag the density of the real table down.
* **Row density** — ``density(r) = non_empty_cells(r) / data_col_count`` over
  the table column span. A row with ``density < LOW_DENSITY_THRESHOLD`` (0.3) is
  a subtotal/separator candidate. The "only a single column filled"
  (``non_empty == 1``) rule applies **only when the table is >= 3 columns
  wide**, so a 1- or 2-column (key-value / narrow) table's normal rows are not
  misclassified as subtotals (§7.2, MEDIUM #5).
* **Skip keywords** — a row whose **leading (first non-empty) label cell**
  matches any :data:`~excel_inspector.heuristics.SKIP_KEYWORDS` term is a
  ``skip_rows`` candidate regardless of its density. The leading-label scan is
  anchored at the table's own ``data_left_col`` — not sheet column A — so a
  left-margin note outside the table span never shadows (or fakes) a subtotal
  label (plan v2 Phase 13 Step 3, L7). Multi-character keywords
  match by case-insensitive ``startswith``; the single-character ``"계"``
  matches only on exact equality, so data labels such as ``통계청`` / ``회계팀``
  / ``Total Wine`` never false-match (§7.2, MEDIUM #6). Overridable via
  ``InspectionOptions.skip_keywords``.
* **Blank run** — ``BLANK_RUN`` (2) consecutive fully-empty rows terminate the
  block: ``data_end_row`` is fixed at the last real data row above the run and
  scanning stops (a single sheet contributes only its top block in v1). A
  *single* interior blank row below the run threshold is recorded in
  ``skip_rows`` so it never leaks into the loaded frame as an all-NaN row
  (§7.2, MEDIUM #4).

Merge bridging (plan v2 Task 11.1, Phase 11a): the Merge Scanner collects each
sheet's merged ranges (structure mode, unclassified) **before** this analyzer
runs, on ``context.merge_spans``. Two merge-aware rules apply:

* **Header virtual fill (Step 2)** — when computing the header row's column
  span, the empty cells covered by any merge *intersecting the header row* are
  treated as filled, restoring the contiguous span that the merge collapsed
  (the spec §7.2 "defer to merge analysis" case is thereby resolved in-line).
  The fill is virtual: only the span/gap derivation sees it, never the density
  or keyword scans.
* **Trailing merged-group exclusion (fixture-contradiction guard)** — a
  *multi-row* body merge groups several physical rows into one logical record.
  When the resolved table body is entirely merge-free and a fully
  merge-covered row group *trails* it (extends to the end of the scanned
  data), the group's structure differs from the table's flat one-row-one-record
  body — it is judged a separate block (annotation/demo), ``data_end_row`` is
  clamped to the last flat data row, and a warning makes the exclusion
  visible. A table whose body itself contains merged groups (group-label
  tables, spec §4.4 forward-fill) is **never** clamped: any merged group that
  is followed by a flat data row — or that starts the body — marks the merge
  style as interior and disables the rule.

Outputs on each tabular :class:`SheetProfile` (all 1-based [D1]):
``data_start_row`` (first real data row below the header, or ``None`` when the
sheet has no data), ``data_end_row`` (last real data row, trailing totals/blank
rows excluded), ``skip_rows`` (interior + trailing subtotal/total/separator
rows), and ``data_left_col``/``data_right_col``.

Override [D2]: ``SheetOverride.skip_rows_add`` / ``skip_rows_remove`` are
applied to the heuristic ``skip_rows`` after detection (spec §5.0). A sheet
with a ``header_row`` of ``None`` (estimation failed / declared headerless via
override) is left untouched here — there is no anchor to scan from.

Row windows (plan v2 Task 10.2 Step 1): the detection core
(:meth:`BoundaryDetector._detect_block`) accepts an optional ``row_window`` —
a 1-based inclusive ``(start, end)`` row range — generalizing the v1 "whole
sheet" assumption so Phase 10b can bound each stacked table block
independently. The body scan below the header is clamped to
``min(window_end, max_row)``; ``row_window=None`` (the default, used by
:meth:`BoundaryDetector.analyze`) means the whole sheet and reproduces the v1
behavior exactly. The core accumulates everything — including the
unreliable-span fallback that discards column boundaries — on a block-local
:class:`BlockBoundary`, never on the shared profile (guard 3), and the
sheet-level applier copies the result onto the profile.
"""

from __future__ import annotations

import datetime as _dt
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from ..context import InspectionContext
from .._time_axis import is_time_axis_value
from ..heuristics import (
    BLANK_RUN,
    LOW_DENSITY_THRESHOLD,
    TYPE_SAMPLE_ROWS,
    TYPE_SUCCESS_THRESHOLD,
    WIDE_SPARSE_MIN_POPULATED_COLS,
)
from ..models import SheetProfile
from ..options import get_sheet_override, get_skip_keywords
from ..pipeline import Analyzer

if TYPE_CHECKING:  # pragma: no cover - typing only
    from .merge_analyzer import MergeSpan

#: Marker standing in for a header cell that is *virtually* filled because a
#: merge intersecting the header row covers it (plan v2 Task 11.1 Step 2).
#: Non-empty for :func:`_is_empty`, and used only inside the span/gap
#: derivation — it never reaches density or keyword matching.
_MERGE_FILLED: object = object()


def _is_empty(value: object) -> bool:
    """Whether a sampled cell counts as empty for density purposes.

    ``None`` and the empty string are empty; any other value (including ``0``
    and ``False``) is a populated cell.

    Args:
        value: A cached cell value from a read_only sample.

    Returns:
        ``True`` when the cell is empty.
    """

    return value is None or (isinstance(value, str) and value == "")


def _leading_label_raw(row: list[object], left_col: int = 1) -> str | None:
    """Return the row's leading (first non-empty) string label, original case.

    Identical scan to :func:`_leading_label` — anchored at the table's 1-based
    ``left_col``; only a *string* first-populated cell qualifies as a label (a
    leading number/date or an empty span yields ``None``) — but **without**
    lower-casing. Used to render the excluded subtotal/separator row's label in
    the aggregator's "no silent loss" note (issue #2), where the original case
    (e.g. ``"Total"``) must survive.

    Args:
        row: The sampled row values.
        left_col: The table's 1-based left column boundary [D1]; cells to its
            left (the sheet margin) are ignored. Defaults to 1 (whole row).

    Returns:
        The stripped leading label in its original case, or ``None`` when the
        row has no leading string label.
    """

    for value in row[max(left_col, 1) - 1 :]:
        if value is None or (isinstance(value, str) and value == ""):
            continue
        if isinstance(value, str):
            return value.strip()
        # First populated cell is non-string (number/date) -> not a label row.
        return None
    return None


def _leading_label(row: list[object], left_col: int = 1) -> str | None:
    """Return the row's leading (first non-empty) label cell, stripped (§7.2).

    The scan is anchored at the **table's** 1-based ``left_col`` — not sheet
    column A (plan v2 Phase 13 Step 3, L7): a left-margin filler cell outside
    the table span must never shadow (or stand in for) the table's own leading
    label, otherwise a subtotal row whose '소계' sits at ``data_left_col``
    silently leaks into the data when the margin holds unrelated text.

    Only string-like cells qualify as a *label*. Dates and numbers do not anchor
    a keyword match (a numeric first cell is data, not a "소계"/"Total" label), so
    a leading non-string cell yields ``None`` (no label). Leading empty cells are
    skipped to find the first populated cell.

    Args:
        row: The sampled row values.
        left_col: The table's 1-based left column boundary [D1]; cells to its
            left (the sheet margin) are ignored. Defaults to 1 (whole row —
            the pre-L7 behavior for full-width tables).

    Returns:
        The lower-cased, stripped leading label, or ``None`` when the row has no
        leading string label.
    """

    raw = _leading_label_raw(row, left_col)
    return raw.lower() if raw is not None else None


def _matches_keyword(
    row: list[object], keywords: list[str], left_col: int = 1
) -> bool:
    """Whether the row's leading label is a skip keyword (§7.2, MEDIUM #6).

    Matching is anchored to the row's **leading (first non-empty) label cell
    at or after the table's** ``left_col`` (plan v2 Phase 13 Step 3, L7) —
    never an arbitrary substring of the whole row, and never a left-margin
    cell outside the table span — so data labels like ``통계청`` / ``회계팀`` /
    ``Total Wine`` do not false-match (MEDIUM #6) and a margin note cannot
    shadow the table's own '소계' label (L7):

    * Multi-character keywords match by case-insensitive ``startswith`` (so
      ``소계`` matches ``소계`` and ``소계 합산`` but not ``회계``).
    * The single-character keyword ``"계"`` matches only on an **exact** equality
      (so it never fires inside ``통계청``/``회계팀``).

    Args:
        row: The sampled row values.
        keywords: Effective skip keywords (already merged with overrides).
        left_col: The table's 1-based left column boundary [D1]; defaults to 1
            (whole row).

    Returns:
        ``True`` when the leading label matches a keyword under these rules.
    """

    label = _leading_label(row, left_col)
    if label is None:
        return False
    for keyword in keywords:
        kw = keyword.strip().lower()
        if not kw:
            continue
        if len(kw) == 1:
            if label == kw:
                return True
        elif label.startswith(kw):
            return True
    return False


def _header_column_span(
    header: list[object], max_col: int
) -> tuple[int | None, int | None]:
    """Find the table's left/right column boundaries from the header row (§7.2).

    The boundary is the **longest contiguous run** of non-empty cells in the
    header row, expressed 1-based. Ties keep the first (left-most) run. When the
    header is entirely empty, ``(None, None)`` is returned (all columns).

    Args:
        header: The header row's sampled values.
        max_col: The sheet's used column count (1-based extent).

    Returns:
        ``(data_left_col, data_right_col)`` 1-based inclusive, or
        ``(None, None)`` when no populated cell exists.
    """

    width = max_col if max_col and max_col > 0 else len(header)
    best_start: int | None = None
    best_len = 0

    run_start: int | None = None
    for position in range(width):
        value = header[position] if position < len(header) else None
        if not _is_empty(value):
            if run_start is None:
                run_start = position
        else:
            if run_start is not None:
                run_len = position - run_start
                if run_len > best_len:
                    best_len = run_len
                    best_start = run_start
                run_start = None
    if run_start is not None:
        run_len = width - run_start
        if run_len > best_len:
            best_len = run_len
            best_start = run_start

    if best_start is None:
        return None, None
    return best_start + 1, best_start + best_len  # 1-based inclusive


def _header_populated_positions(
    header: list[object], max_col: int
) -> list[int]:
    """Return the 0-based positions of populated cells in the header row."""

    width = max_col if max_col and max_col > 0 else len(header)
    positions: list[int] = []
    for position in range(width):
        value = header[position] if position < len(header) else None
        if not _is_empty(value):
            positions.append(position)
    return positions


def _header_has_interior_gap(header: list[object], max_col: int) -> bool:
    """Whether the header has a non-contiguous (gapped) populated layout (§7.2).

    A merged header typically leaves only its lead cell populated, so two (or
    more) populated cells are separated by an empty interior cell (e.g.
    ``('이름', None, '점수')``). Such an interior gap is the signal that the
    contiguous-run column boundary is unreliable and should defer to merge
    analysis (MEDIUM #7).

    Args:
        header: The header row's sampled values.
        max_col: The sheet's used column count (1-based extent).

    Returns:
        ``True`` when at least two populated cells exist and they are not all
        contiguous (i.e. there is an empty cell between populated cells).
    """

    positions = _header_populated_positions(header, max_col)
    if len(positions) < 2:
        return False
    # Contiguous iff positions form an unbroken range.
    return positions[-1] - positions[0] + 1 != len(positions)


def _is_merge_narrowed_header(
    header: list[object],
    max_col: int,
    left_col: int,
    right_col: int,
) -> bool:
    """Whether a merged header narrowed the column span to a single column (§7.2).

    The longest contiguous run collapsing to width 1 *while the header carries
    additional populated cells beyond an interior gap* is the merged-header case
    the spec defers to merge analysis: only the merge's lead cell is populated,
    so later group labels are separated by the (empty) merged interior. A
    genuine single-column table has exactly one populated header cell and is
    therefore *not* flagged here (MEDIUM #5 / #7).

    Args:
        header: The header row's sampled values.
        max_col: The sheet's used column count (1-based extent).
        left_col: The contiguous-run left boundary (1-based) just derived.
        right_col: The contiguous-run right boundary (1-based) just derived.

    Returns:
        ``True`` when the span is a single column *and* the header has an
        interior gap (a deferred merged header).
    """

    if right_col - left_col + 1 != 1:
        return False
    return _header_has_interior_gap(header, max_col)


def _bridge_merged_header(
    header: list[object],
    spans: Sequence["MergeSpan"],
    header_row: int,
) -> list[object]:
    """Virtually fill the header cells covered by header-row merges (Step 2).

    A merged header leaves only its lead (anchor) cell populated, collapsing
    the contiguous column run to one cell (spec §7.2). Every empty header cell
    covered by a merge *intersecting the header row* is replaced with the
    :data:`_MERGE_FILLED` marker so the span/gap derivation sees the run the
    merge really spans (plan v2 Task 11.1 Step 2). The original row is never
    mutated; with no intersecting merge it is returned unchanged.

    Args:
        header: The header row's sampled values.
        spans: The sheet's collected merge spans (1-based bounds [D1]).
        header_row: The block's 1-based header row.

    Returns:
        The (possibly) bridged header row for span computation only.
    """

    covered_cols: set[int] = set()
    for span in spans:
        if span.min_row <= header_row <= span.max_row:
            covered_cols.update(range(span.min_col, span.max_col + 1))
    if not covered_cols:
        return header

    width = max(len(header), max(covered_cols))
    bridged: list[object] = list(header) + [None] * (width - len(header))
    for col in covered_cols:
        if _is_empty(bridged[col - 1]):
            bridged[col - 1] = _MERGE_FILLED
    return bridged


def _merged_group_rows(
    spans: Sequence["MergeSpan"], header_row: int
) -> set[int]:
    """Rows grouped by a *multi-row* merge strictly below the header (1-based).

    A merge spanning several rows below the header binds those physical rows
    into one logical record/group — the structural signal behind the trailing
    merged-group exclusion (module docstring). Single-row (horizontal) body
    merges group nothing and merges touching the header row are header
    structure, so neither contributes.

    Args:
        spans: The sheet's collected merge spans (1-based bounds [D1]).
        header_row: The block's 1-based header row.

    Returns:
        The set of 1-based rows covered by multi-row body-side merges.
    """

    rows: set[int] = set()
    for span in spans:
        if span.min_row > header_row and span.max_row > span.min_row:
            rows.update(range(span.min_row, span.max_row + 1))
    return rows


def _span_density(
    row: list[object], left_col: int, right_col: int
) -> tuple[float, int]:
    """Density of a row within the table column span (§7.2).

    Args:
        row: The sampled row values (0-based list, sheet columns).
        left_col: Table left boundary (1-based, inclusive).
        right_col: Table right boundary (1-based, inclusive).

    Returns:
        ``(density, non_empty_count)`` where density is ``non_empty / span``.
    """

    span = right_col - left_col + 1
    if span <= 0:
        return 0.0, 0
    non_empty = 0
    for col in range(left_col - 1, right_col):
        value = row[col] if col < len(row) else None
        if not _is_empty(value):
            non_empty += 1
    return non_empty / span, non_empty


def _cell_at(row: list[object], one_based_col: int) -> object | None:
    """Return a 1-based cell from a sampled row, or ``None`` if absent."""

    index = one_based_col - 1
    return row[index] if 0 <= index < len(row) else None


def _is_wide_sparse_axis_data_row(
    row: list[object], left_col: int, table_width: int, non_empty: int
) -> bool:
    """Whether a low-density row is still a wide time-series data row.

    In issue #22, the row's first table column is the date/period axis and the
    many right-hand series columns are sparsely populated. Such rows are real
    observations, not subtotal/separator rows, even when their density is below
    the generic threshold.
    """

    if table_width < WIDE_SPARSE_MIN_POPULATED_COLS or non_empty == 0:
        return False
    return is_time_axis_value(_cell_at(row, left_col))


def _value_kind(value: object) -> str | None:
    """Coarse non-empty value kind used by the leading-key-column heuristic."""

    if _is_empty(value):
        return None
    if isinstance(value, bool):
        return "other"
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return "date"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    return "other"


@dataclass
class _LeadingKeyColumnStats:
    """Evidence that a blank-header column is really a body key column."""

    observed_rows: int = 0
    populated_rows: int = 0
    kinds: dict[str, int] = field(default_factory=dict)

    def add(self, value: object) -> None:
        """Record one body row for this candidate column."""

        self.observed_rows += 1
        kind = _value_kind(value)
        if kind is None:
            return
        self.populated_rows += 1
        self.kinds[kind] = self.kinds.get(kind, 0) + 1

    def is_consistent_key(self) -> bool:
        """Whether the sampled evidence is strong enough to extend left."""

        if self.observed_rows <= 0 or self.populated_rows <= 0:
            return False
        populated_ratio = self.populated_rows / self.observed_rows
        if populated_ratio < TYPE_SUCCESS_THRESHOLD:
            return False
        kind_ratio = max(self.kinds.values()) / self.populated_rows
        return kind_ratio >= TYPE_SUCCESS_THRESHOLD


def _blank_leading_header_columns(
    header: list[object], left_col: int
) -> list[int]:
    """Blank header columns immediately left of ``left_col`` (nearest first)."""

    columns: list[int] = []
    for col in range(left_col - 1, 0, -1):
        value = header[col - 1] if col - 1 < len(header) else None
        if not _is_empty(value):
            break
        columns.append(col)
    return columns


@dataclass
class BlockBoundary:
    """Block-local boundary detection result (plan v2 Task 10.2 guard 3).

    Everything :meth:`BoundaryDetector._detect_block` derives for one row
    window lives here, so the unreliable-span fallback mutates *this* local
    state and never a shared :class:`SheetProfile` field — Phase 10b runs the
    detection once per block, and one block's fallback must not leak into
    another block's (or the sheet's) boundaries.

    All coordinates are openpyxl 1-based, inspection domain [D1].
    ``skip_rows`` holds the *heuristic* skips only — the [D2]
    ``skip_rows_add`` / ``skip_rows_remove`` overrides are folded in by the
    caller when applying the result. ``warnings`` are forwarded to the context
    by the caller in detection order (deterministic).
    """

    data_start_row: int | None = None
    data_end_row: int | None = None
    data_left_col: int | None = None
    data_right_col: int | None = None
    skip_rows: list[int] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    #: Labels of the *non-blank* skip rows (subtotal/total/low-density), keyed
    #: by 1-based sheet row, for the aggregator's "no silent loss" note (issue
    #: #2). The value is the row's raw leading label, or ``None`` when the
    #: excluded row has no leading string label (a purely sparse row). Interior
    #: blank separator rows carry no data and are deliberately absent.
    subtotal_skip_labels: dict[int, str | None] = field(default_factory=dict)


class BoundaryDetector(Analyzer):
    """Detect the data region of each tabular sheet (spec §4.5, §7.2) [D6]."""

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "boundary_detector"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Populate data-region boundaries on every tabular sheet.

        Non-tabular sheets, sheets with no detected ``header_row`` (estimation
        failed or declared headerless), and sheets reachable only without a
        loader are skipped. For the rest the §7.2 rules are applied below the
        header and the §5.0 ``skip_rows`` overrides folded in [D2].

        Args:
            context: Shared context carrying a ready :class:`Loader` and the
                header-located sheet profiles.

        Returns:
            The same context with boundary fields populated on tabular sheets.
        """

        loader = context.loader
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            if profile.header_row is None:
                # No anchor to scan from (no-header / headerless override).
                continue
            if loader is None:
                context.add_warning(
                    f"boundary_detector: no loader available; cannot detect "
                    f"boundaries for sheet {profile.name!r}"
                )
                continue
            self._detect(context, profile)
        return context

    def _detect(
        self, context: InspectionContext, profile: SheetProfile
    ) -> None:
        """Run the §7.2 boundary detection for one sheet and apply the result.

        Thin applier over the block-local :meth:`_detect_block` core
        (whole-sheet window): the core computes a :class:`BlockBoundary`
        without touching the profile (guard 3); this method copies it onto the
        shared profile, folds in the [D2] skip overrides, and forwards the
        block's warnings to the context.
        """

        header_row = profile.header_row
        assert header_row is not None  # guarded by analyze()

        result = self._detect_block(context, profile, header_row)

        profile.data_left_col = result.data_left_col
        profile.data_right_col = result.data_right_col
        profile.data_start_row = result.data_start_row
        profile.data_end_row = result.data_end_row
        profile.skip_rows = self._apply_skip_overrides(
            context, profile, result.skip_rows
        )
        profile.subtotal_skip_labels = dict(result.subtotal_skip_labels)
        for warning in result.warnings:
            context.add_warning(warning)

    def _detect_block(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        header_row: int,
        row_window: tuple[int, int] | None = None,
    ) -> BlockBoundary:
        """Detect one block's data region below ``header_row`` (§7.2 core).

        Generalizes the v1 whole-sheet scan to a 1-based inclusive
        ``row_window`` (plan v2 Task 10.2 Step 1): the body scan below the
        header stops at ``min(window_end, max_row)`` instead of running to the
        sheet's end. ``row_window=None`` means the whole sheet (v1 behavior).

        The header row is read first (it is near the top). A bounded streaming
        evidence scan may sample up to ``TYPE_SAMPLE_ROWS`` rows to preserve
        blank-header leading key columns (issue #16), then the rows below it
        are processed in a forward streaming pass — no row beyond the header is
        materialized into a list, so a large table is bounded by running state
        only (spec §8; no full materialization) [D3].

        Merge bridging (plan v2 Task 11.1): the Merge Scanner's collected
        spans (``context.merge_spans``) drive the header-span virtual fill and
        the trailing merged-group exclusion (module docstring). With no
        collected spans both rules no-op (pre-11a behavior).

        All derived state — boundaries, skips, the unreliable-span fallback
        that discards the column boundaries, and the deferral warnings — is
        accumulated on the returned :class:`BlockBoundary` only (guard 3);
        nothing on the shared profile is mutated here.

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet being scanned (read for name/dimensions only).
            header_row: The block's 1-based header row (the scan anchor).
            row_window: Optional 1-based inclusive ``(start, end)`` row window;
                ``None`` means the whole sheet (v1 behavior).

        Returns:
            The block-local :class:`BlockBoundary` (heuristic ``skip_rows``,
            [D2] overrides not yet applied).
        """

        result = BlockBoundary()
        window_end = None if row_window is None else row_window[1]

        # Collected (unclassified) merge spans from the Merge Scanner (plan v2
        # Task 11.1 Step 1). Empty when the scanner did not run (standalone /
        # partial-context use) — every merge-aware rule then no-ops, exactly
        # the pre-11a behavior.
        spans = context.merge_spans.get(profile.name, [])

        header_values = self._read_header_row(context, profile, header_row)
        # Virtual fill (Task 11.1 Step 2): bridge the header cells covered by
        # merges intersecting the header row, so a merged header's collapsed
        # column run is restored before the span is derived. Only the span/gap
        # derivation sees the bridged row.
        bridged_header = _bridge_merged_header(header_values, spans, header_row)
        left_col, right_col = _header_column_span(
            bridged_header, profile.max_col
        )

        # spec §7.2: a merged header can leave only its lead cell populated, so
        # the longest contiguous run collapses to a single column even though the
        # header genuinely spans more (an interior gap from the merge separates
        # later populated cells). With the Merge Scanner's spans the gap is
        # bridged above; this fallback remains for a gap *not* explained by any
        # collected merge (or when no scanner ran): the column boundary is left
        # *unresolved* (None) and deferred to merge analysis (§4.4) rather than
        # emitting a broken 1-column span.
        if left_col is not None and _is_merge_narrowed_header(
            bridged_header, profile.max_col, left_col, right_col
        ):
            left_col = None
            right_col = None

        if left_col is not None and right_col is not None:
            left_col = self._extend_left_boundary_for_blank_key_columns(
                context,
                profile,
                header_values,
                left_col,
                right_col,
                header_row,
                window_end,
            )

        # spec §4.5: column boundaries are reported only when the table occupies
        # *part* of the used range. A left-anchored, full-width table (starts at
        # column 1 and reaches max_col) needs no usecols restriction, so its
        # boundaries are left ``None`` ("all columns"); the internal span is
        # still used as the density denominator below.
        max_col = profile.max_col if profile.max_col and profile.max_col > 0 else None
        is_partial_span = left_col is not None and not (
            left_col == 1 and (max_col is None or right_col == max_col)
        )
        result.data_left_col = left_col if is_partial_span else None
        result.data_right_col = right_col if is_partial_span else None

        # Without a usable column span there is no table to bound. A header whose
        # span was discarded as merge-narrowed (above) records the deferral
        # warning so the None/None outcome is explicit (spec §7.2, MEDIUM #7);
        # Phase 6 merge analysis will later resolve the true span.
        if left_col is None or right_col is None:
            result.data_start_row = None
            result.data_end_row = None
            result.skip_rows = []
            if _header_has_interior_gap(bridged_header, profile.max_col):
                result.warnings.append(
                    f"boundary_detector: header column span for sheet "
                    f"{profile.name!r} collapsed to a single column (merged "
                    f"header); column boundaries discarded pending merge "
                    f"analysis"
                )
            return result

        keywords = get_skip_keywords(context.options)
        table_width = right_col - left_col + 1

        data_start: int | None = None
        data_end: int | None = None
        skip_rows: list[int] = []
        # Labels of the non-blank (subtotal/total/low-density) skip rows, keyed
        # by 1-based row, captured at the skip point so the aggregator can name
        # each excluded row in its "no silent loss" note (issue #2).
        subtotal_labels: dict[int, str | None] = {}
        blank_run = 0
        # Blank rows recorded for the *current* (not-yet-terminating) run; held
        # so they can be retracted if the run grows to a terminator (trailing,
        # not interior) — see the ``non_empty == 0`` branch (MEDIUM #4).
        pending_blanks: list[int] = []
        # Whether any *non-blank* row was flagged as a skip (keyword/low-density).
        # Distinguishes the "unreliable span" fallback (real content rows all
        # rejected) from a sheet whose only skips are stray blank rows.
        non_blank_skip = False

        # Trailing merged-group exclusion state (plan v2 Task 11.1; module
        # docstring). ``merged_rows`` are the rows bound into logical groups by
        # multi-row body merges; ``trailing_merged`` collects merge-grouped
        # data rows seen since the last *flat* (ungrouped) data row, and
        # ``merged_interior`` records that a merged group was followed by a
        # flat row (or started the body) — the table's own style then includes
        # merges and the exclusion rule is disabled.
        merged_rows = _merged_group_rows(spans, header_row)
        last_flat: int | None = None
        trailing_merged: list[int] = []
        merged_interior = False

        # Scan strictly below the header (1-based rows header_row+1 .. end,
        # bounded by the row window when one is given), streaming so no row
        # beyond the header is materialized [D3].
        for one_based, row in self._iter_rows_below(
            context, profile, header_row, window_end
        ):
            density, non_empty = _span_density(row, left_col, right_col)

            if non_empty == 0:
                blank_run += 1
                if blank_run >= BLANK_RUN:
                    # Block terminated by the blank run. The blank rows of *this*
                    # run are trailing (beyond the data), not interior, so drop
                    # any that were tentatively recorded below.
                    if pending_blanks:
                        skip_rows = skip_rows[: -len(pending_blanks)]
                    break
                # A *single* interior blank row (below the run threshold) is
                # still recorded as a skip so it never leaks into the loaded
                # frame as an all-NaN row (spec §7.2, MEDIUM #4). It is held in
                # ``pending_blanks`` so it can be retracted if this turns out to
                # be the start of a *terminating* run (the run is trailing, not
                # interior). The blank-run termination logic above is unchanged.
                skip_rows.append(one_based)
                pending_blanks.append(one_based)
                continue
            blank_run = 0
            pending_blanks.clear()

            # L7 (plan v2 Phase 13 Step 3): the keyword scan anchors at the
            # table's own left column, so a left-margin note outside the span
            # can neither shadow a real subtotal label nor fake one.
            is_keyword = _matches_keyword(row, keywords, left_col)
            # The "single column filled" (non_empty == 1) subtotal rule applies
            # only when the table is at least 3 columns wide; a 1- or 2-column
            # (key-value / narrow) table's normal rows would otherwise be
            # misclassified as subtotals (spec §7.2, MEDIUM #5).
            is_axis_data_row = _is_wide_sparse_axis_data_row(
                row, left_col, table_width, non_empty
            )
            is_low_density = False if is_axis_data_row else (
                density < LOW_DENSITY_THRESHOLD
                or (non_empty == 1 and table_width >= 3)
            )

            if is_keyword or is_low_density:
                skip_rows.append(one_based)
                # Record the excluded row's label (original case) for the
                # aggregator's no-silent-loss note (issue #2). A low-density row
                # without a leading string label records None.
                subtotal_labels[one_based] = _leading_label_raw(row, left_col)
                non_blank_skip = True
                continue

            # A genuine data row.
            if data_start is None:
                data_start = one_based
            data_end = one_based
            if one_based in merged_rows:
                trailing_merged.append(one_based)
            else:
                if trailing_merged:
                    # A merged group followed by a flat data row is interior
                    # table structure (group-label style) -> rule disabled.
                    merged_interior = True
                    trailing_merged = []
                last_flat = one_based

        # A single blank row dangling at EOF (no terminating second blank and no
        # data row after it) is trailing, not interior; retract it so it never
        # appears as a skip beyond the data region (MEDIUM #4).
        if pending_blanks:
            skip_rows = skip_rows[: -len(pending_blanks)]

        # Trailing merged-group exclusion (plan v2 Task 11.1; module
        # docstring): the resolved body is merge-free (``merged_interior`` is
        # False) yet a multi-row-merge-grouped row group trails the last flat
        # data row — a structure change marking a separate block
        # (annotation/demo), not more table records. data_end_row is clamped
        # back to the last flat data row and the exclusion is made visible
        # (never a silent loss).
        if trailing_merged and not merged_interior and last_flat is not None:
            result.warnings.append(
                f"boundary_detector: sheet {profile.name!r}: trailing "
                f"merged-row group (rows {trailing_merged[0]}-"
                f"{trailing_merged[-1]}) below a merge-free table body "
                f"excluded from the data region (data_end_row clamped to "
                f"{last_flat}); verify it is not table data"
            )
            data_end = last_flat

        # Trailing skip rows that sit after the last real data row remain in
        # skip_rows but must not extend data_end_row (spec §4.5: data_end_row is
        # the last *real* data row, excluding trailing totals/blank rows).
        if data_start is None and non_blank_skip:
            # Every candidate row was flagged single-column/low-density against
            # the chosen span: the span is unreliable (e.g. an unresolved header
            # merge leaves an interior gap, narrowing the contiguous run). Rather
            # than emit a broken narrow usecols and skip all rows, fall back to
            # "no column restriction" and let a later phase (Merge Analyzer,
            # §4.4) refine. The data region stays unresolved (start/end None).
            # Block-local only (guard 3): the discard lives on this result, not
            # on the shared profile.
            result.data_left_col = None
            result.data_right_col = None
            result.warnings.append(
                f"boundary_detector: no data rows resolved within the header "
                f"column span for sheet {profile.name!r}; column boundaries "
                f"discarded pending merge analysis"
            )
            skip_rows = []

        result.data_start_row = data_start
        result.data_end_row = data_end
        result.skip_rows = skip_rows
        # Keep only labels of rows that survived as final skips (the
        # unreliable-span fallback empties skip_rows; blank-row retraction
        # removes blanks — never present here anyway). Iteration order follows
        # ascending row order, so the aggregator's notes are deterministic.
        result.subtotal_skip_labels = {
            row: subtotal_labels[row]
            for row in skip_rows
            if row in subtotal_labels
        }
        return result

    def _extend_left_boundary_for_blank_key_columns(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        header_values: list[object],
        left_col: int,
        right_col: int,
        header_row: int,
        end_row: int | None,
    ) -> int:
        """Include blank-header leading key columns when body evidence is clear.

        Issue #16: many time-series tables leave the top-left header cell blank
        while column A carries the row key/date axis. The header-only
        contiguous-run rule would choose B:... and silently drop A via
        ``usecols``. To avoid pulling in true left margins, only blank header
        columns immediately adjacent to the resolved header run are considered,
        and each must be populated with one consistent value kind across the
        sampled body rows whose original header span contains data.
        """

        candidate_cols = _blank_leading_header_columns(
            header_values, left_col
        )
        if not candidate_cols:
            return left_col

        stats = {
            col: _LeadingKeyColumnStats() for col in candidate_cols
        }
        blank_run = 0
        observed = 0
        for _, row in self._iter_rows_below(
            context, profile, header_row, end_row
        ):
            _, non_empty = _span_density(row, left_col, right_col)
            if non_empty == 0:
                blank_run += 1
                if blank_run >= BLANK_RUN:
                    break
                continue
            blank_run = 0

            observed += 1
            for col in candidate_cols:
                value = row[col - 1] if col - 1 < len(row) else None
                stats[col].add(value)
            if observed >= TYPE_SAMPLE_ROWS:
                break

        extended_left = left_col
        for col in candidate_cols:
            if not stats[col].is_consistent_key():
                break
            extended_left = col
        return extended_left

    @staticmethod
    def _apply_skip_overrides(
        context: InspectionContext,
        profile: SheetProfile,
        skip_rows: list[int],
    ) -> list[int]:
        """Fold ``skip_rows_add`` / ``skip_rows_remove`` into the result [D2].

        Args:
            context: Shared context (for the per-sheet override).
            profile: The sheet whose overrides apply.
            skip_rows: Heuristic skip rows (1-based) before overrides.

        Returns:
            The sorted, de-duplicated 1-based skip-row list after overrides.
        """

        result = set(skip_rows)
        override = get_sheet_override(context.options, profile.name)
        if override is not None:
            result.update(override.skip_rows_add)
            result.difference_update(override.skip_rows_remove)
        return sorted(result)

    @staticmethod
    def _read_header_row(
        context: InspectionContext, profile: SheetProfile, header_row: int
    ) -> list[object]:
        """Read just the (1-based) ``header_row`` values in data mode [D3].

        Only the single header row is materialized; the body is streamed
        separately by :meth:`_iter_rows_below` (spec §8; no full
        materialization).

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet whose header row is read.
            header_row: The 1-based header row number.

        Returns:
            The header row's values, or ``[]`` when the sheet/row is absent.
        """

        workbook = context.loader.data_workbook()
        try:
            worksheet = workbook[profile.name]
        except KeyError:  # pragma: no cover - defensive
            return []
        for row in worksheet.iter_rows(
            min_row=header_row, max_row=header_row, values_only=True
        ):
            return list(row)
        return []

    @staticmethod
    def _iter_rows_below(
        context: InspectionContext,
        profile: SheetProfile,
        header_row: int,
        end_row: int | None = None,
    ):
        """Yield ``(one_based, row)`` for every row strictly below the header [D3].

        A single forward **streaming** pass (read_only) over rows
        ``header_row + 1 .. end``; each row is yielded and dropped, so the
        body is never materialized into a list (spec §8). The end of the pass
        is ``min(end_row, max_row)`` when a window bound is given (plan v2
        Task 10.2 Step 1); with ``end_row=None`` (whole sheet, v1 behavior) it
        is ``max_row``, and when ``max_row`` is untrusted/zero the stream runs
        to natural EOF (or to ``end_row`` when bounded).

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet whose body rows are streamed.
            header_row: The 1-based header row (rows at/above it are skipped).
            end_row: Optional 1-based inclusive last row of the scan window;
                ``None`` means no window bound (whole sheet).

        Yields:
            ``(one_based, row_values)`` tuples in ascending sheet order.
        """

        workbook = context.loader.data_workbook()
        try:
            worksheet = workbook[profile.name]
        except KeyError:  # pragma: no cover - defensive
            return

        max_row = (
            profile.max_row if profile.max_row and profile.max_row > 0 else None
        )
        if end_row is not None:
            max_row = end_row if max_row is None else min(max_row, end_row)
        for one_based, row in enumerate(
            worksheet.iter_rows(
                min_row=header_row + 1, max_row=max_row, values_only=True
            ),
            start=header_row + 1,
        ):
            yield one_based, list(row)
