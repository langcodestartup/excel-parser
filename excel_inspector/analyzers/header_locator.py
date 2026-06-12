"""Header Locator analyzer (spec §4.3, §7.1) [D2][D4].

Estimates the header row of each tabular sheet. The header is not always the
first row: a title block, an "as-of" date, or unit notes may precede it. This
analyzer reads only the top ``HEADER_SCAN_ROWS`` rows in **data mode**
(read_only streaming) [D3] and scores each candidate row with the §7.1
formula::

    score(r) = 0.5 * non_empty_string_ratio(r)
             + 0.3 * type_consistency(rows r+1 .. r+5) * (n_below / 5)
             + 0.2 * distinctness(r vs rows r+1 .. r+5)

where ``n_below`` is the number of rows actually observed in the lookahead
window. The evidence factor (issue #8) keeps a bottom-of-sample candidate —
whose 1-row window is trivially self-consistent — from outscoring the true
header in small mixed-type tables.

The highest-scoring row becomes ``header_row`` (1-based) with
``header_confidence = score`` and ``header_provenance = "heuristic"``. When the
best score is below the threshold (``InspectionOptions.header_confidence_threshold``,
default :data:`~excel_inspector.heuristics.HEADER_CONFIDENCE_THRESHOLD`) the
sheet is left for manual specification: ``header_row=None``,
``header_confidence=0.0``, ``needs_manual_header=True`` (spec §4.3, §9).

Override [D2]: when ``InspectionOptions`` carries a ``header_row`` override for
the sheet, scoring is skipped entirely and the override value is recorded with
``header_provenance="manual"`` and ``header_confidence=1.0``.

v1 produces a single header row only; ``is_multi_level_header`` is left
``False`` (multi-level headers are deferred to v1+ [D6]).

Row windows (plan v2 Task 10.2 Step 1): the scoring core (:meth:`HeaderLocator.
_locate`) accepts an optional ``row_window`` — a 1-based inclusive ``(start,
end)`` row range — generalizing the v1 "whole sheet" assumption so Phase 10b
can locate a header inside each stacked table block independently. The scan is
clamped to ``min(window_start + HEADER_SCAN_ROWS - 1, window_end)`` and the
resulting header row is the **absolute** sheet row ``window_start +
best_index``. ``row_window=None`` (the default, used by :meth:`HeaderLocator.
analyze`) means the whole sheet and reproduces the v1 behavior exactly.
"""

from __future__ import annotations

import datetime as _dt
from typing import TYPE_CHECKING

from ..context import InspectionContext
from ..heuristics import (
    HEADER_LOOKAHEAD_ROWS,
    HEADER_SCAN_ROWS,
    HEADER_WEIGHT_DISTINCTNESS,
    HEADER_WEIGHT_NON_EMPTY_STRING,
    HEADER_WEIGHT_TYPE_CONSISTENCY,
)
from ..models import SheetProfile
from ..options import (
    get_header_confidence_threshold,
    get_sheet_override,
    has_header_override,
)
from ..pipeline import Analyzer

if TYPE_CHECKING:  # pragma: no cover - typing only
    from openpyxl.worksheet.worksheet import Worksheet

#: Sampled-cell category labels used by the scoring functions.
_CAT_EMPTY = "empty"
_CAT_STRING = "string"
_CAT_NUMBER = "number"
_CAT_DATE = "date"
_CAT_OTHER = "other"


def _categorize(value: object) -> str:
    """Classify one sampled cell value into a coarse type category.

    Booleans are deliberately *not* treated as numbers (``bool`` is a subclass
    of ``int`` in Python) so a boolean cell does not inflate the numeric
    signal; they fall through to ``_CAT_OTHER``.

    Args:
        value: The cached cell value from a read_only sample (already
            ``values_only``).

    Returns:
        One of the ``_CAT_*`` category constants.
    """

    if value is None:
        return _CAT_EMPTY
    if isinstance(value, str):
        return _CAT_EMPTY if value == "" else _CAT_STRING
    if isinstance(value, bool):
        return _CAT_OTHER
    if isinstance(value, (int, float)):
        return _CAT_NUMBER
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return _CAT_DATE
    return _CAT_OTHER


def _non_empty_string_ratio(row: list[object], col_count: int) -> float:
    """Fraction of the used columns whose cell is a non-empty string (§7.1).

    Args:
        row: The candidate row's sampled values.
        col_count: Total used columns (the denominator).

    Returns:
        ``non_empty_string_cells / col_count`` in ``[0, 1]``; ``0.0`` when
        ``col_count`` is zero.
    """

    if col_count <= 0:
        return 0.0
    strings = sum(
        1 for v in row[:col_count] if _categorize(v) == _CAT_STRING
    )
    return strings / col_count


def _type_consistency(
    below: list[list[object]], col_count: int
) -> float:
    """Average per-column type consistency of the rows below a candidate (§7.1).

    For each column, among its non-empty cells in ``below`` we take the
    fraction held by the single most common category; a column with no
    non-empty cells contributes a neutral ``0.0``. The per-column fractions are
    averaged over all ``col_count`` columns.

    Args:
        below: Up to :data:`HEADER_LOOKAHEAD_ROWS` rows immediately below the
            candidate.
        col_count: Total used columns.

    Returns:
        Mean per-column consistency in ``[0, 1]``; ``0.0`` when there are no
        rows below or no columns.
    """

    if col_count <= 0 or not below:
        return 0.0

    total = 0.0
    for col in range(col_count):
        counts: dict[str, int] = {}
        non_empty = 0
        for row in below:
            value = row[col] if col < len(row) else None
            category = _categorize(value)
            if category == _CAT_EMPTY:
                continue
            non_empty += 1
            counts[category] = counts.get(category, 0) + 1
        if non_empty == 0:
            continue
        total += max(counts.values()) / non_empty
    return total / col_count


def _distinctness(
    candidate: list[object], below: list[list[object]], col_count: int
) -> float:
    """How different the candidate row looks from the rows below it (§7.1).

    Two per-column signals are combined and averaged across columns:

    * **type difference** — 1.0 when the candidate's category differs from the
      dominant category of that column below it, else 0.0.
    * **length difference** — the relative difference in string-length between
      the candidate cell and the mean cell length below, clamped to ``[0, 1]``.

    A header row (labels) typically differs in type (text vs numbers) and in
    cell-length pattern from the data rows below it, so a real header scores
    high here while a data-like row scores low.

    Args:
        candidate: The candidate row's sampled values.
        below: Rows immediately below the candidate.
        col_count: Total used columns.

    Returns:
        Mean per-column distinctness in ``[0, 1]``; ``0.0`` when there are no
        rows below or no columns.
    """

    if col_count <= 0 or not below:
        return 0.0

    total = 0.0
    for col in range(col_count):
        cand_value = candidate[col] if col < len(candidate) else None
        cand_cat = _categorize(cand_value)

        below_cats: list[str] = []
        below_lengths: list[int] = []
        for row in below:
            value = row[col] if col < len(row) else None
            category = _categorize(value)
            if category == _CAT_EMPTY:
                continue
            below_cats.append(category)
            below_lengths.append(len(str(value)))

        if not below_cats:
            # Nothing below in this column -> a populated candidate cell is, by
            # itself, a distinguishing signal; an empty candidate cell is not.
            total += 0.0 if cand_cat == _CAT_EMPTY else 1.0
            continue

        dominant = max(set(below_cats), key=below_cats.count)
        type_diff = 1.0 if cand_cat != dominant else 0.0

        cand_len = 0 if cand_cat == _CAT_EMPTY else len(str(cand_value))
        mean_below_len = sum(below_lengths) / len(below_lengths)
        denom = max(cand_len, mean_below_len, 1.0)
        length_diff = abs(cand_len - mean_below_len) / denom

        total += (type_diff + length_diff) / 2.0
    return total / col_count


def _score_row(
    index: int, rows: list[list[object]], col_count: int
) -> float:
    """Score a single candidate row at ``index`` within ``rows`` (§7.1).

    Args:
        index: 0-based position of the candidate within the sampled ``rows``.
        rows: All sampled rows (top ``HEADER_SCAN_ROWS``).
        col_count: Total used columns (scoring denominator).

    Returns:
        The weighted §7.1 score in ``[0, 1]``.
    """

    candidate = rows[index]
    below = rows[index + 1 : index + 1 + HEADER_LOOKAHEAD_ROWS]

    # Issue #8: scale the consistency term by the observed lookahead evidence
    # (n_below / HEADER_LOOKAHEAD_ROWS). A shrunken window is trivially
    # self-consistent — a single row below yields 1.0 per column — which
    # handed bottom-of-sample data rows a free type-consistency win over the
    # true header in small mixed-type tables (spec §7.1).
    evidence = len(below) / HEADER_LOOKAHEAD_ROWS

    return (
        HEADER_WEIGHT_NON_EMPTY_STRING
        * _non_empty_string_ratio(candidate, col_count)
        + HEADER_WEIGHT_TYPE_CONSISTENCY
        * _type_consistency(below, col_count)
        * evidence
        + HEADER_WEIGHT_DISTINCTNESS
        * _distinctness(candidate, below, col_count)
    )


class HeaderLocator(Analyzer):
    """Estimate each tabular sheet's header row (spec §4.3, §7.1) [D4]."""

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "header_locator"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Estimate ``header_row`` for every tabular sheet.

        Non-tabular sheets are skipped (they are excluded from loading, spec
        §9). For each tabular sheet, a ``header_row`` override [D2] short-
        circuits scoring; otherwise the top sample rows are read in data mode
        and scored per §7.1.

        Args:
            context: Shared context carrying a ready :class:`Loader` and the
                enumerated sheet profiles.

        Returns:
            The same context with header fields populated on tabular sheets.
        """

        loader = context.loader
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue

            if has_header_override(context.options, profile.name):
                self._apply_override(context, profile)
                continue

            if loader is None:
                context.add_warning(
                    f"header_locator: no loader available; cannot estimate "
                    f"header for sheet {profile.name!r}"
                )
                continue

            self._estimate(context, profile)
        return context

    def _apply_override(
        self, context: InspectionContext, profile: SheetProfile
    ) -> None:
        """Record a manual ``header_row`` override on ``profile`` [D2].

        The override value (possibly ``None``) is authoritative: scoring is
        skipped and the field is stamped ``header_provenance="manual"`` with
        ``header_confidence=1.0``.
        """

        override = get_sheet_override(context.options, profile.name)
        assert override is not None  # guaranteed by has_header_override
        profile.header_row = override.header_row
        profile.header_confidence = 1.0
        profile.header_provenance = "manual"
        profile.needs_manual_header = False

    def _estimate(
        self, context: InspectionContext, profile: SheetProfile
    ) -> None:
        """Score the top sample rows and record the best header candidate.

        Thin applier over the block-local :meth:`_locate` core (whole-sheet
        window): the core scores without mutating anything, and this method
        applies the estimate (or the manual-needed fallback + warning) to the
        shared profile.
        """

        threshold = get_header_confidence_threshold(context.options)
        header_row, best_score = self._locate(context, profile)

        if header_row is None:
            self._mark_manual(profile)
            context.add_warning(
                f"header_locator: no sampleable rows for sheet "
                f"{profile.name!r}; needs manual header"
            )
            return

        if best_score < threshold:
            self._mark_manual(profile)
            context.add_warning(
                f"header_locator: best header score "
                f"{max(best_score, 0.0):.3f} below threshold {threshold:.3f} "
                f"for sheet {profile.name!r}; needs manual header"
            )
            return

        profile.header_row = header_row  # 1-based [D1]
        profile.header_confidence = best_score
        profile.header_provenance = "heuristic"
        profile.needs_manual_header = False
        profile.is_multi_level_header = False  # v1: single header only [D6]

    def _locate(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        row_window: tuple[int, int] | None = None,
    ) -> tuple[int | None, float]:
        """Score §7.1 header candidates within ``row_window`` (block-local core).

        Generalizes the v1 "whole sheet" scan to a 1-based inclusive row window
        (plan v2 Task 10.2 Step 1) so Phase 10b can score each stacked table
        block independently. No profile field is mutated here — the caller
        decides how to apply the estimate (sheet-level today, per-block next).

        Window semantics: at most :data:`HEADER_SCAN_ROWS` rows are sampled
        from the window's top — the scan end is clamped to
        ``min(window_start + HEADER_SCAN_ROWS - 1, window_end)`` (guard 2).
        With ``row_window=None`` (whole sheet) the sample is rows
        ``1 .. HEADER_SCAN_ROWS``, exactly the v1 read. A windowed call also
        switches the scoring denominator to the window-local used-column count
        (guard 1, see :meth:`_col_count`).

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet whose rows are sampled.
            row_window: Optional 1-based inclusive ``(start, end)`` row window;
                ``None`` means the whole sheet (v1 behavior).

        Returns:
            ``(header_row, best_score)`` where ``header_row`` is the
            **absolute** 1-based sheet row of the best-scoring candidate
            (``window_start + best_index``, guard 2 — no hardcoded sheet-top
            offset), or ``(None, 0.0)`` when the window yields no sampleable
            rows/columns. The threshold judgment is the caller's.
        """

        window_start = 1 if row_window is None else row_window[0]
        window_end = None if row_window is None else row_window[1]

        rows = self._sample_rows(context, profile, window_start, window_end)
        col_count = self._col_count(
            profile, rows, window_scoped=row_window is not None
        )
        if not rows or col_count <= 0:
            return None, 0.0

        best_index = -1
        best_score = -1.0
        for index in range(len(rows)):
            score = _score_row(index, rows, col_count)
            if score > best_score:
                best_score = score
                best_index = index

        if best_index < 0:  # pragma: no cover - scores are >= 0 for any row
            return None, 0.0
        return window_start + best_index, best_score

    @staticmethod
    def _mark_manual(profile: SheetProfile) -> None:
        """Mark a sheet as needing a manual header (estimation failed)."""

        profile.header_row = None
        profile.header_confidence = 0.0
        profile.header_provenance = "heuristic"
        profile.needs_manual_header = True
        profile.is_multi_level_header = False

    def _sample_rows(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        window_start: int = 1,
        window_end: int | None = None,
    ) -> list[list[object]]:
        """Read a row window's top ``HEADER_SCAN_ROWS`` rows in data mode [D3].

        The scan end is clamped to ``min(window_start + HEADER_SCAN_ROWS - 1,
        window_end)`` (plan v2 Task 10.2 guard 2). The defaults
        (``window_start=1``, ``window_end=None``) reproduce the v1 whole-sheet
        read exactly: rows ``1 .. HEADER_SCAN_ROWS``, with no clamp against
        the (possibly untrusted) sheet dimensions.

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet whose rows are sampled.
            window_start: First row of the window (1-based, inclusive).
            window_end: Last row of the window (1-based, inclusive), or
                ``None`` for no window bound (whole sheet).

        Returns:
            A list of rows (each a list of cached cell values). Empty when the
            worksheet cannot be located, the window is empty, or it yields no
            rows.
        """

        workbook = context.loader.data_workbook()
        try:
            worksheet = workbook[profile.name]
        except KeyError:  # pragma: no cover - defensive
            return []

        scan_end = window_start + HEADER_SCAN_ROWS - 1
        if window_end is not None:
            scan_end = min(scan_end, window_end)
        if scan_end < window_start:
            return []

        rows: list[list[object]] = []
        for row in worksheet.iter_rows(
            min_row=window_start, max_row=scan_end, values_only=True
        ):
            rows.append(list(row))
        return rows

    @staticmethod
    def _col_count(
        profile: SheetProfile,
        rows: list[list[object]],
        *,
        window_scoped: bool = False,
    ) -> int:
        """Determine the used-column denominator for scoring (§7.1).

        Whole-sheet scoring (``window_scoped=False``, the v1 path) prefers the
        trusted ``max_col`` from the sheet enumerator and otherwise falls back
        to the widest non-empty extent observed in the sample so the ratio
        denominator stays meaningful.

        **Window-scoped** scoring never uses the sheet-global ``max_col``: in a
        sheet stacking a narrow table next to a wide one, the global column
        count would dilute the narrow band's scores into a "not a table"
        misjudgment (plan v2 Task 10.2 guard 1), so the denominator is always
        the widest non-empty extent within the window's own sample.
        """

        if not window_scoped and profile.max_col and profile.max_col > 0:
            return profile.max_col

        widest = 0
        for row in rows:
            last_non_empty = 0
            for position, value in enumerate(row, start=1):
                if _categorize(value) != _CAT_EMPTY:
                    last_non_empty = position
            widest = max(widest, last_non_empty)
        return widest
