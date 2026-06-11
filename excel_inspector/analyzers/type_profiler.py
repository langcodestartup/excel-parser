"""Type Profiler analyzer (spec §4.6, §7.3) [D4][D5].

Infers each column's data type, missing ratio, and (column) identity from a
deterministic sample of the data region. Runs after the Boundary Detector and
before the Merge Analyzer (spec §3 topology), so the data region
(``data_start_row``/``data_end_row``), the interior ``skip_rows``, and the
table column span (``data_left_col``/``data_right_col``) are already populated.

Sampling (spec §7.3):
    ``TYPE_SAMPLE_ROWS = min(200, data_row_count)`` rows are drawn **evenly and
    deterministically** from the inclusive data region ``data_start_row ..
    data_end_row`` in **data mode** (read_only streaming) [D3], with any
    interior ``skip_rows`` (subtotals / blank separators) excluded *before*
    sampling so they never pollute the type signal or the ``null_ratio``
    denominator. Even, index-based selection (no RNG) keeps the sample
    deterministic across runs (implementation plan §5.3). The sampled row
    numbers are computed *before* any cell read (from the already-resolved
    boundaries), so only the header row and the sampled rows are materialized —
    never the whole sheet (spec §8; no full scan) [D3].

Per-column classification order (spec §7.3), applied to the non-missing sampled
values of each column:

1.  Drop missing cells (``None`` / empty string).
2.  Date parse-success rate ``>= TYPE_SUCCESS_THRESHOLD`` (0.95) -> ``date``.
3.  Numeric parse-success rate ``>= 0.95`` -> if the original *stored* form is a
    string (an Excel text cell holding a numeric-looking value) -> ``numeric_text``,
    otherwise -> ``number``.
4.  Every non-missing value is a non-numeric string -> ``text``.
5.  No single type reaches 0.95 -> ``mixed``.

A column whose sample is entirely missing is left as ``text`` with
``null_ratio == 1.0`` (no type signal to commit to a numeric/date type).

Outputs on each tabular :class:`SheetProfile`: a :class:`ColumnProfile` list
where ``index`` is **0-based from the table top-left** [D5], ``name`` is the
header-row value at that column, ``inferred_type`` is overwritten from the
``"unknown"`` sentinel to the classified type, and ``null_ratio`` uses the
number of *sampled data rows* as its denominator (spec §5.3, §7.3).

A sheet with no resolved data region (``data_start_row`` / ``data_end_row``
``None`` — empty / header-only / merge-deferred) profiles no columns; the
``"unknown"`` sentinel survives so the aggregator skips dtype inference for it
(spec §5.3).

Row windows (plan v2 Task 10.2 Step 1): the profiling core
(:meth:`TypeProfiler._profile_block`) takes every boundary — header row, data
region, skip rows, column span — as an **explicit block-local parameter**
instead of reading the shared profile, generalizing the v1 "the profile's data
region is the sheet's only table" assumption so Phase 10b can profile each
stacked table block independently. An optional ``row_window`` (1-based
inclusive) additionally clamps the sampled region to
``max(data_start_row, window_start) .. min(data_end_row, window_end)``;
``row_window=None`` (the default, used by :meth:`TypeProfiler.analyze`) means
no extra clamp and reproduces the v1 behavior exactly.
"""

from __future__ import annotations

import datetime as _dt

from ..context import InspectionContext
from ..heuristics import TYPE_SAMPLE_ROWS, TYPE_SUCCESS_THRESHOLD
from ..models import ColumnProfile, SheetProfile
from ..pipeline import Analyzer


def _is_missing(value: object) -> bool:
    """Whether a sampled cell counts as missing for type profiling (§7.3).

    ``None`` and the empty string are missing; any other value (including ``0``
    and ``False``) is present.

    Args:
        value: A cached cell value from a read_only sample.

    Returns:
        ``True`` when the cell is missing.
    """

    return value is None or (isinstance(value, str) and value == "")


def _is_date_value(value: object) -> bool:
    """Whether a cached cell value is a date/datetime (§7.3).

    In data mode (``data_only=True``) openpyxl returns native
    :class:`datetime.datetime` / :class:`datetime.date` objects for date-format
    cells, so a successful "date parse" is simply an ``isinstance`` check.
    :class:`datetime.time` is *not* counted as a date (a bare clock time is not
    a calendar date column).

    Args:
        value: A non-missing sampled cell value.

    Returns:
        ``True`` when the value is a date or datetime (but not a bare time).
    """

    return isinstance(value, (_dt.datetime, _dt.date)) and not isinstance(
        value, _dt.time
    )


def _is_numeric_value(value: object) -> bool:
    """Whether a non-missing cell value parses as a number (§7.3).

    A real numeric cell (``int`` / ``float``) counts, as does a *string* whose
    text is a valid number (the ``numeric_text`` case — a digit string stored as
    Excel text such as ``"007"``). Booleans are deliberately excluded (``bool``
    subclasses ``int`` in Python but a logical flag is not a numeric column), and
    dates are excluded (handled by :func:`_is_date_value` first).

    Args:
        value: A non-missing sampled cell value.

    Returns:
        ``True`` when the value is numeric or a numeric-looking string.
    """

    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return False
        try:
            float(text)
        except ValueError:
            return False
        return True
    return False


def _classify_column(values: list[object], sample_row_count: int) -> str:
    """Classify one column's sampled values per the §7.3 judgment order.

    Args:
        values: The column's sampled cell values (one per sampled data row,
            including missing cells so the order matches the sample).
        sample_row_count: The number of sampled data rows (the ``null_ratio``
            denominator; unused here but kept for symmetry with the caller).

    Returns:
        One of ``"number"``, ``"text"``, ``"numeric_text"``, ``"date"``,
        ``"mixed"`` (spec §5.3).
    """

    present = [v for v in values if not _is_missing(v)]
    if not present:
        # No non-missing value to type; an all-missing column commits to text
        # (there is no numeric/date evidence to claim a stronger type).
        return "text"

    total = len(present)
    date_hits = sum(1 for v in present if _is_date_value(v))
    if date_hits / total >= TYPE_SUCCESS_THRESHOLD:
        return "date"

    numeric_hits = sum(1 for v in present if _is_numeric_value(v))
    if numeric_hits / total >= TYPE_SUCCESS_THRESHOLD:
        # Original storage form decides number vs numeric_text: if *every*
        # non-missing value was stored as a string (an Excel text cell), the
        # column is digit-strings-as-text; any native numeric storage makes it a
        # plain number column.
        all_stored_as_string = all(isinstance(v, str) for v in present)
        return "numeric_text" if all_stored_as_string else "number"

    string_hits = sum(1 for v in present if isinstance(v, str))
    if string_hits == total:
        # Every non-missing value is a (non-numeric) string -> text.
        return "text"

    # No single type reached the threshold -> mixed.
    return "mixed"


def _even_sample_indices(count: int, sample_size: int) -> list[int]:
    """Return ``sample_size`` evenly-spaced 0-based indices over ``count`` (§7.3).

    Deterministic, RNG-free even selection so repeated runs draw the same rows
    (implementation plan §5.3). When ``sample_size >= count`` every index is
    returned (the whole population). Indices are unique and ascending.

    Args:
        count: Population size (number of eligible data rows).
        sample_size: Desired sample size (already capped at ``count`` by the
            caller, but defended here too).

    Returns:
        A sorted list of unique 0-based indices into the population.
    """

    if count <= 0 or sample_size <= 0:
        return []
    if sample_size >= count:
        return list(range(count))
    # Evenly spread sample_size picks across [0, count): index i maps to
    # floor(i * count / sample_size). The mapping is strictly increasing for
    # sample_size <= count, so the picks are already unique and ascending.
    return [(i * count) // sample_size for i in range(sample_size)]


class TypeProfiler(Analyzer):
    """Infer per-column types from a deterministic data sample (spec §4.6) [D4]."""

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "type_profiler"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Profile column types on every tabular sheet with a data region.

        Non-tabular sheets, sheets without a resolved data region, and sheets
        reachable only without a loader are skipped (their ``"unknown"`` column
        type sentinel survives so the aggregator omits their dtype inference).

        Args:
            context: Shared context carrying a ready :class:`Loader` and the
                boundary-detected sheet profiles.

        Returns:
            The same context with ``columns`` populated on profiled sheets.
        """

        loader = context.loader
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            if profile.data_start_row is None or profile.data_end_row is None:
                # No resolved data region (empty / header-only / merge-deferred).
                continue
            if loader is None:
                context.add_warning(
                    f"type_profiler: no loader available; cannot profile "
                    f"types for sheet {profile.name!r}"
                )
                continue
            self._profile(context, profile)
        return context

    def _profile(
        self, context: InspectionContext, profile: SheetProfile
    ) -> None:
        """Run the §7.3 type inference for one sheet and apply the result.

        Thin applier over the block-local :meth:`_profile_block` core: the
        core profiles from explicit boundary parameters (here the profile's
        own fields — the whole-sheet single-table case) without mutating
        anything, and this method assigns the produced columns to the shared
        profile.
        """

        assert profile.data_start_row is not None  # guarded by analyze()
        assert profile.data_end_row is not None

        columns = self._profile_block(
            context,
            profile,
            header_row=profile.header_row,
            data_start_row=profile.data_start_row,
            data_end_row=profile.data_end_row,
            skip_rows=profile.skip_rows,
            data_left_col=profile.data_left_col,
            data_right_col=profile.data_right_col,
        )
        if columns is None:
            # No header and no sampled data rows -> nothing to profile.
            return
        profile.columns = columns

    def _profile_block(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        *,
        header_row: int | None,
        data_start_row: int,
        data_end_row: int,
        skip_rows: list[int],
        data_left_col: int | None,
        data_right_col: int | None,
        row_window: tuple[int, int] | None = None,
    ) -> list[ColumnProfile] | None:
        """Profile one block's columns from explicit boundaries (§7.3 core).

        Every boundary is a block-local parameter (plan v2 Task 10.2 Step 1) so
        Phase 10b can profile each table block independently — nothing is read
        from or written to the shared profile's boundary/column fields here
        (``profile`` supplies only the sheet name and ``max_col`` fallback).

        Only the header row and the deterministic data-row *sample* are read
        from the stream — never the whole sheet — so a large table is profiled
        from at most ``TYPE_SAMPLE_ROWS`` (200) rows plus the header (spec §8;
        no full scan / materialization) [D3].

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet being profiled (name / ``max_col`` only).
            header_row: The block's 1-based header row, or ``None`` (headerless).
            data_start_row: The block's 1-based first data row (inclusive).
            data_end_row: The block's 1-based last data row (inclusive).
            skip_rows: The block's 1-based interior skip rows.
            data_left_col: The block's 1-based left column boundary, or ``None``.
            data_right_col: The block's 1-based right column boundary, or
                ``None`` (the span then falls back to ``1 .. max_col``).
            row_window: Optional 1-based inclusive ``(start, end)`` row window
                further clamping the sampled region to
                ``max(data_start_row, start) .. min(data_end_row, end)``;
                ``None`` means no extra clamp (v1 behavior).

        Returns:
            The block's :class:`ColumnProfile` list (``index`` 0-based from the
            block's table top-left [D5]), or ``None`` when there is nothing to
            profile (no header values and no sampled rows).
        """

        sample_rows = self._sampled_row_numbers(  # 1-based, sorted
            data_start_row, data_end_row, skip_rows, row_window
        )
        # A vertical header merge bridges its anchor value into the leaf-row
        # name: the leaf cell is an empty covered continuation, so the raw
        # header value lives in the anchor cell one or more rows up (issue #1,
        # spec §4.4). Horizontal group merges never bridge (their label belongs
        # to ``resolved_name``, not the raw inspection ``name``).
        header_anchors = self._vertical_header_anchors(
            context, profile, header_row
        )
        header_values, sampled, anchor_values = self._read_sample(
            context, profile, header_row, sample_rows,
            set(header_anchors.values()),
        )
        sample_row_count = len(sampled)
        if sample_row_count == 0 and not header_values:
            # No header and no sampled data rows -> nothing to profile.
            return None

        left_col, right_col = self._table_span(
            data_left_col, data_right_col, profile.max_col
        )

        columns: list[ColumnProfile] = []
        for table_index, sheet_col in enumerate(range(left_col, right_col + 1)):
            cell_index = sheet_col - 1  # 0-based into the row list
            values = [
                row[cell_index] if cell_index < len(row) else None
                for row in sampled
            ]
            inferred = _classify_column(values, sample_row_count)
            null_ratio = self._null_ratio(values, sample_row_count)
            name = self._header_name(header_values, cell_index)
            if name is None and sheet_col in header_anchors:
                # Empty leaf cell covered by a vertical header merge -> take the
                # anchor cell's value (same column, one or more rows up).
                anchor_row = header_anchors[sheet_col]
                name = self._header_name(
                    anchor_values.get(anchor_row, []), cell_index
                )
            columns.append(
                ColumnProfile(
                    index=table_index,
                    name=name,
                    inferred_type=inferred,
                    null_ratio=null_ratio,
                )
            )

        return columns

    @staticmethod
    def _table_span(
        data_left_col: int | None,
        data_right_col: int | None,
        max_col: int | None,
    ) -> tuple[int, int]:
        """Return the 1-based inclusive table column span (left, right) (§7.3).

        When the boundary detector resolved explicit column boundaries
        (``data_left_col``/``data_right_col``, a partial-width table) those are
        used. A full-width / left-anchored table leaves them ``None``; the span
        then defaults to ``1 .. max_col`` (all used columns). The boundaries
        are block-local parameters (plan v2 Task 10.2 Step 1), not profile
        reads, so a Phase 10b block can pass its own span.
        """

        left = data_left_col if data_left_col else 1
        right = (
            data_right_col
            if data_right_col
            else (max_col if max_col and max_col > 0 else left)
        )
        if right < left:
            right = left
        return left, right

    @staticmethod
    def _header_name(header_values: list[object], cell_index: int) -> str | None:
        """Return the column name from the header row, or ``None``.

        Args:
            header_values: The header row's sampled values.
            cell_index: 0-based sheet-column index of the column.

        Returns:
            The header cell's value coerced to ``str`` (stripped) when present
            and non-empty, otherwise ``None``.
        """

        if cell_index >= len(header_values):
            return None
        value = header_values[cell_index]
        if value is None:
            return None
        if isinstance(value, str):
            stripped = value.strip()
            return stripped or None
        return str(value)

    @staticmethod
    def _vertical_header_anchors(
        context: InspectionContext,
        profile: SheetProfile,
        header_row: int | None,
    ) -> dict[int, int]:
        """Map a 1-based column to the anchor row of a bridging vertical merge.

        Issue #1: a header merge that spans **vertically** down onto the leaf
        ``header_row`` leaves that row's cell empty (an openpyxl covered
        continuation), so the column's raw header value really lives in the
        merge anchor one or more rows up. Such a merge bridges column ``c``
        when its anchor column is ``c`` (``min_col == c``) and its row span
        reaches the leaf header from strictly above::

            min_row < header_row <= max_row

        ``min_row == header_row`` (anchor already on the leaf row, value reads
        directly) and merges not reaching the leaf row (e.g. a horizontal group
        header one row up over a present leaf cell) are excluded — only the
        anchor *column* is filled, so a horizontal group label never leaks into
        a different column's raw ``name`` (that flattening is ``resolved_name``
        territory; ``test_merged_header_columns_profiled`` pins the contrast).

        The unclassified spans come from the :class:`MergeScanner` via
        ``context.merge_spans`` (populated before the Type Profiler runs); a
        sheet absent from the map (scanner unavailable, spec §6) simply yields
        no anchors and the v1 ``None`` name survives.

        Args:
            context: Shared context carrying the collected ``merge_spans``.
            profile: The sheet being profiled (its name keys ``merge_spans``).
            header_row: The block's 1-based leaf header row, or ``None``
                (headerless -> no bridge).

        Returns:
            ``{column_1based: anchor_row_1based}``; empty when nothing bridges.
            On the (degenerate) chance of several vertical merges anchoring the
            same column, the lowest anchor (closest above the leaf) wins.
        """

        if header_row is None:
            return {}
        anchors: dict[int, int] = {}
        for span in context.merge_spans.get(profile.name, []):
            if span.min_row < header_row <= span.max_row:
                existing = anchors.get(span.min_col)
                if existing is None or span.min_row > existing:
                    anchors[span.min_col] = span.min_row
        return anchors

    @staticmethod
    def _sampled_row_numbers(
        data_start_row: int,
        data_end_row: int,
        skip_rows: list[int],
        row_window: tuple[int, int] | None = None,
    ) -> list[int]:
        """Compute the deterministic sample's 1-based row numbers (§7.3).

        Eligible data rows are the inclusive span ``data_start_row ..
        data_end_row`` with any interior ``skip_rows`` removed *before* sampling
        (so subtotals / blank separators never enter the type signal or the
        ``null_ratio`` denominator). The sample size is
        ``min(TYPE_SAMPLE_ROWS, eligible_count)``, drawn at even index positions
        for determinism (implementation plan §5.3).

        Crucially this is computed **without reading any cells** — it depends
        only on the already-resolved boundaries — so the caller can then read
        *only* the sampled rows from the stream rather than materializing the
        whole sheet (spec §8; no full scan) [D3].

        The boundaries are block-local parameters (plan v2 Task 10.2 Step 1).
        An optional ``row_window`` clamps the eligible span to
        ``max(data_start_row, window_start) .. min(data_end_row, window_end)``
        (an empty intersection yields no sample); ``None`` means no extra
        clamp (v1 behavior).

        Args:
            data_start_row: The block's 1-based first data row (inclusive).
            data_end_row: The block's 1-based last data row (inclusive).
            skip_rows: The block's 1-based interior skip rows.
            row_window: Optional 1-based inclusive ``(start, end)`` clamp.

        Returns:
            The sorted, ascending 1-based row numbers to read (possibly empty).
        """

        start = data_start_row
        end = data_end_row
        if row_window is not None:
            start = max(start, row_window[0])
            end = min(end, row_window[1])

        skip = set(skip_rows)
        eligible = [
            one_based
            for one_based in range(start, end + 1)
            if one_based not in skip
        ]
        if not eligible:
            return []

        sample_size = min(TYPE_SAMPLE_ROWS, len(eligible))
        picks = _even_sample_indices(len(eligible), sample_size)
        return [eligible[pick] for pick in picks]

    @staticmethod
    def _null_ratio(values: list[object], sample_row_count: int) -> float:
        """Missing ratio over the sampled data rows (spec §5.3 / §7.3).

        Denominator is the number of *sampled data rows* (not just the populated
        ones), so an all-missing column reports ``1.0``.

        Args:
            values: The column's sampled cell values.
            sample_row_count: The number of sampled data rows (denominator).

        Returns:
            ``missing_cells / sample_row_count`` in ``[0, 1]``; ``0.0`` when no
            rows were sampled.
        """

        if sample_row_count <= 0:
            return 0.0
        missing = sum(1 for v in values if _is_missing(v))
        return missing / sample_row_count

    def _read_sample(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        header_row: int | None,
        sample_rows: list[int],
        anchor_rows: set[int] | None = None,
    ) -> tuple[list[object], list[list[object]], dict[int, list[object]]]:
        """Stream only the header row and the sampled data rows in data mode [D3].

        A single forward streaming pass is made over the worksheet, but only the
        header row (when present), the rows whose 1-based number is in
        ``sample_rows``, and any vertical-merge ``anchor_rows`` (issue #1) are
        retained; every other row is read past and dropped. The pass stops as
        soon as everything needed has been collected, so a large table costs at
        most ``TYPE_SAMPLE_ROWS`` (+1) retained data rows plus a handful of
        header-anchor rows (all at or above the header) and never the whole
        sheet (spec §8; no full materialization) [D3].

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet whose header/sample rows are read.
            header_row: The block's 1-based header row, or ``None``
                (headerless) — a block-local parameter, not a profile read
                (plan v2 Task 10.2 Step 1).
            sample_rows: Sorted 1-based data-row numbers to retain (from
                :meth:`_sampled_row_numbers`).
            anchor_rows: 1-based rows of vertical-merge header anchors to retain
                for the merged-header name bridge (issue #1); always at or above
                ``header_row`` and so already within the streamed prefix.

        Returns:
            ``(header_values, sampled_rows, anchor_values)`` where
            ``header_values`` is the header row's values (``[]`` for a headerless
            sheet), ``sampled_rows`` are the retained data rows in ascending
            sheet order, and ``anchor_values`` maps each retained anchor row's
            1-based number to its values (``{}`` when no anchors are requested).
        """

        anchor_rows = anchor_rows or set()
        workbook = context.loader.data_workbook()
        try:
            worksheet = workbook[profile.name]
        except KeyError:  # pragma: no cover - defensive
            return [], [], {}

        wanted = set(sample_rows)
        # The last row we need to touch; once past it we can stop streaming.
        last_needed = max(
            [r for r in sample_rows]
            + ([header_row] if header_row else [])
            + list(anchor_rows),
            default=0,
        )
        if last_needed <= 0:
            return [], [], {}

        header_values: list[object] = []
        sampled: list[list[object]] = []
        anchor_values: dict[int, list[object]] = {}
        for one_based, row in enumerate(
            worksheet.iter_rows(
                min_row=1, max_row=last_needed, values_only=True
            ),
            start=1,
        ):
            if header_row is not None and one_based == header_row:
                header_values = list(row)
            if one_based in wanted:
                sampled.append(list(row))
            if one_based in anchor_rows:
                anchor_values[one_based] = list(row)
            if one_based >= last_needed:
                break
        return header_values, sampled, anchor_values
