"""Formula Detector analyzer (spec §4.7, plan v2 Phase 12) [D6].

Decides, per table column, whether the stored cells are *formulas* and — when
they are — whether loading should take the cached results or the formula
strings:

* ``has_formula=True`` when any sampled data cell, read in **formula mode**
  (``read_only=True, data_only=False``), is a ``str`` starting with ``"="``.
* ``read_hint="as_formula"`` when every one of those formula cells reads
  ``None`` in **data mode** (``data_only=True``) — the cached results are
  empty (typically a file written programmatically and never opened/recalced
  in Excel), so a value-mode load would yield an all-null column. A warning
  ``"column N: formula cache empty (file never opened in Excel?)"`` makes the
  situation visible, and the Plan Aggregator skips dtype inference for the
  column (recording an advisory note on the plan instead).
* ``read_hint="as_value"`` when at least one formula cell has a cached
  result — pandas' normal cached-value load is then trustworthy.

Laziness (plan v2 §6 Step 1/Step 4): opening a third workbook handle costs a
full zip parse, so the formula-mode workbook is opened **only when the file
can possibly contain formulas**. A stored formula is serialized as an ``<f>``
element inside ``xl/worksheets/*.xml`` (OOXML SpreadsheetML); a raw byte scan
for that markup over the worksheet members decides formula presence without
opening any workbook. Text content can never false-positive the scan — a
literal ``"<f"`` inside a cell string is XML-escaped to ``&lt;f`` — and the
``<f[ >/]`` pattern cannot match other worksheet tags (``<formula>``,
``<filterColumn>``, ``<firstHeader>`` … all continue with a letter). When the
scan finds nothing, :meth:`Loader.formula_workbook` is **never called** —
pinned by an open-counter test over the formula-free corpus.

Block coupling (plan v2 §6 Step 2: "블록 구조와 결합"): detection runs per
:class:`~excel_inspector.models.TableBlock`, sampling each block's own data
region (same deterministic row choice as the Type Profiler, §7.3) and
flagging the block's own :class:`~excel_inspector.models.ColumnProfile`
objects ([D5] block-local 0-based index). The flat sheet mirror shares those
instances (Phase 10 mirror rule), so the sheet-level view updates with the
top-most block automatically. A blockless sheet with a resolved data region
(e.g. an explicit headerless override) falls back to the flat fields.

Warning order: sheets in workbook order, blocks top-down, columns ascending —
deterministic, matching the Phase 10 guard-6 convention.

Coordinates [D1]: everything here speaks openpyxl 1-based rows; only the
Plan Aggregator converts to pandas 0-based.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from ..context import InspectionContext
from ..models import ColumnProfile, SheetProfile, TableBlock
from ..pipeline import Analyzer
from .type_profiler import TypeProfiler

#: A stored formula inside a worksheet part: an ``<f>`` element opening
#: (``<f>``, ``<f t="shared" …>``, or self-closing ``<f/>``). Other worksheet
#: tags starting with ``f`` (``<formula>``, ``<formula1>``, ``<filterColumn>``,
#: ``<firstHeader>`` …) continue with a letter and cannot match.
_FORMULA_MARKUP_RE = re.compile(rb"<f[ >/]")

#: Worksheet XML members of an OOXML package (``xl/worksheets/sheet1.xml`` …).
_WORKSHEET_MEMBER_RE = re.compile(r"^xl/worksheets/[^/]+\.xml$")


def _workbook_has_formula_markup(path: str | Path) -> bool:
    """Whether any worksheet part of the package contains formula markup.

    Raw byte scan over the ``xl/worksheets/*.xml`` zip members for the
    ``<f…`` element — the lazy gate that lets a formula-free inspection skip
    the formula-mode workbook entirely (plan v2 §6 Step 4).

    Args:
        path: Path to the ``.xlsx`` package.

    Returns:
        ``True`` when formula markup is present. Undecidable packages
        (unreadable zip — cannot normally happen this deep in the pipeline)
        conservatively return ``True`` so the real openpyxl open, with its
        proper domain-error translation, gets the final say.
    """

    try:
        with zipfile.ZipFile(path) as archive:
            for member in archive.namelist():
                if not _WORKSHEET_MEMBER_RE.match(member):
                    continue
                if _FORMULA_MARKUP_RE.search(archive.read(member)):
                    return True
    except (OSError, zipfile.BadZipFile, KeyError):  # pragma: no cover
        return True  # defensive: let formula_workbook() decide/translate
    return False


def _is_formula_text(value: object) -> bool:
    """Whether a formula-mode cell value is a formula (plan v2 §6 Step 2).

    In formula mode openpyxl returns the stored formula as a string starting
    with ``"="``; anything else (literals, ``None``) is not a formula.
    """

    return isinstance(value, str) and value.startswith("=")


#: Suffix of every cache-empty warning (plan v2 §6 Step 2) — stable so tests
#: and callers can recognize the condition without parsing the whole message.
_CACHE_EMPTY_SUFFIX = "formula cache empty (file never opened in Excel?)"


class FormulaDetector(Analyzer):
    """Flag formula columns and recommend as_value/as_formula (spec §4.7)."""

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "formula_detector"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Run per-block formula detection over every tabular sheet.

        The lazy gate runs first: without formula markup in the package, the
        method returns immediately and the formula-mode workbook is never
        opened (plan v2 §6 Step 4 — pinned by an open-counter test).

        Args:
            context: Shared context carrying a ready loader and the
                block-analyzed sheet profiles (columns final).

        Returns:
            The same context with ``has_formula``/``read_hint`` set on the
            affected :class:`ColumnProfile` objects and cache-empty warnings
            recorded.
        """

        loader = context.loader
        if loader is None:
            # Partial test contexts: nothing to read from.
            return context
        if not _workbook_has_formula_markup(loader.path):
            return context  # lazy guarantee: no formula workbook open

        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            if profile.blocks:
                for block in profile.blocks:
                    self._detect_block(context, profile, block)
            elif (
                profile.data_start_row is not None
                and profile.data_end_row is not None
                and profile.columns
            ):
                # Blockless fallback (e.g. explicit headerless override):
                # the flat fields are the only table description.
                self._detect_columns(
                    context,
                    profile,
                    columns=profile.columns,
                    data_start_row=profile.data_start_row,
                    data_end_row=profile.data_end_row,
                    skip_rows=profile.skip_rows,
                    data_left_col=profile.data_left_col,
                    data_right_col=profile.data_right_col,
                    label=f"sheet {profile.name!r}",
                )
        return context

    def _detect_block(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        block: TableBlock,
    ) -> None:
        """Detect formulas in one table block's columns (plan v2 §6 Step 2).

        The block's :class:`ColumnProfile` instances are shared with the flat
        sheet mirror (Phase 10 mirror rule), so flagging them here updates
        both views consistently. Blocks without a resolved data region or
        without profiled columns have nothing to sample and are skipped.
        """

        if (
            block.data_start_row is None
            or block.data_end_row is None
            or not block.columns
        ):
            return
        self._detect_columns(
            context,
            profile,
            columns=block.columns,
            data_start_row=block.data_start_row,
            data_end_row=block.data_end_row,
            skip_rows=block.skip_rows,
            data_left_col=block.data_left_col,
            data_right_col=block.data_right_col,
            label=(
                f"sheet {profile.name!r} table T{block.block_index + 1}"
            ),
        )

    def _detect_columns(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        *,
        columns: list[ColumnProfile],
        data_start_row: int,
        data_end_row: int,
        skip_rows: list[int],
        data_left_col: int | None,
        data_right_col: int | None,
        label: str,
    ) -> None:
        """Sample one table's data region in both modes and flag its columns.

        Row choice reuses the Type Profiler's deterministic even sampling
        (§7.3) so repeated inspections see identical evidence. The same row
        numbers are then streamed from the formula-mode and data-mode
        workbooks — index ``i`` of both result lists is the same sheet row —
        and each column is judged:

        * any formula-mode cell ``str`` starting with ``"="`` ->
          ``has_formula=True``;
        * cached (data-mode) values **at exactly those formula cells** all
          ``None`` -> ``read_hint="as_formula"`` + cache-empty warning;
          otherwise ``read_hint="as_value"``.

        Args:
            context: Shared context with a ready loader.
            profile: The owning sheet (name / ``max_col`` fallback only).
            columns: The table's column profiles ([D5] table-local index),
                mutated in place.
            data_start_row: First data row (1-based, inclusive) [D1].
            data_end_row: Last data row (1-based, inclusive) [D1].
            skip_rows: Interior skip rows (1-based) excluded from sampling.
            data_left_col: Left table column (1-based), or ``None``.
            data_right_col: Right table column (1-based), or ``None``.
            label: Human-readable scope prefix for warnings (sheet / table).
        """

        sample_rows = TypeProfiler._sampled_row_numbers(  # noqa: SLF001
            data_start_row, data_end_row, skip_rows
        )
        if not sample_rows:
            return

        formula_rows = self._read_rows(
            context.loader.formula_workbook(), profile.name, sample_rows
        )
        cached_rows = self._read_rows(
            context.loader.data_workbook(), profile.name, sample_rows
        )
        if not formula_rows:  # pragma: no cover - defensive (sheet vanished)
            return

        left_col, _ = TypeProfiler._table_span(  # noqa: SLF001
            data_left_col, data_right_col, profile.max_col
        )
        for column in columns:
            cell_index = (left_col + column.index) - 1  # 0-based in row tuple
            formula_positions = [
                i
                for i, row in enumerate(formula_rows)
                if cell_index < len(row) and _is_formula_text(row[cell_index])
            ]
            if not formula_positions:
                continue
            column.has_formula = True
            cached_values = [
                cached_rows[i][cell_index]
                if i < len(cached_rows) and cell_index < len(cached_rows[i])
                else None
                for i in formula_positions
            ]
            if all(value is None for value in cached_values):
                column.read_hint = "as_formula"
                context.add_warning(
                    f"formula_detector: {label}: column {column.index}: "
                    f"{_CACHE_EMPTY_SUFFIX}"
                )
            else:
                column.read_hint = "as_value"

    @staticmethod
    def _read_rows(
        workbook: object, sheet_name: str, sample_rows: list[int]
    ) -> list[list[object]]:
        """Stream only ``sample_rows`` from one worksheet (single pass) [D3].

        Mirrors the Type Profiler's streaming discipline: one forward pass up
        to the last needed row, retaining only the wanted rows, so a large
        sheet never materializes (spec §8). Works identically for the
        formula-mode and data-mode read_only workbooks.

        Args:
            workbook: An open (read_only) openpyxl workbook.
            sheet_name: Worksheet to read.
            sample_rows: Sorted 1-based row numbers to retain.

        Returns:
            The retained rows (ascending sheet order), each a value list.
        """

        try:
            worksheet = workbook[sheet_name]
        except KeyError:  # pragma: no cover - defensive
            return []

        wanted = set(sample_rows)
        last_needed = max(sample_rows)
        retained: list[list[object]] = []
        for one_based, row in enumerate(
            worksheet.iter_rows(
                min_row=1, max_row=last_needed, values_only=True
            ),
            start=1,
        ):
            if one_based in wanted:
                retained.append(list(row))
            if one_based >= last_needed:
                break
        return retained
