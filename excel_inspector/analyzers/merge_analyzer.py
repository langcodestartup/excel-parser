"""Merge Scanner + Merge Analyzer (spec §4.4, plan v2 Task 11.1) [D3].

Phase 11a splits the v1 single-step merge analysis into two stages
(plan v2 §5 Task 11.1 Step 1):

* **Collection** (:class:`MergeScanner`) — runs once per sheet **before** the
  Boundary Detector, reading the **structure-mode** workbook only [D3]
  (read_only worksheets have no ``merged_cells``). It records every merged
  range's *structural bounds* (:class:`MergeSpan`, 1-based [D1]) on
  ``context.merge_spans`` without classifying anything — the header is not
  resolved yet. The spans are sorted by position because openpyxl does not
  guarantee a stable ``merged_cells.ranges`` order; sorting makes every
  downstream note/warning deterministic.
* **Classification** (:class:`MergeAnalyzer`) — keeps its original pipeline
  slot *after* the header/boundary analyzers, so ``header_row`` is final when
  each collected span is classified as ``header`` or ``body`` and the
  multi-level flag is judged. When the scanner did not run (standalone /
  partial-context use), the analyzer falls back to collecting the ranges
  itself, preserving the v1 single-step behavior.

The collected (unclassified) spans are what let the Boundary Detector bridge a
merged header: a merge intersecting the header row marks its empty covered
cells as *virtually filled*, restoring the contiguous header column span that
the merge had collapsed (plan v2 Task 11.1 Step 2; spec §7.2 deferral
resolved).

Classification rules (spec §4.4):

* **header** — the merge overlaps the header row *or* sits entirely above it
  (``min_row <= header_row``). A merge above the header is part of the title /
  grouping band and is treated as header structure.
* **body** — the merge sits entirely below the header row (``min_row >
  header_row``). These are in-body merges (e.g. a group label spanning several
  data rows).

When ``header_row`` is unknown (estimation failed / declared headerless) there
is no anchor to classify against; every merge is conservatively recorded as
``body`` (the safe default — a body merge only adds a forward-fill *note*, it
never reclassifies header structure).

``is_multi_level_header`` (spec §4.4, §5.2): ``True`` when at least one header
merge spans a row **strictly above** the resolved ``header_row`` — i.e. an extra
merged header band sits over the leaf header (the two-level
``상반기``/``하반기`` over ``1월``..``4월`` case). v1 only *judges* this flag; it
does **not** branch the load path (multi-level header loading is deferred to
v1+ [D6]). When ``header_row`` is unknown the flag is left ``False``.

Body-merge forward-fill recommendation (spec §4.4): the actual fill is the
loader's job, so v1 only *records* the recommendation. Each ``body`` merge
becomes a note on the eventual :class:`~excel_inspector.models.ReadPlan` — but
because the read plan does not exist until the Plan Aggregator runs *after* this
analyzer, the recommendation is carried on the sheet's ``merges`` list and the
aggregator translates each ``body`` region into a ``ReadPlan.notes`` entry (see
``aggregator._body_merge_notes``). Per-block plans attribute each merge to the
block whose band rows it intersects (plan v2 Task 11.1 Step 1, block
attribution — see ``aggregator.build_block_read_plan``).

Override note: there is no merge-specific override channel in
:class:`~excel_inspector.models.InspectionOptions` (spec §5.0), so neither
stage has an override branch — merges are always derived from the structure
workbook.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..context import InspectionContext
from ..models import MergeRegion, SheetProfile
from ..pipeline import Analyzer

if TYPE_CHECKING:  # pragma: no cover - typing only
    from openpyxl.worksheet.worksheet import Worksheet

#: The two merge-region classifications (spec §5.4).
_KIND_HEADER = "header"
_KIND_BODY = "body"


@dataclass(frozen=True)
class MergeSpan:
    """One *collected* (unclassified) merged range (plan v2 Task 11.1 Step 1).

    Structural bounds only — classification into ``header``/``body`` happens
    later, once the header row is final (:class:`MergeAnalyzer`). All bounds
    are **openpyxl 1-based**, inclusive (inspection domain [D1]).

    Attributes:
        range: The merge range in A1 notation (e.g. ``"A1:B1"``).
        min_row: Top row of the merge (1-based).
        min_col: Left column of the merge (1-based).
        max_row: Bottom row of the merge (1-based, inclusive).
        max_col: Right column of the merge (1-based, inclusive).
    """

    range: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int


def _collect_spans(worksheet: "Worksheet") -> list[MergeSpan]:
    """Collect a sheet's merged ranges as sorted :class:`MergeSpan` items.

    openpyxl guarantees no stable ``merged_cells.ranges`` order, so the spans
    are sorted by position (top-down, left-right) — every downstream
    note/warning derived from them is then deterministic (plan v2 Task 11.1).

    Args:
        worksheet: A **structure-mode** worksheet (read_only worksheets lack
            ``merged_cells`` [D3]).

    Returns:
        The sorted span list (possibly empty).
    """

    spans = [
        MergeSpan(
            range=str(cell_range),
            min_row=cell_range.min_row,
            min_col=cell_range.min_col,
            max_row=cell_range.max_row,
            max_col=cell_range.max_col,
        )
        for cell_range in worksheet.merged_cells.ranges
    ]
    spans.sort(key=lambda s: (s.min_row, s.min_col, s.max_row, s.max_col))
    return spans


def _classify_kind(merge_min_row: int, header_row: int | None) -> str:
    """Classify one merge region as ``header`` or ``body`` (spec §4.4).

    Args:
        merge_min_row: The merge region's top (smallest) 1-based row.
        header_row: The resolved header row (1-based), or ``None`` when the
            header is unknown.

    Returns:
        ``"header"`` when the merge overlaps or sits above the header row
        (``min_row <= header_row``); otherwise ``"body"``. When ``header_row``
        is ``None`` the conservative ``"body"`` default is returned (a body
        merge only adds a forward-fill note; it never alters header structure).
    """

    if header_row is None:
        return _KIND_BODY
    return _KIND_HEADER if merge_min_row <= header_row else _KIND_BODY


class MergeScanner(Analyzer):
    """Collect merged ranges per sheet, classification-free (Task 11.1) [D3].

    Runs right after the Sheet Enumerator — once per sheet, **before** the
    Boundary Detector — so the collected spans can bridge a merged header's
    collapsed column span (virtual fill) during boundary detection. The spans
    land on ``context.merge_spans`` keyed by sheet name; classification waits
    for the resolved header (:class:`MergeAnalyzer`).
    """

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "merge_scanner"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Populate ``context.merge_spans`` for every tabular sheet.

        Reads the **structure-mode** workbook only [D3]. Non-tabular sheets
        are skipped (they are excluded from loading, spec §9).

        Args:
            context: Shared context carrying a ready :class:`Loader` and the
                enumerated sheet profiles.

        Returns:
            The same context with ``merge_spans`` populated (an empty list for
            a scanned sheet without merges).
        """

        loader = context.loader
        if loader is None:
            context.add_warning(
                "merge_scanner: no loader available; skipping merge collection"
            )
            return context

        workbook = loader.structure_workbook()
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            try:
                worksheet = workbook[profile.name]
            except KeyError:  # pragma: no cover - defensive
                continue
            context.merge_spans[profile.name] = _collect_spans(worksheet)
        return context


class MergeAnalyzer(Analyzer):
    """Classify merge regions and flag multi-level headers (spec §4.4) [D3].

    Phase 11a (plan v2 Task 11.1 Step 1): classification consumes the spans
    already collected by :class:`MergeScanner` (``context.merge_spans``) — the
    structure workbook is *not* re-read on the pipeline path. A sheet that was
    never scanned (standalone / partial-context use without the scanner) falls
    back to the v1 single-step behavior of collecting the ranges itself.
    """

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "merge_analyzer"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Populate ``merges`` and ``is_multi_level_header`` per tabular sheet.

        Classification anchors on the (now final) ``header_row``. Non-tabular
        sheets are skipped (they are excluded from loading, spec §9).

        Args:
            context: Shared context carrying the collected ``merge_spans``
                (and a ready :class:`Loader` for the no-scanner fallback) and
                the header-resolved sheet profiles.

        Returns:
            The same context with ``merges`` (classified) and
            ``is_multi_level_header`` set on every tabular sheet.
        """

        fallback_workbook = None
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            if profile.name in context.merge_spans:
                spans = context.merge_spans[profile.name]
            else:
                # Fallback (sheet never scanned): collect now, structure mode
                # only [D3] — read_only worksheets lack merged_cells.
                if fallback_workbook is None:
                    loader = context.loader
                    if loader is None:  # pragma: no cover - wiring guard
                        context.add_warning(
                            "merge_analyzer: no loader available; skipping "
                            "merge analysis"
                        )
                        return context
                    fallback_workbook = loader.structure_workbook()
                try:
                    worksheet = fallback_workbook[profile.name]
                except KeyError:  # pragma: no cover - defensive
                    continue
                spans = _collect_spans(worksheet)
            self._classify_sheet(spans, profile)
        return context

    @staticmethod
    def _classify_sheet(
        spans: list[MergeSpan], profile: SheetProfile
    ) -> None:
        """Classify one sheet's collected spans and record them (spec §4.4)."""

        header_row = profile.header_row

        merges: list[MergeRegion] = []
        has_band_above_header = False
        for span in spans:
            kind = _classify_kind(span.min_row, header_row)
            merges.append(MergeRegion(range=span.range, kind=kind))
            if (
                kind == _KIND_HEADER
                and header_row is not None
                and span.min_row < header_row
            ):
                # A header merge spanning a row strictly above the resolved
                # header is an extra header band -> multi-level header.
                has_band_above_header = True

        profile.merges = merges
        # v1 only judges the flag; the load path is not branched [D6].
        profile.is_multi_level_header = has_band_above_header
