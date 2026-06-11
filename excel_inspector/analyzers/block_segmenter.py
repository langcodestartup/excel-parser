"""Block Segmenter — row-band splitting + multi-block warning (plan v2 §4, Phase 10a).

A real-world sheet can stack several independent tables vertically. v1 loads
only one block per sheet and silently drops the rest (spec §10). Phase 10a is
the *safety net* half of the fix: the sheet's rows are split into **row bands**
separated by ``BLANK_RUN`` (2) or more consecutive blank rows, and when a band
beyond the first contains a header-candidate row (§7.1 score at or above the
header-confidence threshold) an explicit warning is recorded — the silent loss
becomes visible. Full per-band extraction is Phase 10b (Task 10.2).

Coordinate contract [D1]: :class:`RowBand` rows are **openpyxl 1-based**
inclusive, like every other inspection-domain coordinate. The input ``rows``
of :func:`split_row_bands` are the same data-mode row tuples the Boundary
Detector consumes, aligned so that ``rows[r - 1]`` is sheet row ``r``.

Pipeline placement: right after the Sheet Enumerator (plan v2 §4 Task 10.1
Step 3). The computed bands are stored on ``context.row_bands`` keyed by sheet
name so Phase 10b's per-block analysis can reuse them.
"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass

from ..context import InspectionContext
from ..heuristics import BLANK_RUN, HEADER_SCAN_ROWS
from ..models import SheetProfile
from ..options import get_header_confidence_threshold
from ..pipeline import Analyzer
from .header_locator import _CAT_EMPTY, _categorize, _score_row


@dataclass(frozen=True)
class RowBand:
    """One contiguous row band inside a sheet (1-based inspection coordinates).

    A band starts and ends on a *non-blank* row; interior blank rows shorter
    than the splitting run stay inside the band. Bands never overlap and are
    ordered top-down.

    Attributes:
        start_row: First non-blank row of the band (1-based, inclusive).
        end_row: Last non-blank row of the band (1-based, inclusive).
    """

    start_row: int
    end_row: int


def _is_blank_row(row: Sequence[object]) -> bool:
    """Whether every cell of a sampled row is empty (``None`` / ``""``).

    Mirrors the Boundary Detector's emptiness rule so band splitting and the
    §7.2 blank-run terminator agree on what "blank" means.
    """

    return all(
        value is None or (isinstance(value, str) and value == "")
        for value in row
    )


def split_row_bands(
    rows: Iterable[Sequence[object]], blank_run: int = BLANK_RUN
) -> list[RowBand]:
    """Split data-mode rows into bands separated by blank-row runs (§7.2).

    A run of ``blank_run`` or more consecutive fully-blank rows separates two
    bands; a shorter blank run stays *inside* its band (consistent with the v1
    boundary rule where a single blank row is an interior skip, not a
    terminator). Leading/trailing blank rows belong to no band, so every band
    starts and ends on a non-blank row.

    Args:
        rows: Row value tuples in sheet order, aligned 1-based
            (``rows[r - 1]`` is sheet row ``r``) — the same data-mode tuples
            the Boundary Detector consumes. Any iterable is accepted, so a
            streaming ``iter_rows(values_only=True)`` generator works without
            materializing the sheet.
        blank_run: Number of consecutive blank rows that splits bands
            (default :data:`~excel_inspector.heuristics.BLANK_RUN`).

    Returns:
        The top-down list of :class:`RowBand` (possibly empty for an all-blank
        or empty sheet).
    """

    if blank_run < 1:
        raise ValueError(f"blank_run must be >= 1, got {blank_run}")

    bands: list[RowBand] = []
    band_start: int | None = None
    band_end: int | None = None
    run = 0

    for one_based, row in enumerate(rows, start=1):
        if _is_blank_row(row):
            run += 1
            if band_start is not None and run >= blank_run:
                # The blank run reached the splitting threshold: close the
                # open band at its last non-blank row.
                assert band_end is not None
                bands.append(RowBand(band_start, band_end))
                band_start = None
                band_end = None
            continue
        run = 0
        if band_start is None:
            band_start = one_based
        band_end = one_based

    if band_start is not None:
        assert band_end is not None
        bands.append(RowBand(band_start, band_end))
    return bands


def _band_col_count(rows: list[list[object]]) -> int:
    """Band-local used-column count (denominator for header scoring).

    Task 10.2 latent-bug guard #1 (plan v2 §4): scoring a band against the
    sheet-global ``max_col`` dilutes a narrow table stacked next to a wide one,
    so the denominator is the widest non-empty extent *within the band*.
    """

    widest = 0
    for row in rows:
        last_non_empty = 0
        for position, value in enumerate(row, start=1):
            if _categorize(value) != _CAT_EMPTY:
                last_non_empty = position
        widest = max(widest, last_non_empty)
    return widest


def _has_header_candidate(rows: list[list[object]], threshold: float) -> bool:
    """Whether any sampled band row scores as a §7.1 header candidate."""

    col_count = _band_col_count(rows)
    if not rows or col_count <= 0:
        return False
    return any(
        _score_row(index, rows, col_count) >= threshold
        for index in range(len(rows))
    )


def _suspected_block_warning(sheet_name: str, band: RowBand) -> str:
    """Render the Phase 10a "additional block suspected" warning for one band.

    Single source of the message text: the Block Analyzer (Phase 10b) rewrites
    this exact string into an "extracted as '<sheet>!T<n>'" notice once the
    band has actually been extracted, so the wording must match between the
    two analyzers (plan v2 §4 Task 10.1 Step 3 / W-A review LOW #8).
    """

    return (
        f"sheet '{sheet_name}': additional table block "
        f"suspected at rows {band.start_row}-{band.end_row}"
    )


class BlockSegmenter(Analyzer):
    """Detect multiple table blocks per sheet and warn (plan v2 Phase 10a).

    For every tabular sheet the rows are streamed once to compute the
    :class:`RowBand` list (stored on ``context.row_bands`` for Phase 10b).
    When two or more bands exist and a band beyond the first contains a
    header-candidate row, a warning is recorded so the v1 single-block
    extraction no longer drops table blocks *silently* (spec §10 / §11).
    """

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "block_segmenter"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Compute row bands per tabular sheet and warn on extra table blocks.

        Args:
            context: Shared context carrying a ready :class:`Loader` and the
                enumerated sheet profiles.

        Returns:
            The same context with ``row_bands`` populated and any multi-block
            warnings accumulated (sheet order, then band top-down — the
            warning order is deterministic).
        """

        loader = context.loader
        threshold = get_header_confidence_threshold(context.options)
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            if loader is None:
                context.add_warning(
                    f"block_segmenter: no loader available; cannot segment "
                    f"row bands for sheet {profile.name!r}"
                )
                continue

            bands = split_row_bands(self._iter_sheet_rows(context, profile))
            context.row_bands[profile.name] = bands
            if len(bands) < 2:
                continue

            for band in bands[1:]:
                sample = self._band_top_rows(context, profile, band)
                if _has_header_candidate(sample, threshold):
                    context.add_warning(
                        _suspected_block_warning(profile.name, band)
                    )
        return context

    @staticmethod
    def _iter_sheet_rows(
        context: InspectionContext, profile: SheetProfile
    ) -> Iterable[tuple[object, ...]]:
        """Stream every row of a sheet in data mode (read_only) [D3].

        A single forward pass over rows ``1 .. max_row`` (natural EOF when the
        dimensions are untrusted); rows are yielded straight into
        :func:`split_row_bands`, so the sheet is never materialized (spec §8).
        """

        workbook = context.loader.data_workbook()
        try:
            worksheet = workbook[profile.name]
        except KeyError:  # pragma: no cover - defensive
            return iter(())
        max_row = (
            profile.max_row if profile.max_row and profile.max_row > 0 else None
        )
        return worksheet.iter_rows(
            min_row=1, max_row=max_row, values_only=True
        )

    @staticmethod
    def _band_top_rows(
        context: InspectionContext, profile: SheetProfile, band: RowBand
    ) -> list[list[object]]:
        """Read the top ``HEADER_SCAN_ROWS`` rows of one band in data mode.

        Only the band's top sample is materialized (bounded by
        ``HEADER_SCAN_ROWS`` per band), mirroring the Header Locator's
        top-of-sheet sampling but scoped to the band (plan v2 §4: each band is
        treated like an independent mini-sheet).
        """

        workbook = context.loader.data_workbook()
        try:
            worksheet = workbook[profile.name]
        except KeyError:  # pragma: no cover - defensive
            return []
        top_end = min(band.start_row + HEADER_SCAN_ROWS - 1, band.end_row)
        return [
            list(row)
            for row in worksheet.iter_rows(
                min_row=band.start_row, max_row=top_end, values_only=True
            )
        ]
