"""Block Analyzer — per-band table extraction loop (plan v2 §4, Phase 10b).

Phase 10a (Block Segmenter) split each sheet into row bands and made extra
table blocks *visible* via warnings. This analyzer is the Phase 10b half: every
band is analyzed like an independent mini-sheet — Header Locator → Boundary
Detector → Type Profiler, each scoped by ``row_window`` — and every band judged
to be a table becomes a :class:`~excel_inspector.models.TableBlock` on
``SheetProfile.blocks`` (top-down order). The Plan Aggregator then derives one
:class:`~excel_inspector.models.ReadPlan` per block, so a vertically stacked
multi-table sheet loses **no** table silently (spec §10/§11).

Judgments and conventions (plan v2 §4 Task 10.2):

* **Not a table** — a band with a *heuristic* (non-manual) header is judged
  not a table when its best §7.1 header score misses the threshold **or**
  the Boundary Detector resolves no data row below the candidate header
  (plan v2 §4: "헤더 신뢰도 임계 미달 + 데이터 미해소"; W-A review HIGH).
  Boundary detection always runs first, so the judgment — and its warning —
  reflects the *measured* data status, never just the score: a 1-row title
  band scoring exactly at the threshold (0.500) is still rejected because no
  data resolves beneath it. The band enters no block — a warning records the
  skip instead (guard 7: ``blocks`` holds only judged tables, so the flat
  mirror is always the top-most *table*). A *manual* header override is
  authoritative and is never rejected here; its read plan is band-clamped by
  the aggregator instead (defense line).
* **Guard 1 (denominator)** — band-scoped header scoring uses the band-local
  used-column count, never the sheet-global ``max_col``, so a narrow table
  stacked with a wide one is not diluted into a "not a table" misjudgment.
* **Guard 3 (state isolation)** — all per-band boundary state lives on the
  block-local :class:`~excel_inspector.analyzers.boundary_detector.
  BlockBoundary` result; nothing leaks across blocks through the shared
  profile.
* **Guard 4 (override semantics)** — ``SheetOverride.header_row`` and
  ``skip_rows_add`` are **absolute 1-based sheet coordinates** and apply only
  to the block whose band contains the row. An explicit *headerless* override
  (``header_row=None``) is a sheet-wide declaration: per-band analysis is
  skipped entirely and the v1 headerless flat path is preserved.
* **Guard 6 (warning order)** — warnings accumulate sheet order first, then
  band top-down within a sheet (deterministic JSON ``warnings``).
* **Mirror rule** — when blocks exist, the sheet's flat header/boundary/column
  fields are overwritten with ``blocks[0]`` (the top-most table), restoring
  the spec §10 "top-most block" intent (v1 picked the best-*scoring* block).

Single-band sheets keep their v1 whole-sheet analysis bit-identical: the flat
fields (already computed by the sheet-level analyzers) are *copied into* a
mirror block rather than re-derived through a row window, so the existing v1
golden corpus cannot drift. A single-band sheet with no header and no data
(headerless / needs-manual) contributes no block — the v1 fallback plan path
stays in charge ("블록 없으면 기존 동작 유지").

Pipeline placement: after the Type Profiler (the sheet-level flat fields are
final) and before the Merge Analyzer / Plan Aggregator.
"""

from __future__ import annotations

from ..context import InspectionContext
from ..models import SheetOverride, SheetProfile, TableBlock
from ..options import (
    get_header_confidence_threshold,
    get_sheet_override,
    has_header_override,
)
from ..pipeline import Analyzer
from .block_segmenter import RowBand, _suspected_block_warning
from .boundary_detector import BoundaryDetector
from .header_locator import HeaderLocator
from .type_profiler import TypeProfiler


class BlockAnalyzer(Analyzer):
    """Per-band Header→Boundary→Type loop producing ``SheetProfile.blocks``."""

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "block_analyzer"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Build :class:`TableBlock` lists for every multi-band tabular sheet.

        Sheets are processed in workbook order and bands top-down (guard 6:
        deterministic warning accumulation). Single-band sheets get a mirror
        block copied from their (v1) flat fields; multi-band sheets are
        re-analyzed per band through the row-window cores.

        Args:
            context: Shared context carrying a ready loader, the enumerated
                profiles (flat fields final), and ``row_bands`` from the
                Block Segmenter.

        Returns:
            The same context with ``blocks`` populated (and, for multi-band
            sheets, the flat fields mirrored from ``blocks[0]``).
        """

        loader = context.loader
        for profile in context.workbook_profile.sheets:
            if not profile.is_tabular_candidate:
                continue
            bands = context.row_bands.get(profile.name) or []
            if not bands:
                continue

            # Guard 4: an explicit headerless declaration (header_row=None
            # override) is sheet-wide — there is no per-block header to
            # anchor on, so the v1 headerless flat path stays authoritative
            # and no block is produced.
            override = get_sheet_override(context.options, profile.name)
            if (
                override is not None
                and override.header_row_set
                and override.header_row is None
            ):
                continue

            if len(bands) == 1:
                block = self._mirror_block(profile, bands[0])
                if block is not None:
                    profile.blocks = [block]
                continue

            if loader is None:
                context.add_warning(
                    f"block_analyzer: no loader available; cannot analyze "
                    f"table blocks for sheet {profile.name!r}"
                )
                continue

            # Guard 4 / W-A review MEDIUM #5: an absolute header_row override
            # that points at no detected band (e.g. a blank separator row)
            # can anchor no block — say so instead of dropping it silently.
            if (
                override is not None
                and override.header_row_set
                and isinstance(override.header_row, int)
                and not any(
                    b.start_row <= override.header_row <= b.end_row
                    for b in bands
                )
            ):
                context.add_warning(
                    f"block_analyzer: sheet {profile.name!r}: header_row "
                    f"override {override.header_row} falls inside no detected "
                    f"table band; ignored"
                )

            blocks: list[TableBlock] = []
            for band in bands:
                block = self._analyze_band(
                    context, profile, band, block_index=len(blocks)
                )
                if block is not None:
                    blocks.append(block)
            profile.blocks = blocks
            if blocks:
                self._mirror_to_profile(profile, blocks[0])
                self._mark_extracted_bands(context, profile, blocks)
        return context

    @staticmethod
    def _mirror_block(
        profile: SheetProfile, band: RowBand
    ) -> TableBlock | None:
        """Build the single-band mirror block from the (v1) flat fields.

        The flat fields were produced by the unwindowed sheet-level analyzers,
        so copying them — instead of re-deriving through a degenerate
        whole-sheet window — keeps the v1 behavior bit-identical for the
        single-table corpus. A sheet with neither a header nor a resolved data
        region (headerless / needs-manual / empty band) is not a table block;
        returning ``None`` leaves ``blocks`` empty so the v1 fallback read
        plan path stays in charge.

        Args:
            profile: The sheet whose flat fields are mirrored.
            band: The sheet's only row band.

        Returns:
            The mirror :class:`TableBlock`, or ``None`` (no table block).
        """

        if profile.header_row is None and profile.data_start_row is None:
            return None
        return TableBlock(
            block_index=0,
            band_start_row=band.start_row,
            band_end_row=band.end_row,
            header_row=profile.header_row,
            header_confidence=profile.header_confidence,
            header_provenance=profile.header_provenance,
            data_start_row=profile.data_start_row,
            data_end_row=profile.data_end_row,
            data_left_col=profile.data_left_col,
            data_right_col=profile.data_right_col,
            skip_rows=list(profile.skip_rows),
            columns=list(profile.columns),
            read_plan=None,
            subtotal_skip_labels=dict(profile.subtotal_skip_labels),
        )

    def _analyze_band(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        band: RowBand,
        block_index: int,
    ) -> TableBlock | None:
        """Run Header→Boundary→Type for one band and judge it (plan v2 §4).

        The §7.1 header score is computed band-locally (guard 1/2 via the
        ``row_window`` cores), then the Boundary Detector always runs so the
        judgment measures the band's data status instead of assuming it
        (W-A review HIGH). A band with a *heuristic* header is judged **not a
        table** (title/footnote band) when its best score misses the threshold
        or no data row resolves below the candidate header — the recorded
        warning states exactly the reasons that actually held (guard 7 /
        review LOW #6). A [D2] ``header_row`` override whose absolute row
        falls inside this band short-circuits the scoring (guard 4) and is
        authoritative: a manual block is kept even with unresolved data (its
        plan is band-clamped by the aggregator).

        Args:
            context: Shared context with a ready data-mode loader.
            profile: The sheet being analyzed.
            band: The band to analyze (1-based inclusive [D1]).
            block_index: The 0-based index this block will take in ``blocks``.

        Returns:
            The band's :class:`TableBlock`, or ``None`` when the band is not
            a table.
        """

        window = (band.start_row, band.end_row)
        threshold = get_header_confidence_threshold(context.options)
        override = get_sheet_override(context.options, profile.name)

        header_row: int | None
        below_threshold = False
        if (
            has_header_override(context.options, profile.name)
            and override is not None
            and isinstance(override.header_row, int)
            and band.start_row <= override.header_row <= band.end_row
        ):
            # Guard 4: the absolute override row lives in this band -> manual.
            header_row = override.header_row
            confidence = 1.0
            provenance = "manual"
        else:
            header_row, score = HeaderLocator()._locate(  # noqa: SLF001
                context, profile, row_window=window
            )
            if header_row is None:
                context.add_warning(
                    f"block_analyzer: sheet {profile.name!r} rows "
                    f"{band.start_row}-{band.end_row}: no header candidate "
                    f"row found; band judged not a table (skipped)"
                )
                return None
            confidence = score
            provenance = "heuristic"
            below_threshold = score < threshold

        boundary = BoundaryDetector()._detect_block(  # noqa: SLF001
            context, profile, header_row, row_window=window
        )
        # Guard 6: the band's boundary warnings are forwarded immediately, in
        # detection order, so the accumulated list is sheet -> band top-down.
        for warning in boundary.warnings:
            context.add_warning(warning)

        # Not-a-table judgment (plan v2 §4 / W-A review HIGH): a heuristic
        # header below the score threshold OR with no resolved data row marks
        # a title/footnote band, not a table. The warning lists only the
        # reasons that actually held (review LOW #6), so it is always
        # truthful — and a manual override is never second-guessed here.
        if provenance != "manual":
            reasons: list[str] = []
            if below_threshold:
                reasons.append(
                    f"best header score {max(confidence, 0.0):.3f} below "
                    f"threshold {threshold:.3f}"
                )
            if boundary.data_start_row is None:
                reasons.append(
                    f"no data rows resolved below the candidate header "
                    f"(row {header_row})"
                )
            if reasons:
                context.add_warning(
                    f"block_analyzer: sheet {profile.name!r} rows "
                    f"{band.start_row}-{band.end_row}: {' and '.join(reasons)}"
                    f"; band judged not a table (skipped)"
                )
                return None

        skip_rows = self._fold_skip_overrides(
            override, band, boundary.skip_rows
        )

        columns = []
        if (
            boundary.data_start_row is not None
            and boundary.data_end_row is not None
        ):
            profiled = TypeProfiler()._profile_block(  # noqa: SLF001
                context,
                profile,
                header_row=header_row,
                data_start_row=boundary.data_start_row,
                data_end_row=boundary.data_end_row,
                skip_rows=skip_rows,
                data_left_col=boundary.data_left_col,
                data_right_col=boundary.data_right_col,
                row_window=window,
            )
            columns = profiled if profiled is not None else []

        return TableBlock(
            block_index=block_index,
            band_start_row=band.start_row,
            band_end_row=band.end_row,
            header_row=header_row,
            header_confidence=confidence,
            header_provenance=provenance,
            data_start_row=boundary.data_start_row,
            data_end_row=boundary.data_end_row,
            data_left_col=boundary.data_left_col,
            data_right_col=boundary.data_right_col,
            skip_rows=skip_rows,
            columns=columns,
            read_plan=None,
            subtotal_skip_labels=dict(boundary.subtotal_skip_labels),
        )

    @staticmethod
    def _fold_skip_overrides(
        override: SheetOverride | None,
        band: RowBand,
        skip_rows: list[int],
    ) -> list[int]:
        """Fold [D2] skip overrides into one block's heuristic skips (guard 4).

        ``skip_rows_add`` rows are absolute 1-based sheet coordinates and are
        folded only when they fall inside this block's band; ``skip_rows_remove``
        is a plain set-difference (removing a row not present is a no-op, so no
        band filter is needed).

        Args:
            override: The sheet's :class:`SheetOverride`, or ``None``.
            band: The block's enclosing band.
            skip_rows: The block's heuristic skip rows (1-based).

        Returns:
            The sorted, de-duplicated folded skip rows.
        """

        result = set(skip_rows)
        if override is not None:
            result.update(
                row
                for row in override.skip_rows_add
                if band.start_row <= row <= band.end_row
            )
            result.difference_update(override.skip_rows_remove)
        return sorted(result)

    @staticmethod
    def _mark_extracted_bands(
        context: InspectionContext,
        profile: SheetProfile,
        blocks: list[TableBlock],
    ) -> None:
        """Rewrite "suspected" notices into "extracted" ones (review LOW #8).

        The Block Segmenter (Phase 10a) records ``additional table block
        suspected at rows S-E`` for every band beyond the first that carries a
        header candidate. Once Phase 10b has actually extracted a band, the
        suspicion is stale: the notice is rewritten **in place** (preserving
        the guard-6 deterministic warning order) to name the resulting table
        id. A band extracted *without* a prior suspicion (e.g. anchored only
        by a manual override) appends the notice instead, so every extracted
        non-top band is recorded. Bands judged not a table keep their
        "suspected" notice alongside the not-a-table warning.
        """

        for block in blocks:
            band = RowBand(block.band_start_row, block.band_end_row)
            suspected = _suspected_block_warning(profile.name, band)
            extracted = (
                f"sheet '{profile.name}': additional table block extracted "
                f"as '{profile.name}!T{block.block_index + 1}' "
                f"(rows {block.band_start_row}-{block.band_end_row})"
            )
            for position, warning in enumerate(context.warnings):
                if warning == suspected:
                    context.warnings[position] = extracted
                    break
            else:
                if block.block_index > 0:
                    context.add_warning(extracted)

    @staticmethod
    def _mirror_to_profile(profile: SheetProfile, block: TableBlock) -> None:
        """Overwrite the flat fields with the top-most block (mirror rule).

        Plan v2 §4.0: the flat fields mirror ``blocks[0]``. Only table-judged
        blocks exist in ``blocks``, so the mirror is always the top-most
        *table* (guard 7) — this intentionally replaces v1's "best-scoring
        block" selection with the spec §10 "top-most block" intent.
        """

        profile.header_row = block.header_row
        profile.header_confidence = block.header_confidence
        profile.header_provenance = block.header_provenance
        profile.needs_manual_header = False
        profile.data_start_row = block.data_start_row
        profile.data_end_row = block.data_end_row
        profile.data_left_col = block.data_left_col
        profile.data_right_col = block.data_right_col
        profile.skip_rows = list(block.skip_rows)
        profile.columns = list(block.columns)
        profile.subtotal_skip_labels = dict(block.subtotal_skip_labels)
