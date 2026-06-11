"""Pipeline shared context (spec §6).

The :class:`InspectionContext` is threaded through every analyzer: each one
reads fields populated by earlier stages and fills in its own. For unit tests,
synthesize a *partial* context with only the fields under test (see
``tests/conftest.py``); all fields therefore have sensible defaults so the
context can be partially filled [spec §6].

The ``loader`` is injected later (Phase 1), so it is typed loosely here to keep
Phase 0 free of a hard dependency on ``loader.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from .models import InspectionOptions, WorkbookProfile

if TYPE_CHECKING:  # pragma: no cover - typing only
    from .analyzers.block_segmenter import RowBand
    from .analyzers.merge_analyzer import MergeSpan
    from .loader import Loader


@dataclass
class InspectionContext:
    """Shared, incrementally-populated pipeline context (spec §6).

    Attributes:
        options: Inspection overrides injected at the entry point [D2].
        loader: Mode-aware workbook loader (injected in Phase 1). ``Any`` until
            ``loader.py`` exists; ``None`` in partial test contexts.
        workbook_profile: The workbook profile being progressively filled.
        warnings: Accumulated low-confidence / undecidable notices (spec §6).
        row_bands: Per-sheet row bands computed by the Block Segmenter (plan
            v2 Phase 10a), keyed by sheet name; 1-based inclusive coordinates
            [D1]. Consumed by the Phase 10b per-block analysis.
        merge_spans: Per-sheet merged-cell ranges collected (structure mode,
            *unclassified*) by the Merge Scanner (plan v2 Task 11.1 Step 1),
            keyed by sheet name; 1-based bounds [D1], sorted for determinism.
            Consumed by the Boundary Detector (header-span virtual fill) and
            classified later by the Merge Analyzer. A sheet name *absent* from
            the dict means "never scanned" (the Merge Analyzer then collects
            on its own); an empty list means "scanned, no merges".
    """

    options: InspectionOptions = field(default_factory=InspectionOptions)
    loader: "Loader | Any | None" = None
    workbook_profile: WorkbookProfile = field(default_factory=WorkbookProfile)
    warnings: list[str] = field(default_factory=list)
    row_bands: dict[str, list["RowBand"]] = field(default_factory=dict)
    merge_spans: dict[str, list["MergeSpan"]] = field(default_factory=dict)

    def add_warning(self, message: str) -> None:
        """Append a warning message to the context (spec §6)."""

        self.warnings.append(message)
