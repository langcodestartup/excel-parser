"""Analyzer strategies for the inspection pipeline (spec §4).

Individual analyzers (sheet enumerator, header locator, boundary detector,
type profiler, merge analyzer, and the v1+ formula detector stub) are added in
later phases. This package is established in Phase 0 to fix the module layout
(implementation plan §2). Phase 1 adds the sheet enumerator; Phase 2 adds the
header locator; Phase 3 adds the boundary detector; Phase 5 adds the type
profiler; Phase 6 adds the merge analyzer.
"""

from __future__ import annotations

from .boundary_detector import BoundaryDetector
from .header_locator import HeaderLocator
from .merge_analyzer import MergeAnalyzer
from .sheet_enumerator import SheetEnumerator
from .type_profiler import TypeProfiler

__all__ = [
    "BoundaryDetector",
    "HeaderLocator",
    "MergeAnalyzer",
    "SheetEnumerator",
    "TypeProfiler",
]
