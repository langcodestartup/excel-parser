"""v1 heuristic constants for the Excel Structure Inspector (spec §7) [D4].

These constants are fixed for v1. External configurability (beyond the few
fields exposed via :class:`~excel_inspector.models.InspectionOptions`) is
deferred to v1+. All values are calibrated against the fixture corpus.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# §7.1 Header detection scoring
# ---------------------------------------------------------------------------

#: Number of top rows sampled when scoring header candidates (spec §7.1).
HEADER_SCAN_ROWS: int = 20

#: Weight on the non-empty string ratio of the candidate row.
HEADER_WEIGHT_NON_EMPTY_STRING: float = 0.5

#: Weight on the type consistency of the 5 rows below the candidate.
HEADER_WEIGHT_TYPE_CONSISTENCY: float = 0.3

#: Weight on the distinctness of the candidate row vs the rows below.
HEADER_WEIGHT_DISTINCTNESS: float = 0.2

#: Number of rows below a header candidate examined for consistency/distinctness.
HEADER_LOOKAHEAD_ROWS: int = 5

#: Default header confidence threshold; below this -> ``needs_manual_header``.
#: Overridable via ``InspectionOptions.header_confidence_threshold`` (spec §7.1).
HEADER_CONFIDENCE_THRESHOLD: float = 0.5

# ---------------------------------------------------------------------------
# §7.2 Boundary detection rules
# ---------------------------------------------------------------------------

#: Number of consecutive blank rows (density == 0) marking a data end / block
#: separation (spec §7.2).
BLANK_RUN: int = 2

#: Rows with density below this are subtotal / separator-row candidates. The
#: complementary "only a single column filled" (``non_empty == 1``) rule applies
#: only to tables >= 3 columns wide (spec §7.2; see ``boundary_detector``).
LOW_DENSITY_THRESHOLD: float = 0.3

#: Default boundary keywords. A row is a ``skip_rows`` candidate when its
#: **leading (first non-empty) label cell** matches one of these: multi-char
#: keywords match by case-insensitive ``startswith``, and the single-char
#: ``"계"`` matches only on exact equality (so ``통계청``/``회계팀``/``Total
#: Wine`` do not false-match). Overridable/extendable via
#: ``InspectionOptions.skip_keywords`` (spec §7.2).
SKIP_KEYWORDS: list[str] = [
    "합계",
    "소계",
    "총계",
    "계",
    "Total",
    "Subtotal",
    "Grand Total",
]

# ---------------------------------------------------------------------------
# §7.3 Type inference
# ---------------------------------------------------------------------------

#: Sample size for type inference: ``min(TYPE_SAMPLE_ROWS, data_row_count)``,
#: evenly drawn from the data region (spec §7.3).
TYPE_SAMPLE_ROWS: int = 200

#: Minimum per-type parse success rate to commit to a single type; below this
#: across all candidate types -> ``mixed`` (spec §7.3).
TYPE_SUCCESS_THRESHOLD: float = 0.95
