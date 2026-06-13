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

#: Additive score cap for a wide time-series code header signal (issue #23).
#: This rescues leaf-code rows such as ``Period, Q:...`` that sit under several
#: all-string metadata rows. The final header score remains clamped to 1.0.
HEADER_TIMESERIES_CODE_BONUS: float = 0.25

#: Minimum share of non-axis labels in a candidate row that must look like
#: compact series/code tokens before the time-series bonus can apply.
HEADER_CODE_TOKEN_RATIO_THRESHOLD: float = 0.6

#: Minimum number of code-like labels required for the time-series bonus.
HEADER_TIMESERIES_MIN_CODE_LABELS: int = 3

#: Minimum observed date-like values in the candidate's axis column below it.
HEADER_TIMESERIES_MIN_DATE_AXIS_VALUES: int = 2

#: Minimum share of observed axis values below the candidate that must be
#: date-like for the time-series bonus.
HEADER_TIMESERIES_DATE_AXIS_RATIO_THRESHOLD: float = 0.8

#: Minimum share of populated non-axis cells below the candidate that must be
#: numeric-like for the time-series bonus. Empty cells are neutral.
HEADER_TIMESERIES_VALUE_RATIO_THRESHOLD: float = 0.8

#: Maximum length of a compact series/code token.
HEADER_CODE_TOKEN_MAX_LENGTH: int = 40

#: Default header confidence threshold; below this -> ``needs_manual_header``.
#: Overridable via ``InspectionOptions.header_confidence_threshold`` (spec §7.1).
HEADER_CONFIDENCE_THRESHOLD: float = 0.5

# ---------------------------------------------------------------------------
# §4.2 Tabular-candidate (non-tabular sheet) detection
# ---------------------------------------------------------------------------

#: Number of top rows sampled when judging whether a sheet is tabular (spec
#: §4.2). Matched to ``HEADER_SCAN_ROWS`` so the tabular gate and the header
#: scan look at the same top-of-sheet window.
NON_TABULAR_SAMPLE_ROWS: int = 20

#: A sheet whose content sample populates at most this many distinct columns is
#: non-tabular (a cover / description sheet), regardless of which column the
#: text starts in (issue #3). The original ``max_col``-only gate was sensitive
#: to the leftmost empty columns; counting *populated* columns is offset-free.
MIN_TABULAR_POPULATED_COLS: int = 1

#: When a sheet populates >= 2 columns but its sample cell density
#: (filled / (populated_cols * populated_rows)) is below this, it is still
#: treated as non-tabular (a multi-column but scattered cover). Calibrated
#: against the corpus: the lowest density among regression-pinned corpus tables
#: is 0.688 (stacked_uneven_width); a lower 0.648 occurs in the demo-only sheet
#: 지역별매출 (complex_demo.xlsx, not test-pinned), so the narrowest known margin
#: above this threshold is 0.148. The ``sparse_real_table`` fixture (density
#: 0.583) pins that margin in the test suite so the threshold cannot creep up.
NON_TABULAR_DENSITY_THRESHOLD: float = 0.5

#: Minimum sampled populated columns for the wide-sparse matrix escape hatch
#: (issue #22). Below this, the existing density rule remains authoritative so
#: scattered cover sheets like ``cover_sparse`` stay non-tabular.
WIDE_SPARSE_MIN_POPULATED_COLS: int = 8

#: A candidate header row in a wide sparse sheet must populate most sampled
#: columns. BIS ``Quarterly Series`` has dense metadata/header rows over
#: hundreds of columns, while the data rows are intentionally sparse.
WIDE_SPARSE_DENSE_ROW_RATIO: float = 0.8

#: Number of rows immediately below a dense candidate row examined for a time
#: axis. Kept small so title/metadata rows above the real header do not borrow
#: date evidence from far below.
WIDE_SPARSE_AXIS_LOOKAHEAD_ROWS: int = 3

#: Minimum immediate below rows whose first column looks like a date/period
#: value before a low-density wide sheet is preserved as tabular.
WIDE_SPARSE_MIN_AXIS_ROWS: int = 2

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
