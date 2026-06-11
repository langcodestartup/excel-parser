"""Override / options helpers (spec §5.0, §4.8) [D2].

These helpers provide a uniform way for analyzers and the aggregator to look
up per-sheet, per-field overrides from :class:`InspectionOptions`. The
override contract [D2]: when an option is present for a field an analyzer owns,
the analyzer skips its own computation, records the override value, and marks
``provenance="manual"`` with ``confidence=1.0`` where applicable.
"""

from __future__ import annotations

from .heuristics import HEADER_CONFIDENCE_THRESHOLD, SKIP_KEYWORDS
from .models import InspectionOptions, SheetOverride


def get_sheet_override(
    options: InspectionOptions | None, sheet_name: str
) -> SheetOverride | None:
    """Return the :class:`SheetOverride` for ``sheet_name`` if present.

    Args:
        options: The inspection options, or ``None``.
        sheet_name: Target sheet name.

    Returns:
        The matching :class:`SheetOverride`, or ``None`` when no options are
        supplied or no override exists for the sheet.
    """

    if options is None:
        return None
    return options.sheet_overrides.get(sheet_name)


def has_header_override(
    options: InspectionOptions | None, sheet_name: str
) -> bool:
    """Whether ``sheet_name`` has an explicit ``header_row`` override (HIGH #2).

    A header override of ``None`` is a *meaningful* declaration ("this sheet has
    no header"), distinct from "not overridden". The :class:`SheetOverride`
    sentinel (:data:`~excel_inspector.models._UNSET`) lets us tell the two
    apart: ``header_row_set`` is ``True`` iff the caller actually specified
    ``header_row`` (an int *or* an explicit ``None``).

    Crucially this is **not** simply "a :class:`SheetOverride` exists": a user
    who registers a :class:`SheetOverride` to set only ``dtype_force`` /
    ``is_tabular`` / ``skip_rows_*`` has *not* overridden the header, so the
    heuristic header locator must still run for that sheet (HIGH #2 regression).

    Args:
        options: The inspection options, or ``None``.
        sheet_name: Target sheet name.

    Returns:
        ``True`` only when a :class:`SheetOverride` exists for the sheet *and*
        its ``header_row`` was explicitly specified (its value — possibly
        ``None`` — is then authoritative).
    """

    override = get_sheet_override(options, sheet_name)
    return override is not None and override.header_row_set


def get_header_confidence_threshold(options: InspectionOptions | None) -> float:
    """Return the effective header-confidence threshold (spec §7.1)."""

    if options is None:
        return HEADER_CONFIDENCE_THRESHOLD
    return options.header_confidence_threshold


def get_skip_keywords(options: InspectionOptions | None) -> list[str]:
    """Return the effective skip-keyword list (spec §7.2).

    When ``options.skip_keywords`` is ``None`` the v1 default
    ``heuristics.SKIP_KEYWORDS`` is used; otherwise the provided list replaces
    the default.
    """

    if options is None or options.skip_keywords is None:
        return list(SKIP_KEYWORDS)
    return list(options.skip_keywords)


def get_dtype_force(
    options: InspectionOptions | None, sheet_name: str
) -> dict[str, str]:
    """Return forced dtypes for ``sheet_name`` keyed by 0-based position [D5]."""

    override = get_sheet_override(options, sheet_name)
    if override is None:
        return {}
    return dict(override.dtype_force)


def get_is_tabular_override(
    options: InspectionOptions | None, sheet_name: str
) -> bool | None:
    """Return the forced ``is_tabular`` flag, or ``None`` to defer."""

    override = get_sheet_override(options, sheet_name)
    if override is None:
        return None
    return override.is_tabular
