"""SheetOverride header-channel sentinel tests (spec §5.0, §6) [D2] / HIGH #2.

The override channel must distinguish three header states:

* ``header_row`` unspecified (sentinel)  -> defer to the heuristic locator.
* ``header_row=<int>``                    -> force the header to that row.
* ``header_row=None``                     -> declare the sheet has no header.

Before the sentinel, *any* registered :class:`SheetOverride` defaulted
``header_row`` to ``None`` and so read as a headerless declaration, which
destroyed heuristic header detection for users who only set ``dtype_force`` /
``is_tabular`` / ``skip_rows_*``. These tests pin the corrected semantics.
"""

from __future__ import annotations

from excel_inspector import InspectionOptions, SheetOverride
from excel_inspector.options import has_header_override


def test_default_override_does_not_set_header() -> None:
    """A bare SheetOverride leaves the header channel unset (deferred)."""

    override = SheetOverride()
    assert override.header_row_set is False
    # Collapsed to None for downstream type simplicity, but not "specified".
    assert override.header_row is None


def test_dtype_only_override_does_not_set_header() -> None:
    """Setting only dtype_force does not specify the header (HIGH #2)."""

    override = SheetOverride(dtype_force={"1": "string"})
    assert override.header_row_set is False


def test_is_tabular_only_override_does_not_set_header() -> None:
    """Setting only is_tabular does not specify the header (HIGH #2)."""

    override = SheetOverride(is_tabular=True)
    assert override.header_row_set is False


def test_skip_rows_only_override_does_not_set_header() -> None:
    """Setting only skip_rows_* does not specify the header (HIGH #2)."""

    override = SheetOverride(skip_rows_add=[5], skip_rows_remove=[8])
    assert override.header_row_set is False


def test_explicit_int_header_is_set() -> None:
    """An int header_row is recorded as specified."""

    override = SheetOverride(header_row=4)
    assert override.header_row_set is True
    assert override.header_row == 4


def test_explicit_none_header_is_set() -> None:
    """An explicit header_row=None is a meaningful headerless declaration."""

    override = SheetOverride(header_row=None)
    assert override.header_row_set is True
    assert override.header_row is None


def test_has_header_override_false_for_dtype_only() -> None:
    """has_header_override is False when only dtype_force is set (HIGH #2)."""

    options = InspectionOptions(
        sheet_overrides={"S": SheetOverride(dtype_force={"0": "string"})}
    )
    assert has_header_override(options, "S") is False


def test_has_header_override_true_for_explicit_int() -> None:
    """has_header_override is True for an explicit int header_row."""

    options = InspectionOptions(
        sheet_overrides={"S": SheetOverride(header_row=2)}
    )
    assert has_header_override(options, "S") is True


def test_has_header_override_true_for_explicit_none() -> None:
    """has_header_override is True for an explicit headerless declaration."""

    options = InspectionOptions(
        sheet_overrides={"S": SheetOverride(header_row=None)}
    )
    assert has_header_override(options, "S") is True


def test_has_header_override_false_when_no_override() -> None:
    """has_header_override is False when the sheet has no override at all."""

    assert has_header_override(InspectionOptions(), "S") is False
    assert has_header_override(None, "S") is False
