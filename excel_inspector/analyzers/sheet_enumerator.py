"""Sheet Enumerator analyzer (spec §4.2).

Collects per-sheet metadata from the loader's **structure mode** workbook
[D3]: sheet name, visibility, used range, max row/column, dimension trust, and
a tabular-candidate guess. Dimensions are taken from structure mode because
read_only dimensions may be reset/unreliable (spec §4.2); when structure-mode
dimensions look untrustworthy this analyzer marks
``used_range_trusted=False`` and records a warning rather than failing (spec
§9). The ``is_tabular`` override [D2] short-circuits the tabular guess.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from openpyxl.utils import get_column_letter

from ..context import InspectionContext
from ..models import SheetProfile
from ..options import get_is_tabular_override
from ..pipeline import Analyzer

if TYPE_CHECKING:  # pragma: no cover - typing only
    from openpyxl.worksheet.worksheet import Worksheet

#: Sheets whose used range is at most this many columns wide are treated as
#: non-tabular (description/README) candidates unless overridden.
_MAX_NON_TABULAR_COLS = 1


class SheetEnumerator(Analyzer):
    """Enumerate sheets and their basic structural metadata (spec §4.2)."""

    def name(self) -> str:
        """Return the analyzer identifier."""

        return "sheet_enumerator"

    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Populate ``workbook_profile.sheets`` from the structure workbook.

        For each worksheet a :class:`SheetProfile` is created with ``name``,
        ``is_visible``, ``used_range``, ``max_row``/``max_col``,
        ``used_range_trusted``, and ``is_tabular_candidate``. The
        ``is_tabular`` override [D2] takes precedence over the heuristic.

        Args:
            context: Shared context carrying a ready :class:`Loader`.

        Returns:
            The same context with ``workbook_profile.sheets`` populated.
        """

        loader = context.loader
        if loader is None:  # pragma: no cover - guarded by pipeline wiring
            context.add_warning(
                "sheet_enumerator: no loader available; skipping enumeration"
            )
            return context

        workbook = loader.structure_workbook()
        profiles: list[SheetProfile] = []
        for worksheet in workbook.worksheets:
            profiles.append(self._profile_sheet(worksheet, context))

        context.workbook_profile.sheets = profiles
        return context

    def _profile_sheet(
        self, worksheet: "Worksheet", context: InspectionContext
    ) -> SheetProfile:
        """Build a :class:`SheetProfile` for a single worksheet."""

        name = worksheet.title
        is_visible = worksheet.sheet_state == "visible"

        max_row, max_col, used_range, trusted = self._dimensions(worksheet)
        if not trusted:
            context.add_warning(
                f"sheet_enumerator: untrusted dimensions for sheet "
                f"{name!r}; used_range marked unreliable"
            )

        is_tabular, provenance = self._is_tabular_candidate(
            context, name, max_row, max_col
        )

        return SheetProfile(
            name=name,
            is_visible=is_visible,
            is_tabular_candidate=is_tabular,
            is_tabular_provenance=provenance,
            used_range=used_range,
            used_range_trusted=trusted,
            max_row=max_row,
            max_col=max_col,
        )

    def _dimensions(
        self, worksheet: "Worksheet"
    ) -> tuple[int, int, str, bool]:
        """Return ``(max_row, max_col, used_range, trusted)`` (1-based).

        Dimensions come from structure mode. They are considered untrusted
        when openpyxl reports ``None`` (which can occur for files with a reset
        dimension record); in that case the values are coerced to ``1`` and the
        used range to the single anchor cell ``A1``.
        """

        max_row = worksheet.max_row
        max_col = worksheet.max_column

        trusted = max_row is not None and max_col is not None
        if max_row is None:
            max_row = 1
        if max_col is None:
            max_col = 1

        used_range = (
            f"A1:{get_column_letter(max_col)}{max_row}"
            if max_row >= 1 and max_col >= 1
            else "A1"
        )
        return max_row, max_col, used_range, trusted

    def _is_tabular_candidate(
        self,
        context: InspectionContext,
        sheet_name: str,
        max_row: int,
        max_col: int,
    ) -> tuple[bool, str]:
        """Decide whether a sheet looks like a data table (spec §4.2).

        The ``is_tabular`` override [D2] wins outright and is recorded with
        ``provenance="manual"``. Otherwise a sheet with no usable area, or only
        a single populated column (e.g. a README / description sheet), is
        flagged non-tabular with ``provenance="heuristic"``.

        Returns:
            ``(is_tabular_candidate, provenance)`` where provenance is
            ``"manual"`` for an override and ``"heuristic"`` otherwise [D2].
        """

        override = get_is_tabular_override(context.options, sheet_name)
        if override is not None:
            return override, "manual"

        if max_row < 1 or max_col < 1:
            return False, "heuristic"
        return max_col > _MAX_NON_TABULAR_COLS, "heuristic"
