"""Excel Structure Inspector.

Read-only inspection of ``.xlsx`` workbooks that produces a pandas-consumable
:class:`~excel_inspector.models.ReadPlan`. See ``docs/`` for the spec and
implementation plan.

The public surface exposes the data models, options, context, the
analyzer/pipeline interfaces, the loader, domain exceptions, the v1 entry point
:func:`inspect`, and the read-side adapter (:func:`load_dataframe` /
:func:`read_plan_to_kwargs`) that turns a :class:`~excel_inspector.models.ReadPlan`
into loaded data.

v1 public API:

    >>> from excel_inspector import inspect, load_dataframe
    >>> profile = inspect("book.xlsx")                 # WorkbookProfile
    >>> sheet = profile.sheets[0]
    >>> if sheet.read_plan is not None:
    ...     df = load_dataframe("book.xlsx", sheet.read_plan)  # pandas DataFrame

v2 result layer (Phase 9): the one-call :func:`extract` returns a
:class:`~excel_inspector.results.WorkbookResult` with every table loaded:

    >>> from excel_inspector import extract
    >>> result = extract("book.xlsx")                  # WorkbookResult
    >>> result.tables[0].dataframe                     # pandas DataFrame
    >>> result.to_json(indent=2)                       # JSON schema v1.0
    >>> result.to_markdown()                           # human-readable tables
"""

from __future__ import annotations

from pathlib import Path

from .adapters.pandas_loader import load_dataframe, read_plan_to_kwargs
from .aggregator import PlanAggregator
from .analyzers.block_analyzer import BlockAnalyzer
from .analyzers.block_segmenter import BlockSegmenter, RowBand, split_row_bands
from .analyzers.boundary_detector import BoundaryDetector
from .analyzers.formula_detector import FormulaDetector
from .analyzers.header_locator import HeaderLocator
from .analyzers.merge_analyzer import MergeAnalyzer, MergeScanner, MergeSpan
from .analyzers.sheet_enumerator import SheetEnumerator
from .analyzers.type_profiler import TypeProfiler
from .context import InspectionContext
from .exceptions import (
    CorruptWorkbookError,
    EncryptedWorkbookError,
    InspectorError,
)
from .loader import Loader
from .models import (
    ColumnProfile,
    InspectionOptions,
    MergeRegion,
    ReadPlan,
    SheetOverride,
    SheetProfile,
    TableBlock,
    WorkbookProfile,
)
from .pipeline import Analyzer, Pipeline
from .results import (
    SheetResultEntry,
    TableResult,
    WorkbookResult,
    build_workbook_result,
)


def inspect(
    path: str | Path, options: InspectionOptions | None = None
) -> WorkbookProfile:
    """Inspect a workbook and return its :class:`WorkbookProfile` (spec §3).

    Wires the pipeline ``Loader -> [SheetEnumerator -> MergeScanner ->
    BlockSegmenter -> HeaderLocator -> BoundaryDetector -> TypeProfiler ->
    BlockAnalyzer -> MergeAnalyzer -> FormulaDetector] -> PlanAggregator``
    and runs it over a shared context. The Merge Scanner (plan v2 Task 11.1)
    collects each sheet's merged ranges once — *before* boundary detection —
    so a merged header's collapsed column span can be bridged (virtual fill);
    the Merge Analyzer later classifies the collected ranges against the
    resolved header. The Block Segmenter (Phase 10a) splits each sheet into
    row bands right after the scan; the Block Analyzer (Phase 10b) then
    re-runs Header/Boundary/Type per band so every stacked table becomes a
    :class:`~excel_inspector.models.TableBlock` with its own read plan — a
    multi-table sheet loses no table silently. The Formula Detector (Phase
    12) then flags formula columns per block and recommends
    ``as_value``/``as_formula``, opening its third (formula-mode) workbook
    handle lazily — only for files that actually contain formula markup. The
    loader is opened read-only and every handle is released before returning
    (spec §4.1, §8) [D3], so inspection is side-effect free and idempotent.

    Args:
        path: Path to the ``.xlsx`` file to inspect.
        options: Optional inspection overrides [D2]; defaults to a fresh
            :class:`InspectionOptions`.

    Returns:
        The populated :class:`WorkbookProfile` (``file_path`` set, per-sheet
        profiles and v1 read plans attached).

    Raises:
        CorruptWorkbookError: The file is corrupt / not a valid ``.xlsx``.
        EncryptedWorkbookError: The file is password/encryption protected.
    """

    resolved = Path(path)
    profile = WorkbookProfile(file_path=str(resolved))
    context = InspectionContext(
        options=options or InspectionOptions(),
        workbook_profile=profile,
    )
    pipeline = Pipeline(
        [
            SheetEnumerator(),
            MergeScanner(),
            BlockSegmenter(),
            HeaderLocator(),
            BoundaryDetector(),
            TypeProfiler(),
            BlockAnalyzer(),
            MergeAnalyzer(),
            FormulaDetector(),
            PlanAggregator(),
        ]
    )

    with Loader(resolved) as loader:
        context.loader = loader
        context = pipeline.run(context)

    context.workbook_profile.open_errors.extend(context.warnings)
    return context.workbook_profile


def extract(
    path: str | Path, options: InspectionOptions | None = None
) -> WorkbookResult:
    """One-call API: inspect the workbook, then load every table per its ReadPlan.

    Phase 9 result layer (plan v2 §3): runs :func:`inspect` and translates the
    resulting :class:`WorkbookProfile` into a :class:`WorkbookResult` whose
    per-table DataFrames are already loaded, ready for ``to_dict()`` /
    ``to_json()`` (schema v1.0) / ``to_markdown()`` serialization.

    Args:
        path: Path to the ``.xlsx`` file to extract.
        options: Optional inspection overrides [D2]; defaults to a fresh
            :class:`InspectionOptions`.

    Returns:
        The loaded :class:`WorkbookResult` (one :class:`TableResult` per
        tabular sheet; non-tabular sheets recorded as skipped entries).

    Raises:
        CorruptWorkbookError: The file is corrupt / not a valid ``.xlsx``.
        EncryptedWorkbookError: The file is password/encryption protected.
    """

    return build_workbook_result(path, inspect(path, options))


__all__ = [
    "Analyzer",
    "BlockAnalyzer",
    "BlockSegmenter",
    "BoundaryDetector",
    "ColumnProfile",
    "CorruptWorkbookError",
    "EncryptedWorkbookError",
    "FormulaDetector",
    "HeaderLocator",
    "InspectionContext",
    "InspectionOptions",
    "InspectorError",
    "Loader",
    "MergeAnalyzer",
    "MergeRegion",
    "MergeScanner",
    "MergeSpan",
    "Pipeline",
    "PlanAggregator",
    "ReadPlan",
    "RowBand",
    "SheetEnumerator",
    "SheetOverride",
    "SheetProfile",
    "SheetResultEntry",
    "TableBlock",
    "TableResult",
    "TypeProfiler",
    "WorkbookProfile",
    "WorkbookResult",
    "build_workbook_result",
    "extract",
    "inspect",
    "load_dataframe",
    "read_plan_to_kwargs",
    "split_row_bands",
]

__version__ = "0.1.0"
