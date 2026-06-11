"""Phase 0 smoke test: empty pipeline yields an empty WorkbookProfile.

Completion criterion (implementation plan §3, Phase 0): pytest runs the
pipeline with an empty analyzer list without raising and returns an empty
``WorkbookProfile``.
"""

from __future__ import annotations

from excel_inspector import (
    InspectionContext,
    InspectionOptions,
    Pipeline,
    WorkbookProfile,
)


def test_empty_pipeline_returns_empty_workbook_profile() -> None:
    """Running an empty pipeline leaves an empty WorkbookProfile and no error."""
    context = InspectionContext()
    pipeline = Pipeline([])

    result = pipeline.run(context)

    assert result is context
    assert isinstance(result.workbook_profile, WorkbookProfile)
    # "Empty": no sheets, no open errors, no warnings.
    assert result.workbook_profile.sheets == []
    assert result.workbook_profile.open_errors == []
    assert result.workbook_profile.file_path == ""
    assert result.warnings == []


def test_empty_pipeline_supports_no_analyzers_argument() -> None:
    """Pipeline() with no analyzers argument also supports an empty run."""
    pipeline = Pipeline()

    assert pipeline.analyzers == []

    result = pipeline.run(InspectionContext(options=InspectionOptions()))

    assert result.workbook_profile.sheets == []


def test_pipeline_defaults_are_sensible() -> None:
    """A default InspectionContext carries default options and empty state."""
    context = InspectionContext()

    assert isinstance(context.options, InspectionOptions)
    assert context.options.header_confidence_threshold == 0.5
    assert context.options.skip_keywords is None
    assert context.options.sheet_overrides == {}
    assert context.loader is None
    assert context.warnings == []
