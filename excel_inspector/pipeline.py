"""Analyzer interface and pipeline runner (spec §3, §6).

The inspector follows a pipeline architecture: analyzers implementing a common
contract run in topological order (spec §3), each enriching the shared
:class:`~excel_inspector.context.InspectionContext`.

Robustness policy (spec §6, §9): an individual analyzer failure must not halt
the whole pipeline — such failures are absorbed into ``context.warnings``.
The loader's corruption/encryption domain exceptions
(:class:`~excel_inspector.exceptions.InspectorError`) are the explicit
exception to this rule and propagate immediately.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Iterable

from .context import InspectionContext
from .exceptions import InspectorError


class Analyzer(ABC):
    """Common analyzer contract (spec §6).

    Every analyzer takes the shared context, enriches it with its own results,
    and returns it.
    """

    @abstractmethod
    def analyze(self, context: InspectionContext) -> InspectionContext:
        """Enrich and return the shared context.

        Args:
            context: The shared, partially-populated inspection context.

        Returns:
            The (same) context, enriched with this analyzer's results.
        """

    @abstractmethod
    def name(self) -> str:
        """Return a stable identifier for logging/diagnostics."""


class Pipeline:
    """Runs analyzers in order over a shared context (spec §3, §6).

    Args:
        analyzers: Ordered analyzers to run. An empty list is supported and
            simply returns the context unchanged.
    """

    def __init__(self, analyzers: Iterable[Analyzer] | None = None) -> None:
        self._analyzers: list[Analyzer] = list(analyzers or [])

    @property
    def analyzers(self) -> list[Analyzer]:
        """The ordered analyzers this pipeline will run."""

        return list(self._analyzers)

    def run(self, context: InspectionContext) -> InspectionContext:
        """Execute every analyzer in order against ``context``.

        Analyzer exceptions are absorbed into ``context.warnings`` so the
        pipeline continues. Loader domain exceptions
        (:class:`InspectorError`) propagate immediately (spec §6, §9).

        Args:
            context: The shared context (already carrying options/loader).

        Returns:
            The enriched context after all analyzers have run.

        Raises:
            InspectorError: Propagated unchanged from any analyzer (loader
                corruption/encryption domain errors must stop the pipeline).
        """

        for analyzer in self._analyzers:
            try:
                context = analyzer.analyze(context)
            except InspectorError:
                # Loader domain errors (corrupt/encrypted) must halt — re-raise.
                raise
            except Exception as exc:  # noqa: BLE001 - intentional absorption
                context.add_warning(
                    f"analyzer '{analyzer.name()}' failed: "
                    f"{type(exc).__name__}: {exc}"
                )
        return context
