"""CLI: ``python -m excel_inspector <file.xlsx> [--format json|markdown] [--max-rows N]``.

Plan v2 Phase 13 Step 1 (L8): the one-shot ergonomic entry point over
:func:`excel_inspector.extract` — every table of the workbook is extracted and
printed to stdout as human-readable Markdown tables (default) or as the
schema-v1.0 JSON document.

Output / exit-code contract:

* success -> exit code ``0``; the rendered document on **stdout** only.
* corrupt / encrypted / unreadable input -> exit code ``1``; nothing on
  stdout, an explicit ``error: ...`` line (the domain exception's message) on
  **stderr** — so shell pipelines never mistake a failure for empty output.
* a nonexistent path -> ``error: file not found: ...``; a directory (or any
  non-regular-file path) -> ``error: not a file: ...`` — both exit code ``1``
  (review LOW: previously both surfaced as a misleading "corrupt" error).
* ``--max-rows`` bounds the rows rendered per table in *markdown* mode
  (default 20, must be ``>= 0`` — a negative value is an argparse usage error,
  exit code ``2``). JSON mode always emits every row (the plan §7 Step 1
  contract) so the document stays a faithful, machine-consumable extraction.

The advisory ``warnings`` (multi-table notices, formula-cache gaps, ...) ride
inside the rendered document itself — JSON's ``warnings`` array / the
markdown trailer — exactly as :class:`~excel_inspector.results.WorkbookResult`
serializes them.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import extract
from .exceptions import InspectorError


def _non_negative_int(text: str) -> int:
    """argparse type for ``--max-rows``: an int ``>= 0`` (review LOW).

    A negative row budget is meaningless (``DataFrame.head(-n)`` silently
    *drops* the last n rows instead of failing), so it is rejected at parse
    time with an explicit message rather than producing surprising output.

    Args:
        text: The raw CLI token.

    Returns:
        The parsed non-negative int.

    Raises:
        argparse.ArgumentTypeError: Not an int, or negative — argparse turns
            this into the standard ``error: argument --max-rows: ...`` usage
            failure (exit code 2).
    """

    try:
        value = int(text)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"invalid int value: {text!r}"
        ) from exc
    if value < 0:
        raise argparse.ArgumentTypeError(
            f"must be a non-negative integer, got {value}"
        )
    return value


def main(argv: list[str] | None = None) -> int:
    """Run the CLI and return the process exit code.

    Args:
        argv: Argument vector (without the program name); ``None`` reads
            ``sys.argv[1:]`` (the ``python -m excel_inspector`` path).

    Returns:
        ``0`` on success, ``1`` when the workbook cannot be opened/read (the
        explicit error goes to stderr).
    """

    parser = argparse.ArgumentParser(
        prog="excel_inspector",
        description=(
            "Extract every table of an .xlsx workbook and print it as "
            "Markdown tables or as the schema-v1.0 JSON document."
        ),
    )
    parser.add_argument("path", help="path to the .xlsx workbook")
    parser.add_argument(
        "--format",
        choices=["markdown", "json"],
        default="markdown",
        help="output format (default: markdown)",
    )
    parser.add_argument(
        "--max-rows",
        type=_non_negative_int,
        default=20,
        help=(
            "rows rendered per table in markdown output (default: 20, "
            "must be >= 0); json output always contains every row"
        ),
    )
    args = parser.parse_args(argv)

    # Path pre-checks (review LOW): a missing path and a directory both used
    # to fall through to openpyxl and surface as a misleading "corrupt or not
    # a valid .xlsx" error. Distinguish them explicitly before extraction.
    path = Path(args.path)
    if not path.exists():
        print(f"error: file not found: {args.path}", file=sys.stderr)
        return 1
    if not path.is_file():
        print(f"error: not a file: {args.path}", file=sys.stderr)
        return 1

    try:
        result = extract(args.path)
    except InspectorError as exc:
        # Corrupt / encrypted workbooks (spec §9): explicit stderr error +
        # non-zero exit so callers and pipelines can rely on the exit code.
        print(f"error: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:  # pragma: no cover - defensive (loader translates)
        print(f"error: cannot read {args.path!r}: {exc}", file=sys.stderr)
        return 1

    if args.format == "json":
        print(result.to_json(indent=2, max_rows=None))
    else:
        print(result.to_markdown(max_rows=args.max_rows))
    return 0


if __name__ == "__main__":
    sys.exit(main())
