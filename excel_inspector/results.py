"""Result layer: per-table DataFrames + JSON/Markdown serialization (plan v2 Phase 9).

This module consumes only the public contracts (WorkbookProfile, ReadPlan,
adapters.pandas_loader) — it must not reach into analyzer internals.

Serialization rules (fixed contract, plan v2 §3.0):

* dates/datetimes -> ISO 8601 strings;
* missing (``NaN``/``NA``/``NaT``/``None``) -> ``null``;
* ``numeric_text`` stays a string (leading zeros survive);
* numpy scalars -> native Python values;
* headerless (``header=None``) column names -> ``"col_0".."col_n"``;
* non-string column labels (date/int header cells) -> strings via the same
  scalar contract (dates become ISO 8601 strings), so JSON object keys and
  DataFrame columns are always ``str``;
* MultiIndex columns (multi-level headers, plan v2 Task 11.2) -> flattened to
  ``"상위 / 하위"`` strings; unfilled group/leaf cells (pandas
  ``Unnamed: N_level_M`` placeholders or missing values) contribute an empty
  string and are dropped from the join, so a column under an unnamed group
  keeps just its leaf name;
* duplicate column names -> uniquified with ``.1``/``.2`` suffixes;
* ``bytes``/``timedelta``/``Decimal`` -> explicit string fallback (never a
  silent ``json.dumps`` crash);
* each ``columns[]`` item exposes both ``name`` (the raw header cell seen at
  inspection time) and ``resolved_name`` (the post-load flattened/deduped
  column name at that position — the exact ``records`` key, so the two are
  joinable; adversarial review MEDIUM #2).

Path contract (plan v2 §3.0): the JSON ``file`` field is the **absolute** path
of the inspected workbook — :func:`build_workbook_result` normalizes the
caller's path with ``Path.absolute()`` (no symlink resolution, so fixture
paths survive macOS ``/var`` -> ``/private/var`` rewriting).
"""

from __future__ import annotations

import datetime as _dt
import decimal as _decimal
import json
import re
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any

import pandas as pd

from .adapters.pandas_loader import load_dataframe
from .models import ColumnProfile, ReadPlan, WorkbookProfile

#: JSON schema version emitted by :meth:`WorkbookResult.to_dict` (plan v2 §3.0).
#: The shape is stable across later phases — only ``tables`` items grow.
SCHEMA_VERSION = "1.0"


def _jsonify_scalar(value: Any) -> Any:
    """Convert a pandas/numpy/datetime scalar to a JSON-compatible value.

    Rules (fixed contract): missing (NaN/NA/NaT/None) -> None; datetimes/dates
    -> ISO 8601 strings; numpy scalars -> native Python; bytes -> UTF-8 text
    (replacement on undecodable bytes); timedelta/Decimal -> ``str`` (explicit
    fallback so ``json.dumps`` never crashes on them); everything else as-is.
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass  # arrays/odd types: not a missing scalar
    if isinstance(value, (pd.Timestamp, _dt.datetime, _dt.date)):
        return value.isoformat()
    # Explicit string fallbacks (plan v2 §3 review checklist): a silent
    # json.dumps TypeError on bytes/timedelta/Decimal is forbidden.
    if isinstance(value, _dt.timedelta):  # includes pd.Timedelta
        return str(value)
    if isinstance(value, (bytes, bytearray)):
        return bytes(value).decode("utf-8", errors="replace")
    if isinstance(value, _decimal.Decimal):
        return str(value)
    if hasattr(value, "item"):  # numpy scalar
        # Recurse: np.datetime64 -> datetime, np.timedelta64 -> timedelta/int,
        # np.bytes_ -> bytes all need a second pass through the rules above.
        return _jsonify_scalar(value.item())
    return value


#: pandas' placeholder for an empty header cell: ``"Unnamed: 3"`` for a flat
#: header, ``"Unnamed: 0_level_0"`` / ``"Unnamed: 2_level_1"`` for a
#: multi-level (list) header (measured against pandas 3.0.3, plan v2 Task
#: 11.2 Step 0 spike). These are *unfilled* cells, not real labels.
_UNNAMED_LABEL_RE = re.compile(r"^Unnamed: \d+(?:_level_\d+)?$")

#: Separator joining MultiIndex levels into one flat column name
#: (plan v2 Task 11.2 Step 3: ``"상위 / 하위"``).
_LEVEL_JOIN = " / "


def _flatten_column_tuple(parts: tuple[Any, ...]) -> str:
    """Flatten one MultiIndex column tuple to ``"상위 / 하위"`` (Task 11.2).

    Each level goes through the scalar label contract (dates -> ISO 8601,
    ``str`` fallback). Unfilled cells — missing values or pandas'
    ``Unnamed: N[_level_M]`` placeholders — are treated as the empty string
    (plan rule) and therefore dropped from the join, so a column whose group
    cell is empty keeps just its leaf name (and vice versa). A tuple with no
    real label at all flattens to ``""`` (later uniquified by
    :func:`_dedupe_columns`).
    """

    rendered: list[str] = []
    for part in parts:
        jsonified = _jsonify_scalar(part)
        if jsonified is None:
            continue  # missing level cell -> empty contribution
        text = jsonified if isinstance(jsonified, str) else str(jsonified)
        if not text or _UNNAMED_LABEL_RE.match(text):
            continue  # unfilled group/leaf cell -> empty contribution
        rendered.append(text)
    return _LEVEL_JOIN.join(rendered)


def _stringify_label(column: Any) -> str:
    """Render one column label as a string (fixed serialization contract).

    Column labels read from a worksheet header are not always strings — a
    header cell holding a date or an int yields a non-string pandas label,
    which would crash ``json.dumps`` when used as a records key (P9 review,
    HIGH). Labels reuse the scalar contract of :func:`_jsonify_scalar` (dates
    -> ISO 8601), with a final ``str`` fallback for anything non-string.
    A tuple label (a MultiIndex column on a directly-constructed
    :class:`TableResult`) flattens through :func:`_flatten_column_tuple` so
    JSON keys never read as ``"('상위', '하위')"``.
    """

    if isinstance(column, str):
        return column
    if isinstance(column, tuple):
        return _flatten_column_tuple(column)
    jsonified = _jsonify_scalar(column)
    return jsonified if isinstance(jsonified, str) else str(jsonified)


def _dedupe_columns(names: list[str]) -> list[str]:
    """Make column names unique with '.N' suffixes (JSON object keys must be unique).

    Collision-safe (plan v2 §3 review checklist): an input like
    ``["a", "a.1", "a"]`` must not produce a second ``"a.1"`` — the suffix
    counter keeps advancing until the candidate is genuinely unused.
    """
    seen: dict[str, int] = {}
    used: set[str] = set()
    out: list[str] = []
    for n in names:
        if n not in used:
            seen.setdefault(n, 0)
            used.add(n)
            out.append(n)
            continue
        counter = seen.get(n, 0)
        candidate = n
        while candidate in used:
            counter += 1
            candidate = f"{n}.{counter}"
        seen[n] = counter
        used.add(candidate)
        out.append(candidate)
    return out


def _md_cell(value: Any) -> str:
    """Render one scalar as a Markdown table cell (escaped, never breaks rows).

    Missing values render empty; ``|`` is escaped and newlines collapse to a
    single space so a cell can never break the table structure (plan v2 §3
    review checklist).
    """
    jsonified = _jsonify_scalar(value)
    if jsonified is None:
        return ""
    text = str(jsonified)
    return (
        text.replace("|", "\\|")
        .replace("\r\n", " ")
        .replace("\n", " ")
        .replace("\r", " ")
    )


@dataclass
class TableResult:
    """One extracted table: cleaned DataFrame + the inspection metadata behind it."""

    sheet_name: str
    table_id: str                      # "<sheet>!T<n>" (1-based block order, top-down)
    dataframe: pd.DataFrame
    header_row: int | None             # 1-based (inspection domain), None = headerless
    header_confidence: float
    header_provenance: str
    columns: list[ColumnProfile]
    notes: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        """Resolve each column's post-load name positionally (review MEDIUM #2).

        ``ColumnProfile.name`` is the raw header cell seen at inspection time;
        after loading, flattening (``"상위 / 하위"``) and deduping (``.N``)
        the ``records`` keys no longer match it, so the two could not be
        joined. ``resolved_name`` is filled here — at TableResult creation —
        by position: ``ColumnProfile.index`` is the 0-based position within
        the usecols-selected frame [D5], which is exactly the DataFrame column
        position. Labels go through :func:`_stringify_label`, the same
        contract ``to_dict`` applies to record keys, so
        ``columns[i].resolved_name`` always equals the records key of that
        column. Profiles are *copied* (``dataclasses.replace``), never mutated
        in place — the inspection-side ``SheetProfile``/``TableBlock`` column
        lists share these instances. A caller-supplied ``resolved_name`` is
        kept; an out-of-range index (column-count mismatch) stays ``None``.
        """

        keys = [_stringify_label(col) for col in self.dataframe.columns]
        self.columns = [
            c
            if c.resolved_name is not None
            else replace(
                c,
                resolved_name=(
                    keys[c.index] if 0 <= c.index < len(keys) else None
                ),
            )
            for c in self.columns
        ]

    def to_dict(self, max_rows: int | None = None) -> dict[str, Any]:
        df = self.dataframe if max_rows is None else self.dataframe.head(max_rows)
        # Records keys go through _stringify_label too: a TableResult is a
        # public dataclass constructible without the builder, so a directly
        # attached DataFrame with non-string labels must not crash to_json().
        keys = [_stringify_label(col) for col in df.columns]
        records = [
            {key: _jsonify_scalar(v) for key, v in zip(keys, row)}
            for row in df.itertuples(index=False, name=None)
        ]
        return {
            "table_id": self.table_id,
            "header_row": self.header_row,
            "header_confidence": round(self.header_confidence, 4),
            "header_provenance": self.header_provenance,
            "columns": [
                {"index": c.index, "name": c.name,
                 "resolved_name": c.resolved_name,
                 "inferred_type": c.inferred_type, "null_ratio": round(c.null_ratio, 4)}
                for c in self.columns
            ],
            "row_count": len(self.dataframe),
            "records": records,
            "notes": list(self.notes),
        }

    def to_json(self, max_rows: int | None = None, **dumps_kwargs: Any) -> str:
        dumps_kwargs.setdefault("ensure_ascii", False)
        return json.dumps(self.to_dict(max_rows=max_rows), **dumps_kwargs)

    def to_markdown(self, max_rows: int = 20) -> str:
        df = self.dataframe.head(max_rows)
        headers = [_md_cell(c) for c in df.columns]
        out = ["| " + " | ".join(headers) + " |",
               "| " + " | ".join("---" for _ in headers) + " |"]
        for row in df.itertuples(index=False, name=None):
            cells = [_md_cell(v) for v in row]
            out.append("| " + " | ".join(cells) + " |")
        if len(self.dataframe) > max_rows:
            out.append(f"\n… {len(self.dataframe) - max_rows} more rows")
        return "\n".join(out)


@dataclass
class SheetResultEntry:
    """Per-sheet grouping inside WorkbookResult (mirrors the JSON 'sheets' items)."""

    name: str
    is_visible: bool
    tables: list[TableResult]
    skipped: bool = False
    skip_reason: str | None = None


@dataclass
class WorkbookResult:
    """Whole-workbook extraction result (JSON schema v1.0; plan v2 §3.0)."""

    file_path: str
    sheets: list[SheetResultEntry]
    warnings: list[str] = field(default_factory=list)

    @property
    def tables(self) -> list[TableResult]:
        return [t for s in self.sheets for t in s.tables]

    def to_dict(self, max_rows: int | None = None) -> dict[str, Any]:
        return {
            "schema_version": SCHEMA_VERSION,
            "file": self.file_path,
            "sheets": [
                {"name": s.name, "is_visible": s.is_visible,
                 "tables": [t.to_dict(max_rows=max_rows) for t in s.tables],
                 "skipped": s.skipped, "skip_reason": s.skip_reason}
                for s in self.sheets
            ],
            "warnings": list(self.warnings),
        }

    def to_json(self, max_rows: int | None = None, **dumps_kwargs: Any) -> str:
        dumps_kwargs.setdefault("ensure_ascii", False)
        return json.dumps(self.to_dict(max_rows=max_rows), **dumps_kwargs)

    def to_markdown(self, max_rows: int = 20) -> str:
        parts: list[str] = []
        for s in self.sheets:
            for t in s.tables:
                parts.append(f"### {t.table_id}\n\n{t.to_markdown(max_rows=max_rows)}")
            if s.skipped:
                parts.append(f"### {s.name} — skipped ({s.skip_reason})")
        if self.warnings:
            parts.append("> ⚠ " + "\n> ⚠ ".join(self.warnings))
        return "\n\n".join(parts)


def _postprocess_dataframe(df: pd.DataFrame, plan: ReadPlan) -> pd.DataFrame:
    """Apply result-layer column-name contracts: stringify + headerless + dedupe.

    Every label is stringified via :func:`_stringify_label` (a date/int header
    cell otherwise leaves a non-string label that crashes ``to_json()``; P9
    review, HIGH). A MultiIndex (multi-level header plan, ``header=[..]``)
    yields tuple labels which flatten to ``"상위 / 하위"`` strings via the same
    helper (plan v2 Task 11.2 Step 3), then deduplicate like any other name.
    The rename guard compares the *raw* labels against the stringified names,
    so it fires for any non-string label while staying a no-op for the
    all-string common case.
    """
    if plan.header is None:
        df = df.copy()
        df.columns = [f"col_{i}" for i in range(len(df.columns))]
        return df
    names = _dedupe_columns([_stringify_label(c) for c in df.columns])
    if list(df.columns) != names:
        df = df.copy()
        df.columns = names
    return df


def build_workbook_result(
    file_path: str | Path, profile: WorkbookProfile
) -> WorkbookResult:
    """Translate an inspected WorkbookProfile into loaded per-table results.

    The result's ``file_path`` (the JSON ``file`` field) is normalized to an
    **absolute** path via ``Path.absolute()`` — matching the plan v2 §3.0
    example — without symlink resolution (P9 review, LOW).

    Multi-table sheets (plan v2 Task 10.2 Step 4): when ``sp.blocks`` is
    populated, every block yields one :class:`TableResult` with the id
    ``"{sheet}!T{n}"`` (``n`` 1-based, top-down). A sheet without blocks
    (headerless / fallback) keeps the v1 single-table path driven by the flat
    ``read_plan``.
    """
    sheets: list[SheetResultEntry] = []
    warnings: list[str] = list(profile.open_errors)
    for sp in profile.sheets:
        if not sp.is_tabular_candidate or sp.read_plan is None:
            sheets.append(SheetResultEntry(
                name=sp.name, is_visible=sp.is_visible, tables=[],
                skipped=True, skip_reason="non-tabular"))
            continue
        tables: list[TableResult] = []
        if sp.blocks:
            for number, block in enumerate(sp.blocks, start=1):
                plan = block.read_plan
                if plan is None:  # pragma: no cover - aggregator fills plans
                    continue
                df = _postprocess_dataframe(
                    load_dataframe(file_path, plan), plan
                )
                tables.append(TableResult(
                    sheet_name=sp.name, table_id=f"{sp.name}!T{number}",
                    dataframe=df, header_row=block.header_row,
                    header_confidence=block.header_confidence,
                    header_provenance=block.header_provenance,
                    columns=list(block.columns), notes=list(plan.notes),
                ))
        else:
            df = _postprocess_dataframe(
                load_dataframe(file_path, sp.read_plan), sp.read_plan
            )
            tables.append(TableResult(
                sheet_name=sp.name, table_id=f"{sp.name}!T1", dataframe=df,
                header_row=sp.header_row,
                header_confidence=sp.header_confidence,
                header_provenance=sp.header_provenance,
                columns=list(sp.columns), notes=list(sp.read_plan.notes),
            ))
        sheets.append(SheetResultEntry(
            name=sp.name, is_visible=sp.is_visible, tables=tables))
    return WorkbookResult(
        file_path=str(Path(file_path).absolute()), sheets=sheets, warnings=warnings
    )
