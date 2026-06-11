# 엑셀 구조 검사기 v2 구현 플랜 — 테이블별 결과 출력(JSON) & 전상황 커버리지

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 어떤 형태의 실무 엑셀이 들어와도, 시트 안의 **표 하나하나를 독립된 데이터 테이블(DataFrame)로 추출**하고 그 결과를 **JSON/마크다운으로 직렬화**하는 단일 호출 API(`extract()`)와 CLI를 완성한다.

**Architecture:** v1 파이프라인(`inspect()` → `WorkbookProfile`/`ReadPlan`)은 그대로 두고, 그 위에 **결과 계층(results)** 을 얹는다. 이후 시트를 **행 밴드(블록)** 로 분할해 블록마다 Header→Boundary→Type을 독립 실행함으로써 "시트당 표 1개" 가정을 제거한다(다중 표). 병합/다단 헤더·수식은 후속 단계로 해소한다.

**Tech Stack:** Python 3.14 (.venv), openpyxl 3.1.5, pandas 3.0.3, pytest 9.0.3. 신규 런타임 의존성 없음(마크다운 렌더러는 자체 구현).

| 항목 | 내용 |
| --- | --- |
| 문서 종류 | 구현 플랜 (v2) |
| 선행 문서 | `excel-structure-inspector-spec.md`(개정 1판), `excel-structure-inspector-implementation-plan.md`(v1, Phase 0~8 완료) |
| 작성일 | 2026-06-11 |
| 현재 상태 | v1 완료 — 383 tests green, 픽스처 17개 |

> ⚠️ **git 미사용**: 본 프로젝트는 git 저장소가 아니다. 각 태스크의 체크포인트는 커밋이 아니라 **전체 스위트 green**(`383+ passed`, 기존 테스트 무파손)이다. 버전 관리를 원하면 착수 전 `git init`을 별도 결정으로 진행한다.

---

## 0. 배경 — v1이 남긴 것

v1(Phase 0~8)으로 검사→ReadPlan→pandas 적재까지 완성됐다. 남은 한계는 모두 문서에 기록되어 있다:

| # | 한계 | 출처 | 해소 단계 |
| --- | --- | --- | --- |
| L1 | 결과를 받으려면 `inspect()` + `load_dataframe()`를 직접 조합해야 하고, JSON/표 출력 계층이 없음 | 사용자 요구 | **Phase 9** |
| L2 | 한 시트 다중 표: 신뢰도 최고 블록 1개만 적재, 나머지 **조용히 누락**(경고 없음) | 스펙 §10 | **Phase 10** |
| L3 | 병합 헤더(`merged_header`)에서 헤더 열 구간이 1칸으로 좁아져 경계 미해소(`data_start=None`) | 스펙 §7.2, openIssue | **Phase 11a** |
| L4 | 다단 헤더: `is_multi_level_header` 판정만 하고 적재 분기 없음(`header: list[int]` 미사용) | 스펙 [D6] | **Phase 11b** |
| L5 | 수식 탐지 미구현(`has_formula` 항상 False) | 스펙 [D6], v1 플랜 Phase 9 | **Phase 12** |
| L6 | 헤더리스 시트는 타입 프로파일·dtype_map이 비어 silent 손실 (LOW) | W3 리뷰 | **Phase 13** |
| L7 | 키워드 선두 라벨 스캔이 표의 `left_col`이 아닌 시트 A열 기준 (openIssue) | W3 리뷰 | **Phase 13** |
| L8 | CLI/원샷 진입점 없음 — "모든 상황에서 사용"의 인체공학 | 사용자 요구 | **Phase 13** |

**우선순위 근거**: 사용자의 직접 요구가 "각각의 테이블 + JSON"이므로 결과 계층(Phase 9)이 최우선. "각각의"가 성립하려면 다중 표(Phase 10)가 본질이므로 그다음. Phase 9의 데이터 모델을 처음부터 **테이블 목록(list)** 으로 설계해 Phase 10이 끼워져도 JSON 스키마가 변하지 않게 한다.

---

## 1. 파일 맵

```
excel_inspector/
├── results.py                  # [신규 P9] TableResult / WorkbookResult / JSON·MD 직렬화
├── __init__.py                 # [수정 P9] extract() 공개, [P13] __all__ 정리
├── __main__.py                 # [신규 P13] CLI (python -m excel_inspector)
├── models.py                   # [수정 P10] TableBlock 추가, SheetProfile.blocks
├── analyzers/
│   ├── block_segmenter.py      # [신규 P10] 행 밴드 분할 (BLANK_RUN 기준)
│   ├── header_locator.py       # [수정 P10] row_window 스코프 지원
│   ├── boundary_detector.py    # [수정 P10] row_window, [P11a] 병합 가상 채움, [P13] left_col 키워드
│   ├── merge_analyzer.py       # [수정 P11a] 스캔(전단)/분류(후단) 분리
│   ├── type_profiler.py        # [수정 P10] row_window/블록 단위
│   └── formula_detector.py     # [신규 P12] 수식 탐지 (v1 플랜 Phase 9 승계)
├── loader.py                   # [수정 P12] formula_workbook() 모드 활성화
├── aggregator.py               # [수정 P10] 블록별 ReadPlan, [P11b] header: list[int]
└── adapters/pandas_loader.py   # [수정 P11b] header 리스트 → MultiIndex 처리

tests/
├── test_results.py             # [신규 P9]
├── test_block_segmenter.py     # [신규 P10]
├── test_multi_table.py         # [신규 P10] 골든
├── test_merge_bridge.py        # [신규 P11a]
├── test_multi_level_load.py    # [신규 P11b]
├── test_formula_detector.py    # [신규 P12]
├── test_cli.py                 # [신규 P13]
└── fixtures/generate.py        # [수정] multi_table_stacked / formulas / wide_two_tables 추가
```

**경계 원칙**: `results.py`는 `WorkbookProfile`+`ReadPlan`+어댑터만 소비한다(분석기 내부에 의존 금지). 좌표 변환은 v1과 동일하게 `aggregator.py` 단독 책임을 유지한다.

---

## 2. 단계 개요와 의존성

```
Phase 9 (결과 계층: extract/JSON/MD)        ← 즉시 착수 가능 (v1 위에 순수 추가)
   ↓
Phase 10a (다중 블록 감지 + 경고)            ← 조용한 누락 제거 (안전장치)
Phase 10b (블록별 완전 추출)                 ← SheetProfile.blocks, 표마다 ReadPlan
   ↓
Phase 11a (병합 헤더 경계 브리지)            ← merged_header 해소
Phase 11b (다단 헤더 적재: header=list[int]) ← MultiIndex + 평탄화
   ↓
Phase 12 (수식 탐지)                         ← 독립적 (11과 병렬 가능)
Phase 13 (CLI + 헤더리스/키워드/마무리)      ← 최종
```

각 Phase는 단독으로 배포 가능한 동작 산출물을 남긴다. **모든 Phase의 공통 완료 기준: 전체 스위트 green + 기존 383개 무파손 + 신규 테스트 추가.**

---

## 3. Phase 9 — 결과 계층: `extract()` / TableResult / JSON·마크다운 (M)

### 3.0 설계 계약

**공개 API** (시그니처는 현 코드 실측과 일치시킨다):

```python
from excel_inspector import extract
result = extract("file.xlsx", options=None)   # -> WorkbookResult

result.tables                  # list[TableResult] — 표 하나당 1개 (P10 이후 시트당 N개)
result.tables[0].dataframe     # pandas DataFrame (정제 적재 완료)
result.tables[0].to_dict()     # JSON 호환 dict
result.to_json(indent=2)       # 워크북 전체 JSON 문자열
result.to_markdown()           # 사람이 읽는 표 출력
```

**JSON 스키마 v1.0** (결정적·버전 명시, P10 이후에도 형태 불변 — tables 항목만 늘어남):

```json
{
  "schema_version": "1.0",
  "file": "/abs/path/file.xlsx",
  "sheets": [
    {
      "name": "매출",
      "is_visible": true,
      "tables": [
        {
          "table_id": "매출!T1",
          "header_row": 4,
          "header_confidence": 0.88,
          "header_provenance": "heuristic",
          "columns": [
            {"index": 0, "name": "지역", "inferred_type": "text", "null_ratio": 0.0}
          ],
          "row_count": 6,
          "records": [
            {"지역": "서울", "제품코드": "00123", "출시일": "2026-01-05T00:00:00", "수량": 50, "매출액": 20000, "비고": null}
          ],
          "notes": []
        }
      ],
      "skipped": false,
      "skip_reason": null
    },
    {"name": "안내", "is_visible": true, "tables": [], "skipped": true, "skip_reason": "non-tabular"}
  ],
  "warnings": []
}
```

**직렬화 규칙(고정)**: 날짜 → ISO 8601 문자열 / `NaN`·`NA`·`NaT` → `null` / `numeric_text`는 문자열 유지(선행 0 보존) / numpy 스칼라 → 파이썬 기본형 / 헤더리스(`header=None`) 컬럼명 → `"col_0".."col_n"` / 중복 컬럼명 → `이름`, `이름.1` 접미사로 유일화.

### Task 9.1: TableResult / WorkbookResult 모델 + 직렬화 헬퍼

**Files:**
- Create: `excel_inspector/results.py`
- Test: `tests/test_results.py`

- [ ] **Step 1: 실패하는 테스트 작성** — 스칼라 직렬화 규칙부터 고정

```python
# tests/test_results.py
"""Result layer (extract / TableResult / WorkbookResult) tests — Phase 9."""
from __future__ import annotations

import datetime as dt
import json
import math
from pathlib import Path

import pandas as pd
import pytest

from excel_inspector.results import (
    SCHEMA_VERSION,
    _dedupe_columns,
    _jsonify_scalar,
)


def test_jsonify_scalar_rules() -> None:
    assert _jsonify_scalar(None) is None
    assert _jsonify_scalar(float("nan")) is None
    assert _jsonify_scalar(pd.NA) is None
    assert _jsonify_scalar(pd.NaT) is None
    assert _jsonify_scalar(pd.Timestamp("2026-01-05")) == "2026-01-05T00:00:00"
    assert _jsonify_scalar(dt.date(2026, 1, 5)) == "2026-01-05"
    assert _jsonify_scalar("00123") == "00123"          # numeric_text 문자열 유지
    import numpy as np
    assert _jsonify_scalar(np.int64(7)) == 7            # numpy -> python int
    assert type(_jsonify_scalar(np.int64(7))) is int


def test_dedupe_columns() -> None:
    assert _dedupe_columns(["a", "b", "a", "a"]) == ["a", "b", "a.1", "a.2"]
```

- [ ] **Step 2: 실패 확인**

Run: `.venv/bin/python -m pytest tests/test_results.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'excel_inspector.results'`

- [ ] **Step 3: 최소 구현**

```python
# excel_inspector/results.py
"""Result layer: per-table DataFrames + JSON/Markdown serialization (plan v2 Phase 9).

This module consumes only the public contracts (WorkbookProfile, ReadPlan,
adapters.pandas_loader) — it must not reach into analyzer internals.
"""
from __future__ import annotations

import datetime as _dt
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from .models import ColumnProfile, ReadPlan, WorkbookProfile

SCHEMA_VERSION = "1.0"


def _jsonify_scalar(value: Any) -> Any:
    """Convert a pandas/numpy/datetime scalar to a JSON-compatible value.

    Rules (fixed contract): missing (NaN/NA/NaT/None) -> None; datetimes/dates
    -> ISO 8601 strings; numpy scalars -> native Python; everything else as-is.
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
    if hasattr(value, "item"):  # numpy scalar
        return value.item()
    return value


def _dedupe_columns(names: list[str]) -> list[str]:
    """Make column names unique with '.N' suffixes (JSON object keys must be unique)."""
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        if n in seen:
            seen[n] += 1
            out.append(f"{n}.{seen[n]}")
        else:
            seen[n] = 0
            out.append(n)
    return out
```

- [ ] **Step 4: 통과 확인**

Run: `.venv/bin/python -m pytest tests/test_results.py -v`
Expected: 2 PASS

- [ ] **Step 5: TableResult/WorkbookResult 실패 테스트 추가**

```python
# tests/test_results.py 에 추가
from excel_inspector.results import TableResult, WorkbookResult


def _toy_table() -> TableResult:
    df = pd.DataFrame({"지역": ["서울"], "수량": [50]})
    return TableResult(
        sheet_name="매출", table_id="매출!T1", dataframe=df,
        header_row=4, header_confidence=0.88, header_provenance="heuristic",
        columns=[], notes=[],
    )


def test_table_result_to_dict_shape() -> None:
    d = _toy_table().to_dict()
    assert d["table_id"] == "매출!T1"
    assert d["row_count"] == 1
    assert d["records"] == [{"지역": "서울", "수량": 50}]


def test_workbook_result_to_json_roundtrip() -> None:
    wr = WorkbookResult(file_path="/x.xlsx", sheets=[], warnings=[])
    parsed = json.loads(wr.to_json())
    assert parsed["schema_version"] == SCHEMA_VERSION
    assert parsed["file"] == "/x.xlsx"


def test_table_result_to_markdown_has_separator() -> None:
    md = _toy_table().to_markdown()
    lines = md.splitlines()
    assert lines[0].startswith("|") and "지역" in lines[0]
    assert set(lines[1].replace("|", "").replace(" ", "")) == {"-"}
```

- [ ] **Step 6: 모델/직렬화 본체 구현**

```python
# excel_inspector/results.py 에 추가
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

    def to_dict(self, max_rows: int | None = None) -> dict[str, Any]:
        df = self.dataframe if max_rows is None else self.dataframe.head(max_rows)
        records = [
            {col: _jsonify_scalar(v) for col, v in zip(df.columns, row)}
            for row in df.itertuples(index=False, name=None)
        ]
        return {
            "table_id": self.table_id,
            "header_row": self.header_row,
            "header_confidence": round(self.header_confidence, 4),
            "header_provenance": self.header_provenance,
            "columns": [
                {"index": c.index, "name": c.name,
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
        headers = [str(c) for c in df.columns]
        out = ["| " + " | ".join(headers) + " |",
               "| " + " | ".join("---" for _ in headers) + " |"]
        for row in df.itertuples(index=False, name=None):
            cells = ["" if _jsonify_scalar(v) is None else str(_jsonify_scalar(v)) for v in row]
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
```

- [ ] **Step 7: 통과 확인 + 전체 스위트**

Run: `.venv/bin/python -m pytest tests/test_results.py -v && .venv/bin/python -m pytest -q`
Expected: 신규 전부 PASS, 전체 `383+N passed`

### Task 9.2: `build_workbook_result()` + `extract()` 진입점

**Files:**
- Modify: `excel_inspector/results.py`(빌더 추가), `excel_inspector/__init__.py`(extract 공개)
- Test: `tests/test_results.py`

- [ ] **Step 1: 실패 테스트 — 실제 픽스처 end-to-end**

```python
# tests/test_results.py 에 추가 (conftest의 fixture_path fixture 사용)
from excel_inspector import extract


def test_extract_mixed_sheets(fixture_path) -> None:
    """mixed_sheets: 표 시트는 TableResult 1개, README는 skipped로 기록."""
    wr = extract(fixture_path("mixed_sheets"))
    assert [s.name for s in wr.sheets] == ["README", "Data"]
    readme, data = wr.sheets
    assert readme.skipped and readme.tables == [] and readme.skip_reason == "non-tabular"
    assert len(data.tables) == 1
    assert data.tables[0].table_id == "Data!T1"
    assert len(data.tables[0].dataframe) > 0


def test_extract_offset_plus_subtotals_records(fixture_path) -> None:
    """키스톤: 소계/합계가 records에 새지 않고 정확히 6행."""
    wr = extract(fixture_path("offset_plus_subtotals"))
    (table,) = wr.tables
    d = table.to_dict()
    assert d["row_count"] == 6
    assert sum(r["amount"] for r in d["records"]) == 590
    assert all("소계" not in str(r.values()) and "합계" not in str(r.values())
               for r in d["records"])


def test_extract_types_mixed_json_values(fixture_path) -> None:
    """numeric_text 선행 0 보존, date ISO 문자열, JSON 파싱 가능."""
    wr = extract(fixture_path("types_mixed"))
    parsed = json.loads(wr.to_json())
    recs = parsed["sheets"][0]["tables"][0]["records"]
    assert any(isinstance(r["code"], str) and r["code"].startswith("0") for r in recs)
    assert all(isinstance(r["date"], str) and r["date"][:4].isdigit() for r in recs)


def test_extract_headerless_override_col_names(fixture_path) -> None:
    """no_header + headerless override → 컬럼명 col_0..col_n."""
    from excel_inspector import InspectionOptions, SheetOverride
    opts = InspectionOptions(sheet_overrides={"Sheet1": SheetOverride(header_row=None)})
    wr = extract(fixture_path("no_header"), options=opts)
    (table,) = wr.tables
    assert list(table.dataframe.columns) == [f"col_{i}" for i in range(len(table.dataframe.columns))]
```

> 주의: `no_header` 픽스처의 시트명·`mixed_sheets`의 데이터 행 수는 `tests/fixtures/generate.py`의 FIXTURES 설명이 단일 출처다. 단언 값이 어긋나면 **픽스처 설명을 먼저 확인**하고 테스트를 맞춘다(픽스처를 바꾸지 말 것).

- [ ] **Step 2: 실패 확인**

Run: `.venv/bin/python -m pytest tests/test_results.py -k extract -v`
Expected: FAIL — `ImportError: cannot import name 'extract'`

- [ ] **Step 3: 빌더 + extract 구현**

```python
# excel_inspector/results.py 에 추가
from .adapters.pandas_loader import load_dataframe


def _postprocess_dataframe(df: pd.DataFrame, plan: ReadPlan) -> pd.DataFrame:
    """Apply result-layer column-name contracts: headerless naming + dedupe."""
    if plan.header is None:
        df = df.copy()
        df.columns = [f"col_{i}" for i in range(len(df.columns))]
        return df
    names = _dedupe_columns([str(c) for c in df.columns])
    if names != [str(c) for c in df.columns]:
        df = df.copy()
        df.columns = names
    return df


def build_workbook_result(
    file_path: str | Path, profile: WorkbookProfile
) -> WorkbookResult:
    """Translate an inspected WorkbookProfile into loaded per-table results."""
    sheets: list[SheetResultEntry] = []
    warnings: list[str] = list(profile.open_errors)
    for sp in profile.sheets:
        if not sp.is_tabular_candidate or sp.read_plan is None:
            sheets.append(SheetResultEntry(
                name=sp.name, is_visible=sp.is_visible, tables=[],
                skipped=True, skip_reason="non-tabular"))
            continue
        df = _postprocess_dataframe(load_dataframe(file_path, sp.read_plan), sp.read_plan)
        table = TableResult(
            sheet_name=sp.name, table_id=f"{sp.name}!T1", dataframe=df,
            header_row=sp.header_row, header_confidence=sp.header_confidence,
            header_provenance=sp.header_provenance,
            columns=list(sp.columns), notes=list(sp.read_plan.notes),
        )
        sheets.append(SheetResultEntry(
            name=sp.name, is_visible=sp.is_visible, tables=[table]))
    return WorkbookResult(file_path=str(file_path), sheets=sheets, warnings=warnings)
```

```python
# excel_inspector/__init__.py 에 추가 (기존 inspect 정의 아래)
from .results import TableResult, WorkbookResult, build_workbook_result


def extract(
    path: str | Path, options: InspectionOptions | None = None
) -> WorkbookResult:
    """One-call API: inspect the workbook, then load every table per its ReadPlan."""
    return build_workbook_result(path, inspect(path, options))
```

`__all__`에 `extract`, `TableResult`, `WorkbookResult` 추가.

- [ ] **Step 4: 통과 + 전체 스위트 green 확인**

Run: `.venv/bin/python -m pytest -q`
Expected: `383+N passed` (기존 무파손)

- [ ] **Step 5: 결정성 확인** — `extract()` 2회 호출의 `to_json()` 문자열 동일 단언 테스트 추가 후 전체 스위트 5회 반복

```python
def test_extract_json_is_deterministic(fixture_path) -> None:
    p = fixture_path("offset_plus_subtotals")
    assert extract(p).to_json() == extract(p).to_json()
```

Run: `for i in 1 2 3 4 5; do .venv/bin/python -m pytest -q | tail -1; done`
Expected: 5회 모두 동일 카운트 passed

**Phase 9 완료 기준**: `extract()` 한 호출로 모든 tabular 시트의 표가 `TableResult`로 나오고, `to_json()`이 위 스키마 v1.0을 결정적으로 산출하며, 마크다운 출력이 가능하다. 기존 383개 무파손.

**Phase 9 리뷰 체크리스트(단일 적대 리뷰 1회, 반례 명시 — 사전 평가에서 식별):**
- `_dedupe_columns(["a", "a.1", "a"])` → 충돌 없이 전부 유일한가(나이브 구현은 `a.1` 중복 → records dict가 컬럼 값을 조용히 덮어씀).
- 셀 값에 `|`/개행 포함 시 `to_markdown` 표 구조가 깨지지 않는가(이스케이프 또는 치환).
- `bytes`/`timedelta`/`Decimal` 스칼라에 대한 `_jsonify_scalar` 정책(명시적 str 폴백 또는 명시적 에러 — 조용한 `json.dumps` 크래시 금지).
- `to_json()` 5회 반복 문자열 동일(결정성).
- (P13 합류 시) `_leading_label` 좌표를 좌측 여백+소계 **변형 픽스처로 실측** — 현 픽스처는 좌측 여백이 빈 셀뿐이라 A열 기준 구현도 green인 함정.

---

## 4. Phase 10 — 한 시트 다중 표 (L)

### 4.0 설계

**핵심 아이디어**: 시트의 행을 한 번 스캔해 **연속 빈 행 `BLANK_RUN`(=2) 이상으로 분리된 "행 밴드"** 로 나누고, 각 밴드를 독립된 미니 시트처럼 Header→Boundary→Type 분석한다. 헤더 점수가 임계 미달이고 데이터도 없는 밴드(제목/각주)는 표가 아니라고 판정하고 경고만 남긴다.

```
multi_table.xlsx:  rows 1-4 (표1) | rows 5-6 빈 줄×2 | rows 7-10 (표2)
  → 밴드 [1..4], [7..10] → TableBlock 2개 → "다중표!T1", "다중표!T2"
messy_sales.xlsx:  rows 1-13 (제목+표, 내부 빈 줄은 1개라 분리 안 됨) | 14-15 빈 줄×2 | row 16 각주
  → 밴드 [1..13], [16..16] → 표 1개 + "각주 밴드" 경고 0건(표 아님이 자명) 
```

**모델** (`models.py`에 추가):

```python
@dataclass
class TableBlock:
    """One detected table block inside a sheet (1-based inspection coordinates)."""
    block_index: int                   # 0-based, top-down order
    band_start_row: int                # 밴드 경계(1-based, 포함)
    band_end_row: int
    header_row: int | None
    header_confidence: float
    header_provenance: str
    data_start_row: int | None
    data_end_row: int | None
    data_left_col: int | None
    data_right_col: int | None
    skip_rows: list[int]
    columns: list[ColumnProfile]
    read_plan: ReadPlan | None
```

`SheetProfile.blocks: list[TableBlock]` 추가. **하위 호환**: 기존 평면 필드(`header_row` 등)는 **최상위(top-most) 블록**의 미러로 유지한다 — v1의 "신뢰도 최고 선택"을 "최상위 선택"으로 바꾸는 동작 변경이며, 스펙 §10의 원래 의도를 복원한다(어차피 모든 블록이 추출되므로 정보 손실 없음). 픽스처 코퍼스는 전부 단일 표라 기존 383개에 영향 없다.

**ReadPlan 산출**: 좌표 변환은 기존 `aggregator`의 함수를 블록 필드로 호출하면 된다 — `skiprows`는 절대 0-based이므로 "헤더 위 모든 행 흡수" 규칙(1..header_row-1)이 블록 위치와 무관하게 그대로 성립한다([D1] 불변). `nrows = data_end - data_start + 1` 역시 동일.

### Task 10.1 (= Phase 10a): 밴드 분할기 + 다중 밴드 경고(안전장치)

**Files:**
- Create: `excel_inspector/analyzers/block_segmenter.py`, `tests/test_block_segmenter.py`
- Modify: `tests/fixtures/generate.py`(픽스처 `multi_table_stacked.xlsx` 추가 — 위 multi_table 데모와 동일 구조: 표1 rows1-4 / 빈 2행 / 표2 rows7-10, FIXTURES에 1-based 좌표 문서화)

- [ ] **Step 1: 실패 테스트**

```python
# tests/test_block_segmenter.py
from excel_inspector.analyzers.block_segmenter import RowBand, split_row_bands


def test_split_two_stacked_tables() -> None:
    rows = [("부서", "인원"), ("영업", 12), ("개발", 20), ("관리", 5),
            (None, None), (None, None),
            ("제품명", "단가"), ("키보드", 30000), ("마우스", 15000)]
    bands = split_row_bands(rows)
    assert bands == [RowBand(1, 4), RowBand(7, 9)]


def test_single_blank_does_not_split() -> None:
    rows = [("h",), ("a",), (None,), ("b",)]
    assert split_row_bands(rows) == [RowBand(1, 4)]
```

- [ ] **Step 2: 구현** — `split_row_bands(rows, blank_run=BLANK_RUN) -> list[RowBand]`; 행 전체가 None이면 빈 행, `blank_run` 이상 연속 시 밴드 경계. 입력 `rows`는 BoundaryDetector가 쓰는 것과 동일한 data 모드 행 튜플(1-based 정렬: `rows[r-1]` = r행).

- [ ] **Step 3: 경고 배선** — 파이프라인(`inspect`)에서 SheetEnumerator 직후 밴드를 계산해 `context`에 저장. 밴드가 2개 이상이고 두 번째 이후 밴드에 헤더 점수 임계 이상 행이 있으면:
  `warnings.append(f"sheet '{name}': additional table block suspected at rows {b.start_row}-{b.end_row}")`
  (Task 10.2 완료 후 이 경고는 "추출됨" 정보로 바뀐다 — 경고 문구를 테스트로 고정하지 말고 존재만 단언.)

- [ ] **Step 4: 픽스처 추가 + 경고 단언 테스트** — `multi_table_stacked.xlsx`에 대해 `extract()` 결과 `warnings`가 비어있지 않음 단언. 전체 스위트 green.

### Task 10.2 (= Phase 10b): 블록별 완전 추출

**Files:**
- Modify: `excel_inspector/models.py`(TableBlock), `analyzers/header_locator.py`·`boundary_detector.py`·`type_profiler.py`(행 윈도 스코프), `aggregator.py`(블록별 ReadPlan), `results.py`(블록당 TableResult), `excel_inspector/pipeline.py` 또는 `__init__.py`(블록 루프)
- Test: `tests/test_multi_table.py`

- [ ] **Step 1: 행 윈도 리팩터** — 세 분석기의 "시트 전체" 가정을 `row_window: tuple[int, int]`(1-based 포함) 파라미터로 일반화한다. 기본값 = 시트 전체(기존 동작 보존). 헤더 스캔은 `min(window_start+HEADER_SCAN_ROWS-1, window_end)`까지. **기존 테스트가 전부 green인 상태를 유지하면서 리팩터를 먼저 끝낸다(동작 불변 단계).**
- [ ] **Step 2: 블록 루프** — 밴드마다 Header→Boundary→Type을 실행해 `TableBlock`을 만들고, 헤더 신뢰도 임계 미달 + 데이터 미해소 밴드는 표가 아님(경고만). `SheetProfile.blocks`에 top-down 순서로 적재, 평면 필드는 `blocks[0]` 미러.
- [ ] **Step 3: 블록별 ReadPlan** — aggregator가 각 블록에 대해 기존 좌표 변환을 수행해 `block.read_plan`을 채운다. 미러 블록의 plan은 기존 `sheet.read_plan`과 동일해야 한다(호환성 단언).
- [ ] **Step 4: results 연동** — `build_workbook_result`가 `sp.blocks`를 순회해 `"{sheet}!T{n}"` ID로 TableResult를 N개 생성(블록이 없으면 기존 단일 경로 폴백).
- [ ] **Step 5: 골든 테스트** (`tests/test_multi_table.py`)

```python
def test_two_stacked_tables_both_extracted(fixture_path) -> None:
    wr = extract(fixture_path("multi_table_stacked"))
    ids = [t.table_id for t in wr.tables]
    assert len(ids) == 2 and ids[0].endswith("!T1") and ids[1].endswith("!T2")
    t1, t2 = wr.tables
    assert list(t1.dataframe.columns) == ["부서", "인원", "예산"]
    assert len(t1.dataframe) == 3            # 조용한 누락 제거 — 표1 복원!
    assert list(t2.dataframe.columns) == ["제품명", "단가", "재고", "비고"]
    assert len(t2.dataframe) == 3
```

- [ ] **Step 6: 전체 스위트 5회 반복 green** (기존 383+ 무파손 — 특히 `offset_plus_subtotals`/`blank_run_terminates` 골든이 밴드 도입 후에도 동일해야 함)

**Task 10.2 잠복 버그 가드(사전 평가에서 식별된 green-잠복형 — 구현·리뷰 시 정조준할 것):**
1. **점수 분모 희석**: 헤더 점수/밀도의 분모(`_col_count`)가 시트 전역 `max_col`이면 폭이 다른 적층 표에서 좁은 표의 점수가 희석되어 "표 아님" 오판 → 조용한 누락 재도입. **분모는 밴드 내 사용 열 수** 기준으로. 폭 불균등 적층 픽스처(3열 표 + 8열 표)로 단언.
2. **좌표 오프셋**: `header_row = best_index + 1`(header_locator.py의 시작행 1 하드코딩)을 `window_start + best_index`로 바꿀 때 off-by-one. 윈도 클램프는 `min(window_start + HEADER_SCAN_ROWS - 1, window_end)`.
3. **블록 간 상태 누출**: boundary_detector의 unreliable-span 폴백이 공유 profile 필드를 직접 변이하지 않도록 블록 로컬 상태로 격리.
4. **블록별 override 의미**: `SheetOverride.header_row`/`skip_rows_add`는 **시트 절대 1-based 좌표**로 해석해 해당 행을 포함하는 블록에만 적용한다(규약 명문화 + 블록 2 대상 override 테스트).
5. **블록 2의 pandas 계약 실측**: "헤더 위 전부 흡수 규칙이 블록 위치와 무관하게 성립"은 가정이다 — 아래쪽 블록의 ReadPlan을 실제 `read_excel` 왕복으로 골든 고정(미실측 가정 금지, v1 nrows 사고 교훈).
6. **warnings 결정성**: 경고 축적 순서를 시트 순서 → 블록 top-down으로 고정(JSON `warnings` 배열 결정성).
7. **미러 명확화**: `blocks`에는 표로 판정된 블록만 들어가므로 미러는 항상 "최상위 표"다 — 제목/각주 밴드는 blocks에 넣지 않는다.

**Phase 10 완료 기준**: 수직 적층 다중 표가 각각의 TableResult로 추출되고(JSON `tables` 배열에 N개), 표가 아닌 밴드는 경고로 가시화된다. **조용한 데이터 손실 0.** 비목표: 가로 나란한(side-by-side) 표 — §11에 기존대로 유지.

---

## 5. Phase 11 — 병합·다단 헤더 (M)

### Task 11.1 (= 11a): 병합 헤더 경계 브리지

현재 `merged_header.xlsx`는 헤더 병합(A1:B1)이 헤더 행의 연속 열 구간을 1칸으로 좁혀 경계가 미해소된다(`data_start=None` + 경고, characterization 테스트로 고정됨).

**Files:** Modify: `analyzers/merge_analyzer.py`, `analyzers/boundary_detector.py`, `excel_inspector/__init__.py`(파이프라인 순서); Test: `tests/test_merge_bridge.py` + 기존 characterization 테스트 갱신

- [ ] **Step 1: MergeAnalyzer 분리** — 병합 **수집**(구조 모드, 분류 없음)을 BoundaryDetector **앞**으로 이동(`MergeScanner`), 헤더/본문 **분류**는 헤더 확정 후 기존 위치에서 수행. 수집은 시트당 1회(블록 공통).
- [ ] **Step 2: 가상 채움** — BoundaryDetector가 헤더 행 열 구간을 계산할 때, 헤더 행과 교차하는 병합 범위의 빈 셀을 "채워진 것"으로 간주(가상 채움)해 연속 구간을 복원한다.
- [ ] **Step 3: characterization 뒤집기** — 기존 "None/None + discarded pending merge analysis" 단언을 `data_start_row=2, data_end_row=5`(픽스처 문서값)로 교체. `extract()`로 본문 병합 forward-fill 노트가 notes에 들어있는지 단언.

> ⚠ **픽스처 모순 가드(사전 평가에서 적발)**: `generate.py`의 `merged_header` 문서값(`data_end_row=5`)과 밀도 규칙의 나이브 산출(가상 채움 후 7)이 **이미 모순**일 수 있다. **정답은 5다** — A6:A7 body-merge 데모 블록은 표 본문 외(저밀도 행)로 남아야 한다. 구현 산출이 7이면 **단언을 7로 굳히지 말고** 밀도/스킵 규칙과 픽스처 의도를 재검토하라. 잘못된 단언을 green으로 굳힌 v1 nrows 사고와 동형 위험이며, 리뷰 렌즈가 이 지점을 반드시 독립 실측한다.

- [ ] **Step 4: 전체 스위트 green.**

### Task 11.2 (= 11b): 다단 헤더 적재 (`header: list[int]`)

**Files:** Modify: `aggregator.py`, `adapters/pandas_loader.py`, `results.py`; Test: `tests/test_multi_level_load.py`

- [ ] **Step 0: pandas 3.0.3 실측 선행(가정 금지)** — 구현 전에 다음 4건을 스파이크로 실측해 결과를 테스트에 박는다: (a) MultiIndex에서 그룹 셀 미충전 시 `Unnamed:` 류 라벨이 어떻게 나오는지(평탄화 규칙에 반영), (b) [D5] 위치 문자열 dtype 키가 `header=list`에서도 위치 기준으로 유효한지, (c) `header=[0,1]`일 때 `nrows`가 헤더 행을 소비 예산에 포함하는지, (d) `_dedupe_columns`와 MultiIndex 평탄화의 상호작용.
- [ ] **Step 1**: `is_multi_level_header=True`이고 헤더 위 병합 행이 헤더 행과 **연속**일 때, aggregator가 `ReadPlan.header`를 0-based 연속 리스트(post-skip)로 산출. 비연속/판정 불가 시 단일 헤더 유지 + 경고(보수적).
- [ ] **Step 2**: 어댑터는 리스트 헤더를 그대로 `read_excel`에 전달(MultiIndex). `read_plan_to_kwargs` 수정 불필요 여부를 테스트로 확인.
- [ ] **Step 3**: results 계층에서 MultiIndex 컬럼을 `"상위 / 하위"`로 평탄화(JSON 키는 문자열이어야 함). `multi_level_header.xlsx` 골든: 평탄화된 컬럼명과 데이터 값 단언.
- [ ] **Step 4**: 전체 스위트 green. dtype_map 키([D5] 위치 문자열)가 MultiIndex에서도 위치 기준으로 유효함을 단언.

---

## 6. Phase 12 — 수식 탐지 (S, v1 플랜 Phase 9 승계)

**Files:** Create: `analyzers/formula_detector.py`, `tests/test_formula_detector.py`; Modify: `loader.py`(수식 모드 활성화), `tests/fixtures/generate.py`(`formulas.xlsx`: D열 `=B{r}*C{r}` 수식, openpyxl로 생성 — 캐시값 없음이 정상)

- [ ] **Step 1**: `loader.formula_workbook()` — `read_only=True, data_only=False` 별도 인스턴스, lazy 개방(수식 탐지가 실행될 때만), close 정리 동일 보장.
- [ ] **Step 2**: FormulaDetector — 데이터 구간 표본 행에서 셀 값이 `str`이고 `"="`로 시작하면 해당 컬럼 `has_formula=True`. 캐시값(데이터 모드)이 전부 None인 수식 컬럼은 경고 `"column N: formula cache empty (file never opened in Excel?)"` + `read_hint='as_formula'`, 캐시가 있으면 `'as_value'`.
- [ ] **Step 3**: aggregator/notes — `as_formula` 컬럼은 dtype 추론을 건너뛰고 `ReadPlan.notes`에 권고 기록.
- [ ] **Step 4**: 골든 — `formulas.xlsx`에서 `has_formula=True`, 캐시 공백 경고, 전체 스위트 green. 수식 없는 기존 픽스처에서 수식 모드 워크북이 **아예 열리지 않음**(lazy)을 단언해 성능 회귀 방지.

**Phase 12 리뷰 프로브(사전 평가 식별 — 리뷰어에게 명시 지시할 것, 셋 다 green-잠복형):**
1. `loader.close()`의 핸들 튜플에 `_formula_wb`가 포함됐는가 — 기존 weakref 누수 테스트를 formula 핸들로 확장(빠뜨려도 기존 7개 close 테스트는 green인 함정).
2. **캐시값을 보유한** xlsx를 실측 제작(openpyxl은 캐시를 쓰지 않으므로 XML 수작업 또는 별도 도구)해 `as_value` 분기를 실제로 실행 — 안 하면 이 분기는 영원히 사문(死文)인 채 green.
3. 수식 없는 픽스처 전체에서 formula 워크북 미개방(lazy) 단언이 실제로 의미 있는 방식(개방 카운터/mock)으로 검증되는가.

---

## 7. Phase 13 — CLI + 잔여 마감 (S)

**Files:** Create: `excel_inspector/__main__.py`, `tests/test_cli.py`; Modify: `analyzers/type_profiler.py`(헤더리스), `analyzers/boundary_detector.py`(키워드 left_col), `README.md`

- [ ] **Step 1: CLI**

```python
# excel_inspector/__main__.py
"""CLI: python -m excel_inspector <file.xlsx> [--format json|markdown] [--max-rows N]"""
from __future__ import annotations

import argparse
import sys

from . import extract


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="excel_inspector")
    parser.add_argument("path")
    parser.add_argument("--format", choices=["markdown", "json"], default="markdown")
    parser.add_argument("--max-rows", type=int, default=20)
    args = parser.parse_args(argv)
    result = extract(args.path)
    if args.format == "json":
        print(result.to_json(indent=2, max_rows=None))
    else:
        print(result.to_markdown(max_rows=args.max_rows))
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

테스트(`tests/test_cli.py`): `main([str(path), "--format", "json"])`을 capsys로 잡아 `json.loads` 성공 + `schema_version` 단언; markdown 모드에서 `| --- |` 구분선 단언; 손상 파일 → 비정상 종료 코드 또는 명시 에러 메시지.

- [ ] **Step 2: 헤더리스 dtype 가시화(L6)** — headerless 시트에서 TypeProfiler가 위치 기반으로 타입을 추론하게 하거나, 최소한 `ReadPlan.notes`에 `"headerless sheet: dtype inference skipped"` 기록. 테스트로 고정.
- [ ] **Step 3: 키워드 선두 라벨 좌표(L7)** — `_leading_label`이 시트 A열이 아닌 표의 `data_left_col`부터 첫 비어있지 않은 셀을 읽도록 수정. `left_margin_cols`에 소계 행을 추가한 변형 픽스처로 회귀 단언.
- [ ] **Step 4: 대용량 성능 스모크(v1 openIssue 승계)** — `generate.py`에 100k행 빌더 추가(기본 코퍼스에서 제외, `@pytest.mark.slow` 마커). `tracemalloc`으로 `inspect()` 피크 메모리 ≤ 200MB(스펙 §8) 단언. `pyproject.toml`에 `markers = ["slow"]` 등록, 기본 실행에서 `-m "not slow"`는 강제하지 않음(전체 실행 시간 허용 범위 확인 후 결정).
- [ ] **Step 5: README 갱신** — `extract()`/CLI/JSON 스키마 사용 예시. 전체 스위트 최종 5회 반복 green.

---

## 8. 테스트 전략 (v1 원칙 승계 + 추가)

- **픽스처 단일 출처 유지**: 신규 픽스처(`multi_table_stacked`, `formulas`, 변형들)는 전부 `generate.py` 빌더로 결정적 생성(타임스탬프 핀 규약 준수), FIXTURES dict에 1-based 좌표 문서화.
- **골든은 부분 단언 우선**(취약성 완화): records 전체 비교 대신 행 수·합계·경계 컬럼명 위주.
- **JSON 결정성**: 동일 입력 2회 `to_json()` 문자열 동일을 모든 신규 픽스처에 파라미터화.
- **mtime 단언 금지**(v1 교훈): 읽기전용 검증은 SHA-256만.
- **호환성 가드**: Phase 10 미러 규칙(평면 필드 = blocks[0]) 덕분에 기존 383개가 무수정 통과해야 한다 — 깨지면 미러 구현 버그로 간주.

## 9. 위험과 완화

| 위험 | 영향 | 완화 |
| --- | --- | --- |
| 밴드 분할 오탐(희소 표 내부의 빈 2행) | 중 | v1도 BLANK_RUN=2에서 데이터를 종료시키므로 동작 일관 — 분할은 기존 종료 규칙의 일반화. 희소 표 픽스처로 경계 고정 |
| 평면 필드 미러 변경(최고점→최상위) | 중 | 코퍼스가 전부 단일표라 기존 테스트 영향 0. 스펙 §10 의도 복원임을 문서화. 다중표에서는 어차피 전 블록 추출 |
| 행 윈도 리팩터 회귀 | 고 | "동작 불변 리팩터 먼저, 블록 루프는 그다음" 2단계 분리 + 기존 383 green 게이트 |
| MultiIndex/JSON 키 충돌 | 저 | 평탄화 규칙(`"상위 / 하위"`) + `_dedupe_columns` 재사용 |
| 수식 모드 추가 개방 비용 | 저 | lazy 개방 + "수식 없으면 안 연다" 단언 테스트 |
| JSON 대용량(수만 행 records) | 저 | `max_rows` 파라미터 제공(기본 전체). 스트리밍 직렬화는 비목표로 명시 |

## 10. 완료 정의 (v2 DoD)

- `extract()` 한 호출로 워크북의 **모든 표가 각각의 TableResult**(DataFrame)로 나오고, `to_json()`(스키마 v1.0)·`to_markdown()`이 결정적으로 동작한다.
- 한 시트 다중 표(수직 적층)가 전부 추출된다 — **조용한 누락 0** (표가 아닌 밴드는 경고로 가시화).
- `merged_header`가 경계 해소되고, 다단 헤더가 MultiIndex로 적재·평탄화된다.
- 수식 컬럼이 `has_formula`/`read_hint`로 표시되고 캐시 공백이 경고된다.
- CLI(`python -m excel_inspector`)로 어떤 `.xlsx`든 표/JSON을 즉시 확인할 수 있다.
- 전체 스위트가 기존 383개 무파손 + 신규 테스트 포함 green, 5회 반복 flaky 0.

## 11. 비목표 (v2에서도 하지 않는 것)

- 가로 나란한(side-by-side) 표 분할 — 스펙 §11 잔류.
- `.xls`/CSV 입력, 검사 결과 캐싱 — 스펙 §11 잔류.
- JSON 스트리밍 직렬화(행 단위) — `max_rows`로 충분.

## 12. 실행 순서 요약

| 순서 | Phase | 규모 | 산출물 핵심 |
| --- | --- | --- | --- |
| 1 | **9** 결과 계층 | M | `extract()`, TableResult/WorkbookResult, JSON v1.0, 마크다운 |
| 2 | **10a** 밴드 감지 | S | 다중 블록 경고(조용한 누락 제거) |
| 3 | **10b** 다중 표 추출 | L | `SheetProfile.blocks`, 표마다 ReadPlan/TableResult |
| 4 | **11a** 병합 브리지 | S | `merged_header` 경계 해소 |
| 5 | **11b** 다단 헤더 | M | `header: list[int]` → MultiIndex → 평탄화 |
| 6 | **12** 수식 탐지 | S | `has_formula`, `as_formula` 권고, 캐시 공백 경고 |
| 7 | **13** CLI·마감 | S | `python -m excel_inspector`, 헤더리스/키워드 마감 |
