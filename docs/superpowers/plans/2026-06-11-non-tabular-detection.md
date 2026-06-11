# 비표(non-tabular) 판정 내용 인식형 휴리스틱 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `sheet_enumerator`의 `max_col > 1` 단일 게이트를, 상단 샘플 기반 "채워진 열 개수 + 밀도" 판정으로 교체해 텍스트 시작 열 위치와 무관하게 표지/안내 시트를 일관되게 비표로 분류한다 (이슈 #3).

**Architecture:** 비표 판정의 단일 권위 지점인 `_is_tabular_candidate`만 내용 인식형으로 교체한다. 하류 스테이지(MergeScanner/BlockSegmenter/HeaderLocator/PlanAggregator)는 이미 `is_tabular_candidate=False`를 존중하므로 추가 배선이 없다. 블록 단위 표 생성과 `[D2]` 오버라이드는 불변.

**Tech Stack:** Python 3.14, openpyxl 3.1.5, pandas 3.0.3, pytest. 작업 디렉터리: 워크트리 `/.worktrees/issue-3-non-tabular`. 인터프리터: 메인 venv `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python`.

> 모든 명령은 워크트리 루트(`.worktrees/issue-3-non-tabular`)에서 실행한다. 패키지는 미설치 상태이며 pytest는 루트 conftest로 `excel_inspector`를 import한다(메인 체크아웃과 동일).

---

## File Structure

- **Modify** `excel_inspector/heuristics.py` — 새 상수 3개 추가 (`[D4]`).
- **Modify** `excel_inspector/analyzers/sheet_enumerator.py` — `_is_tabular_candidate` 교체 + 헬퍼 `_is_non_empty`, `_sample_density`, `_dims_tabular`.
- **Modify** `tests/fixtures/generate.py` — fixture 2개(`cover_offset`, `cover_sparse`) 추가 (`FIXTURES` + builder + `BUILDERS`).
- **Modify** `tests/test_sheet_enumerator.py` — 규칙 단위 테스트 + 분류 회귀 파라미터 테스트.
- **Modify** `tests/test_results.py` — 이슈 #3 end-to-end 가드.

---

## Task 1: 휴리스틱 상수 추가

**Files:**
- Modify: `excel_inspector/heuristics.py`

- [ ] **Step 1: 상수 추가**

`excel_inspector/heuristics.py`에서 헤더 관련 상수 블록(`HEADER_CONFIDENCE_THRESHOLD` 정의 직후, 라인 31 부근) 아래에 다음을 추가한다:

```python
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
```

- [ ] **Step 2: import 가능 확인**

Run: `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python -c "from excel_inspector.heuristics import NON_TABULAR_SAMPLE_ROWS, MIN_TABULAR_POPULATED_COLS, NON_TABULAR_DENSITY_THRESHOLD; print(NON_TABULAR_SAMPLE_ROWS, MIN_TABULAR_POPULATED_COLS, NON_TABULAR_DENSITY_THRESHOLD)"`
Expected: `20 1 0.5`

- [ ] **Step 3: Commit**

```bash
git add excel_inspector/heuristics.py
git commit -m "feat: 비표 판정용 휴리스틱 상수 추가 (#3)"
```

---

## Task 2: 신규 fixture 2개 추가

**Files:**
- Modify: `tests/fixtures/generate.py`

- [ ] **Step 1: `FIXTURES` 메타데이터 추가**

`tests/fixtures/generate.py`의 `FIXTURES` 딕셔너리에서 `"mixed_sheets"` 항목 바로 뒤(라인 234 이후)에 두 항목을 추가한다:

```python
    "cover_offset": FixtureSpec(
        "cover_offset.xlsx",
        True,
        "Single-sheet cover '표지' whose sparse text starts in column B "
        "(B2/B4/B6). max_col=2 fooled the legacy max_col>1 gate, but only one "
        "column is populated -> is_tabular_candidate expected False (issue #3).",
    ),
    "cover_sparse": FixtureSpec(
        "cover_sparse.xlsx",
        True,
        "Single-sheet cover '표지' with scattered cells across 3 columns "
        "(B2/E4/C6): populated_cols=3, populated_rows=3, filled=3, "
        "density=0.333 -> is_tabular_candidate expected False via the density "
        "rule (issue #3).",
    ),
    "sparse_real_table": FixtureSpec(
        "sparse_real_table.xlsx",
        True,
        "Genuine 4-column table with many missing cells: header row + 5 data "
        "rows, populated_cols=4, populated_rows=6, filled=14, density=0.583 -> "
        "ABOVE NON_TABULAR_DENSITY_THRESHOLD (0.5) so is_tabular_candidate "
        "expected True. Pins the density-rule margin so the threshold cannot "
        "be raised past 0.583 without a red test (issue #3).",
    ),
```

- [ ] **Step 2: builder 함수 추가**

`build_mixed_sheets` 정의 바로 뒤(라인 851 이후)에 추가한다:

```python
def build_cover_offset() -> bytes:
    """Cover sheet whose sparse text starts in column B (issue #3).

    Reproduces the original bug: text offset to column B leaves max_col=2, so
    the legacy ``max_col > 1`` gate misclassified it as tabular. Content-wise it
    is a single populated column -> expected non-tabular.
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "표지"
    ws["B2"] = "2026년 영업 보고서"
    ws["B4"] = "작성: 영업팀"
    ws["B6"] = "기밀 — 외부 배포 금지"
    return _save_bytes(wb)


def build_cover_sparse() -> bytes:
    """Multi-column but scattered cover sheet (issue #3, density rule).

    Three cells in three different columns and rows -> populated_cols=3,
    populated_rows=3, filled=3, density=3/9=0.333. Above MIN_TABULAR_POPULATED_
    COLS but below NON_TABULAR_DENSITY_THRESHOLD -> expected non-tabular.
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "표지"
    ws["B2"] = "분기 실적 요약"
    ws["E4"] = "2026-03-31"
    ws["C6"] = "재무팀"
    return _save_bytes(wb)


def build_sparse_real_table() -> bytes:
    """Genuine 4-column table with many missing cells (issue #3, density rule).

    Header + 5 data rows over columns A-D, with scattered blanks so
    populated_cols=4, populated_rows=6, filled=14, density=14/24=0.583. Above
    NON_TABULAR_DENSITY_THRESHOLD (0.5) -> expected tabular. Pins the density
    margin: raising the threshold past 0.583 would wrongly skip this real table.
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    _write_rows(
        ws,
        1,
        [
            ["품목", "1월", "2월", "3월"],
            ["연필", 100, None, None],
            ["지우개", None, 50, None],
            ["공책", None, None, 210],
            ["펜", None, None, 90],
            ["자", 30, None, None],
        ],
    )
    return _save_bytes(wb)
```

- [ ] **Step 3: `BUILDERS` 등록**

`BUILDERS` 딕셔너리에서 `"mixed_sheets": build_mixed_sheets,` 줄 뒤(라인 1332 이후)에 추가한다:

```python
    "cover_offset": build_cover_offset,
    "cover_sparse": build_cover_sparse,
    "sparse_real_table": build_sparse_real_table,
```

- [ ] **Step 4: 코퍼스 재생성 + 지표 검증**

Run:
```bash
/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python - <<'PY'
import sys; sys.path.insert(0, "tests/fixtures")
import generate, openpyxl
paths = generate.generate_all("tests/fixtures")
for fid in ("cover_offset", "cover_sparse", "sparse_real_table"):
    sheet = "Sheet1" if fid == "sparse_real_table" else "표지"
    wb = openpyxl.load_workbook(paths[fid], read_only=True, data_only=True)
    ws = wb[sheet]
    cols=set(); prows=0; filled=0
    for r in ws.iter_rows(min_row=1, max_row=20, values_only=True):
        rf=False
        for ci,v in enumerate(r):
            if v is not None and not (isinstance(v,str) and v==""):
                cols.add(ci); filled+=1; rf=True
        prows += 1 if rf else 0
    wb.close()
    pc=len(cols); dens=filled/(pc*prows) if pc and prows else 0.0
    print(fid, "pop_cols", pc, "pop_rows", prows, "filled", filled, "density", round(dens,3), "max_col", openpyxl.load_workbook(paths[fid]).active.max_column)
PY
```
Expected:
```
cover_offset pop_cols 1 pop_rows 3 filled 3 density 1.0 max_col 2
cover_sparse pop_cols 3 pop_rows 3 filled 3 density 0.333 max_col 5
sparse_real_table pop_cols 4 pop_rows 6 filled 14 density 0.583 max_col 4
```

- [ ] **Step 5: Commit**

```bash
git add tests/fixtures/generate.py tests/fixtures/cover_offset.xlsx tests/fixtures/cover_sparse.xlsx tests/fixtures/sparse_real_table.xlsx
git commit -m "test: 비표 회귀 fixture cover_offset/cover_sparse/sparse_real_table 추가 (#3)"
```

---

## Task 3: 내용 인식형 `_is_tabular_candidate` (TDD)

**Files:**
- Test: `tests/test_sheet_enumerator.py`
- Modify: `excel_inspector/analyzers/sheet_enumerator.py:126-151`

- [ ] **Step 1: 실패 테스트 작성**

`tests/test_sheet_enumerator.py` 끝에 추가한다. (파일 맨 위 import 블록에 `import pytest`와 `from excel_inspector.exceptions import CorruptWorkbookError`를 추가하고, `make_context`는 이미 import되어 있다.)

```python
import pytest

from excel_inspector.exceptions import CorruptWorkbookError


class _FakeWorksheet:
    """Minimal read_only-style worksheet yielding canned rows for sampling."""

    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows

    def iter_rows(self, min_row=1, max_row=None, values_only=False):
        end = len(self._rows) if max_row is None else max_row
        for row in self._rows[min_row - 1 : end]:
            yield tuple(row)


class _FakeWorkbook:
    def __init__(self, sheets: dict[str, _FakeWorksheet]) -> None:
        self._sheets = sheets

    def __getitem__(self, name: str) -> _FakeWorksheet:
        return self._sheets[name]


class _FakeLoader:
    def __init__(self, sheets: dict[str, _FakeWorksheet]) -> None:
        self._wb = _FakeWorkbook(sheets)

    def data_workbook(self) -> _FakeWorkbook:
        return self._wb


class _BoomLoader:
    def data_workbook(self):
        raise RuntimeError("boom")


class _CorruptLoader:
    def data_workbook(self):
        raise CorruptWorkbookError("boom")


def test_offset_cover_is_non_tabular(fixture_path) -> None:
    """cover_offset: text in column B is one populated column -> non-tabular."""

    ctx = _run_on(fixture_path("cover_offset"))
    s = ctx.workbook_profile.sheets[0]
    assert s.max_col == 2  # the offset that fooled the legacy max_col>1 gate
    assert s.is_tabular_candidate is False
    assert s.is_tabular_provenance == "heuristic"


def test_sparse_cover_is_non_tabular(fixture_path) -> None:
    """cover_sparse: 3 scattered columns, density 0.333 < 0.5 -> non-tabular."""

    ctx = _run_on(fixture_path("cover_sparse"))
    s = ctx.workbook_profile.sheets[0]
    assert s.max_col == 5
    assert s.is_tabular_candidate is False


@pytest.mark.parametrize(
    "fixture_id,sheet,expected",
    [
        ("stacked_uneven_width", "Sheet1", True),  # corpus-floor density (0.688)
        ("sparse_real_table", "Sheet1", True),  # sparse real table density 0.583 > 0.5
        ("no_header", "Sheet1", True),  # all-numeric real table (pop_cols=3)
        ("header_only", "Sheet1", True),  # header-only real table (pop_cols=3)
        ("hidden_sheet", "Hidden", True),  # 2-column table, high density
        ("mixed_sheets", "README", False),  # single sparse text column
        ("empty_sheet", "Sheet1", False),  # empty sheet
        ("cover_offset", "표지", False),  # issue #3: B-offset single column
        ("cover_sparse", "표지", False),  # multi-column but sparse (density 0.333)
    ],
)
def test_tabular_classification(fixture_path, fixture_id, sheet, expected) -> None:
    """Whole-sheet tabular gate stays correct across the corpus (regression)."""

    ctx = _run_on(fixture_path(fixture_id))
    sheets = {s.name: s for s in ctx.workbook_profile.sheets}
    assert sheets[sheet].is_tabular_candidate is expected


def test_empty_sample_falls_back_to_dimensions() -> None:
    """No content in the sample -> defer to the legacy max_col>1 dims rule."""

    loader = _FakeLoader({"S": _FakeWorksheet([])})
    ctx = make_context(loader=loader)
    result, prov = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=10, max_col=3
    )
    assert result is True and prov == "heuristic"


def test_sampling_failure_falls_back_with_warning() -> None:
    """A sampling exception must not break enumeration (spec §6): fall back."""

    ctx = make_context(loader=_BoomLoader())
    result, prov = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=10, max_col=1
    )
    assert result is False and prov == "heuristic"  # dims fallback: 1 not > 1
    assert any("sheet_enumerator" in w and "S" in w for w in ctx.warnings)


def test_inspector_error_propagates_not_swallowed() -> None:
    """Loader domain errors (corrupt/encrypted) must propagate, not be absorbed
    into a warning + dims fallback (spec §6/§9; consistent with pipeline.py)."""

    ctx = make_context(loader=_CorruptLoader())
    with pytest.raises(CorruptWorkbookError):
        SheetEnumerator()._is_tabular_candidate(ctx, "S", max_row=10, max_col=2)
    assert ctx.warnings == []  # not downgraded to a warning


def test_density_rule_counts_only_content_rows() -> None:
    """density = filled/(populated_cols*populated_rows); a blank middle row is
    NOT counted in populated_rows. Here pc=2, content_rows=2, filled=3 ->
    0.75 >= 0.5 -> tabular (pins the populated_rows denominator term)."""

    rows = [["a", "b"], [None, None], ["c", None]]
    loader = _FakeLoader({"S": _FakeWorksheet(rows)})
    ctx = make_context(loader=loader)
    result, prov = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=3, max_col=2
    )
    assert result is True and prov == "heuristic"


def test_density_rule_low_density_is_non_tabular() -> None:
    """pc=3, content_rows=3, filled=3 -> 3/9=0.333 < 0.5 -> non-tabular
    (exercises the density branch directly via the fake loader)."""

    rows = [["a", None, None], [None, None, "b"], [None, "c", None]]
    loader = _FakeLoader({"S": _FakeWorksheet(rows)})
    ctx = make_context(loader=loader)
    result, _ = SheetEnumerator()._is_tabular_candidate(
        ctx, "S", max_row=3, max_col=3
    )
    assert result is False
```

- [ ] **Step 2: 실패 확인**

Run: `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python -m pytest tests/test_sheet_enumerator.py -q`
Expected: FAIL — 다음이 실패한다(현재 코드는 셀 내용을 안 보고 `max_col>1`만 보므로 cover_*가 tabular로 분류되고, 폴백/예외/density 분기가 없음):
`test_offset_cover_is_non_tabular`, `test_sparse_cover_is_non_tabular`, `test_tabular_classification[cover_offset...]`/`[cover_sparse...]`, `test_sampling_failure_falls_back_with_warning`(현재 코드는 loader를 안 건드려 warning 미기록), `test_inspector_error_propagates_not_swallowed`(현재 코드는 sampling 자체가 없어 raise 안 함), `test_density_rule_low_density_is_non_tabular`(현재 코드는 density 분기 없음 → max_col 3>1로 True 반환).

> 주의: `test_empty_sample_falls_back_to_dimensions`(입력 max_col=3)와 `test_density_rule_counts_only_content_rows`(max_col=2, 결과 True)는 레거시 `max_col>1` 규칙에서도 통과하므로 red가 아니라 **그대로 PASS**한다. 이들은 신 로직의 빈-샘플 폴백/density 경계 분기를 보증하는 green 가드다(회귀 시 red).

- [ ] **Step 3: 헬퍼 + 판정 교체 구현**

`excel_inspector/analyzers/sheet_enumerator.py`를 수정한다.

(a) import 블록(라인 20 부근, `from ..options import get_is_tabular_override` 뒤)에 추가:

```python
from ..exceptions import InspectorError
from ..heuristics import (
    MIN_TABULAR_POPULATED_COLS,
    NON_TABULAR_DENSITY_THRESHOLD,
    NON_TABULAR_SAMPLE_ROWS,
)
```

(b) 모듈 상수 `_MAX_NON_TABULAR_COLS = 1`(라인 28)은 그대로 둔다(dims 폴백에서 사용). 그 아래에 모듈 함수를 추가:

```python
def _is_non_empty(value: object) -> bool:
    """True when a sampled cell holds content (not None, not the empty string).

    Matches the Header Locator's notion of an empty cell so the tabular gate and
    header scoring agree on what counts as populated.
    """

    return value is not None and not (isinstance(value, str) and value == "")
```

(c) 라인 126-151의 `_is_tabular_candidate`를 통째로 아래로 교체한다:

```python
    def _is_tabular_candidate(
        self,
        context: InspectionContext,
        sheet_name: str,
        max_row: int,
        max_col: int,
    ) -> tuple[bool, str]:
        """Decide whether a sheet looks like a data table (spec §4.2) [D4].

        The ``is_tabular`` override [D2] wins outright (``provenance="manual"``).
        Otherwise the top :data:`~excel_inspector.heuristics.NON_TABULAR_SAMPLE_
        ROWS` rows are sampled in data mode and judged on *content*, not on the
        rightmost-column dimension (issue #3 — the legacy ``max_col`` gate was
        sensitive to which column the text started in):

        * an empty sample defers to the legacy dimension rule (data may begin
          below the window; a truly empty sheet stays non-tabular);
        * at most :data:`~excel_inspector.heuristics.MIN_TABULAR_POPULATED_COLS`
          populated columns -> non-tabular (a single-column cover, any offset);
        * >= 2 populated columns but sample density below
          :data:`~excel_inspector.heuristics.NON_TABULAR_DENSITY_THRESHOLD`
          -> non-tabular (a scattered multi-column cover).

        Robustness (spec §6): a loader domain error (:class:`InspectorError` —
        corrupt/encrypted) propagates so the pipeline aborts (consistent with
        ``pipeline.py``); any other sampling failure falls back to the legacy
        dimension rule with a warning so enumeration never breaks.

        Returns:
            ``(is_tabular_candidate, provenance)`` where provenance is
            ``"manual"`` for an override and ``"heuristic"`` otherwise [D2].
        """

        override = get_is_tabular_override(context.options, sheet_name)
        if override is not None:
            return override, "manual"

        try:
            populated_cols, populated_rows, filled = self._sample_density(
                context, sheet_name
            )
        except InspectorError:
            # Loader domain errors (corrupt/encrypted) are NOT absorbed: they
            # must abort the pipeline (spec §6/§9, like pipeline.py).
            raise
        except Exception as exc:  # noqa: BLE001 - robustness policy (spec §6)
            context.add_warning(
                f"sheet_enumerator: tabular sampling failed for sheet "
                f"{sheet_name!r} ({exc!r}); falling back to dimension heuristic"
            )
            return self._dims_tabular(max_row, max_col), "heuristic"

        if populated_cols == 0:
            return self._dims_tabular(max_row, max_col), "heuristic"
        if populated_cols <= MIN_TABULAR_POPULATED_COLS:
            return False, "heuristic"
        density = filled / (populated_cols * populated_rows)
        if density < NON_TABULAR_DENSITY_THRESHOLD:
            return False, "heuristic"
        return True, "heuristic"

    @staticmethod
    def _dims_tabular(max_row: int, max_col: int) -> bool:
        """Legacy dimension-only tabular rule (pre-issue-#3 fallback).

        A sheet with no usable area, or only a single populated column by
        dimension, is non-tabular; otherwise tabular.
        """

        if max_row < 1 or max_col < 1:
            return False
        return max_col > _MAX_NON_TABULAR_COLS

    def _sample_density(
        self, context: InspectionContext, sheet_name: str
    ) -> tuple[int, int, int]:
        """Sample the top rows in data mode and summarize populated content.

        Reads the top :data:`~excel_inspector.heuristics.NON_TABULAR_SAMPLE_ROWS`
        rows of ``sheet_name`` in data mode [D3] and returns
        ``(populated_cols, populated_rows, filled)`` where ``populated_cols`` is
        the number of distinct columns holding any non-empty cell,
        ``populated_rows`` the number of rows with any non-empty cell, and
        ``filled`` the total non-empty cell count.

        Raises:
            InspectorError: A loader domain error; the caller re-raises it.
            Exception: Any other sampling failure; the caller absorbs it into a
                warning and falls back (spec §6).
        """

        loader = context.loader
        if loader is None:
            raise RuntimeError("no loader available for tabular sampling")
        worksheet = loader.data_workbook()[sheet_name]

        populated_cols: set[int] = set()
        populated_rows = 0
        filled = 0
        for row in worksheet.iter_rows(
            min_row=1, max_row=NON_TABULAR_SAMPLE_ROWS, values_only=True
        ):
            row_has_content = False
            for col_index, value in enumerate(row):
                if _is_non_empty(value):
                    populated_cols.add(col_index)
                    filled += 1
                    row_has_content = True
            if row_has_content:
                populated_rows += 1
        return len(populated_cols), populated_rows, filled
```

- [ ] **Step 4: 통과 확인**

Run: `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python -m pytest tests/test_sheet_enumerator.py -q`
Expected: PASS (전체).

- [ ] **Step 5: Commit**

```bash
git add excel_inspector/analyzers/sheet_enumerator.py tests/test_sheet_enumerator.py
git commit -m "fix: 비표 판정을 채워진 열 개수+밀도 기반 내용 인식형으로 교체 (#3)"
```

---

## Task 4: 이슈 #3 end-to-end 가드 (TDD)

**Files:**
- Test: `tests/test_results.py`

> Task 3 이후 동작은 이미 고쳐졌으므로 이 테스트는 통과한다. Task 3 이전이라면 `extract`가 `표지`에 `columns=[]` 테이블을 만들어 실패했을 가드다(이슈 재현을 결과 레이어에 고정).

- [ ] **Step 1: end-to-end 테스트 작성**

`tests/test_results.py` 끝에 추가한다(파일은 이미 `from excel_inspector import extract`, `import json`, `fixture_path` 픽스처를 사용한다):

```python
def test_offset_cover_sheet_skipped_not_empty_table(fixture_path) -> None:
    """issue #3: B열에서 시작하는 표지는 columns=[] 빈 테이블이 아니라 skip."""

    wr = extract(fixture_path("cover_offset"))
    assert [s.name for s in wr.sheets] == ["표지"]
    cover = wr.sheets[0]
    assert cover.skipped is True
    assert cover.skip_reason == "non-tabular"
    assert cover.tables == []


def test_sparse_cover_sheet_skipped(fixture_path) -> None:
    """issue #3: 다중 열이지만 희소한 표지도 skip(non-tabular)으로 끝난다."""

    wr = extract(fixture_path("cover_sparse"))
    cover = wr.sheets[0]
    assert cover.skipped is True and cover.skip_reason == "non-tabular"
    assert cover.tables == []
```

- [ ] **Step 2: 통과 확인**

Run: `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python -m pytest tests/test_results.py -q`
Expected: PASS (신규 2개 포함 전체).

- [ ] **Step 3: Commit**

```bash
git add tests/test_results.py
git commit -m "test: 이슈 #3 표지 시트 skip end-to-end 가드 (#3)"
```

---

## Task 5: 전체 회귀 + 데모 파일 확인

**Files:** 없음 (검증 전용)

- [ ] **Step 1: 기본 스위트 전체 실행**

Run: `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python -m pytest -q`
Expected: 전부 PASS, 회귀 없음. (느린 100k perf smoke는 addopts로 기본 제외)

- [ ] **Step 2: 느린 perf smoke 확인(메모리/스트리밍 보장 불변)**

Run: `/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python -m pytest -m slow -q`
Expected: PASS (~14s, tracemalloc peak <= 200MB on 100k rows).

- [ ] **Step 3: 데모 파일 표지가 이제 skip 되는지 확인**

Run:
```bash
/Users/daniel/Documents/project/sk-ax/excel-parser/.venv/bin/python - <<'PY'
import json
from excel_inspector import extract
d = json.loads(extract("tests/fixtures/complex_demo.xlsx").to_json())
for sh in d["sheets"]:
    print(sh["name"], "skipped=", sh["skipped"], "reason=", sh["skip_reason"],
          "tables=", [(t["table_id"], len(t["columns"])) for t in sh["tables"]])
PY
```
Expected: `표지 skipped= True reason= non-tabular tables= []`, 나머지 시트는 기존대로 표 유지.

> `complex_demo.xlsx`는 빌더(`build_complex_demo.py`)로 만든 데모 파일이라 위 출력이 빌더가 만든 파일과 다르면 빌더로 재생성해야 할 수 있다. 검증 테스트는 없으므로 출력 확인만으로 충분하다.

- [ ] **Step 4: 최종 상태 확인**

Run: `git -C . status --short && git -C . log --oneline -6`
Expected: 워킹트리 클린, 커밋 5개(docs 스펙 + 상수 + fixtures + 휴리스틱 + e2e 가드)가 `issue-3-non-tabular` 브랜치에 존재.

---

## Self-Review (작성자 점검 결과)

- **스펙 커버리지:** §3.1(위치)→Task 3, §3.2(규칙 1~6)→Task 3 + 단위테스트(규칙 2/6) + 분류테스트(규칙 3/4/5), §3.3(상수)→Task 1, §5(테스트/fixture)→Task 2/3/4, §6(리스크)는 상수 보정으로 반영. 빠진 요구사항 없음.
- **플레이스홀더:** 없음(모든 코드/명령/기대출력 구체화).
- **타입/이름 일관성:** `_is_tabular_candidate`/`_sample_density`/`_dims_tabular`/`_is_non_empty`, 상수 `NON_TABULAR_SAMPLE_ROWS`/`MIN_TABULAR_POPULATED_COLS`/`NON_TABULAR_DENSITY_THRESHOLD` 가 Task 1·3·테스트 전반에서 동일하게 사용됨. `_is_tabular_candidate` 시그니처`(context, sheet_name, max_row, max_col)`는 기존과 동일(호출부 `_profile_sheet` 불변).
