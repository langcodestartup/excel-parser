# 블록 단위 오버라이드 채널 [D7] 구현 계획 (issue #9)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 적층(multi-block) 시트의 개별 밴드에 header_row를 강제하거나 headerless로 선언할 수 있는 `BlockOverride` 채널을 추가한다 (issue #9, 설계 문서 `docs/superpowers/specs/2026-06-12-block-override-design.md`).

**Architecture:** 앵커 행(대상 밴드 내 임의의 1-based 행)으로 키된 `SheetOverride.block_overrides: dict[int, BlockOverride]`를 추가하고, `options.resolve_block_overrides`가 앵커→밴드 해석·충돌 검출을 전담한다. `BlockAnalyzer`가 해석 결과를 밴드별로 적용하며(특이성 우선: 블록 > 시트 > 휴리스틱), headerless 블록은 밴드 전체를 데이터 구간으로 갖는 manual 블록이 된다. `PlanAggregator.build_read_plan`은 `declared_headerless` 키워드로 선행 행 전체를 skiprows로 흡수한다. 모든 충돌·오류는 경고로 흡수, 예외 없음 (spec §6).

**Tech Stack:** Python 3.14, openpyxl 3.1.5, pandas 3.0.3, pytest.

---

## 작업 환경 (모든 Task 공통)

- **작업 디렉터리**: `/Users/daniel/Documents/project/sk-ax/excel-parser/.worktrees/issue-9-block-override` (브랜치 `issue-9-block-override`). 주 체크아웃을 절대 수정하지 말 것.
- **인터프리터**: 워크트리에는 venv가 없다. 주 체크아웃의 venv를 쓴다: `../../.venv/bin/python`. 모든 pytest는 워크트리 루트에서 `../../.venv/bin/python -m pytest …`로 실행한다 (`python -m pytest`가 cwd를 sys.path에 올리므로 워크트리의 `excel_inspector`가 import된다).
- **주의**: 테스트 실행 후 `tests/fixtures/encrypted.xlsx`가 dirty로 보이면 절대 스테이징하지 말 것 (비결정적 재생성 노이즈). 커밋 전 `git checkout -- tests/fixtures/encrypted.xlsx`로 되돌린다.
- 커밋 메시지는 한국어, 본문 끝에 `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`.
- 좌표계: 검사 도메인은 openpyxl **1-based**, `ReadPlan`은 pandas **0-based**, 변환은 PlanAggregator 단독 [D1].

---

### Task 1: `BlockOverride` 모델 + `SheetOverride.block_overrides`

**Files:**
- Modify: `excel_inspector/models.py` (SheetOverride는 53행 부근; `BlockOverride`를 그 직전에 추가)
- Modify: `excel_inspector/__init__.py` (57행 부근 `.models` import 목록, 164행 부근 `__all__`)
- Create: `tests/test_block_override.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_block_override.py` 생성:

```python
"""Block-level override channel [D7] (issue #9).

Covers the BlockOverride model 3-state sentinel contract, the
options.resolve_block_overrides anchor resolution rules, the BlockAnalyzer
specificity chain, and the declared-headerless block read plan.
"""

from __future__ import annotations

from excel_inspector.models import BlockOverride, InspectionOptions, SheetOverride


def test_block_override_three_states() -> None:
    """_UNSET sentinel keeps int / explicit-None / unspecified distinct (HIGH #2)."""

    unspecified = BlockOverride()
    assert unspecified.header_row_set is False
    assert unspecified.header_row is None  # sentinel collapsed for readers

    forced = BlockOverride(header_row=7)
    assert forced.header_row_set is True
    assert forced.header_row == 7

    headerless = BlockOverride(header_row=None)
    assert headerless.header_row_set is True
    assert headerless.header_row is None


def test_sheet_override_carries_block_overrides() -> None:
    """SheetOverride.block_overrides defaults empty; keys are anchor rows."""

    bare = SheetOverride()
    assert bare.block_overrides == {}

    override = SheetOverride(
        block_overrides={19: BlockOverride(header_row=None)}
    )
    assert override.block_overrides[19].header_row_set is True
    # The block channel alone is NOT a sheet-level header declaration.
    assert override.header_row_set is False


def test_block_override_is_publicly_exported() -> None:
    """BlockOverride is part of the public API surface."""

    import excel_inspector

    assert excel_inspector.BlockOverride is BlockOverride
    assert "BlockOverride" in excel_inspector.__all__
```

- [ ] **Step 2: 실패 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: FAIL — `ImportError: cannot import name 'BlockOverride'`

- [ ] **Step 3: 구현**

`excel_inspector/models.py` — `_UNSET` 정의(49행)와 `class SheetOverride`(53행) 사이에 추가:

```python
@dataclass
class BlockOverride:
    """Per-block manual override, anchored by a row inside the target band [D7].

    Registered on :attr:`SheetOverride.block_overrides` keyed by an *anchor
    row* — any 1-based absolute row inside the target band [D1] (e.g. any row
    from a ``rows 19-23`` warning). The header channel reuses the
    :data:`_UNSET` sentinel contract of :class:`SheetOverride` (HIGH #2), so
    three states are representable:

    * ``header_row`` left at :data:`_UNSET` — defer to the heuristic locator.
    * ``header_row=<int>`` — force this block's header to that 1-based row
      (must fall inside the anchored band; validated by
      ``options.resolve_block_overrides``).
    * ``header_row=None`` — declare this block headerless: the band is a
      table whose data region is the whole band (conservative analysis;
      boundary/type profiling skipped, same contract as the sheet-level
      headerless declaration).

    Attributes:
        header_row: Forced header row (1-based), ``None`` for a headerless
            declaration, or the :data:`_UNSET` sentinel. Whether the field
            was actually specified is exposed via :attr:`header_row_set`.
    """

    header_row: int | None | _Unset = _UNSET
    #: Whether ``header_row`` was explicitly specified (set in
    #: ``__post_init__``; not a constructor argument).
    header_row_set: bool = field(init=False, default=False)

    def __post_init__(self) -> None:
        """Record whether ``header_row`` was specified [D7] (HIGH #2 pattern)."""

        self.header_row_set = self.header_row is not _UNSET
        if self.header_row is _UNSET:
            self.header_row = None
```

`SheetOverride`에 필드 추가 — `is_tabular: bool | None = None`(85행) 바로 다음, `header_row_set` 필드 선언 앞에:

```python
    #: Per-block overrides keyed by anchor row (any 1-based row inside the
    #: target band) [D7]. Resolution / conflict policy lives in
    #: ``options.resolve_block_overrides``.
    block_overrides: dict[int, BlockOverride] = field(default_factory=dict)
```

(`header_row_set: bool = field(init=False, ...)`은 init 인자가 아니므로 dataclass 필드 순서 제약에 걸리지 않는다. `block_overrides`는 기본값이 있어 어디든 가능하지만 가독성을 위해 `is_tabular` 다음에 둔다.)

`SheetOverride`의 docstring Attributes에 한 줄 추가:

```
        block_overrides: Per-block overrides keyed by anchor row [D7]; see
            :class:`BlockOverride`.
```

`excel_inspector/__init__.py` — `.models` import 목록(57행 부근, 알파벳순)에 `BlockOverride,` 추가, `__all__`(164행 부근, 알파벳순)에 `"BlockOverride",` 추가.

- [ ] **Step 4: 통과 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 3 passed

- [ ] **Step 5: 기존 모델 계약 회귀 확인**

Run: `../../.venv/bin/python -m pytest tests/test_options_sentinel.py -v`
Expected: all passed

- [ ] **Step 6: 커밋**

```bash
git add excel_inspector/models.py excel_inspector/__init__.py tests/test_block_override.py
git commit -m "feat: BlockOverride 모델과 SheetOverride.block_overrides 채널 추가 [D7] (#9)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 2: `options.resolve_block_overrides` — 앵커 해석·충돌 검출

**Files:**
- Modify: `excel_inspector/options.py`
- Test: `tests/test_block_override.py` (추가)

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_block_override.py`에 추가:

```python
from excel_inspector.analyzers.block_segmenter import RowBand
from excel_inspector.options import resolve_block_overrides

#: Bands mirroring the multi_table_stacked fixture: [1..4] and [7..10].
_BANDS = [RowBand(1, 4), RowBand(7, 10)]


def _opts(block_overrides: dict[int, BlockOverride]) -> InspectionOptions:
    return InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(block_overrides=block_overrides)}
    )


def test_resolver_no_options_or_no_block_overrides() -> None:
    assert resolve_block_overrides(None, "Sheet1", _BANDS) == ({}, [])
    assert resolve_block_overrides(
        InspectionOptions(), "Sheet1", _BANDS
    ) == ({}, [])


def test_resolver_anchor_maps_any_row_inside_band() -> None:
    """Anchor 9 (mid-band) resolves to band start 7; keyed by band_start_row."""

    resolved, warnings = resolve_block_overrides(
        _opts({9: BlockOverride(header_row=None)}), "Sheet1", _BANDS
    )
    assert warnings == []
    assert set(resolved) == {7}
    assert resolved[7].header_row is None and resolved[7].header_row_set


def test_resolver_anchor_outside_all_bands_warns() -> None:
    """Row 5 is the blank separator -> no band -> warned and ignored."""

    resolved, warnings = resolve_block_overrides(
        _opts({5: BlockOverride(header_row=None)}), "Sheet1", _BANDS
    )
    assert resolved == {}
    assert len(warnings) == 1
    assert "anchor row 5" in warnings[0]
    assert "no detected table band" in warnings[0]


def test_resolver_duplicate_anchors_lowest_wins() -> None:
    resolved, warnings = resolve_block_overrides(
        _opts(
            {
                7: BlockOverride(header_row=7),
                9: BlockOverride(header_row=None),
            }
        ),
        "Sheet1",
        _BANDS,
    )
    assert set(resolved) == {7}
    assert resolved[7].header_row == 7  # lowest anchor's override won
    assert len(warnings) == 1
    assert "anchor row 9" in warnings[0] and "anchor row 7" in warnings[0]


def test_resolver_int_header_outside_anchored_band_warns() -> None:
    resolved, warnings = resolve_block_overrides(
        _opts({7: BlockOverride(header_row=2)}), "Sheet1", _BANDS
    )
    assert resolved == {}
    assert len(warnings) == 1
    assert "header_row 2" in warnings[0]
    assert "outside the anchored band" in warnings[0]


def test_resolver_empty_override_warns_and_does_not_claim() -> None:
    """An empty BlockOverride is a no-op; a later valid anchor still claims."""

    resolved, warnings = resolve_block_overrides(
        _opts(
            {
                7: BlockOverride(),
                9: BlockOverride(header_row=None),
            }
        ),
        "Sheet1",
        _BANDS,
    )
    assert set(resolved) == {7}
    assert resolved[7].header_row is None  # anchor 9's override claimed the band
    assert len(warnings) == 1
    assert "no override field specified" in warnings[0]
```

- [ ] **Step 2: 실패 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 새 테스트 6개 FAIL — `ImportError: cannot import name 'resolve_block_overrides'`

- [ ] **Step 3: 구현**

`excel_inspector/options.py` — import 블록을 다음으로 교체 (10–13행):

```python
from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from .heuristics import HEADER_CONFIDENCE_THRESHOLD, SKIP_KEYWORDS
from .models import BlockOverride, InspectionOptions, SheetOverride

if TYPE_CHECKING:  # runtime import would cycle through the analyzers package
    from .analyzers.block_segmenter import RowBand
```

파일 끝에 추가:

```python
def resolve_block_overrides(
    options: InspectionOptions | None,
    sheet_name: str,
    bands: Sequence[RowBand],
) -> tuple[dict[int, BlockOverride], list[str]]:
    """Resolve anchor-keyed block overrides onto row bands [D7] (issue #9).

    Anchors are processed in ascending order (deterministic). Resolution
    rules (design doc §4): an anchor inside no band, a duplicate anchor to a
    band already claimed by a valid override, an empty override (no field
    specified), and an int ``header_row`` outside the anchored band are each
    warned and ignored — never raised (spec §6). A band whose override was
    ignored falls back to the specificity chain (sheet-level override where
    its absolute row anchors the band, else the heuristic).

    Args:
        options: The inspection options, or ``None``.
        sheet_name: Target sheet name.
        bands: The sheet's detected row bands (1-based inclusive [D1]).

    Returns:
        ``(resolved, warnings)`` where ``resolved`` maps each claimed band's
        ``start_row`` to its winning :class:`BlockOverride`.
    """

    override = get_sheet_override(options, sheet_name)
    if override is None or not override.block_overrides:
        return {}, []

    resolved: dict[int, BlockOverride] = {}
    claimed: dict[int, int] = {}  # band start_row -> winning anchor row
    warnings: list[str] = []
    for anchor in sorted(override.block_overrides):
        block_override = override.block_overrides[anchor]
        band = next(
            (b for b in bands if b.start_row <= anchor <= b.end_row), None
        )
        if band is None:
            warnings.append(
                f"block_override: sheet {sheet_name!r}: anchor row {anchor} "
                f"falls inside no detected table band; override ignored"
            )
            continue
        if band.start_row in claimed:
            warnings.append(
                f"block_override: sheet {sheet_name!r}: anchor row {anchor} "
                f"targets the same band (rows {band.start_row}-"
                f"{band.end_row}) as anchor row {claimed[band.start_row]}; "
                f"override ignored"
            )
            continue
        if not block_override.header_row_set:
            warnings.append(
                f"block_override: sheet {sheet_name!r}: anchor row {anchor}: "
                f"no override field specified; override ignored"
            )
            continue
        if isinstance(block_override.header_row, int) and not (
            band.start_row <= block_override.header_row <= band.end_row
        ):
            warnings.append(
                f"block_override: sheet {sheet_name!r}: anchor row {anchor}: "
                f"header_row {block_override.header_row} falls outside the "
                f"anchored band (rows {band.start_row}-{band.end_row}); "
                f"override ignored"
            )
            continue
        claimed[band.start_row] = anchor
        resolved[band.start_row] = block_override
    return resolved, warnings
```

- [ ] **Step 4: 통과 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 9 passed

- [ ] **Step 5: 커밋**

```bash
git add excel_inspector/options.py tests/test_block_override.py
git commit -m "feat: resolve_block_overrides 앵커 해석·충돌 검출 헬퍼 [D7] (#9)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 3: PlanAggregator — `declared_headerless` 변환 경로

**Files:**
- Modify: `excel_inspector/aggregator.py` (`build_read_plan` 168행 부근, `build_block_read_plan` 622행 부근)
- Test: `tests/test_block_override.py` (추가)

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_block_override.py`에 추가:

```python
from excel_inspector.aggregator import build_block_read_plan
from excel_inspector.models import TableBlock
from tests.conftest import make_sheet_profile


def _headerless_block() -> TableBlock:
    """A declared-headerless block on band rows 17-21 (design doc §6)."""

    return TableBlock(
        block_index=2,
        band_start_row=17,
        band_end_row=21,
        header_row=None,
        header_confidence=1.0,
        header_provenance="manual",
        data_start_row=17,
        data_end_row=21,
        data_left_col=None,
        data_right_col=None,
        skip_rows=[],
        columns=[],
        read_plan=None,
        subtotal_skip_labels={},
    )


def test_declared_headerless_block_plan() -> None:
    """header=None, rows 1-16 absorbed into skiprows, nrows = band length."""

    profile = make_sheet_profile(name="Sheet1", max_row=21, max_col=3)
    plan = build_block_read_plan(
        profile, _headerless_block(), None, None, band_scoped=True
    )
    assert plan.header is None
    assert plan.skiprows == list(range(16))  # 1-based rows 1-16 -> 0-based [D1]
    assert plan.nrows == 5
    assert plan.usecols is None
    assert plan.dtype_map == {}
    assert "headerless sheet: dtype inference skipped" in plan.notes


def test_sheet_level_headerless_plan_unchanged() -> None:
    """The sheet-level headerless path never sets data_start_row -> no
    skiprows absorption; the [D7] rule must not disturb it."""

    from excel_inspector.aggregator import build_read_plan

    profile = make_sheet_profile(name="Sheet1", max_row=10, max_col=3)
    opts = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    plan = build_read_plan(profile, opts, None)
    assert plan.header is None
    assert plan.skiprows == []
    assert plan.nrows is None
```

(파일 상단 import에 합치지 말고 위처럼 추가해도 되지만, 깔끔하게 하려면 모듈 상단 import 블록으로 옮겨도 좋다 — 어느 쪽이든 동작은 같다.)

- [ ] **Step 2: 실패 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py::test_declared_headerless_block_plan -v`
Expected: FAIL — 현재 코드는 `header_row=None`이면 v1 fallback `header=0`을 내고 skiprows를 흡수하지 않는다 (`assert plan.header is None` 실패).

- [ ] **Step 3: 구현**

`excel_inspector/aggregator.py` — `build_read_plan` 시그니처(168–172행)를 keyword 전용 인자로 확장:

```python
def build_read_plan(
    profile: SheetProfile,
    options: InspectionOptions | None = None,
    warnings: list[str] | None = None,
    *,
    declared_headerless: bool = False,
) -> ReadPlan:
```

docstring Args에 추가:

```
        declared_headerless: ``True`` for a [D7] block-scoped headerless
            declaration (``BlockOverride(header_row=None)``). Forces the
            headerless plan shape (``header=None``) and — because a block's
            data region IS its band — absorbs every row above
            ``data_start_row`` into ``skiprows``. The sheet-level headerless
            path never sets ``data_start_row``, so it is unaffected.
```

`headerless_override` 계산(202–205행)을 교체:

```python
    headerless_override = declared_headerless or (
        has_header_override(options, profile.name)
        and profile.header_row is None
    )
```

Rule 1의 headerless 분기(244행 부근 `if headerless_override: header = None`)를 교체:

```python
    if headerless_override:
        header = None
        # [D7] declared-headerless block: the band's data region is known,
        # so every leading row (1 .. data_start_row-1) is absorbed into
        # skiprows — the [D1] rule-1 analogue for a block with no header
        # anchor. The sheet-level headerless path never sets data_start_row
        # (boundary analysis is skipped), so it stays bit-identical.
        if (
            declared_headerless
            and profile.data_start_row is not None
            and profile.data_start_row > 1
        ):
            skiprows.extend(range(0, profile.data_start_row - 1))
```

`build_block_read_plan`의 plan 생성(740행 부근 `plan = build_read_plan(synthetic, None, warnings)`)을 교체:

```python
    # [D7] a declared-headerless block: header_row None with manual
    # provenance is impossible for any other block kind (a manual block has
    # always carried an int header; a heuristic band with no header is
    # judged not-a-table), so the pair is a safe discriminator.
    declared_headerless = (
        block.header_row is None and block.header_provenance == "manual"
    )
    plan = build_read_plan(
        synthetic, None, warnings, declared_headerless=declared_headerless
    )
```

- [ ] **Step 4: 통과 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 11 passed

- [ ] **Step 5: 집계기 회귀 확인**

Run: `../../.venv/bin/python -m pytest tests/test_aggregator.py tests/test_multi_table.py tests/test_multi_level_load.py -v`
Expected: all passed (declared_headerless 기본값 False라 기존 경로 비트 동일)

- [ ] **Step 6: 커밋**

```bash
git add excel_inspector/aggregator.py tests/test_block_override.py
git commit -m "feat: declared-headerless 블록의 ReadPlan 변환 경로 추가 [D7][D1] (#9)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 4: BlockAnalyzer — 특이성 체인 + headerless 블록 생성

**Files:**
- Modify: `excel_inspector/analyzers/block_analyzer.py` (`analyze` 97–159행, `_analyze_band` 202행 부근)
- Test: `tests/test_block_override.py` (추가)

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_block_override.py`에 추가. `multi_table_stacked` fixture(밴드 [1..4]와 [7..10], T1 헤더 1행 '부서/인원/예산', T2 헤더 7행 '제품명/단가/재고/비고')와 `header_simple` fixture(단일 밴드)를 사용한다. `fixture_path`는 `tests/conftest.py`의 세션 fixture.

```python
from excel_inspector import inspect


def test_block_int_override_beats_sheet_int_for_same_band(fixture_path) -> None:
    """Sheet header_row=8 and a block override both target band [7..10];
    the block override (header_row=7) wins, with a conflict warning."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                header_row=8,
                block_overrides={9: BlockOverride(header_row=7)},
            )
        }
    )
    profile = inspect(fixture_path("multi_table_stacked"), opts)
    sheet = profile.sheets[0]
    b1, b2 = sheet.blocks
    assert b1.header_provenance == "heuristic"  # band 1 untouched
    assert b2.header_row == 7
    assert b2.header_provenance == "manual"
    assert any(
        "the block override wins" in w for w in profile.warnings
    )


def test_block_headerless_override_creates_manual_band_block(fixture_path) -> None:
    """BlockOverride(header_row=None) on band [7..10]: the band becomes a
    manual headerless block spanning the whole band."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={8: BlockOverride(header_row=None)}
            )
        }
    )
    sheet = inspect(fixture_path("multi_table_stacked"), opts).sheets[0]
    assert len(sheet.blocks) == 2
    b2 = sheet.blocks[1]
    assert b2.header_row is None
    assert b2.header_provenance == "manual"
    assert b2.header_confidence == 1.0
    assert (b2.data_start_row, b2.data_end_row) == (7, 10)
    assert (b2.data_left_col, b2.data_right_col) == (None, None)
    assert b2.columns == []
    # [D7] declared-headerless plan shape (Task 3).
    assert b2.read_plan is not None
    assert b2.read_plan.header is None
    assert b2.read_plan.skiprows == list(range(6))  # rows 1-6 absorbed
    assert b2.read_plan.nrows == 4


def test_sheet_headerless_plus_block_overrides_block_channel_wins(
    fixture_path,
) -> None:
    """Sheet-wide header_row=None + block_overrides: contradiction warned,
    per-band analysis proceeds (band 1 heuristic, band 2 headerless)."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                header_row=None,
                block_overrides={7: BlockOverride(header_row=None)},
            )
        }
    )
    profile = inspect(fixture_path("multi_table_stacked"), opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 2
    assert sheet.blocks[0].header_provenance == "heuristic"
    assert sheet.blocks[1].header_row is None
    assert sheet.blocks[1].header_provenance == "manual"
    assert any("contradicts block_overrides" in w for w in profile.warnings)


def test_sheet_headerless_without_block_overrides_unchanged(fixture_path) -> None:
    """Without block_overrides the sheet-wide headerless gate is intact
    (guard 4): no per-band analysis, no blocks."""

    opts = InspectionOptions(
        sheet_overrides={"Sheet1": SheetOverride(header_row=None)}
    )
    sheet = inspect(fixture_path("multi_table_stacked"), opts).sheets[0]
    assert sheet.blocks == []


def test_block_overrides_on_single_band_sheet_warn_and_ignore(
    fixture_path,
) -> None:
    """header_simple is single-band: block_overrides are ignored with a
    pointer to the sheet-level channel; the mirror block is untouched."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={1: BlockOverride(header_row=None)}
            )
        }
    )
    profile = inspect(fixture_path("header_simple"), opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 1
    assert sheet.blocks[0].header_row == 1  # heuristic mirror intact
    assert any(
        "single-band sheet" in w and "sheet-level SheetOverride" in w
        for w in profile.warnings
    )


def test_resolver_warnings_surface_through_inspect(fixture_path) -> None:
    """An anchor in the blank separator (row 5) surfaces its warning."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={5: BlockOverride(header_row=None)}
            )
        }
    )
    profile = inspect(fixture_path("multi_table_stacked"), opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 2  # both bands fall back to the heuristic
    assert all(b.header_provenance == "heuristic" for b in sheet.blocks)
    assert any(
        "anchor row 5" in w and "no detected table band" in w
        for w in profile.warnings
    )
```

- [ ] **Step 2: 실패 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 새 테스트 6개 중 `test_sheet_headerless_without_block_overrides_unchanged`만 PASS(기존 동작), 나머지 5개 FAIL (block_overrides가 아직 무시되므로 manual 블록·경고가 생기지 않음)

- [ ] **Step 3: 구현**

`excel_inspector/analyzers/block_analyzer.py`:

(a) import 갱신 — 58–63행:

```python
from ..models import BlockOverride, SheetOverride, SheetProfile, TableBlock
from ..options import (
    get_header_confidence_threshold,
    get_sheet_override,
    has_header_override,
    resolve_block_overrides,
)
```

(b) `analyze()`의 105–121행(시트 전역 headerless 게이트 + 단일 밴드 분기)을 교체:

```python
            # Guard 4: an explicit headerless declaration (header_row=None
            # override) is sheet-wide — there is no per-block header to
            # anchor on, so the v1 headerless flat path stays authoritative
            # and no block is produced. [D7] exception: when the more
            # specific block channel is also present it wins — the
            # contradiction is warned and per-band analysis proceeds.
            override = get_sheet_override(context.options, profile.name)
            block_override_map = (
                override.block_overrides if override is not None else {}
            )
            if (
                override is not None
                and override.header_row_set
                and override.header_row is None
            ):
                if not block_override_map:
                    continue
                context.add_warning(
                    f"block_analyzer: sheet {profile.name!r}: sheet-wide "
                    f"headerless override (header_row=None) contradicts "
                    f"block_overrides; the block channel wins and per-band "
                    f"analysis proceeds [D7]"
                )

            if len(bands) == 1:
                # [D7] the block channel presupposes multi-band analysis;
                # a single-band sheet is fully covered by the sheet-level
                # channel, and routing block overrides into the mirror path
                # would break the mirror-plan == flat-plan invariant.
                if block_override_map:
                    context.add_warning(
                        f"block_analyzer: sheet {profile.name!r}: "
                        f"block_overrides ignored on a single-band sheet; "
                        f"use the sheet-level SheetOverride channel instead "
                        f"[D7]"
                    )
                block = self._mirror_block(profile, bands[0])
                if block is not None:
                    profile.blocks = [block]
                continue
```

(주의: 기존 105–115행에 있던 `override = get_sheet_override(...)` 줄이 위로 올라왔으므로 중복 선언이 남지 않게 한다.)

(c) `analyze()`의 기존 "no detected band" 경고 블록(130–146행) **다음**, `blocks: list[TableBlock] = []` **앞**에 추가:

```python
            # [D7] anchor-keyed block overrides: resolution + conflict
            # warnings are the resolver's job; guard 6 ordering holds (the
            # warnings are sheet-scoped and precede the band loop).
            resolved, override_warnings = resolve_block_overrides(
                context.options, profile.name, bands
            )
            for warning in override_warnings:
                context.add_warning(warning)

            # [D7] specificity: a sheet-level int header_row and a block
            # override claiming the same band — the block override wins.
            if (
                override is not None
                and override.header_row_set
                and isinstance(override.header_row, int)
            ):
                for band in bands:
                    if (
                        band.start_row <= override.header_row <= band.end_row
                        and band.start_row in resolved
                    ):
                        context.add_warning(
                            f"block_analyzer: sheet {profile.name!r}: "
                            f"sheet-level header_row {override.header_row} "
                            f"and a block override both target band rows "
                            f"{band.start_row}-{band.end_row}; the block "
                            f"override wins [D7]"
                        )
```

(d) 밴드 루프의 `_analyze_band` 호출(150–152행)을 교체:

```python
                block = self._analyze_band(
                    context,
                    profile,
                    band,
                    block_index=len(blocks),
                    block_override=resolved.get(band.start_row),
                )
```

(e) `_analyze_band` 시그니처에 인자 추가:

```python
    def _analyze_band(
        self,
        context: InspectionContext,
        profile: SheetProfile,
        band: RowBand,
        block_index: int,
        block_override: BlockOverride | None = None,
    ) -> TableBlock | None:
```

docstring Args에 추가:

```
            block_override: The band's resolved [D7] block override (already
                validated by ``options.resolve_block_overrides``), or
                ``None``. A headerless declaration short-circuits the whole
                Header→Boundary→Type loop; an int header beats the
                sheet-level channel.
```

(f) `_analyze_band` 본문 — `window = ...` / `threshold = ...` / `override = ...` 세 줄(234–236행) **직후**에 headerless 단락을 삽입하고, 이어지는 manual 분기 조건을 확장:

```python
        # [D7] block-scoped headerless declaration: the band's blank-run
        # boundaries ARE the data region; header/boundary/type analysis is
        # skipped (conservative — same contract as the sheet-level
        # headerless path: columns stay unprofiled, the aggregator attaches
        # the headerless note) and the manual declaration is never judged
        # not-a-table. data_left/right_col stay None (full width): the
        # column span is undetected, a documented limitation.
        if block_override is not None and block_override.header_row is None:
            return TableBlock(
                block_index=block_index,
                band_start_row=band.start_row,
                band_end_row=band.end_row,
                header_row=None,
                header_confidence=1.0,
                header_provenance="manual",
                data_start_row=band.start_row,
                data_end_row=band.end_row,
                data_left_col=None,
                data_right_col=None,
                skip_rows=self._fold_skip_overrides(override, band, []),
                columns=[],
                read_plan=None,
                subtotal_skip_labels={},
            )

        header_row: int | None
        below_threshold = False
        if block_override is not None:
            # [D7] block-scoped manual header; the resolver validated that
            # the row sits inside this band. Beats the sheet-level channel.
            assert isinstance(block_override.header_row, int)
            header_row = block_override.header_row
            confidence = 1.0
            provenance = "manual"
        elif (
            has_header_override(context.options, profile.name)
            and override is not None
            and isinstance(override.header_row, int)
            and band.start_row <= override.header_row <= band.end_row
        ):
            # Guard 4: the absolute override row lives in this band -> manual.
            header_row = override.header_row
            confidence = 1.0
            provenance = "manual"
        else:
            ...  # (기존 휴리스틱 분기 그대로)
```

(기존 `if has_header_override(...)` 분기와 `else:` 휴리스틱 분기는 그대로 두고, 그 앞에 `if block_override is not None:` 분기만 끼워 넣는 구조다.)

- [ ] **Step 4: 통과 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 17 passed

- [ ] **Step 5: 멀티테이블/E2E 회귀 확인**

Run: `../../.venv/bin/python -m pytest tests/test_multi_table.py tests/test_end_to_end.py tests/test_results.py -v`
Expected: all passed (특히 `test_headerless_override_sheet_has_no_blocks` — 시트 전역 headerless 단독 사용 계약 불변)

- [ ] **Step 6: 커밋**

```bash
git add excel_inspector/analyzers/block_analyzer.py tests/test_block_override.py
git commit -m "feat: BlockAnalyzer 특이성 체인과 headerless 블록 생성 [D7] (#9)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 5: fixture `stacked_headerless_band` + E2E 골든

**Files:**
- Modify: `tests/fixtures/generate.py` (FIXTURES dict 357행 부근, 빌더 1188행 부근, 빌더 레지스트리 1483행 부근)
- Test: `tests/test_block_override.py` (추가)

- [ ] **Step 1: fixture 추가**

`tests/fixtures/generate.py`의 `FIXTURES` dict에서 `"multi_table_stacked"` 항목 **다음**에 추가:

```python
    "stacked_headerless_band": FixtureSpec(
        "stacked_headerless_band.xlsx",
        True,
        "Sheet 'Sheet1'. [D7 block override, issue #9] Two headered tables "
        "stacked above a pure-numeric HEADERLESS band, each separated by a "
        "BLANK_RUN (2) of empty rows. Table 1: header row 1 (A1:C1 = "
        "'분기','매출','비고'), data rows 2-5 (4 rows), columns A-C. Blank "
        "rows 6-7. Table 2: header row 8 (A8:C8 = '부서','인원','예산'), "
        "data rows 9-14 (6 rows), columns A-C. Blank rows 15-16. Band 3: "
        "rows 17-21 are 5 pure-numeric rows (no header; values "
        "[r, r*11, r*111] for r in 1..5). Row bands (1-based, inclusive): "
        "[1..5], [8..14], [17..21]. max_row=21, max_col=3. Without overrides "
        "band 3 must be judged not a table (numeric rows score below the "
        "0.5 header threshold) with a visible warning; with "
        "BlockOverride(header_row=None) anchored anywhere in 17..21 it must "
        "load as a 5x3 headerless table (col_0..col_2) with no row loss.",
    ),
```

빌더 — `build_multi_table_stacked()` 함수 정의 다음에 추가:

```python
def build_stacked_headerless_band() -> bytes:
    """Two headered tables above a pure-numeric headerless band [D7] (#9).

    Table 1: header row 1, data rows 2-5. Blank rows 6-7. Table 2: header
    row 8, data rows 9-14. Blank rows 15-16. Band 3: pure-numeric rows
    17-21 (no header). See ``FIXTURES`` for the band/boundary expectations.
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    _write_rows(
        ws,
        1,
        [
            ["분기", "매출", "비고"],  # row 1 header (table 1)
            ["1분기", 100, "a"],  # rows 2-5 data
            ["2분기", 200, "b"],
            ["3분기", 300, "c"],
            ["4분기", 400, "d"],
            # rows 6-7 left fully empty -> BLANK_RUN band separator
        ],
    )
    _write_rows(
        ws,
        8,
        [
            ["부서", "인원", "예산"],  # row 8 header (table 2)
            ["영업", 10, 1000],  # rows 9-14 data
            ["개발", 20, 2000],
            ["기획", 5, 500],
            ["재무", 3, 300],
            ["인사", 4, 400],
            ["총무", 2, 200],
            # rows 15-16 left fully empty -> BLANK_RUN band separator
        ],
    )
    _write_rows(
        ws,
        17,
        [
            [1, 11, 111],  # rows 17-21: pure-numeric headerless band
            [2, 22, 222],
            [3, 33, 333],
            [4, 44, 444],
            [5, 55, 555],
        ],
    )
    return _save_bytes(wb)
```

빌더 레지스트리(1483행 부근)의 `"multi_table_stacked": build_multi_table_stacked,` 다음 줄에 추가:

```python
    "stacked_headerless_band": build_stacked_headerless_band,
```

- [ ] **Step 2: 실패하는 E2E 테스트 작성**

`tests/test_block_override.py`에 추가:

```python
from excel_inspector import extract


def test_headerless_band_skipped_without_override(fixture_path) -> None:
    """Baseline: the numeric band 17-21 is judged not a table, with a
    visible warning (no silent loss; the recovery channel is [D7])."""

    path = fixture_path("stacked_headerless_band")
    wr = extract(path)
    assert len(wr.tables) == 2
    assert any(
        "17-21" in w and "band judged not a table" in w for w in wr.warnings
    )


def test_block_override_recovers_headerless_band(fixture_path) -> None:
    """The issue #9 scenario end-to-end: anchor 19 declares band [17..21]
    headerless; all three tables load with no row loss."""

    opts = InspectionOptions(
        sheet_overrides={
            "Sheet1": SheetOverride(
                block_overrides={19: BlockOverride(header_row=None)}
            )
        }
    )
    path = fixture_path("stacked_headerless_band")

    profile = inspect(path, opts)
    sheet = profile.sheets[0]
    assert len(sheet.blocks) == 3
    b3 = sheet.blocks[2]
    assert b3.header_row is None and b3.header_provenance == "manual"
    assert (b3.data_start_row, b3.data_end_row) == (17, 21)
    assert b3.read_plan is not None
    assert b3.read_plan.header is None
    assert b3.read_plan.skiprows == list(range(16))
    assert b3.read_plan.nrows == 5
    assert "headerless sheet: dtype inference skipped" in b3.read_plan.notes

    wr = extract(path, options=opts)
    t1, t2, t3 = wr.tables
    assert t3.table_id == "Sheet1!T3"
    assert t3.header_row is None
    assert list(t3.dataframe.columns) == ["col_0", "col_1", "col_2"]
    assert len(t3.dataframe) == 5
    assert list(t3.dataframe["col_0"]) == [1, 2, 3, 4, 5]
    assert list(t3.dataframe["col_2"]) == [111, 222, 333, 444, 555]
    # Neighbor blocks untouched.
    assert list(t1.dataframe.columns) == ["분기", "매출", "비고"]
    assert len(t1.dataframe) == 4
    assert list(t2.dataframe.columns) == ["부서", "인원", "예산"]
    assert len(t2.dataframe) == 6
```

- [ ] **Step 3: 실행 — fixture 재생성 포함 확인**

Run: `../../.venv/bin/python -m pytest tests/test_block_override.py -v`
Expected: 19 passed (세션 fixture `fixture_corpus`가 generate.py로 코퍼스를 재생성하므로 새 fixture가 자동 생성된다). 만약 baseline 테스트가 "band judged not a table" 경고 문구 불일치로 실패하면 실제 경고 문자열을 확인해(`extract(...).warnings` 출력) 단정문의 부분 문자열을 실제 형식에 맞춘다 — 휴리스틱 점수에 따라 "no header candidate row found" 변형일 수 있으며, 두 변형 모두 "band judged not a table (skipped)"를 포함한다.

- [ ] **Step 4: 새 fixture 파일 커밋 여부 확인**

`git status`로 `tests/fixtures/stacked_headerless_band.xlsx` 생성 확인. 코퍼스 파일은 생성기가 결정적으로 만들므로 **커밋한다** (기존 코퍼스 파일들과 동일 정책). 단 `encrypted.xlsx`가 함께 dirty면 그것만 되돌린다:

```bash
git checkout -- tests/fixtures/encrypted.xlsx
```

- [ ] **Step 5: 커밋**

```bash
git add tests/fixtures/generate.py tests/fixtures/stacked_headerless_band.xlsx tests/test_block_override.py
git commit -m "test: stacked_headerless_band fixture와 issue #9 E2E 골든 추가 [D7]

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 6: 문서 — spec [D7], §5.0, §4.8, CLAUDE.md

**Files:**
- Modify: `docs/excel-structure-inspector-spec.md` (§0 결정 테이블 ~19-25행, §4.8 변환 규칙 ~120-121행, §5.0 ~140행)
- Modify: `CLAUDE.md` ("Override channel [D2]" 단락)
- Modify: `docs/superpowers/specs/2026-06-12-block-override-design.md` (상태 줄)

- [ ] **Step 1: spec §0 결정 테이블에 [D7] 행 추가**

`| **[D6]** | ...` 행 바로 다음에 추가:

```markdown
| **[D7]** | **블록 단위 오버라이드 채널** (issue #9): `SheetOverride.block_overrides: dict[int, BlockOverride]` — 키는 대상 밴드에 포함된 임의의 1-based 앵커 행. `BlockOverride.header_row`는 int(블록 헤더 강제) \| 명시적 None(블록 headerless 선언) \| 미지정(휴리스틱 위임)의 3-상태(`_UNSET` 센티널). 충돌·오류는 특이성 우선(블록 > 시트 > 휴리스틱) + 경고로 처리하고 예외는 던지지 않는다(§6). headerless 블록은 보수적으로 분석한다: 데이터 구간 = 밴드 전체, 컬럼 경계 미검출(전체 폭), 타입 프로파일링 생략. | 적층 시트에서 개별 블록 header_row 지정 불가 / 시트 전역 headerless 선언이 적층 테이블을 파괴 |
```

- [ ] **Step 2: spec §4.8 변환 규칙에 5번 추가**

`  4. \`nrows\`는 데이터 구간의 전체 행 수 ...` 항목 바로 다음에 추가:

```markdown
  5. **headerless 블록 변환 [D7]**: 명시적 headerless 블록(`BlockOverride(header_row=None)`)의 플랜은 `header=None`이며, 밴드 시작 위 모든 행(`1 .. data_start_row-1`)을 0-based `skiprows`로 흡수한다. `nrows`는 규칙 4 그대로 밴드 전체 행 수. 시트 단위 headerless 선언은 경계 분석을 생략해 `data_start_row`가 미설정이므로 이 규칙의 영향을 받지 않는다(기존 경로 불변).
```

- [ ] **Step 3: spec §5.0에 BlockOverride 계약 추가**

`` `SheetOverride`: `header_row: int | None`... `` 줄(140행)의 끝에 `, block_overrides: dict[int, BlockOverride]` **[D7]**`을 추가하고, 바로 다음 줄에 새 단락 추가:

```markdown
`BlockOverride` **[D7]**: `header_row: int | None`(1-based; int = 블록 헤더 강제 — 앵커 밴드 내여야 함, 명시적 None = 블록 headerless 선언, 미지정 = 휴리스틱 위임). `block_overrides`의 키는 대상 밴드에 포함된 임의의 1-based 절대 앵커 행 **[D1]**. 밴드 밖 앵커·같은 밴드 중복 앵커(낮은 앵커 승리)·앵커 밴드 밖 int header_row·빈 BlockOverride는 각각 경고 후 무시되며, 무시된 밴드는 특이성 체인(시트 → 휴리스틱)으로 폴백한다. 시트 전역 `header_row=None`과 block_overrides가 공존하면 모순 경고 후 블록 채널이 이긴다. 단일 밴드 시트의 block_overrides는 경고 후 무시(시트 채널 사용 안내).
```

- [ ] **Step 4: CLAUDE.md 갱신**

"### Override channel [D2]" 단락 끝에 한 문장 추가:

```markdown
Block-level channel [D7] (issue #9): `SheetOverride.block_overrides: dict[int, BlockOverride]` anchors a per-band override by any 1-based row inside the target band; `BlockOverride(header_row=None)` declares one band headerless (data region = whole band, profiling skipped) without collapsing the sheet, and conflicts resolve by specificity (block > sheet > heuristic) with warnings, never exceptions. Resolution lives in `options.resolve_block_overrides`.
```

- [ ] **Step 5: 설계 문서 상태 갱신**

`docs/superpowers/specs/2026-06-12-block-override-design.md`의 `- 상태: 사용자 승인 대기`를 `- 상태: 승인됨 (2026-06-12), 구현 계획 docs/superpowers/plans/2026-06-12-block-override.md`로 교체.

- [ ] **Step 6: 커밋**

```bash
git add docs/excel-structure-inspector-spec.md CLAUDE.md docs/superpowers/specs/2026-06-12-block-override-design.md
git commit -m "docs: spec [D7] 블록 단위 오버라이드 계약 명문화 + CLAUDE.md 갱신 (#9)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 7: 전체 검증

- [ ] **Step 1: 기본 스위트 전체 실행**

Run: `../../.venv/bin/python -m pytest`
Expected: all passed, 0 failed (slow 마커는 addopts로 자동 제외)

- [ ] **Step 2: slow 포함 전체 실행**

Run: `../../.venv/bin/python -m pytest -m "slow or not slow"`
Expected: all passed (perf smoke ~14s 포함)

- [ ] **Step 3: 워킹 트리 정리 확인**

Run: `git status`
Expected: clean (또는 `encrypted.xlsx`만 dirty — 그 경우 `git checkout -- tests/fixtures/encrypted.xlsx`)

- [ ] **Step 4: 마무리**

superpowers:finishing-a-development-branch 스킬로 머지/PR 여부를 사용자에게 확인한다. PR 생성 시 본문에 `Closes #9` 포함.

---

## Self-Review 결과 (계획 작성 후 점검 완료)

- **스펙 커버리지**: 설계 문서 §3→Task 1, §4→Task 2, §5→Task 4, §6→Task 3, §7 경고 7종→Task 2(4종)+Task 4(3종) 테스트, §8 한계→Task 4 주석+Task 6 spec 문구, §9→Task 6, §10→Task 1–5 테스트. 갭 없음.
- **타입 일관성**: `resolve_block_overrides(options, sheet_name, bands) -> (dict[int, BlockOverride], list[str])` 시그니처가 Task 2 정의·Task 4 호출에서 일치. `_analyze_band(..., block_override: BlockOverride | None = None)` 일치. `build_read_plan(..., *, declared_headerless: bool = False)` 일치.
- **순서 의존성**: Task 3(집계기)이 Task 4(분석기)보다 앞이라, Task 4의 통합 테스트가 read_plan 형태까지 바로 단정할 수 있다.
