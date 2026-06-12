# 블록 단위 오버라이드 채널 설계 [D7] (issue #9)

- 날짜: 2026-06-12
- 상태: 승인됨 (2026-06-12), 구현 계획 docs/superpowers/plans/2026-06-12-block-override.md
- 관련: issue #9, spec §5.0 [D2], §4.8 [D1], §6, §8; plan v2 §4 guard 4
- 선행 검토: issue #9 재현 완료 (아래 §1)

## 0. 요약

`SheetOverride`는 시트 단위로만 적용되므로, 적층(multi-block) 시트에서 특정 블록만
headerless로 선언하거나 header_row를 강제할 채널이 없다. 앵커 행으로 대상 밴드를
지정하는 `BlockOverride`를 `SheetOverride.block_overrides: dict[int, BlockOverride]`로
추가하고, BlockAnalyzer가 앵커를 해석해 밴드 단위로 적용한다. 새 결정 ID **[D7]**.

## 1. 문제 (재현으로 확인됨)

시나리오: 한 시트에 T1(행 1–5, 헤더 있음), T2(행 9–15, 헤더 있음),
T3(행 19–23, 순수 숫자·헤더 없음)이 적층.

| 케이스 | 실측 결과 |
| --- | --- |
| 오버라이드 없음 | T3 `best header score 0.300 below threshold 0.500` 경고와 함께 탈락. 복구 채널 없음 |
| `SheetOverride(header_row=None)` | 시트 전체가 23×3 단일 headerless 테이블로 합쳐져 T1/T2 파괴 (`block_analyzer.py` 시트 전역 headerless 게이트) |
| `SheetOverride(header_row=19)` | T3 추출되나 데이터 행 `(1, 11, 111)`이 헤더로 소비 — 컬럼명 `['1','11','111']`, 5행 중 4행만 잔존 |

근본 원인: `SheetOverride.header_row`는 시트당 단일 슬롯. int 오버라이드는 guard 4로
이미 밴드 귀속 의미를 갖지만, `None`은 좌표가 없어 밴드에 귀속시킬 수 없고
시트 전역 선언으로 남아 있다 (multi-block 이전의 계약, spec §5.0 / HIGH #3).

## 2. 확정된 설계 결정

브레인스토밍에서 사용자가 확정한 사항:

1. **필드 범위**: `header_row`만 (YAGNI). 블록별 `dtype_force` 등 인접 격차는 별도 이슈로 분리.
2. **API 형태**: `dict[앵커행, BlockOverride]`. 앵커 행 = 대상 밴드에 포함된 임의의
   1-based 절대 행 [D1]. 경고 메시지의 `rows 19-23`을 보고 그대로 쓸 수 있다.
   (블록 인덱스 방식은 판정 변화에 따라 번호가 밀려 불안정하므로 기각.)
3. **headerless 블록 분석 깊이**: 보수적 — 경계/타입 분석 생략, 데이터 구간 = 밴드 전체,
   기존 headerless 노트 부착. 시트 단위 headerless 계약과 일관.
4. **충돌 정책**: 특이성 우선 + 경고. 더 구체적인 선언이 이기고, 잘못된 입력은
   경고 후 무시. 예외는 절대 던지지 않음 (spec §6 흡수-지속 정책 유지).
5. **구현 위치**: BlockAnalyzer 중심 (접근 A). "테이블 아님" 판정이 일어나는 자리에서
   오버라이드를 확인해야 판정 뒤집기(T3 복구)가 가능 — 집계기 중심 안은 판정 탈락
   밴드가 블록이 되지 못해 복구 대상이 없으므로 기각.
6. **단일 밴드 시트**: block_overrides가 오면 경고 + 무시, 시트 채널 안내.
   미러 경로(v1 골든 코퍼스 비트 동일성) 불변식을 건드리지 않는다.

## 3. 데이터 모델 (`models.py`)

```python
@dataclass
class BlockOverride:
    """Per-block manual override, anchored by a row inside the target band [D7]."""
    header_row: int | None | _Unset = _UNSET
    header_row_set: bool = field(init=False, default=False)  # __post_init__
```

- `_UNSET` 센티널·`header_row_set` 패턴은 `SheetOverride`(HIGH #2)와 동일한 3-상태 계약:
  미지정(휴리스틱 위임) / int(강제 헤더) / 명시적 `None`(headerless 선언).
- `SheetOverride`에 추가:

```python
block_overrides: dict[int, BlockOverride] = field(default_factory=dict)
# key = 앵커 행 (1-based 절대 행, 대상 밴드 내 임의 행) [D1][D7]
```

기존 필드·계약은 모두 불변.

## 4. 옵션 헬퍼 (`options.py`)

`resolve_block_overrides(options, sheet_name, bands)` 추가 —
앵커→밴드 해석과 충돌 검출을 한 곳에서 수행하고
`(dict[밴드 시작행, BlockOverride], 경고 리스트)`를 반환한다.

해석 규칙 (결정적):

- 앵커가 어느 밴드에도 안 속함 → 경고 + 무시 (기존 guard 4 no-band 경고와 동일 패턴).
- 두 앵커가 같은 밴드 → **낮은 앵커 행 승리**, 나머지 경고 + 무시.
- `header_row=int`가 앵커 밴드 밖 → 경고 + 해당 오버라이드 무시. 무시된 밴드는
  §5 우선순위 체인의 다음 단계로 떨어진다 (시트 int가 그 밴드에 앵커되면 그것이 적용,
  아니면 휴리스틱).
- `header_row` 미지정(빈 BlockOverride) → 경고 + 무시 (no-op 선언). 폴백은 위와 동일.

## 5. BlockAnalyzer (`block_analyzer.py`)

우선순위(특이성 우선): **블록 오버라이드 > 시트 `header_row`(guard 4 밴드 한정) > 휴리스틱**.

- **시트 전역 headerless 게이트 수정**: `header_row=None` 시트 선언이 있어도
  `block_overrides`가 비어 있지 않으면 모순 경고를 내고 per-band 분석을 진행한다
  (블록 채널 승리). `block_overrides`가 없으면 현행 유지 —
  `tests/test_multi_table.py`의 guard 4 headerless 계약 테스트 불변.
- **단일 밴드 시트**: block_overrides 존재 시 경고("시트 채널을 사용하라") + 무시.
- **`_analyze_band`**: 해석된 `BlockOverride(header_row=None)`이 있는 밴드는
  HeaderLocator/BoundaryDetector를 건너뛰고 즉시 headerless 블록을 생성:
  - `header_row=None, header_confidence=1.0, header_provenance="manual"`
  - `data_start_row=band.start_row, data_end_row=band.end_row`
    (블랭크 런 경계가 곧 데이터 구간)
  - `data_left_col/right_col=None` (전체 폭; §8 한계 참고)
  - `columns=[]`, `subtotal_skip_labels={}`
  - `skip_rows` = 기존 `_fold_skip_overrides` 재사용 (밴드 내 `skip_rows_add` 귀속)
  - manual 선언은 "테이블 아님" 판정을 받지 않는다 (기존 원칙 유지)
- `BlockOverride(header_row=int)`는 기존 guard 4 manual 분기와 동일 처리 —
  이로써 시트당 **여러** 밴드의 헤더 강제가 가능해진다 (기존 단일 슬롯 한계 해소).
- 미러 규칙·`_mark_extracted_bands`는 변경 없이 headerless 블록에도 그대로 적용
  (headerless 블록이 `blocks[0]`이면 평면 필드가 그것을 미러).

## 6. PlanAggregator (`aggregator.py`)

`build_read_plan`에 keyword 전용 파라미터 `declared_headerless: bool = False` 추가:

- **[D1] 변환 규칙 1 확장**: `declared_headerless`이고 `data_start_row`가 설정돼 있으면
  행 `1..data_start_row-1` 전체를 0-based `skiprows`로 흡수한다.
  시트 단위 headerless는 `data_start_row` 미설정(경계 분석 생략)이므로
  기존 경로가 비트 동일하게 유지된다.
- `build_block_read_plan`은 선언적 headerless 블록을
  `block.header_row is None and block.header_provenance == "manual"`로 판별한다
  (현재 이 조합은 생성 불가능하므로 안전한 판별식; manual 블록은 항상 int 헤더였다).
- 결과 플랜: `header=None`, `skiprows=[0..band_start-2]`,
  `nrows=밴드 길이`(기존 `_compute_nrows`가 산출), 기존 `_HEADERLESS_NOTE`
  문자열 재사용(안정 계약), body-merge 노트·1-column 노트·`dtype_force`는
  기존 로직 그대로 통과. nrows가 확정되므로 밴드 클램프 불필요.
- `results.py`는 변경 없음 — `header=None` 플랜의 프레임은 이미
  `col_0..col_n`으로 명명된다 (README 직렬화 계약).

기대 출력 (§1 시나리오 + `block_overrides={19: BlockOverride(header_row=None)}`):
T1 4×3, T2 6×3, **T3 5×3 `col_0..col_2`** — 행 손실 없음.

## 7. 경고 목록 (모두 예외 없음)

| 상황 | 처리 |
| --- | --- |
| 밴드 밖 앵커 | 경고 + 무시 |
| 같은 밴드 중복 앵커 | 낮은 앵커 승리, 나머지 경고 + 무시 |
| 앵커 밴드 밖 int `header_row` | 경고 + 무시 |
| 시트 전역 `None` + block_overrides 공존 | 모순 경고, 블록 채널 승리 |
| 시트 int와 블록 오버라이드가 같은 밴드 | 블록 승리 + 경고 |
| 단일 밴드 시트의 block_overrides | 경고 + 무시 (시트 채널 안내) |
| 빈 BlockOverride (`header_row` 미지정) | 경고 + 무시 |

경고 누적 순서는 guard 6(시트 순서 → 밴드 top-down)을 따른다.

## 8. 문서화된 한계

- headerless 블록의 컬럼 경계는 검출하지 않는다 (`data_left/right_col=None` → 전체 폭).
  시트보다 좁은 headerless 밴드는 로드 프레임 우측에 NaN `col_N` 컬럼이 생긴다.
  보이는 결과이므로 침묵 손실은 아니며, 필요 시 후속 개선으로 분리한다.
- headerless 블록은 타입 프로파일링을 생략하므로 `dtype_map`은 시트 단위
  `dtype_force`만 반영한다 (시트 단위 headerless와 동일 계약).
- 블록별 `dtype_force`는 범위 외 (별도 이슈로 분리).

## 9. 스펙 문서 변경

- spec §0 리비전 테이블에 **[D7]** 결정 행 추가
  (블록 단위 오버라이드: 앵커 행 dict, 특이성 우선 + 경고, headerless 블록은 보수적 분석).
- §5.0에 `BlockOverride` 표와 `SheetOverride.block_overrides` 행 추가.
- §4.8에 headerless 블록 변환 규칙(선행 행 전체 skiprows 흡수) 추가.

## 10. 테스트 전략

- **fixture**: `tests/fixtures/generate.py`에 `stacked_headerless_band.xlsx` 추가
  (§1 시나리오 — T1/T2 헤더 + T3 순수 숫자). 코퍼스는 생성기로만 관리.
- **단위 (BlockAnalyzer / options)**: partial-context synthesis(`make_context`)로
  §7 경고 케이스 전부 + headerless 블록 필드 검증.
- **집계 (aggregator)**: headerless 블록 플랜 골든 —
  `header=None / skiprows / nrows / _HEADERLESS_NOTE`.
- **E2E**: `extract` + pandas 왕복 — T3 5행 보존, `col_0..col_2`,
  table_id `…!T3`, 다른 블록 비간섭.
- **회귀**: `block_overrides` 부재 시 기존 코퍼스 결과 비트 동일;
  시트 전역 headerless 단독 사용 계약(`test_multi_table.py`) 불변;
  read-only/멱등성 보증은 기존 하니스가 커버.
