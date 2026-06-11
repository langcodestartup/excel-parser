# 비표(non-tabular) 판정의 내용 인식형 휴리스틱 — 설계

- **이슈:** #3 — 비표 판정이 텍스트 시작 열에 민감: B열 시작 표지가 `columns=[]` 빈 테이블 생성
- **결정 방향:** 옵션 B (휴리스틱 개선) + 판정 신호로 "채워진 열 개수 + 밀도 임계" 둘 다 사용
- **대상 코드:** `excel_inspector/analyzers/sheet_enumerator.py`, `excel_inspector/heuristics.py`
- **관련 결정 ID:** `[D2]`(오버라이드 채널), `[D4]`(휴리스틱 상수), spec §4.2, §6, §9

## 1. 배경 / 문제

`sheet_enumerator._is_tabular_candidate`의 현재 비표 판정은 한 줄이다:

```python
return max_col > _MAX_NON_TABULAR_COLS, "heuristic"   # _MAX_NON_TABULAR_COLS = 1
```

`max_col`은 openpyxl이 보고하는 **최우측 채워진 열**일 뿐 밀도·구조와 무관하다. 그 결과:

- 텍스트가 A열에서 시작하는 표지 → `max_col=1` → 정상 스킵.
- 동일한 표지를 B열부터 쓰면(셀 병합/들여쓰기로 A열을 비운 흔한 레이아웃) → `max_col=2` → 표로 통과.

통과한 시트는 HeaderLocator에서 헤더 점수 미달로 `header_row=None, needs_manual_header=True`가 되고, 컬럼이 프로파일되지 않아 결국 `columns=[]`인데 `row_count>0`인 **모순된 빈 껍데기 테이블**이 결과/JSON에 노출된다. 이는 README 계약(비표 시트 → `skipped=True, skip_reason="non-tabular", tables=[]`)과 어긋난다.

재현(이슈 본문과 동일):

```
[A열 텍스트] skipped= True  reason= non-tabular  tables= []
[B열 시작 ] skipped= False reason= None         tables= [('표지!T1', cols=0, rows=5)]
```

`tests/fixtures/complex_demo.xlsx`의 `표지` 시트(텍스트가 B열에서 시작)에 동일 증상이 살아 있다.

## 2. 목표 / 비목표

### 목표
- A열을 비웠는지 여부와 무관하게, "희소 텍스트만 있고 표 구조가 없는" 시트를 일관되게 `skipped=True, skip_reason="non-tabular", tables=[]`로 판정한다.
- 다중 열이지만 밀도가 매우 낮은 표지(예: 제목 B2 + 날짜 D5)도 비표로 잡는다.

### 비목표 (YAGNI / 사용자 선택 존중)
- `needs_manual_header`인 **진짜 표**(예: 전부 숫자라 헤더가 모호한 표)가 `columns=[]`를 내는 별개 증상(이슈의 결함 ②)은 **건드리지 않는다**. 이런 표는 채워진 열이 2개 이상이라 본 변경 이후에도 표로 유지되며 동작이 불변한다. (옵션 C가 아니라 B를 선택)
- 블록 단위(스택 표) 산출물은 건드리지 않는다. 본 변경은 **시트 단위** 게이트만 바꾼다.

## 3. 설계

### 3.1 위치 — 단일 권위 지점 유지

판정을 `sheet_enumerator._is_tabular_candidate` 안에서 **내용 인식형으로 교체**한다. 이유:

- 비표 판정의 단일 권위 위치를 유지하면, 하류 스테이지(MergeScanner / BlockSegmenter / HeaderLocator 등)가 이미 `is_tabular_candidate=False`를 존중해 자동으로 스킵한다(추가 배선 불필요).
- `[D2]` 오버라이드는 그대로 최우선으로 동작한다.
- `sheet_enumerator`는 Loader 스테이지 이후에 실행되므로 data 모드 핸들을 사용할 수 있다(HeaderLocator가 이미 `data_workbook()`을 같은 방식으로 사용).

기각한 대안:
- (접근 2) 싼 게이트 유지 + HeaderLocator 이후 강등 → 이미 실행된 merge/block 스테이지에 스킵을 역소급해야 하고 `needs_manual`과 혼동.
- (접근 3) results/aggregator에서 빈 테이블만 억제 → 이건 옵션 A(증상 억제)이며 하류 작업 단축 효과가 없음.

### 3.2 판정 규칙

`sheet_enumerator`가 data 모드로 상단 `NON_TABULAR_SAMPLE_ROWS` 행을 샘플해 다음을 계산한다:

- `pop_cols` = non-empty 셀이 하나라도 있는 distinct 열의 수
- `pop_rows` = non-empty 셀이 하나라도 있는 행의 수
- `filled` = non-empty 셀의 총 개수
- `density = filled / (pop_cols * pop_rows)` (분모가 0이면 0.0)

(non-empty 정의는 기존 헤더 로케이터와 동일: `None`도 빈 문자열도 아닌 값.)

판정 순서:

1. `[D2]` 오버라이드가 있으면 그 값을 그대로 사용하고 `provenance="manual"`.
2. `pop_cols == 0` (샘플에 내용 없음) → **레거시 dims 규칙 `max_col > 1`로 폴백**. (empty_sheet는 비표 유지; "데이터가 샘플 창 아래에서 시작"하는 병적 케이스는 오늘 동작을 보존)
3. `pop_cols <= MIN_TABULAR_POPULATED_COLS` (=1) → **비표**. (이슈의 B열 표지 + 기존 A열 표지를 모두 커버)
4. `pop_cols >= 2` 이고 `density < NON_TABULAR_DENSITY_THRESHOLD` (=0.5) → **비표**. (흩어진 다중 열 표지)
5. 그 외 → **표**.
6. 샘플링 중 예외 발생 시 → 레거시 dims 규칙 + `context.add_warning(...)`. `sheet_enumerator`는 절대 전체 열거를 깨지 않는다(spec §6 견고성).

`provenance`는 휴리스틱 결정에 대해 기존대로 `"heuristic"`을 유지한다. 메서드 시그니처는 이미 `context`를 받으므로 외부 시그니처 변경은 없다(내부에서 loader 사용).

### 3.3 상수 (`heuristics.py`, `[D4]`)

```python
#: 비표 판정 시 상단에서 샘플하는 행 수 (spec §4.2).
NON_TABULAR_SAMPLE_ROWS: int = 20

#: 이 값 이하의 채워진 열만 있으면 비표(표지/안내). 텍스트 시작 열 위치 무관.
MIN_TABULAR_POPULATED_COLS: int = 1

#: 채워진 열이 2개 이상이라도 샘플 밀도가 이 값 미만이면 비표(흩어진 표지).
NON_TABULAR_DENSITY_THRESHOLD: float = 0.5
```

`NON_TABULAR_SAMPLE_ROWS=20`은 `HEADER_SCAN_ROWS`와 동일 값으로 맞춰 상단 스캔 폭을 일치시킨다. (기존 `_MAX_NON_TABULAR_COLS`는 폴백(규칙 2)에서 `max_col > 1` 형태로 계속 쓰이거나 동등 상수로 정리)

## 4. 실증 근거 (임계값 보정)

전체 fixture에 대해 상단 20행 기준으로 지표를 계산한 결과:

| 분류 | 시트 | pop_cols | density |
|---|---|---|---|
| 비표(SKIP 돼야 함) | empty_sheet | 0 | 0.0 |
| | mixed_sheets `README` | **1** | 1.0 |
| | complex_demo `표지` (버그) | **1** | 1.0 |
| 표(유지돼야 함) | 최소 pop_cols (원본 / Hidden / VeryHidden) | **2** | 1.0 |
| | 최소 density (지역별매출) | 6 | **0.648** |

- 모든 비표 시트는 `pop_cols <= 1`, 모든 진짜 표는 `pop_cols >= 2`. → 규칙 3만으로 현재 코퍼스가 100% 분리된다.
- 진짜 표의 density 최저는 0.648 → 임계 0.5와 마진 0.148. density 규칙(규칙 4)은 코퍼스에 없는 "다중 열 흩어진 표지"용 안전망이며, 진짜 표를 오분류하지 않는다.

## 5. 테스트 계획

### 신규 fixture (`tests/fixtures/generate.py`에 추가; 절대 수기 편집 금지)
- `cover_offset.xlsx` — 텍스트가 B열에서 시작하는 단일 열 표지(이슈 재현 자체). 기대: `pop_cols=1` → SKIP. (규칙 3)
- `cover_sparse.xlsx` — 다중 열이지만 흩어진 표지. density ≈ 0.25~0.33이 되도록 설계(임계 0.5보다 충분히 낮게). 기대: SKIP. (규칙 4)

### 단위 테스트 (`make_context`/`make_sheet_profile` 합성)
- 규칙 2~6 각각을 직접 검증: pop_cols=0 폴백, pop_cols=1 비표, 저밀도 비표, 정상 표 유지, 오버라이드 우선, 샘플 예외 시 폴백+warning.

### 회귀 가드
- §4 baseline의 모든 표 시트가 표로 유지되고, 모든 비표 시트가 SKIP로 유지되는지(테이블 단위 산출물 불변) 검증.
- 이슈 재현 시나리오(A열/B열 표지)가 둘 다 SKIP로 끝나는지 end-to-end(`extract().to_json()`) 검증.

### 데모 파일
- `complex_demo.xlsx`는 데모 빌더(`build_complex_demo.py`)로 만든 파일이고 이를 검증하는 테스트가 없어 골든 결합도가 없다. 변경 후 `표지`가 SKIP로 바뀌는 것만 확인.

## 6. 리스크 / 트레이드오프

- **알려진 트레이드오프(사용자 동의됨):** `pop_cols=2`이면서 매우 희소한 *진짜* key-value 표는 density 규칙(규칙 4)에 걸려 비표로 오분류될 수 있다. 임계 0.5로 보수적으로 잡아 위험을 최소화하며, 현재 코퍼스엔 해당 사례가 없다.
- **샘플 창 한계:** 데이터가 상단 `NON_TABULAR_SAMPLE_ROWS` 행 아래에서만 시작하는 시트는 규칙 2 폴백으로 오늘 동작(`max_col>1`)을 보존한다. HeaderLocator도 동일 창만 보므로 일관적이다.
- **성능:** data 모드 상단 20행 스트리밍 추가 — 시트당 비용 미미, spec §8의 스트리밍/메모리 보장에 영향 없음.
