# 엑셀 구조 검사기 구현 플랜 (Implementation Plan)

| 항목 | 내용 |
| --- | --- |
| 문서 종류 | 구현 플랜 (Implementation Plan) |
| 선행 문서 | `excel-structure-inspector-spec.md` (개정 1판) |
| 구현 언어 | Python (3.14 검증 완료, venv 전제) |
| 핵심 의존성 | `openpyxl 3.1.x`, `pandas 3.x`, 테스트: `pytest 9.x` |
| 상태 | 개정 1판(Revised) — 갭 분석 반영 |

---

## 0. 개정 요약

스펙 개정 1판의 결정 `[D1]`~`[D6]`을 구현 단계에 반영했다. 주요 변경:
- **환경 선결 단계(Phase 0)** 에 venv·의존성·픽스처 생성기 포함.
- **경계 탐지를 헤더 직후로 상향**(핵심 가치=집계 중복 방지) **[D6]**.
- **Override/Options 단계 추가** **[D2]**.
- **Formula Detector·다단 헤더는 v1+로 분리** **[D6]**.
- **좌표 변환·정규화를 Aggregator 단일 책임**으로 명시, 골든 회귀 케이스 필수 **[D1]**.
- 픽스처 **생성 스크립트**, 골든 취약성 완화, 멱등성 측정법, 결정성 전략을 테스트 전략에 추가.

---

## 1. 목적과 접근 전략

본 문서는 스펙 문서에서 정의한 검사기를 실제로 구축하기 위한 단계별 계획이다. 구현은 **수직 슬라이스(walking skeleton) 우선 → 분석기 점진 추가** 전략을 따른다. 즉 가장 먼저 "파일을 열어 시트를 나열하고 최소 `ReadPlan`을 산출하는" 가장 얇은 종단 경로를 완성한 뒤, 분석기를 하나씩 끼워 넣으며 프로파일을 풍부하게 만든다.

이 전략의 이점은 다음과 같다. 첫째, 초기에 종단 동작이 검증되어 데이터 모델·파이프라인 계약의 결함을 조기에 발견한다. 둘째, 각 분석기가 독립적으로 추가·테스트되어 회귀 위험이 낮다. 셋째, 어느 단계에서 중단해도 동작하는 산출물이 남는다.

---

## 2. 모듈 구조

```
excel_inspector/
├── __init__.py
├── models.py            # InspectionOptions, WorkbookProfile, SheetProfile, ColumnProfile, MergeRegion, ReadPlan
├── options.py           # InspectionOptions / SheetOverride 와 override 적용 유틸 [D2]
├── context.py           # 파이프라인 공유 컨텍스트(스키마 §6) + 부분 컨텍스트 합성 헬퍼
├── pipeline.py          # Analyzer 인터페이스 + 파이프라인 실행기
├── loader.py            # Loader (구조/데이터 모드 추상화, 핸들 수명 관리) [D3]
├── exceptions.py        # CorruptWorkbookError, EncryptedWorkbookError 등 도메인 예외
├── heuristics.py        # §7 v1 상수(점수 가중치/임계값/키워드) [D4]
├── analyzers/
│   ├── __init__.py
│   ├── sheet_enumerator.py
│   ├── header_locator.py
│   ├── boundary_detector.py     # [D6] 우선순위 상향
│   ├── type_profiler.py
│   ├── merge_analyzer.py
│   └── formula_detector.py      # [D6] v1+ (스텁만)
├── aggregator.py        # Plan Aggregator: 좌표 변환·정규화·usecols·dtype_map·override [D1][D5]
└── adapters/
    └── pandas_loader.py # ReadPlan → pandas.read_excel 변환

tests/
├── fixtures/
│   ├── generate.py      # openpyxl로 표본 .xlsx 합성 (손상/암호 포함) — 코퍼스 생성기
│   └── *.xlsx           # generate.py 산출물(또는 생성 시점 재생)
├── conftest.py          # 부분 컨텍스트 합성 헬퍼, 픽스처 fixture
├── test_loader.py
├── test_<analyzer>.py   # 분석기별 단위 테스트
├── test_aggregator_coords.py  # [D1] 좌표 변환/정규화 회귀
└── test_end_to_end.py   # 골든 ReadPlan + 적재 정렬/집계 검증
```

핵심 경계: `excel_inspector`는 검사만 책임지며, 실제 적재는 `adapters/pandas_loader.py`가 `ReadPlan`을 `pandas` 파라미터로 번역하는 지점에서만 외부 라이브러리와 만난다. **좌표계 변환은 `aggregator.py`에서만 발생한다** **[D1]**.

---

## 3. 단계별 마일스톤

각 단계는 목표·산출물·완료 기준(Exit Criteria)을 가진다. 규모는 상대 티셔츠 사이즈(S/M/L)로 표기하며 절대 일정이 아니다.

### Phase 0 — 환경·스캐폴딩 (M)
- **목표**: 실행 가능한 개발 환경 + 패키지 골격 + 데이터 모델 + 파이프라인 계약 + **픽스처 생성기**.
- **산출물**:
  - `venv` + `requirements.txt`(openpyxl/pandas/pytest 핀) + 설치 확인.
  - `models.py`(5장 데이터 모델 전체, `InspectionOptions` 포함), `options.py`, `context.py`, `pipeline.py`, `exceptions.py`, `heuristics.py`(상수 자리).
  - `tests/fixtures/generate.py`: openpyxl로 §5.1 코퍼스 표본을 합성(손상=잘린 zip, 암호=보호 저장).
  - `tests/conftest.py`: 부분 컨텍스트 합성 헬퍼.
- **완료 기준**: `pytest`가 빈 분석기 목록으로 파이프라인을 예외 없이 실행하고 빈 `WorkbookProfile`을 반환. 픽스처 생성기가 모든 표본 `.xlsx`를 만든다.

### Phase 1 — 종단 경로 (Loader + Sheet Enumerator) (M) **[D3]**
- **목표**: 파일을 읽기 전용으로 열고 시트를 나열해 최소 `ReadPlan`을 산출.
- **산출물**: `loader.py`(구조/데이터 모드, 핸들 `close()` 보장), `analyzers/sheet_enumerator.py`, `aggregator.py` 1차.
- **완료 기준**: 정상 파일에서 시트 목록·사용 범위·max_row/col이 채워지고, 헤더를 첫 데이터 행으로 가정한 기본 `ReadPlan`(0-based)이 나온다. 손상/암호 파일은 명시적 도메인 예외로 처리된다. 핸들 누수 없음(속성 테스트).

### Phase 2 — 헤더 탐지 (L) **[D4]**
- **목표**: 상단 표본으로 헤더 행을 §7.1 점수식으로 추정(1-based).
- **산출물**: `analyzers/header_locator.py`, `heuristics.py` 헤더 상수.
- **완료 기준**: 헤더가 1행이 아닌 표본에서 헤더 행과 신뢰도가 정확히 산출. 실패 시 `None`+신뢰도 0+`needs_manual_header`.

### Phase 3 — 경계 탐지 (L) **[D6] 우선순위 상향**
- **목표**: 데이터 시작·끝, 소계/빈 줄, 좌/우 열 경계 식별(§7.2).
- **산출물**: `analyzers/boundary_detector.py`.
- **완료 기준**: 소계·합계 행이 `skip_rows`로 표시되어 **집계 중복이 방지됨을 골든+적재 테스트로 검증**. 열 경계가 `usecols`로 번역됨.

### Phase 4 — 좌표 변환·정규화 강화 (Aggregator) (M) **[D1]**
- **목표**: 1-based→0-based 변환, 헤더 정규화, `skiprows` 합성, `usecols`/`dtype_map` 도출, override 적용.
- **산출물**: `aggregator.py` 확정, `options.py` override 적용.
- **완료 기준**: `test_aggregator_coords.py`가 "상단 선행행 + 중간 소계" 결합 케이스에서 한 행도 밀리지 않음을 단언. override 지정 시 해당 필드가 `provenance=manual`로 반영.

### Phase 5 — 타입 프로파일 (M) **[D4][D5]**
- **목표**: 컬럼별 자료형·결측 비율 추론(§7.3).
- **산출물**: `analyzers/type_profiler.py`.
- **완료 기준**: `numeric_text`·날짜·혼합 타입이 표본 기반으로 분류. `dtype_map` 키가 0-based 위치 문자열로 산출.

### Phase 6 — 병합 분석 (M) **[D3]**
- **목표**: 병합 영역 분류(구조 모드 로드)와 다단 헤더 플래그.
- **산출물**: `analyzers/merge_analyzer.py`.
- **완료 기준**: 헤더/본문 병합이 분류되고, 본문 병합 fill 권고가 `notes`에 기록.

### Phase 7 — 적재기 연동 (M) **[D1]**
- **목표**: `ReadPlan`을 `pandas.read_excel` 호출로 번역.
- **산출물**: `adapters/pandas_loader.py`.
- **완료 기준**: 검사 → 적재 종단에서, 표본 파일들이 **정렬 오류·집계 중복 없이** 정확한 DataFrame으로 적재됨을 골든으로 고정(pandas 3.x).

### Phase 8 — 견고성·성능 (M)
- **목표**: 엣지 케이스, 대용량, 멱등성 강화.
- **산출물**: 예외 경로 보강, 표본 검사 튜닝, 메모리 측정.
- **완료 기준**: 대용량 표본에서 상주 메모리 ≤ 200MB, 동일 입력 재검사 결과 동일(해시 불변).

### Phase 9 (v1+) — 수식 탐지 + 다단 헤더 **[D6]**
- v1 비목표. 데이터 모델 자리(`has_formula`, `read_hint`, `header: list[int]`)는 이미 마련됨. 별도 워크북(`data_only=False`) 개방, MultiIndex/forward-fill 도입.

---

## 4. 권장 순서와 의존성

```
Phase 0 ─→ Phase 1 ─→ Phase 2(헤더) ─→ Phase 3(경계) ─→ Phase 4(좌표/Aggregator)
                                                          ├─→ Phase 5(타입)
                                                          └─→ Phase 6(병합)
                          Phase 4~6 ─→ Phase 7(적재) ─→ Phase 8(견고성)
                                                          └─→ Phase 9(v1+: 수식/다단)
```

Phase 0·1은 직렬 선행. **Phase 2(헤더)→3(경계)→4(좌표 변환)은 핵심 가치 경로로 우선 직렬 완성**한다 **[D6]**. 분석기 간 데이터 의존성이 실재하므로(헤더→경계/타입, 병합→다단) 초안의 "임의 순서" 서술은 폐기하고 위 토폴로지를 따른다. Phase 5·6은 Phase 4 이후 병렬 가능.

---

## 5. 테스트 전략

### 5.1 픽스처 코퍼스 + 생성기
`tests/fixtures/generate.py`가 openpyxl로 표본 `.xlsx`를 **프로그램적으로 합성**한다(수작업 금지, 재현 가능). 양식:
- 헤더 1행 정상 표
- 상단 제목 후 N행 뒤 헤더가 시작되는 표
- **상단 선행행 + 중간 소계 결합**(좌표 정규화 회귀용) **[D1]**
- 병합 헤더 / 다단 헤더(v1+ 검증용)
- 중간 소계 + 말미 합계 포함 표
- `numeric_text`·날짜·혼합 타입 컬럼
- 좌측 설명 열이 있어 표가 일부 열만 차지하는 표(usecols 검증)
- 비표 설명 시트 + 표 시트 혼합 워크북
- 빈 시트 / 데이터 0행 / 헤더 없는 시트 (음성 테스트)
- 손상 파일(잘린 zip), 암호 보호 파일

### 5.2 테스트 계층
- **단위 테스트**: 분석기별로 `conftest.py`의 **부분 컨텍스트 합성 헬퍼**로 입력을 구성 → 기대 보강 결과 검증.
- **좌표 회귀(`test_aggregator_coords.py`)**: 1-based→0-based 변환·헤더 정규화·skiprows 합성이 한 행도 밀리지 않음 **[D1]**.
- **골든 테스트(종단)**: 픽스처별 기대 `ReadPlan`과 비교. **취약성 완화** — 전체 dict 동등 비교 대신 핵심 필드(`header`/`skiprows`/`usecols`) 부분 단언을 우선하고, 풀 골든은 안정화 후 도입.
- **속성 테스트**: 멱등성(검사 전후 **파일 바이트 해시 동일**), 읽기 전용(mtime·해시 불변), 핸들 누수 없음.
- **적재 검증(Phase 7)**: `ReadPlan`으로 적재한 DataFrame의 행 수·집계가 기대치와 일치. pandas 3.x 환경 고정, 결정성 확보.

### 5.3 결정성·전제
- 표본 검사 비결정성과 결정적 골든 단언의 충돌을 막기 위해 **표본 추출을 결정적으로 고정**(시드 또는 균등 인덱스 규칙).
- 테스트 전제(venv·의존성 설치·Python 3.14 휠·pytest 설정)를 `Phase 0` 산출물로 명문화.

### 5.4 회귀 방지
새 양식이 들어올 때마다 `generate.py`에 합성 규칙과 골든 `ReadPlan`을 추가해 코퍼스를 누적 확장한다.

---

## 6. 위험과 완화책

| 위험 | 영향 | 완화책 |
| --- | --- | --- |
| **좌표계 변환 누락 → 한 행 밀림** **[D1]** | 높음 | 변환을 Aggregator 단일 책임으로 고정, 필드별 좌표계 주석, `test_aggregator_coords.py` 회귀 필수 |
| **read_only에 merged_cells 부재** **[D3]** | 높음 | 병합·치수는 구조 모드(일반 로드)로 분리 수집 |
| 헤더 탐지 휴리스틱 정확도 부족 | 높음 | 점수 가중치를 `heuristics.py` 상수로 분리, 신뢰도 노출 + override 경로, 코퍼스 보정 |
| pandas 3.x 동작 변경(dtype 기본값 등) | 중간 | 적재 검증을 3.x 고정, 골든으로 동작 핀 |
| 대용량 파일 메모리 초과 | 중간 | 스트리밍+표본 고정, 전수 스캔 금지, Phase 8 측정 |
| 수식 모드 캐시값 공백 | 중간 | v1+로 연기, `as_value` 기본, 도입 시 `as_formula` 분기 |
| 단일 시트 다중 표 블록 | 중간 | 1차 최상위 1블록만, 미해결 분리 |
| 픽스처 편향 | 중간 | 실제 고객 양식을 익명화해 `generate.py`에 반영 |

---

## 7. 완료 정의 (Definition of Done) — v1

- Phase 0~8 완료 기준 충족, 단위·좌표회귀·골든·속성·적재 테스트 통과.
- 검사 단계가 원본을 변경하지 않음이 해시 기반 속성 테스트로 보장.
- 대표 양식들이 적재기 연동까지 정렬·집계 오류 없이 처리됨(특히 상단 선행행+소계 결합 케이스).
- 저신뢰·판정 불가 상태가 예외 없이 `warnings`/`provenance`로 명시.
- 모든 워크북 핸들이 정리되어 누수 없음.

---

## 8. 미해결 결정 사항

- 헤더 점수 가중치·임계값의 **전면** 외부 설정화 시점(v1은 상수 + 일부 `InspectionOptions` 노출).
- 단일 시트 다중 표 블록 분할 도입 시점.
- 수식 탐지·다단 헤더(Phase 9) 우선순위.
- CSV/TSV·`.xls` 확장.
- 검사 결과 캐싱 도입 여부.

---

## 9. 다음 단계

본 플랜이 확정되면 Phase 0(환경·스캐폴딩)부터 착수한다. `models.py`·`pipeline.py`·`context.py`의 인터페이스가 이후 모든 분석기의 기반이므로 이 계약을 먼저 고정하고, 동시에 `tests/fixtures/generate.py`로 코퍼스를 확보해 첫 골든 테스트의 토대를 만든다.
