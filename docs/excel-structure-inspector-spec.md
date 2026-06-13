# 엑셀 구조 검사기(Excel Structure Inspector) 스펙 문서

| 항목 | 내용 |
| --- | --- |
| 문서 종류 | 기술 스펙 (Technical Specification) |
| 대상 독자 | 데이터 파이프라인 / 백엔드 개발자 |
| 구현 언어 | Python (3.14 검증 완료) |
| 핵심 의존성 | `openpyxl`(구조·서식·수식 검사), `pandas`(후속 적재) |
| 상태 | 개정 1판(Revised) — 갭 분석 반영 |

---

## 0. 개정 이력 / 핵심 설계 결정

본 판은 초안에 대한 갭 분석(High 4 · Medium 23 · Low 38 확정)을 반영해 "어떻게 정확히"의 공백을 메운 것이다. 초안에서 미정의였던 7개 핵심 결정을 아래와 같이 확정한다. 본문 곳곳에 `[D#]` 태그로 적용 위치를 표기한다.

| ID | 결정 | 해소한 갭 |
| --- | --- | --- |
| **[D1]** | **좌표계 계약 분리**: `SheetProfile`의 모든 행/열 위치는 **openpyxl 1-based**(검사 도메인), `ReadPlan`의 모든 위치는 **pandas 0-based**(적재 도메인). 1→0 변환은 **Plan Aggregator 단독 책임**. 헤더 위 선행 행은 전부 `skiprows`로 흡수해 헤더를 정규화하고, 변환 결과는 골든 테스트로 고정한다. | 좌표계 미정의로 인한 한 행 밀림(정렬 오류) |
| **[D2]** | **재정의(override) 채널 도입**: 검사 진입점이 `InspectionOptions`를 받아 컨텍스트 초기 상태로 주입. 모든 추정 필드는 `provenance`(`heuristic`/`manual`/`default`)를 동반. override된 필드는 분석기가 산출을 생략하고 `confidence=1.0, provenance=manual`로 통과. | 핵심 목표인 override의 입력 채널 부재 |
| **[D3]** | **로더 모드 정책**: 병합·치수 등 **구조 메타데이터는 일반(비 read_only) 모드로 1회** 로드, **행 데이터는 read_only 스트리밍 표본**으로 읽는다. `data_only=True/False`는 서로 다른 워크북 인스턴스이며 필요 시 지연 개방한다. 모든 핸들은 명시적 `close()`로 정리한다. | read_only에 `merged_cells` 부재 / 2회 열기 비용 / 핸들 누수 |
| **[D4]** | **휴리스틱 상수 명시**: 헤더 점수식·표본 크기, 타입 판정 임계값, 경계 탐지 밀도·키워드를 v1 상수로 고정(§7). 외부 설정화는 v1+. | 산출식/임계값 전무 |
| **[D5]** | **컬럼 정체성 단일화**: `ColumnProfile.index`는 **표 좌상단을 0으로 하는 0-based 위치**. `ReadPlan.dtype_map`의 키는 **컬럼 위치(0-based)를 문자열화한 값**으로 고정(`None` 컬럼명 문제 회피). | dtype_map 키 모호 / 컬럼 정체성 불일치 |
| **[D6]** | **v1 범위 축소·우선순위 교정**: Formula Detector와 다단 헤더(`header: list[int]`)는 **v1+로 연기**. 핵심 가치인 경계 탐지를 헤더 직후 우선순위로 상향. | 과대범위 / 핵심 가치 우선순위 역전 |
| **[D7]** | **블록 단위 오버라이드 채널** (issue #9): `SheetOverride.block_overrides: dict[int, BlockOverride]` — 키는 대상 밴드에 포함된 임의의 1-based 앵커 행. `BlockOverride.header_row`는 int(블록 헤더 강제) \| 명시적 None(블록 headerless 선언) \| 미지정(휴리스틱 위임)의 3-상태(`_UNSET` 센티널). 충돌·오류는 특이성 우선(블록 > 시트 > 휴리스틱) + 경고로 처리하고 예외는 던지지 않는다(§6). headerless 블록은 보수적으로 분석한다: 데이터 구간 = 밴드 전체, 컬럼 경계 미검출(전체 폭), 타입 프로파일링 생략. | 적층 시트에서 개별 블록 header_row 지정 불가 / 시트 전역 headerless 선언이 적층 테이블을 파괴 |

---

## 1. 개요

엑셀 구조 검사기는 임의의 `.xlsx` 파일을 **읽기 전용으로 사전 검사**하여, 후속 적재 단계가 그대로 사용할 수 있는 **읽기 계획(Read Plan)** 을 산출하는 시스템이다. 실무 엑셀은 헤더 위치, 병합 셀, 소계·합계 행, 자료형이 파일마다 제각각이어서, 검사 없이 곧바로 표로 읽으면 정렬이 어긋나거나 집계가 중복된다. 본 시스템은 "검사(inspection)"와 "적재(loading)"를 분리하고, 검사 단계가 부작용 없이 동작하도록 하는 것을 목표로 한다.

---

## 2. 목표와 비목표

### 2.1 목표 (Goals)
- 워크북의 시트·헤더·병합·데이터 경계·자료형을 자동으로 파악한다.
- 검사 결과를 적재기가 직접 소비할 수 있는 단일 계약(`ReadPlan`)으로 산출한다.
- 대용량 파일에서도 메모리 부담 없이 동작하도록 스트리밍·표본 검사를 기본으로 한다.
- 휴리스틱 판정에 신뢰도를 부여하고, 외부에서 재정의(override) 가능하게 한다. **[D2]**
- 모든 추정 산출에 출처(provenance)를 기록해 자동/수동/기본값을 추적 가능하게 한다. **[D2]**

### 2.2 비목표 (Non-goals)
- 실제 데이터의 전수 적재·변환·집계는 본 시스템의 책임이 아니다(후속 적재 단계의 몫).
- 원본 파일의 수정·생성은 수행하지 않는다(검사는 철저히 읽기 전용).
- `.xls`(구형 바이너리), 비암호 외 보호 형식, 차트 데이터 복원은 1차 범위에서 제외한다.
- **[D6]** 수식 탐지(Formula Detector)와 다단 헤더 처리는 v1 비목표(v1+로 연기). 데이터 모델·계약에는 자리만 남기고 v1 분석기는 단일 헤더(`int | None`)만 산출한다.

---

## 3. 시스템 아키텍처

검사기는 **파이프라인** 구조를 따른다. 입력 파일을 연 뒤, 공통 인터페이스를 구현한 분석기(Analyzer)들이 순차적으로 공유 컨텍스트(프로파일 객체)를 점진적으로 채워 나가고, 마지막에 집계기가 읽기 계획을 확정한다.

```
[Loader]
   → [Sheet Enumerator]
   → [Header Locator]
   → [Boundary Detector]      # [D6] 핵심 가치(집계 중복 방지) 우선순위 상향
   → [Type Profiler]
   → [Merge Analyzer]
   → ([Formula Detector])     # [D6] v1+ 연기
   → [Plan Aggregator]        # [D1] 좌표 변환·정규화·override 적용을 여기서 단일 수행
   → ReadPlan
```

각 분석기는 전략 단위로 분리되어, 특정 고객 양식 전용 탐지기 추가나 휴리스틱 교체가 용이하다. 분석기 간 데이터 의존성(헤더 → 경계/타입, 병합 → 다단 판정)이 존재하므로 **실행 순서는 위 토폴로지 정렬을 따른다**(초안의 "임의 순서" 서술은 폐기). 의존성 없는 분석기끼리는 병렬화 가능하나 v1은 직렬 실행을 기본으로 한다.

---

## 4. 구성 요소 명세

### 4.1 Loader (로더) **[D3]**
- **책임**: 워크북을 읽기 전용으로 연다. 두 가지 개방 모드를 추상화한다.
  - **구조 모드**: 일반(비 `read_only`) `load_workbook(..., read_only=False, data_only=True)` 1회 개방. 병합 영역(`merged_cells`)·치수·시트 메타데이터를 얻는다. *read_only 워크시트에는 `merged_cells`가 존재하지 않으므로 병합 분석은 반드시 이 모드를 사용한다.*
  - **데이터 모드**: `read_only=True` 스트리밍 개방. 표본 행을 전진 단일 패스로 읽는다. 캐시값(`data_only=True`) 기준.
  - **수식 모드(v1+)**: `data_only=False`는 별도 워크북 인스턴스를 요구한다. v1에서는 개방하지 않는다.
- **입력**: 파일 경로 또는 바이트 스트림, `InspectionOptions`.
- **출력**: 모드별 워크북 핸들, 열기 메타데이터.
- **수명·멱등성**: 모든 핸들은 컨텍스트 종료 시 명시적으로 `close()`한다(임시파일·핸들 누수 방지). 동일 입력에 대해 동일한 핸들 집합을 산출하며 원본을 변경하지 않는다.
- **예외 처리**: 손상 파일, 암호 보호 파일, 비 `.xlsx` 형식을 이 계층에서 조기 차단한다. 암호 보호는 openpyxl 단독으로는 부분 탐지에 그치므로, `BadZipFile`/`InvalidFileException` 등을 명시적 도메인 예외(`EncryptedWorkbookError`, `CorruptWorkbookError`)로 변환한다.

### 4.2 Sheet Enumerator (시트 열거기)
- **책임**: 시트 이름, 가시성(숨김 여부), 사용 범위, 최대 행·열을 수집한다.
- **출력**: 시트별 기본 메타데이터(`name`, `is_visible`, `used_range`, `max_row`, `max_col`). 비표 형태(설명 시트 등) 후보를 `is_tabular_candidate=False`로 표시한다.
- **주의**: `read_only` 모드의 `max_row`/`max_col`/`calculate_dimension`은 차원 정보가 리셋된 파일에서 `None`이거나 부정확할 수 있다. 따라서 치수는 구조 모드(일반 로드)에서 수집하고, 신뢰 불가 시 표본 스캔으로 보정한다(`used_range_trusted: bool` 플래그 동반).

### 4.3 Header Locator (헤더 탐지기) **[D4]**
- **책임**: 헤더 행 위치를 추정한다. 헤더는 항상 1행이 아니며, 상단 제목·작성일 등이 선행할 수 있다.
- **방법**: 상단 표본 행만 읽어 §7.1 점수식으로 후보 행을 점수화한다.
- **출력**: `header_row`(1-based, `int | None`) + `header_confidence`(0~1) + `provenance`. 단정하지 않고 재정의 가능 상태로 남긴다. 추정 실패 시 `header_row=None, header_confidence=0, needs_manual_header=True`.
- **v1 제약**: 단일 헤더 행만 산출한다(다단 헤더는 v1+).

### 4.4 Merge Analyzer (병합 분석기) **[D3]**
- **책임**: 병합 영역을 헤더 병합과 본문 병합으로 분류하고, 다단 헤더 여부를 판정한다.
- **입력 제약**: 병합 정보는 **구조 모드(일반 로드)** 의 `worksheet.merged_cells.ranges`에서만 얻는다.
- **출력**: 병합 영역 목록(`MergeRegion[]`), `is_multi_level_header` 플래그.
- **본문 병합 처리 권고**: 본문 병합(`kind=body`)은 적재 시 좌상단 값을 나머지 셀에 전파(forward-fill)할 것을 `ReadPlan.notes`로 권고한다. v1은 권고만 기록하고 실제 fill은 적재기 몫.

### 4.5 Boundary Detector (경계 탐지기) **[D4][D6]**
- **책임**: 데이터 시작 행, 종료 행, 중간 소계 행, 빈 분리 줄을 식별한다. *집계 중복 방지의 핵심으로, v1 우선순위 상위.*
- **방법**: §7.2 규칙(행 밀도 분석 + 키워드 매칭).
- **출력**: `data_start_row`/`data_end_row`(1-based), `skip_rows`(1-based 소계·빈 줄 목록), 그리고 표가 `used_range`의 일부 열만 차지할 때의 좌/우 열 경계(`data_left_col`/`data_right_col`, 1-based)를 함께 산출한다. 열 경계는 Aggregator가 `usecols`로 번역한다.

### 4.6 Type Profiler (타입 프로파일러) **[D4][D5]**
- **책임**: 컬럼별 추정 자료형, 결측 비율, 대표값을 산출한다.
- **방법**: §7.3 규칙으로 표본 행만 검사. "숫자로 보이나 문자열로 저장된 값", 날짜 서식, 혼합 타입을 식별한다.
- **출력**: 컬럼별 `ColumnProfile`. `index`는 표 좌상단 기준 0-based **[D5]**.

### 4.7 Formula Detector (수식 탐지기) — **v1+ 연기 [D6]**
- **책임**: 컬럼에 수식이 포함되는지 판정하고, 적재 시 결과값/수식 문자열 중 무엇을 취할지 권고한다.
- **방법**: 수식 모드(별도 워크북, `data_only=False`) 표본 읽기.
- **연기 사유**: 두 번째 워크북 개방 비용 + "한 번도 Excel에서 열린 적 없는 파일은 캐시값 공백" 가정으로 v1 가치가 낮음. v1은 모든 컬럼 `has_formula=False, read_hint=as_value`로 기본 산출.

### 4.8 Plan Aggregator (집계기) **[D1][D2][D5]**
- **책임**: 모든 분석 결과를 종합해 `ReadPlan`을 확정한다. 본 검사기에서 **유일하게 좌표계를 변환**하는 지점이다.
- **좌표 변환 규칙 [D1]**:
  1. 헤더 위 모든 선행 행(`1 .. header_row-1`)을 `skiprows`(0-based)로 흡수한다.
  2. 헤더 아래 소계/합계/빈 줄(`skip_rows`, 1-based)을 0-based 절대 인덱스로 변환해 `skiprows`에 합친다.
  3. `header`는 위 `skiprows` 적용 후 프레임에서 헤더가 차지하는 위치로 정규화한다(헤더가 선행행을 모두 건너뛴 직후이면 0).
  4. `nrows`는 데이터 구간의 전체 행 수 = `data_end_row - data_start_row + 1`(1-based 포함 구간). **내부 소계/빈 줄 skip을 차감하지 않는다.** pandas `nrows`는 헤더 이후 *소비할 원본 행 수*를 세며, 내부 `skiprows`는 출력에서 제거되지만 nrows 예산은 소비한다. 차감하면 읽기 창이 데이터 끝까지 닿지 못해 마지막 데이터 행이 누락된다(pandas 3.0.3 실측 확인). 이 규칙은 golden 왕복 테스트로 고정한다.
  5. **headerless 블록 변환 [D7]**: 명시적 headerless 블록(`BlockOverride(header_row=None)`)의 플랜은 `header=None`이며, 밴드 시작 위 모든 행(`1 .. data_start_row-1`)을 0-based `skiprows`로 흡수한다. `nrows`는 규칙 4 그대로 밴드 전체 행 수. 시트 단위 headerless 선언은 경계 분석을 생략해 `data_start_row`가 미설정이므로 이 규칙의 영향을 받지 않는다(기존 경로 불변).
- **usecols 도출**: `data_left_col`~`data_right_col`(1-based)을 엑셀 열문자 범위 문자열(예: `"B:H"`)로 변환. 경계 미검출 시 `None`(전체 열).
- **dtype_map 도출 [D5]**: 키 = usecols 선택 프레임 기준 0-based 컬럼 위치의 문자열(`"0"`, `"1"`, ...), 값 = pandas dtype 문자열. 적재기는 키를 정수로 환원해 적용한다.
- **override 적용 [D2]**: `InspectionOptions`에 명시된 필드는 분석기 산출을 무시하고 override 값으로 덮어쓰며 `provenance=manual`로 기록한다.
- **출력**: 적재기가 직접 소비 가능한 `ReadPlan`.

---

## 5. 데이터 모델

검사 결과는 명시적 구조체로 정의한다. 최상위 `WorkbookProfile`이 여러 `SheetProfile`을 담고, 각 시트가 자신의 `ReadPlan`을 가진다. **좌표계 규약 [D1]**: `SheetProfile`·`ColumnProfile`의 행은 openpyxl **1-based**, `ReadPlan`의 행은 pandas **0-based**, `ColumnProfile.index`는 표 좌상단 기준 **0-based**.

### 5.0 InspectionOptions (입력 재정의 계약) **[D2]**
| 필드 | 타입 | 설명 |
| --- | --- | --- |
| `sheet_overrides` | dict[str, SheetOverride] | 시트명별 재정의 |
| `header_confidence_threshold` | float | 헤더 신뢰도 임계값(기본 0.5) |
| `skip_keywords` | list[str] \| None | 경계 키워드 추가/대체 |

`SheetOverride`: `header_row: int | None`(1-based 강제 지정), `skip_rows_add: list[int]`, `skip_rows_remove: list[int]`, `dtype_force: dict[str, str]`, `is_tabular: bool | None`, `block_overrides: dict[int, BlockOverride]` **[D7]**.

`BlockOverride` **[D7]**: `header_row: int | None`(1-based; int = 블록 헤더 강제 — 앵커 밴드 내여야 함, 명시적 None = 블록 headerless 선언, 미지정 = 휴리스틱 위임). `block_overrides`의 키는 대상 밴드에 포함된 임의의 1-based 절대 앵커 행 **[D1]**. 해석은 `options.resolve_block_overrides`가 전담하며 경고 접두사는 `block_override:`로 고정한다(분석기 모듈 접두사 규약의 의도적 예외 — 채널 이름이 곧 출처). 밴드 밖 앵커·같은 밴드 중복 앵커(낮은 앵커 승리)·앵커 밴드 밖 int header_row·빈 BlockOverride는 각각 경고 후 무시되며, 무시된 밴드는 특이성 체인(시트 → 휴리스틱)으로 폴백한다. 시트 전역 `header_row=None`과 block_overrides가 공존하면 모순 경고 후 블록 채널이 이긴다. 단일 밴드 시트의 block_overrides는 경고 후 무시(시트 채널 사용 안내).

### 5.1 WorkbookProfile
| 필드 | 타입 | 설명 |
| --- | --- | --- |
| `file_path` | str | 원본 파일 경로 |
| `sheets` | list[SheetProfile] | 시트별 프로파일 |
| `open_errors` | list[str] | 열기 단계 경고·오류 |

### 5.2 SheetProfile  *(행 위치는 openpyxl 1-based)*
| 필드 | 타입 | 설명 |
| --- | --- | --- |
| `name` | str | 시트 이름 |
| `is_visible` | bool | 가시성 |
| `is_tabular_candidate` | bool | 표 형태 여부 추정 |
| `used_range` | str | 사용 범위 (예: `A1:H120`) |
| `used_range_trusted` | bool | 치수 신뢰 여부(read_only 보정 플래그) |
| `max_row` / `max_col` | int | 최대 행·열(구조 모드 수집) |
| `header_row` | int \| None | 추정 헤더 행 (1-based) |
| `header_confidence` | float | 헤더 추정 신뢰도(0~1) |
| `header_provenance` | str | `heuristic` / `manual` / `default` |
| `needs_manual_header` | bool | 헤더 추정 실패 → 수동 지정 요구 |
| `is_multi_level_header` | bool | 다단 헤더 여부 (v1은 항상 False) |
| `merges` | list[MergeRegion] | 병합 영역 |
| `data_start_row` / `data_end_row` | int \| None | 데이터 시작/종료 행 (1-based) |
| `data_left_col` / `data_right_col` | int \| None | 데이터 좌/우 열 경계 (1-based) |
| `skip_rows` | list[int] | 소계·빈 줄 등 제외 행 (1-based) |
| `columns` | list[ColumnProfile] | 컬럼 프로파일 |
| `read_plan` | ReadPlan | 최종 읽기 계획 |

### 5.3 ColumnProfile  *(index는 표 좌상단 기준 0-based)*
| 필드 | 타입 | 설명 |
| --- | --- | --- |
| `index` | int | 컬럼 위치(0-based) |
| `name` | str \| None | 컬럼명 |
| `inferred_type` | str | `number` / `text` / `numeric_text` / `date` / `mixed` |
| `null_ratio` | float | 결측 비율(분모=표본 데이터 행 수) |
| `has_formula` | bool | 수식 포함 여부 (v1은 항상 False) |
| `read_hint` | str | `as_value` / `as_formula` (v1은 항상 `as_value`) |

### 5.4 MergeRegion
| 필드 | 타입 | 설명 |
| --- | --- | --- |
| `range` | str | 병합 범위 (예: `A1:C1`) |
| `kind` | str | `header` / `body` |

### 5.5 ReadPlan (검사 ↔ 적재 단일 계약)  *(행은 pandas 0-based)*
| 필드 | 타입 | 설명 |
| --- | --- | --- |
| `sheet_name` | str | 대상 시트 |
| `engine` | str | pandas 엔진(고정 `openpyxl`) |
| `header` | int \| list[int] | 헤더 행(0-based, post-skip 정규화). v1은 항상 int |
| `usecols` | str \| None | 엑셀 열문자 범위(예 `B:H`) 또는 None(전체) |
| `skiprows` | list[int] | 건너뛸 행(0-based 절대 인덱스) |
| `nrows` | int \| None | 읽을 행 수 |
| `dtype_map` | dict[str, str] | 키=0-based 컬럼 위치 문자열, 값=pandas dtype **[D5]** |
| `notes` | list[str] | 적재기 참고 사항(본문 병합 fill 권고 등) |

> **계약 주의**: `header`/`skiprows`/`usecols`/`dtype_map`은 pandas `read_excel` 파라미터로 직번역되며, 그 정확한 정렬은 **골든 테스트로 고정**한다(라이브러리 내부 동작을 기억에 의존해 단정하지 않는다).

---

## 6. 공통 인터페이스 **[D2]**

모든 분석기는 동일한 계약을 구현하여 파이프라인에 끼워 넣는다.

| 메서드 | 시그니처(개념) | 설명 |
| --- | --- | --- |
| `analyze` | `analyze(context) -> context` | 공유 컨텍스트를 입력받아 자신의 분석 결과로 보강 후 반환 |
| `name` | `-> str` | 분석기 식별자(로깅·진단용) |

**공유 컨텍스트(Context) 스키마**: `options: InspectionOptions`, `loader: Loader`(모드별 핸들 접근), `workbook_profile: WorkbookProfile`(점진 채움), `warnings: list[str]`. 분석기는 이전 단계가 채운 필드를 읽고 자신의 필드를 채운다. **단위 테스트 시에는 필요한 필드만 채운 부분 컨텍스트를 합성해 주입한다**(테스트 격리). 컨텍스트 합성 헬퍼는 `tests/conftest.py`가 제공한다.

**override 규약 [D2]**: 각 분석기는 자신이 담당하는 필드에 대해 `options`에 override가 있으면 산출을 생략하고 override 값을 기록하며 `provenance=manual`로 표시한다.

판정 불가·저신뢰 상태는 예외가 아닌 명시적 값(`None` + 신뢰도, 또는 `warnings` 누적)으로 표현하여 파이프라인이 중단 없이 진행되게 한다. 단, 로더의 손상·암호 차단(§4.1)은 예외로 즉시 중단한다.

---

## 7. 휴리스틱 명세 (v1 상수) **[D4]**

> 아래 상수는 v1 고정값이다. 외부 설정화(`InspectionOptions` 일부 노출 외)는 v1+. 모든 값은 픽스처 코퍼스로 보정한다.

### 7.1 헤더 탐지 점수화
- **표본**: 상단 `HEADER_SCAN_ROWS = 20`행.
- **점수식**: `score(r) = min(1.0, 0.5·non_empty_string_ratio(r) + 0.3·type_consistency(rows r+1..r+5)·(n_below/5) + 0.2·distinctness(r vs r+1..r+5) + time_series_code_header_bonus(r))`
  - `non_empty_string_ratio`: 행 내 비어있지 않은 **문자열** 셀 / 전체 사용 열 수.
  - `type_consistency`: 바로 아래 5개 행의 컬럼별 자료형 일관성 평균(0~1).
  - `n_below`: lookahead 창에서 **실제 관측된** 아래 행 수(0~5). 증거 가중(issue #8):
    창이 1행뿐이면 컬럼별 일관성이 자명하게 1.0이 되어 표본 하단의 데이터 행이
    진짜 헤더를 이기는 편향이 생기므로, 일관성 항은 관측된 증거량에 비례해서만
    인정한다.
  - `distinctness`: 헤더 후보 행과 아래 행들의 셀 길이·타입 패턴 차이(0~1).
  - `time_series_code_header_bonus`: 첫 non-empty 셀이 `Period`/`Date`/`Time`/`Year`/`Quarter`/`Month` 같은 시간축 라벨이고, 같은 행의 다수 라벨이 짧은 코드형 토큰이며, 아래 행들이 날짜형 축 + 숫자/공백 관측값 패턴을 보이면 최대 `0.25` 가점(issue #23). 파일명·시트명·벤더명은 보지 않는다.
- **판정**: 최고점 행을 헤더로 추정, `header_confidence = score`. `score < 0.5`(임계값, `InspectionOptions`로 조정 가능)이면 `needs_manual_header=True`.

### 7.2 경계 탐지 규칙
- **행 밀도**: `density(r) = non_empty_cells(r) / table_col_count`(표 열 구간 기준).
- **연속 빈 행**: `density=0`인 행이 `BLANK_RUN = 2`개 연속 → 데이터 종료/블록 구분. 단, 종료에 못 미치는 단발 빈 행도 `skip_rows`에 기록해 적재 프레임에 NaN 행이 새어들지 않게 한다.
- **저밀도 행**: `density < 0.3` → 소계/구분 행 후보. "단일 열만 채워진 행(`non_empty == 1`)" 규칙은 **표 폭이 3열 이상일 때만** 적용한다(1~2열 키-값/협폭 표의 정상 행 오탐 방지). 단, 첫 표 컬럼이 `Period`/`Date`/`Time` 계열 시간축이거나 빈 헤더 아래 실제 날짜/시간 키가 있고, 해당 행의 첫 표 셀이 실제 날짜/시간 값, 숫자 연도, 또는 한국식/미국식 날짜 문자열이면 wide sparse 시계열의 정상 관측치로 보존한다(issue #24). 날짜 문자열 정규식은 한국식/ISO 계열(`YYYY-MM-DD`, `YYYY.MM.DD`, `YYYY년 M월 D일`)과 미국식(`M/D/YYYY`, `Month D, YYYY`)까지만 인정한다. 이 예외는 소계/합계 키워드 매칭보다 약하므로 문자열 선두 라벨 subtotal은 계속 제외된다.
- **키워드**: `SKIP_KEYWORDS = ["합계","소계","총계","계","Total","Subtotal","Grand Total"]`. 매칭은 **행의 선두(첫 비어있지 않은) 라벨 셀** 기준으로, 다중자 키워드는 (대소문자 무시) `startswith`, 단일자 `"계"`는 **정확히 일치**할 때만 인정한다(임의 부분일치 금지 — `통계청`·`회계팀`·`Total Wine` 등 오탐 방지). `InspectionOptions.skip_keywords`로 대체/추가 가능.
- **열 경계**: 헤더 행에서 연속으로 채워진 열 구간을 표의 좌/우 경계로 본다. *병합 헤더로 선두 셀만 채워져 열 구간이 1칸으로 좁아지는 경우, v1에서는 경계 미해소(`None`)로 두고 병합 분석(§4.4) 이후로 미룬다.*

### 7.3 타입 추론
- **표본**: `TYPE_SAMPLE_ROWS = min(200, 데이터 행 수)`, 데이터 구간에서 균등 추출.
- **판정 순서**: (1) 결측 제외 → (2) 전부 날짜 서식/파싱 성공률 ≥ `0.95` → `date`, (3) 전부 숫자 파싱 성공 → 원 저장형이 문자열이면 `numeric_text`, 아니면 `number`, (4) 전부 비숫자 문자열 → `text`, (5) 어느 단일 타입도 ≥ `0.95` 미달 → `mixed`.
- **`null_ratio`** 분모 = 표본 데이터 행 수(소계·헤더 제외).

---

## 8. 횡단 관심사

- **성능**: 행 데이터는 read_only 스트리밍 + 표본 검사를 기본으로 한다. **구조 메타데이터(병합·치수)만 일반 모드로 1회 로드** **[D3]**. 전수 스캔은 회피한다. 정량 목표: 10만 행 파일에서 상주 메모리 증가 ≤ 200MB(표본 검사 기준), Phase 8에서 측정.
- **견고성**: 각 분석기 실패가 전체를 멈추지 않도록 저신뢰·판정 불가 상태를 `warnings`로 명시한다. 손상·암호 파일은 로더에서 조기 차단한다.
- **확장성**: 분석기를 전략 패턴으로 분리해 전용 탐지기 추가·교체를 지원한다.
- **읽기 전용성·멱등성**: 검사 단계는 원본을 변경하지 않으며(파일 바이트 해시 불변), 동일 입력에 동일 결과를 보장한다. 모든 워크북 핸들은 `close()`로 정리한다 **[D3]**.

---

## 9. 오류 처리 정책

| 상황 | 처리 |
| --- | --- |
| 파일 열기 실패(손상) | `CorruptWorkbookError`, `open_errors`에 기록 후 중단 |
| 암호 보호 파일 | `EncryptedWorkbookError` 명시적 반환, 사용자에게 안내 |
| 헤더 추정 실패 | `header_row=None`, 신뢰도 0, `needs_manual_header=True` |
| 비표 시트 | `is_tabular_candidate=False`, 적재 대상 제외 |
| 빈 시트 / 데이터 0행 | `data_start_row=None`, `ReadPlan`은 헤더만 또는 빈 계획, `warnings`에 기록 |
| 헤더 없는 데이터 시트 | override 없으면 `needs_manual_header=True`; override로 `header=None` 지정 시 컬럼명 없이 적재 |
| 치수 신뢰 불가(read_only) | `used_range_trusted=False`, 표본 스캔 보정 |

---

## 10. 제약 및 가정

- 입력은 `.xlsx`(OOXML)로 한정한다.
- **단일 시트 다중 표 (v1 한계, 실측 동작)**: 한 시트에 여러 표 블록이 있으면 v1은 **신뢰도 최고 헤더 1개 블록만** 적재하고 나머지는 `skiprows`로 흡수해 버린다. ⚠️ 현재 구현은 위치(최상위)가 아니라 **점수**로 헤더를 고르므로 — 본 절의 초기 의도("최상위 1개")와 달리 — *두 번째/아래쪽 표가 선택될 수 있고*, 누락된 블록에 대한 **경고를 내지 않는다(조용한 누락)**. 다중 표가 의심되면 시트를 분리하거나 `InspectionOptions`로 헤더/경계를 수동 지정해야 한다. 근본 해소는 §11.
- 적재 라이브러리는 **pandas 3.x**(검증 환경 3.0.3)를 가정한다. 2.x 대비 문자열 dtype 기본값·Copy-on-Write 등 동작 차이가 있으므로 적재 검증은 3.x에서 수행한다.
- 실행 환경은 외부 관리형 Python일 수 있어 **venv 사용을 전제**한다(Python 3.14 검증 완료).
- read_only 워크시트에는 `merged_cells`·신뢰 가능한 치수가 없을 수 있다 **[D3]**.

---

## 11. 향후 확장 / 미해결 사항

- 수식 탐지(Formula Detector) 및 캐시값 공백 분기 **[D6]**
- 다단 헤더(`header: list[int]`, MultiIndex, 병합 forward-fill) **[D6]**
- **단일 시트 내 다중 표 블록 분할 인식**: 블록 세분화기(빈 줄 런 `BLANK_RUN`·헤더 재출현·컬럼 시그니처 변화로 used_range를 N개 블록으로 분할) → 블록마다 Header/Boundary/Type 독립 실행 → `SheetProfile.blocks: list[TableBlock]`(각자 ReadPlan) → 어댑터가 블록당 DataFrame 산출. 가로 나란한(side-by-side) 표, "두 번째 표 vs 첫 표의 연속" 구분이 난점. *전면 도입 전 최소 안전장치*로, 추가 블록 감지 시 `warnings`에 명시하고 단일 선택을 '최상위 블록'으로 결정화하는 것을 우선 고려(현재의 조용한 누락 제거).
- 신뢰도 임계값·휴리스틱 가중치의 전면 외부 설정화
- `.xls` 및 CSV/TSV 입력으로의 검사기 확장
- 검사 결과 캐싱(동일 파일 재검사 비용 절감)
