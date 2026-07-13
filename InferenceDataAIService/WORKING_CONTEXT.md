# InferenceDataAIService — 목표 및 작업 이력

최종 갱신: 2026-07-11

이 문서는 이 서비스의 현재 상태, 사용자와 합의한 목표, 그리고 다음 세션이 바로 이어서 작업할 수 있는 재개 지점이다. `GENERALIZATION_PLAN.md`는 설계 원칙, `IMPLEMENTATION_INTENT.md`는 초기 의도를 설명하며, 이 문서는 실제 진행 이력과 실행 우선순위를 기록한다.

## 1. 사용자 목표

약 900개 이상의 서로 다른 Excel 검토보고서에서 다음을 일관된 근거와 함께 DB에 저장하고 활용한다.

- 검토 목적과 보고서 유형
- 비교군(Test/variant)과 대조군(Control/baseline)
- 각 군의 조건: 모델, 공정, 자재, 공급처, Lot, 기간, 설비, 표본 등
- 검증 항목과 지표: NG, Input, NG rate/ppm, 측정값, 규격, 결론
- 비교 차이와 제한 사항
- 모든 주장에 연결된 원본 Excel 시트·셀 범위

목표는 특정 Excel 양식용 parser가 아니다. Excel의 원본 구조를 보존하고, AI가 보고서별 의미와 비교 설계를 제안하며, 프로그램이 산술·근거·저장을 검증하는 재현 가능한 분석 파이프라인이다.

## 2. 합의된 목표 아키텍처

`CLI`와 `AI`는 같은 것이 아니다.

```text
Excel 입력
  → Python CLI: Excel COM으로 원본 grid/병합/셀을 추출하고 DB에 저장
  → Python CLI: 원본과 후보 힌트를 포함한 AI 입력 packet 생성
  → AI: 검토 목적, 비교군/대조군, 지표, 근거, 제한을 analysis-plan JSON 초안으로 제안
  → Python CLI: JSON의 구조, 원본 근거, NG/Input/ppm, delta 산식을 검증
  → DB: NEEDS_REVIEW 상태로 저장
  → 사람(필요 시 2차 AI): 비교 설계의 업무적 타당성을 승인/수정/거절
  → DB: VERIFIED / REJECTED / EXCLUDED
```

역할 분리:

- **AI**: “무엇을 비교해야 하는가?”를 원본 근거와 함께 제안한다. 근거가 없으면 `UNKNOWN` 또는 `NEEDS_REVIEW`로 남긴다.
- **Python CLI**: “어떻게 정확히 읽고, 계산하고, 검증하고, 저장하는가?”를 수행한다.
- **사람**: 모델/Lot/기간/라인/표본선정 등 업무상 비교 타당성을 최종 승인한다.

AI가 Excel마다 Python 코드를 새로 생성하는 방식은 운영 목표가 아니다. 개발 중 새 패턴을 발견하면 프로그램과 스키마를 개선할 수 있지만, 운영 중에는 AI가 **코드가 아닌 versioned analysis-plan JSON**을 생성해야 한다.

## 3. 현재 CLI가 실제로 하는 일

### 자동 구현됨

- `com-index`
  - Excel COM으로 workbook/sheet/row/cell/merge 정보를 fixed grid로 추출한다.
  - 원본 JSON과 universal-grid SQLite에 저장한다.
  - 원본 변경 fingerprint, 재개 처리, 적재 후 구조 검증을 지원한다.
  - **의미상 Test/Control 또는 결론을 자동 확정하지 않는다.**
- `quick-index`
  - 기존 MicroSpeaker indexer를 실행해 Test/Normal/NG/Input 등 후보를 빠르게 찾는다.
  - 후보 힌트이며 최종 분석은 아니다.
- `analysis-import`
  - 이미 작성된 `universal-analysis-v1` manifest JSON을 analysis 테이블에 저장한다.
- `analysis-verify`
  - evidence range 존재, 원본 fingerprint freshness, NG/Input/ppm 산식, comparison delta/상대차를 검증한다.
- `build-packet`
  - 원본 workbook의 행·셀·병합 정보를 AI에 전달할 packet으로 export한다.
- `inference_data_ai_ui.py`
  - 기존 CLI를 다시 구현하지 않는 Windows Tkinter 드래그앤드롭 운영 화면이다.
  - `com-index`(원본 universal DB/COM JSON)와 선택된 `quick-index`(후보 DB/HTML)를 같은 설정으로 순차 실행한다.
  - 파일/폴더를 목록에 드롭하면 원본 절대 경로를 유지하며, fingerprint 재개 처리로 기존 파일은 CLI가 skip한다.
  - `outputs/ui-state.json`에 마지막 목록/설정, `outputs/ui-run-history.jsonl`에 실행 이력을 저장한다.

### 아직 구현되지 않음

- 원본 grid/quick-index 후보를 AI에 자동 전달하고, AI가 `universal-analysis-v1` analysis-plan JSON 초안을 생성하는 CLI 명령
- AI 초안의 schema 후보 필드를 모아 스키마 확장을 관리하는 기능
- 사람 승인 UI/CLI 및 `NEEDS_REVIEW → VERIFIED/REJECTED` 승인 이력
- 989개 전체에 대한 자동 packet 생성·AI 초안·배치 검증 실행

## 4. 스키마 원칙

세 층을 유지한다.

```text
1. 원본 구조층 (안정적): workbooks, worksheets, grid_sheet_rows, grid_sheet_cells, merge_ranges
2. 공통 분석층 (신중히 확장): analysis_reports, review_items, cohorts, metrics, values, comparisons, conclusions, evidence
3. 보고서별 확장층 (우선 JSON): attributes_json, details_json, notes_json, limitations_json
```

900개 분석에서 새 조건이 나오면 AI는 임의로 DB DDL을 변경하지 않는다.

1. 새 속성을 analysis JSON의 확장 필드와 `schema candidate`로 기록한다.
2. 여러 보고서에서 반복되는지, 검색/집계/검증이 필요한지 평가한다.
3. 승인된 반복 속성만 migration으로 공통 스키마에 승격한다.

우선 보강 후보:

- `comparison_design`: test/control, before/after, paired, multi-group 등의 비교 설계와 계산 방법
- `cohort_selection`: 선정/제외 기준, 모델, Lot, 라인, 기간, 표본 단위, 매칭 기준, 교란 요인
- 검색 가능한 `cohort_attributes`: 현재 JSON에만 있는 반복 속성의 정규화 저장
- AI 초안/승인 이력: 생성 모델·프롬프트 버전·검토자·승인 시각·사유

## 5. 현재 DB 및 산출물 상태

DB:

`outputs/universal-grid/InputDataFinish.sqlite`

2026-07-11 확인 수치:

- 원본 workbook: 5건
- analysis report: 4건
- review item: 11건
- cohort: 29건
- metric: 29건
- metric value: 69건
- comparison: 40건
- conclusion: 15건
- evidence: 146건

분석 report:

| ID | Workbook | 분석 키 | 상태 | 결정 |
|---:|---|---|---|---|
| 1 | BRS-2015 New Bond | `brs2015-g06-0003-new-bond-validation` | VERIFIED | CAN_USE |
| 2 | TIU 본딩량 감소 | `tiu-l5s3-01-reduce-insulation-bonding-bako` | VERIFIED | IMPROVED |
| 3 | BRS-161016 YK 공급처 | `brs161016-yk-supplier-qualification` | VERIFIED | CAN_NOT_USE |
| 4 | MSU-L20S15-07 Bonding amount | `msu-l20s15-07-bonding-amount-function-draft` | NEEDS_REVIEW | NEEDS_REVIEW |

대표 manifest:

- `outputs/analysis-manifests/BRS2015_G06_0003_analysis.json`
- `outputs/analysis-manifests/TIU_L5S3_01_reduce_bonding_analysis.json`
- `outputs/analysis-manifests/BRS161016_YK_supplier_qualification_analysis.json`
- `outputs/analysis-manifests/MSU_L20S15_07_bonding_change_draft_analysis.json`

## 6. 이번 세션 이력

1. 기존 CLI, 스키마, 문서, SQLite를 읽기 전용으로 점검했다.
2. 기존 3개 analysis report는 CLI가 AI 의미분석을 자동 실행한 결과가 아니라, 이미 작성된 manifest를 `analysis-import`로 저장·검증한 것임을 확인했다.
3. 별도 Excel `00. Report TEST new dry machine...xlsx`를 COM으로 원본 grid만 적재했다(workbook ID 4). 지표 단위와 명확한 대조군이 없어 analysis manifest는 만들지 않았다.
4. 비교군이 명확한 `009.MSU-L20S15-07 Report test change bonding amount VP+CD, Coil+CD...xlsx`를 COM으로 적재했다(workbook ID 5, Sheet1/Sheet2, 2,211 cells).
5. AI 보조로 해당 workbook을 읽어 `MSU_L20S15_07_bonding_change_draft_analysis.json` 초안을 만들었다.
   - Sheet1: Max/Min/1.5mg 조건을 Sheet1 Normal line과 각각 비교
   - Sheet2: Nozzle VZ05, Plasma CD 2 edge, reduced bonding 조건을 Sheet2 Normal line과 각각 비교
   - 서로 다른 Sheet의 Normal line은 합치지 않았다.
   - 원본의 cohort matching 근거가 부족하므로 결론은 `NEEDS_REVIEW`로만 저장했다.
6. 해당 manifest는 `analysis-import` 및 `analysis-verify`를 통과했다(근거 range, ppm, delta 오류 0).
7. 결과를 기존 검증 대시보드와 같은 표 중심 HTML 형식으로 만들었다.
   - `outputs/MSU_L20S15_07_bonding_change_draft_dashboard.html`
   - 이것은 CLI의 자동 HTML renderer가 아니라 이번 세션에서 만든 정적 표시용 파일이다.
8. 900개와 이후 추가 Excel을 같은 방식으로 적재하기 위한 `inference_data_ai_ui.py`를 추가했다.
   - 원본 DB/JSON과 quick-index 후보 HTML을 공통 CLI로 실행한다.
   - AI 의미 분석 JSON과 분석 결론 HTML은 아직 자동 생성하지 않으며, UI도 이를 자동 결과라고 표시하지 않는다.
9. `InferenceDataAIServiceUI.exe`로 GUI 런처를 PyInstaller 패키징했다.
   - EXE는 GUI를 Python 명령 없이 시작하지만, 기존 `inference_data_ai_cli.py`와 Microsoft Excel COM 추출을 호출하므로 서비스 폴더에 유지하고 로컬 Python·Excel 환경을 사용한다.
   - Python 자동 탐색이 안 되면 `INFERENCE_DATA_AI_PYTHON` 환경 변수로 `python.exe`를 지정한다.
10. Browse 의존을 제거하고 파일 목록·드래그앤드롭 방식으로 UI를 갱신했다.
    - 새 실행 파일은 `InferenceDataAIServiceUI_DragDrop.exe`이다.
    - 기존 EXE가 열려 있어 덮어쓰지 않고 새 이름으로 배포했다.
11. 파일 목록 UI에서 단일 Excel을 선택했을 때 quick-index가 폴더 전용 `--input-dir`를 받아 실패하던 문제를 수정했다.
    - `quick-index` 래퍼는 파일이면 하위 indexer의 `--input-file`, 폴더면 `--input-dir`를 사용한다.

## 7. 다음 구현 순서

1. `analysis-plan JSON`의 명시적 계약을 v2로 설계한다.
   - `comparisonDesign`, `cohortSelection`, `unknown/missingConditions`, `schemaCandidates`, AI 생성 메타데이터를 포함한다.
2. universal DB migration을 추가한다.
   - 우선 `comparison_design`과 `cohort_selection`을 근거와 연결한다.
3. `build-analysis-input` CLI를 추가한다.
   - 원본 grid와 quick-index 후보를 제한된 packet으로 만든다.
4. AI analysis runner를 추가한다.
   - packet → analysis-plan JSON 초안; 직접 DB를 쓰지 않는다.
5. import/verify를 강화한다.
   - 비교군 선정 근거, 필수 조건 누락, 단위/분모 혼합, 허용되지 않은 상태 전이를 검사한다.
6. 10~20개 서로 다른 유형을 먼저 처리해 schema candidate를 수집·검토한다.
7. 그 후 989개 전체를 배치 처리한다. 자동 결과는 기본적으로 `NEEDS_REVIEW`이며, 검증/승인을 거쳐야 `VERIFIED`가 된다.
8. 현재 UI는 원본 적재/후보 HTML 운영 화면이다. 위 3~7 단계의 AI 분석 runner가 안정화되면 UI에는 그 검증된 단일 명령만 추가한다.

## 8. 재개 시 주의사항

- 아직 JinoSupporter Web에 직접 연결하지 않는다.
- 원본 Excel은 read-only로만 열고 저장하지 않는다.
- raw source DB에는 `--covered-cell-mode blank`를 사용하고 `--sparse`를 사용하지 않는다.
- 출력은 이 서비스 폴더 하위에만 둔다.
- CLI에 상대 `--db` 경로를 넘기면 서비스 폴더 기준으로 다시 붙을 수 있으므로, 현재는 absolute DB path 또는 기본값을 사용한다.
- `NEEDS_REVIEW` 분석은 숫자/근거 검증 통과를 뜻할 뿐, 업무적 비교 타당성이나 인과 결론의 승인이 아니다.
