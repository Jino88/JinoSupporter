# 구조 재사용 기반 증분 Excel 분석 설계

## 1. 설계 상태

- 상태: 신규 361건 반복 정량 구조 batch 및 canonical 운영 DB 반영 완료, DB coverage 361/361
- 대상: 기존 989개 Excel 분석 자산을 재사용하여 신규 361개 Excel을 증분 처리하는 파이프라인
- 핵심 결정:
  - AI는 **재사용할 구조/추출 레시피의 선택**과 **새 변형 구조의 레시피 제안**만 담당한다.
  - 셀 값, 수식, 단위, 통계값, 근거 셀 주소의 추출은 프로그램이 담당한다.
  - 이미 검증된 구조와 정확히 일치하는 파일은 AI를 호출하지 않는다.
  - 구조를 확신할 수 없거나 검증에 실패한 파일은 잘못된 결과를 강행하지 않고 격리한다.

## 2. 문제 정의

기존 989개 분석 결과는 파일별 결과 캐시로는 재사용할 수 있어도, 신규 파일과 파일명이 다르기 때문에 직접 캐시 적중은 발생하지 않는다. 그러나 989개를 처리하면서 얻은 다음 자산은 신규 파일에 재사용할 수 있다.

- 시트 구성과 표 배치
- 제목, 헤더, 행·열 역할을 나타내는 구조적 앵커
- 반복 블록과 데이터 영역의 경계
- 어떤 셀을 어떤 파라미터로 변환하는지에 대한 추출 규칙
- 수식, 단위, 근거 셀을 검증하는 규칙

현재 파이프라인은 사전 구조 분류 결과를 생성하지만, 이후 단계에는 주로 파일 경로만 전달한다. 그 결과 `formSignatureId`, `formFamilyId`, registry 결정과 구조 계약이 실제 추출 단계에서 소실되고, 신규 파일마다 큰 컨텍스트를 다시 AI에 전달하여 DRAFT를 생성한다. 이 설계는 그 단절을 제거한다.

## 3. 목표와 비목표

### 목표

1. 기존 989개의 검증 결과로 재사용 가능한 구조 템플릿과 실행 가능한 추출 레시피를 만든다.
2. 신규 파일은 가장 가까운 템플릿을 먼저 찾고, 검증된 레시피를 프로그램으로 실행한다.
3. AI 호출을 파일당 반복 분석이 아니라 구조 불확실성 해소에만 사용한다.
4. 모든 추출값에 원본 파일 SHA, 시트, 셀 주소 또는 범위가 연결되게 한다.
5. 기존 table-first history와 canonical DB 양쪽에 동일한 정규화 결과를 적재한다.
6. 잘못된 템플릿 자동 적용을 막는 fail-closed 검증 절차를 둔다.

### 비목표

- 기존 989개 분석 결과를 폐기하거나 처음부터 다시 AI 분석하지 않는다.
- AI가 원본 수치나 수식을 읽어 최종 값을 직접 작성하게 하지 않는다.
- 모든 Excel을 하나의 고정 좌표 규칙으로 처리하지 않는다.
- 구조 검증 실패 시 전체 워크북 AI DRAFT로 자동 폴백하지 않는다.

## 4. 전체 아키텍처

```text
기존 989개
  capture + table-first request/result + history
          |
          v
구조 지문 v2 생성 및 군집화
          |
          v
FormTemplate + ExtractionRecipe 생성
          |
          v
과거 파일 replay 검증 ---- 실패 ----> 레시피 수정/격리
          |
       승인 registry
          |
          +--------------------------------------+
                                                 |
신규 Excel                                      |
  |                                              |
  v                                              v
capture -> 구조 지문 -> Top-K 템플릿 검색 -> 매칭 결정
                                      |          |
                       정확 일치: AI 0회          |
                       모호함: AI 최대 1회         |
                                      v          |
                              레시피 프로그램 실행
                                      |
                                      v
                               결정론적 검증
                              /             \
                          성공               실패
                           |                  |
                           v                  v
                    정규화 추출 결과      패치 후보/격리
                       /       \
             table-first       canonical DB
```

### 4.1 실제 989건 검증 후 계층 보정

실제 데이터로 검증한 결과, workbook 전체 fingerprint는 시트 조합과 보고서별 부가 표 때문에 과도하게 세분화됐다.

- 과거 989 workbooks에서 exact workbook structures는 973개였다.
- 신규 361 workbooks의 과거 최상위 workbook 유사도는 평균 0.4490, 최대 0.7375였다.
- 반면 과거 table-first가 분리한 5,546 tables에서는 508개 반복 block structures가 2,633 tables와 859 workbooks를 포괄했다.
- 비-TEXT 정량 반복 구조도 357개, 1,284 tables, 415 workbooks에서 확인됐다.

따라서 구조 재사용 단위를 다음 3계층으로 확정한다.

```text
1. Source cache
   source SHA + recipe version 일치
   -> 결과 전체 재사용

2. Workbook profile
   sheet 구성과 주요 anchor
   -> 후보 recipe 묶음 검색과 처리 라우팅에 사용
   -> 전체 workbook 결과의 자동 재사용 기준으로 사용하지 않음

3. Table/block recipe structure
   상대 cell-kind layout + merge 기하 + numeric column role
   + 상대 metric 열별 정규화 header signature
   -> 실제 extraction recipe 선택과 replay의 기본 단위
```

한 workbook이 과거와 완전히 같은 시트 조합을 갖지 않더라도, 내부의 function result, NG rate, dimension, NTI, reliability 같은 정량 block은 기존 recipe를 독립적으로 재사용할 수 있다. 단, 기하 구조 digest가 같아도 metric header가 다른 표가 실제로 존재하므로 digest만으로 recipe를 공유하지 않는다. 동일 block fingerprint를 상대 metric 열별 header signature로 한 번 더 분리한 **recipe structure당 한 번**만 의미 계약을 결정한다.

신규 361건 파일럿에서 하나의 기하 구조에 S931의 `Bako/Hearing/Hearing 2/Hearing 3/Air leak`, X626의 `Gauss/Air leak/Frequency/Bako/Hearing`, X526의 `Gauss 1st/Gauss 2nd/...`가 함께 묶이는 사례가 발견됐다. metric header signature를 추가한 뒤 첫 AI 판단은 의미가 일치하는 S931 variant 6개에만 적용됐고 6/6 deterministic replay를 통과했다. 이 분리 없이 기하 digest만 재사용하는 것은 금지한다.

실제 첫 검증에서는 과거 14 workbooks의 동일 COMPARISON block 15개에서 다음 9개 metric 열의 상대 위치와 의미가 일치했다.

- SPL
- THD
- SPL+THD
- SPL+THD+F0
- NOISE
- TOUCH
- HOHD
- TOTAL NG
- TOTAL NG RATE

이 구조에서 프로그램 소유 numeric facts와 evidence range가 15/15 replay를 통과했다. 이는 workbook 전체가 달라도 table/block recipe를 재사용할 수 있다는 첫 승인 근거다.

## 5. 핵심 데이터 계약

모든 계약은 버전을 명시하고, 변경 시 기존 결과를 재현할 수 있어야 한다.

### 5.1 `StructureFingerprintV2`

원본 값 자체가 아니라 구조 비교에 필요한 특징을 담는다.

```json
{
  "schemaVersion": "excel-structure-fingerprint-v2",
  "sourceSha256": "...",
  "workbook": {
    "sheetCount": 4,
    "visibleSheetCount": 4,
    "sheetOrderRoles": ["cover", "input", "calculation", "result"]
  },
  "sheets": [
    {
      "sheetKey": "normalized-sheet-1",
      "titleTokens": ["result"],
      "usedRangeBucket": {"rows": "65-96", "columns": "9-16"},
      "mergedGeometry": ["R1C1:R2C8"],
      "anchorSketches": [
        {"token": "시험결과", "relativeBand": "top", "occurrences": 1}
      ],
      "headerRoleSketch": ["label", "unit", "value", "criterion"],
      "tableRegionSketches": [
        {"headerDepth": 2, "columnCount": 8, "rowMode": "repeated-items"}
      ],
      "formulaPatternHashes": ["..."],
      "numberFormatRoleSketch": ["text", "decimal-2", "percent"]
    }
  ],
  "fingerprintSha256": "..."
}
```

현재 v1보다 다음을 강화한다.

- 단순 행·열 버킷 외에 정규화된 앵커 위치
- 병합 셀의 상대 기하
- 헤더 깊이와 열 역할
- 반복 데이터 블록의 형태
- 수식의 상대 참조 패턴
- 숫자 형식과 단위 역할
- 숨김 시트와 표 시트의 역할

날짜, 일련번호, 시험값처럼 파일별로 변하는 값은 구조 지문 토큰에서 제외한다.

### 5.2 `FormTemplateV1`

한 구조군의 허용 범위와 검증된 예제를 표현한다.

```json
{
  "schemaVersion": "excel-form-template-v1",
  "templateId": "tpl-...",
  "familyId": "family-...",
  "templateVersion": 1,
  "status": "APPROVED",
  "acceptedFingerprintEnvelope": {
    "requiredAnchors": ["시험결과", "판정"],
    "allowedSheetCount": [3, 5],
    "allowedHeaderDepth": [1, 2],
    "optionalSheetRoles": ["cover"]
  },
  "exampleSources": [
    {"sourceSha256": "...", "quality": "VERIFIED_REPLAY"}
  ],
  "recipeRefs": ["recipe-...@3"],
  "quality": {
    "replayFileCount": 42,
    "replayPassCount": 42,
    "falsePositiveCount": 0
  }
}
```

### 5.3 `ExtractionRecipeV1`

기존 registry의 설명형 계약을 프로그램이 실행 가능한 selector DSL로 바꾼다.

```json
{
  "schemaVersion": "excel-extraction-recipe-v1",
  "recipeId": "recipe-...",
  "recipeVersion": 3,
  "templateId": "tpl-...",
  "sheetSelectors": [
    {
      "id": "resultSheet",
      "titleAliases": ["시험결과", "결과"],
      "requiredAnchors": ["판정"],
      "fallbackRole": "tabular-result",
      "cardinality": "exactly-one"
    }
  ],
  "anchors": [
    {
      "id": "resultHeader",
      "sheet": "resultSheet",
      "textRegex": "^(시험)?결과$",
      "normalized": true,
      "uniqueness": "one"
    }
  ],
  "regions": [
    {
      "id": "resultTable",
      "sheet": "resultSheet",
      "start": {"below": "resultHeader", "rows": 1},
      "stop": {"firstBlankKeyColumnRun": 2},
      "headerDepth": 2,
      "repeatMode": "rows"
    }
  ],
  "axes": {
    "rowKey": {"region": "resultTable", "columnRole": "item-label"},
    "columns": [
      {"role": "unit", "headerAliases": ["단위"]},
      {"role": "value", "headerAliases": ["결과값", "측정값"]},
      {"role": "criterion", "headerAliases": ["기준"]},
      {"role": "judgement", "headerAliases": ["판정"]}
    ]
  },
  "fields": [
    {
      "parameter": "measurement",
      "source": {
        "region": "resultTable",
        "row": "each",
        "valueColumnRole": "value",
        "labelColumnRole": "item-label",
        "unitColumnRole": "unit"
      },
      "valueType": "decimal",
      "evidence": "exact-source-cell",
      "required": true
    }
  ],
  "validationRules": [
    {"rule": "required-field-coverage", "minimum": 1.0},
    {"rule": "evidence-cell-exists"},
    {"rule": "formula-reference-consistency"}
  ]
}
```

레시피의 좌표는 가급적 절대 셀 주소가 아니라 앵커 상대 위치, 헤더 역할, 종료 조건으로 표현한다. 불가피한 고정 좌표는 템플릿 버전에 종속시키고 별도 검증 규칙을 둔다.

지원해야 할 selector 유형:

- 시트: 제목 alias, 순서, 가시성, 구조 역할, 필수 앵커
- 앵커: 정규화 문자열, 정규식, 병합 영역, 고유성
- 영역: 앵커 기준 상대 범위, 빈 행/새 헤더/합계 행 종료 조건
- 반복: 행 반복, 열 반복, 블록 반복
- 필드: scalar, list, table, series, ratio, conclusion
- 값 형식: text, integer, decimal, percent, date, formula-result
- 단위: 원문 단위 보존, 명시된 경우에만 변환
- 근거: 정확한 셀, 병합 대표 셀, 수식 셀과 참조 셀

### 5.4 `TemplateMatchDecisionV1`

AI가 관여하더라도 구조 결정만 기록한다.

```json
{
  "schemaVersion": "excel-template-match-decision-v1",
  "sourceSha256": "...",
  "fingerprintSha256": "...",
  "decision": "EXACT_REUSE",
  "selectedTemplate": "tpl-...@1",
  "selectedRecipe": "recipe-...@3",
  "topCandidates": [
    {"template": "tpl-...@1", "score": 0.981}
  ],
  "requiredAnchorCoverage": 1.0,
  "ai": {
    "used": false,
    "reason": null,
    "model": null,
    "promptVersion": null
  },
  "observedDeviations": []
}
```

허용 결정:

- `EXACT_REUSE`: 결정론적 고신뢰 매칭, AI 0회
- `AI_CONFIRMED_REUSE`: Top-K 후보 중 AI가 선택, AI 최대 1회
- `VARIANT_PATCH_REQUIRED`: 기존 레시피의 구조 변형
- `NEW_TEMPLATE_REQUIRED`: 기존 구조군에 안전하게 포함 불가
- `QUARANTINED`: 입력 손상, DRM/capture 이상 또는 결정 불가

### 5.5 `ExtractionResultV1`

프로그램 실행 결과의 단일 정규화 계약이다.

```json
{
  "schemaVersion": "excel-deterministic-extraction-v1",
  "sourceSha256": "...",
  "template": "tpl-...@1",
  "recipe": "recipe-...@3",
  "engineVersion": "...",
  "parameters": [
    {
      "name": "measurement",
      "label": "압축강도",
      "value": 12.34,
      "unit": "MPa",
      "evidence": [
        {"sheet": "시험결과", "cell": "D18", "rawValue": 12.34},
        {"sheet": "시험결과", "cell": "C18", "rawValue": "MPa"}
      ]
    }
  ]
}
```

숫자, 수식 결과, 통계, aggregate check, 근거 셀은 모두 capture 데이터에서 프로그램이 채운다. AI 출력에서 숫자를 복사하지 않는다.

### 5.6 `ValidationReportV1`

```json
{
  "schemaVersion": "excel-extraction-validation-v1",
  "sourceSha256": "...",
  "recipe": "recipe-...@3",
  "status": "VERIFIED",
  "checks": {
    "anchorUniqueness": "PASS",
    "regionBounds": "PASS",
    "requiredCoverage": "PASS",
    "evidenceExists": "PASS",
    "valueType": "PASS",
    "formulaConsistency": "PASS",
    "unitConsistency": "PASS",
    "unusedQuantitativeCellReview": "PASS"
  },
  "failureCodes": []
}
```

대표 실패 코드:

- `REQUIRED_ANCHOR_MISSING`
- `ANCHOR_NOT_UNIQUE`
- `TABLE_SHAPE_OUT_OF_ENVELOPE`
- `REQUIRED_FIELD_MISSING`
- `EVIDENCE_CELL_NOT_FOUND`
- `VALUE_TYPE_MISMATCH`
- `FORMULA_PATTERN_CHANGED`
- `UNIT_CONFLICT`
- `UNEXPECTED_QUANTITATIVE_REGION`

### 5.7 `RecipePatchV1`

새 변형이 발견됐을 때 기존 레시피 전체를 다시 생성하지 않고 구조 차이만 표현한다.

- 추가/변경된 앵커
- 시트 alias 추가
- 영역 시작·종료 조건 변경
- 열 역할 alias 추가
- 선택적 필드 또는 반복 블록 추가
- 적용 가능한 fingerprint 조건
- replay 검증 결과

패치는 승인 후 새 recipe version으로 승격한다. 기존 버전과 결과는 그대로 보존한다.

## 6. 템플릿 검색과 매칭

### 6.1 1차 후보 검색

전체 registry를 AI에 전달하지 않는다. 프로그램이 다음 hard gate로 후보를 줄인다.

- 필수 시트 역할과 시트 수 범위
- 필수 앵커 존재 여부
- 주요 표의 행·열 형태
- 병합 셀 기하
- 수식 패턴 호환성
- 헤더 깊이와 열 역할

후보 점수의 초기 가중치는 다음처럼 시작하고, 989개 replay 결과로 보정한다.

| 특징 | 초기 가중치 |
|---|---:|
| 필수·선택 앵커 및 상대 위치 | 0.30 |
| 표 영역 기하와 반복 구조 | 0.20 |
| 헤더/행·열 역할 | 0.15 |
| 수식 상대 참조 패턴 | 0.15 |
| 병합 셀 기하 | 0.10 |
| 시트 수와 used-range 형태 | 0.10 |

### 6.2 초기 결정 구간

아래 수치는 고정 정책이 아니라 989개 검증 데이터로 보정할 초기값이다.

- `score >= 0.97`, 필수 앵커 100%, hard gate 통과: `EXACT_REUSE`, AI 0회
- `0.90 <= score < 0.97`: Top-3 후보만 AI에 전달하여 최대 1회 판정
- `0.75 <= score < 0.90`: 기존 family의 variant 가능성 평가, 필요 시 레시피 패치
- `score < 0.75`: `NEW_TEMPLATE_REQUIRED`

AI 입력에는 원본 전체 워크북이 아니라 다음만 제공한다.

- 신규 구조 지문과 축약된 레이아웃
- Top-3 템플릿의 구조 계약
- 두 구조의 차이
- 값이 제거된 소수의 헤더/앵커 예시

AI 출력에는 템플릿/레시피 선택, 구조 차이, 신뢰도만 허용한다. 최종 파라미터 값은 허용하지 않는다.

## 7. 기존 989개 자산의 부트스트랩

### 7.1 입력 자산

- 기존 capture v2 결과
- table-first request와 AI semantic result
- 프로그램이 투영한 table-first 분석 결과
- table-first history DB의 source/evidence
- 현재 canonical DB의 workbook analysis와 knowledge study

### 7.2 부트스트랩 절차

1. 989개 각각에 `StructureFingerprintV2`를 생성한다.
2. hard gate와 구조 점수로 후보 family를 군집화한다.
3. 군집 내 공통 앵커, 표 영역, 행·열 역할을 프로그램으로 추출한다.
4. 기존 table-first 결과에서 의미 역할을 연결하되, 수치와 근거는 capture에서 재구성한다.
5. 실행 가능한 `ExtractionRecipeV1` 초안을 만든다.
6. 동일 family의 모든 과거 예제에 레시피를 replay한다.
7. 과거 table-first의 코드 소유 값·근거와 비교한다.
8. 전부 통과한 범위만 자동 재사용 가능한 `APPROVED` 템플릿으로 승격한다.
9. 하나라도 구조적 충돌이 있으면 family를 분리하거나 variant recipe를 만든다.

AI는 군집의 의미 역할 또는 모호한 열 매핑을 정할 때 **family당 한 번** 사용할 수 있다. 개별 파일의 수치를 재분석하지 않는다.

### 7.3 replay 승인 조건

- 추출 수치와 원본 셀 값 일치율 100%
- 근거 시트·셀 주소 일치율 100%
- 필수 파라미터 누락 0건
- 레시피 적용 대상에서 구조 오탐 0건
- 수식 패턴 변경 미탐 0건
- 기존 query/golden 결과에 회귀 없음

조건을 만족하지 못한 family는 자동 적용 대상에서 제외한다.

## 8. 신규 파일 처리

1. 원본 SHA와 capture 상태를 확인한다.
2. 기존 extraction cache가 있으면 즉시 재사용한다.
3. 구조 지문을 생성한다.
4. registry에서 Top-K 템플릿을 결정론적으로 검색한다.
5. 정확 일치면 AI 없이 레시피를 실행한다.
6. 모호한 경우에만 AI가 Top-K 중 선택한다.
7. 선택한 레시피를 프로그램으로 실행한다.
8. 모든 값에 정확한 evidence cell을 연결한다.
9. 구조·값·수식·단위 검증을 수행한다.
10. 검증 성공 결과만 table-first와 canonical adapter에 전달한다.
11. 검증 실패는 `VARIANT_PATCH_REQUIRED` 또는 `QUARANTINED`로 남긴다.

같은 variant가 여러 파일에서 반복되면 첫 파일에서 승인된 패치를 새 recipe version으로 만들고, 나머지는 AI 없이 처리한다.

## 9. AI 사용량 정책

| 상황 | AI 호출 상한 | 처리 |
|---|---:|---|
| exact fingerprint/고신뢰 template | 0 | 프로그램 추출 |
| Top-K 사이의 모호한 매칭 | recipe structure당 1 | 구조 선택만 수행 |
| 처음 발견된 variant | variant당 1 | recipe patch 제안 |
| 완전히 새로운 family | family당 1 | 초기 구조 계약 제안 |
| 숫자·수식·근거 추출 | 0 | 프로그램 전담 |
| JSON 파싱 실패 | 자동 재호출 없음 | 명시적 실패/격리 |
| 검증 실패 | 전체 DRAFT 폴백 없음 | 패치 또는 격리 |

기본 reasoning effort는 `low`로 둔다. 한 요청 실패 후 동일 파일을 내부 3회, 외부 4회 반복하는 형태의 retry 증폭은 제거한다.

반드시 기록할 원가 지표:

- 파일별·family별 AI 호출 횟수
- 모델, reasoning effort, prompt version
- prompt/response byte와 token 사용량
- exact match cache hit
- AI match cache hit
- extraction cache hit
- recipe version별 처리 파일 수
- family 생성 1회 비용이 몇 파일에 상각됐는지

## 10. 캐시와 멱등성

### 매칭 캐시 키

```text
fingerprintSha256
+ semanticHeaderSha256
+ registryVersion
+ matcherVersion
+ candidateTemplateVersions
+ aiMatchPromptVersion(사용한 경우)
```

### 추출 캐시 키

```text
sourceSha256
+ recipeId
+ recipeVersion
+ extractionEngineVersion
```

### 검증 캐시 키

```text
extractionResultSha256
+ validationRuleSetVersion
```

AI prompt 버전이 바뀌어도 선택된 recipe와 extraction engine이 같다면 프로그램 추출 결과를 무효화하지 않는다. 반대로 recipe가 바뀌면 해당 recipe를 사용한 결과만 다시 검증한다.

source SHA를 기준으로 table-first history와 canonical DB 적재를 멱등 처리하여 재실행 시 중복 study/evidence를 만들지 않는다.

## 11. 안전 장치

- 원본 SHA가 달라지면 기존 추출 결과를 그대로 재사용하지 않는다.
- 필수 앵커가 없거나 중복이면 레시피를 실행하지 않는다.
- 레시피 envelope 밖의 표 크기·수식·단위 변경은 자동 승인하지 않는다.
- 기하 fingerprint가 같아도 상대 metric header signature가 다르면 같은 recipe를 실행하지 않는다.
- evidence cell이 없는 값은 저장하지 않는다.
- 단위 변환은 레시피에 명시된 경우만 허용하고 원문 값과 원문 단위를 함께 보존한다.
- 수식 결과는 저장 값뿐 아니라 수식 패턴과 참조 셀을 검증한다.
- 예상하지 못한 정량 영역이 발견되면 누락 가능성으로 처리한다.
- 승인 전 patch는 production registry에 반영하지 않는다.
- 매칭이 불확실하면 처리량보다 정확성을 우선하여 fail closed한다.

## 12. 기존 코드에 필요한 변경

구현 시 모듈 경계는 다음처럼 둔다.

### 신규 모듈

- `inference_data_ai_structure_fingerprint.py`
  - fingerprint v2 생성과 직렬화
- `inference_data_ai_form_template.py`
  - template 계약, registry 조회, 버전 관리
- `inference_data_ai_extraction_recipe.py`
  - selector DSL 계약과 정적 검증
- `inference_data_ai_recipe_matcher.py`
  - hard gate, Top-K, 점수, AI 판정 adapter
- `inference_data_ai_recipe_executor.py`
  - capture에 레시피를 적용하고 evidence를 생성
- `inference_data_ai_recipe_validation.py`
  - 구조·값·수식·단위 검증
- `inference_data_ai_template_bootstrap.py`
  - 기존 989개에서 template/recipe를 생성하고 replay
- `inference_data_ai_table_recipe_proposal.py`
  - 신규 반복 block을 metric header signature까지 분리하고 redacted 대표 표로 구조당 1회 recipe를 결정·replay
- `inference_data_ai_structure_batch_control.py`
  - exact recipe registry, 구조당 1회·retry 0·run budget 제어, AI-free exact resolver

### 기존 모듈 변경

- `inference_data_ai_form_preflight.py`
  - classifier v2 적용
  - 현재 canonical 자료뿐 아니라 기존 table-first history/capture를 known catalog로 로드
- `inference_data_ai_form_registry.py`
  - 설명형 family contract 외에 실행 가능한 recipe와 승인 상태 저장
- `inference_data_ai_form_pipeline.py`
  - `_manifest_relative_paths`에서 경로만 반환하지 않고 전체 manifest item을 downstream에 전달
  - `formSignatureId`, `formFamilyId`, registry decision, recipe ref 보존
- `inference_data_ai_workflow.py`
  - generic AI workflow보다 recipe 경로를 먼저 실행
  - low-cost mode에서는 검증 실패 시 전체 DRAFT 자동 폴백 금지
- `inference_data_ai_corpus.py`
  - 파일별 구조 매칭/레시피 실행 결과와 비용 지표 집계
  - CLI의 reasoning/worker 설정이 실제 AI 실행까지 전달되도록 단일화
- table-first/canonical 적재 계층
  - `ExtractionResultV1`을 각각의 기존 schema로 바꾸는 adapter 추가

## 13. 단계별 구현·전환 계획

### 0단계: 현 상태 보존

- 기존 989개 table-first batch, history DB, capture, canonical DB를 읽기 전용 기준선으로 고정한다.
- 기존 16-worker 파일별 파이프라인은 사용자 지시에 따라 종료했고 checkpoint는 보존했다.

### 1단계: fingerprint v2와 registry 골격

- fingerprint v2 구현
- 기존 989개와 신규 파일에 지문만 생성
- 유사도 분포와 family 후보 수를 측정
- 이 단계에서는 AI 분석과 DB 적재를 하지 않는다.
- 완료: 과거 989건과 신규 361건 catalog/audit를 AI 0회로 생성했다.

### 2단계: 989개 template/recipe 부트스트랩

- 빈도가 높은 family부터 레시피 생성
- 기존 결과에 전량 replay
- 실패 예제를 기반으로 family 분리 또는 variant 생성

### 3단계: 대표 파일 파일럿

- exact match, 경미한 variant, 신규 family가 섞인 대표 신규 파일을 선택
- 자동 추출과 기존 방식 결과를 나란히 비교
- 오탐 방지를 우선하여 threshold 보정
- 완료: 첫 recipe structure는 AI 1회, retry 0으로 결정했고 6 tables에서 42 column facts와 144 exact cell facts를 프로그램으로 추출해 6/6 통과했다.

### 4단계: 신규 전체 증분 처리

- 승인 template은 AI 0회로 처리
- 모호한 건만 Top-K AI 매칭
- 신규 family/variant만 제한적으로 recipe 생성
- 성공한 patch를 registry에 승격한 뒤 같은 구조 파일을 일괄 replay

### 5단계: 기존 무거운 경로 축소

- 구조 재사용 경로의 정확도와 query 회귀가 통과한 뒤, 매칭된 파일의 staged full-DRAFT 경로를 비활성화한다.
- generic AI workflow는 `NEW_TEMPLATE_REQUIRED`의 조사 도구로만 남긴다.

## 14. 완료 기준

- 승인 recipe의 과거 replay에서 수치·근거 셀 일치 100%
- auto-reuse 구간의 구조 오탐 0건
- 매칭된 기존 구조에서 full-workbook DRAFT 호출 0회
- 모호한 신규 파일도 매칭 AI 호출 최대 1회
- 전체 신규 파일 기준 AI 호출 목표 평균 0.1회/파일 이하
- 동일 family 두 번째 파일부터 recipe 생성 AI 호출 0회
- table-first/canonical 적재 재실행 시 중복 0건
- 기존 golden query 10/10 유지
- 오류 파일은 잘못된 결과를 저장하지 않고 명확한 failure code로 격리

## 15. 구현 순서에 대한 최종 권고

첫 구현은 `fingerprint v2 -> 989개 구조 catalog -> 실행 가능한 recipe 1개 -> 과거 replay`의 얇은 수직 절단으로 진행한다. 처음부터 모든 family를 지원하지 않는다. 가장 많은 파일을 포괄하는 family 하나에서 다음을 입증한 뒤 확장한다.

1. 기존 결과로부터 recipe를 만들 수 있는가
2. 신규 파일과 template을 AI 없이 정확히 매칭할 수 있는가
3. 프로그램이 값과 evidence를 동일하게 재현하는가
4. 실패 시 잘못 저장하지 않고 격리하는가
5. 같은 구조의 후속 파일에서 AI 호출이 0회인가

이 검증이 통과하면 family 단위로 확장하면 되므로, AI 비용은 파일 수가 아니라 실제로 존재하는 구조와 변형 수에 비례하게 된다.

## 16. 신규 361건 실제 실행 결과

### 실행 경로

1. Capture v2와 canonical source bridge에 대상 361개 revision이 모두 존재함을 확인했다. 그중 356개는 table-first semantic request 생성에 성공해 1,981개 표를 프로그램으로 catalog화했고, 5개는 request 생성 단계에서 명시적으로 실패했다.
2. 반복되는 정량 후보를 metric header 의미 서명까지 분리해 90개 구조/307개 표의 recipe queue를 만들었다.
3. 과거 989건에서 일관된 DESCRIPTIVE contract를 찾은 구조는 `HISTORICAL_989_CONSENSUS`로 AI 없이 bootstrap했다.
4. 이미 replay를 통과한 동일 의미 서명은 `VERIFIED_SIGNATURE_PROPAGATION`으로 AI 없이 전파했다.
5. 나머지는 구조당 low-reasoning AI를 최대 1회만 사용하고 retry하지 않았다.
6. AI는 구조와 metric mapping만 정했다. 실제 숫자, 표시 형식, 통계, source range와 evidence cell은 프로그램이 Capture v2에서 추출했다.
7. preview가 잘린 긴 표는 read-only Capture v2 SQLite에서 해당 revision/sheet/range 전체 셀을 다시 읽어 replay했다.
8. title과 비교군 label은 recipe에 원문을 고정하지 않고 상대 text selector로 저장해 각 workbook의 실제 문구를 프로그램이 읽도록 했다.

### 완료 수치

| 항목 | 결과 |
|---|---:|
| 대상 workbook | 361 |
| Capture v2 / source bridge 존재 | 361 |
| table-first request 성공 | 356 |
| 명시적 table-request 실패 | 5 |
| cataloged table | 1,981 |
| 정량 후보 table | 1,328 |
| 반복 정량 recipe queue | 90 structures / 307 tables |
| replay 검증 후 등록 | 56 recipes / 193 tables |
| fail-closed 격리 | 34 structures / 114 tables |
| 미해결 queue | 0 |
| 큐 밖 table | 1,674 |
| AI 호출 | 62 |
| AI 성공/실패 | 56 / 6 |
| retry / 파일별 AI | 0 / 0 |
| prompt/output | 826,844 / 127,495 bytes |
| AI serial duration 합계 | 20.49분 |

307개 표마다 AI를 한 번 호출하는 방식과 비교하면 245회, 79.8%를 피했다. 90개 구조마다 한 번 호출하는 기준에서도 과거 consensus, 검증 recipe 전파, precheck로 28회, 31.1%를 피했다. 다만 최초 bootstrap batch의 평균은 361개 workbook당 0.172회로 목표 0.1회 이하를 아직 충족하지 못했다. 이번에 등록된 56개 recipe와 정확히 일치하는 후속 파일은 AI 0회로 처리되므로 이후 증분 batch에서 상각된다.

### coverage 해석

`final-coverage-report.json`은 다음 합계를 강제 검증한다.

- workbook: `356 table-request success + 5 table-request failure = 361`; Capture/source revision은 `361/361`
- table: `307 repeated queue + 1,674 outside queue = 1,981`
- queue: `193 registered + 114 quarantined = 307`
- structure: `56 registered + 34 quarantined = 90`

큐 밖 1,674개는 누락이 아니라 이번 반복 정량 recipe pass의 명시적 비대상이다. 그 안에는 정량 table 1,021개와 비정량 table 653개가 있다. 이들을 프로그램 추출 완료로 표시하지 않으며, 정량 1,021개 중 여러 workbook에 재사용 가능한 355개는 다음 확대 대상이고 나머지 666개는 단일/비재사용 정량 구조로 별도 정책이 필요하다.

5개 실패는 Capture v2 revision 자체의 누락이나 capture 실패가 아니다. 361건 모두 Capture/source DB에 존재한다. 이 5건은 Capture v2가 선언한 non-empty cell 수와 semantic packet 생성기에 노출된 cell 수가 달라 table-first request를 만들지 못한 경우다. 잘못된 표를 만들어 계속하지 않고 파일명, revision과 오차를 최종 report에 보존했으며 canonical DB에는 `FAILED_TABLE_REQUEST / FAILED` terminal 상태로 계상했다.

## 17. canonical 운영 DB 반영 결과

구조 recipe와 terminal coverage를 실제 AI 질문용 운영 DB `InputDataFinish.sqlite`의 `workbook_analyses → knowledge_studies → knowledge_outcomes → knowledge_observations → evidence_items` 계약으로 적재했다.

| 항목 | 결과 |
|---|---:|
| 대상 source/capture revision | 361 |
| 기존 canonical 분석 보존 | 115 |
| 신규 canonical 분석 적재 | 246 |
| 최종 active canonical coverage | 361 / 361 |
| 누락 / active 중복 | 0 / 0 |
| 신규 recipe replay workbook | 54 |
| 신규 recipe studies / observations | 120 / 595 |
| 신규 distinct evidence | 1,430 |
| recipe 미확보 terminal NEEDS_REVIEW | 186 |
| table-request FAILED terminal | 5 |
| EXCLUDED terminal | 1 |
| canonical import AI 호출 / retry | 0 / 0 |

recipe replay는 program-owned min/max/average/sample size와 정확한 sheet/range evidence만 적재한다. 의미 검토 전에는 모든 신규 recipe study가 `NEEDS_REVIEW`이며 comparison/effect를 생성하지 않는다. recipe가 없는 186건도 운영 DB에서 누락시키지 않고 `NEEDS_REVIEW_NO_VERIFIED_RECIPE` terminal analysis로 남기되, 근거 없는 숫자 claim은 만들지 않는다.

기존 WPF/CLI 질문 경로는 같은 canonical 운영 DB를 조회하므로 별도 side DB 연결 없이 신규 레코드를 검색한다. smoke query `Tape separation`은 신규 `DATA-A6DB8C587B0A`, `Tape separation NG count/rate`, 6개 evidence를 반환했다. 검토 전 데이터이므로 answer-eligible effect 0으로 fail-closed 처리되는 것도 확인했다.

최종 증빙:

- `canonical-import-result.json`: 단일 transaction 실제 적용 결과와 115→361 coverage
- `canonical-import-v1/canonical-import-final-audit.json`: 현재 361/361 coverage, active 유일성, importer entity/evidence 수, AI 0회
- `canonical-import-v1/query-smoke-new-exact.json`: 기존 canonical query 경로의 신규 DB 레코드 검색 결과
