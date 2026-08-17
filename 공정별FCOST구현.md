# 공정별 F-COST 구현안

## 1. 목표

기존 `Test 3`, `Test 4`, `BOM & Drawing`, `F-COST` 기능을 하나의 흐름으로 연결하여 다음 기능을 구현한다.

1. 모델별 공정 순서를 관리한다.
2. 각 공정에서 새로 투입되는 자재와 사용량을 설정한다.
3. 공정별 Input·NG·NG률을 조회한다.
4. 불량 발생 시 폐기되는 누적 자재수량을 계산한다.
5. 기간별 자재 단가를 적용해 공정별 예상 F-COST를 산정한다.
6. 예상 F-COST와 BMES E008의 실제 초과투입 F-COST를 비교한다.

## 2. 현재 기능과 재사용 범위

| 기존 기능 | 재사용할 내용 | 현재 부족한 부분 |
|---|---|---|
| Test 3 | 모델별 공정 순서, 공정번호, 참조공정, 공정→자재 매핑 | JSON 저장이며 버전·유효기간·승인 상태가 없음 |
| Test 4 | 공정별 Input·NG 주차 집계, 실제 F-COST 환산수량 비교 | 자재 사용량과 단가를 적용한 예상금액 계산이 없음 |
| BOM & Drawing | BOM 트리, 자재 사용량, PDM 도면번호와 파일 | BOM 기준수량과 다단계 누적 사용량 확정이 필요함 |
| F-COST | 기간별 실제 F-COST, 원자재 단가, 통화와 환율, 환산수량 | 공정 NG와 연결된 이론 산정 계층이 없음 |

관련 코드:

- `JinoSupporter.Web/Components/Pages/BmesTest3Page.razor`
- `JinoSupporter.Web/Components/Pages/BmesTest4Page.razor`
- `JinoSupporter.Web/Components/Pages/BmesTest5Page.razor`
- `JinoSupporter.Web/Components/Pages/BmesFCostPage.razor`
- `JinoSupporter.Web/Services/ProcessMaterialMappingService.cs`
- `JinoSupporter.Web/Services/ProcessMaterialNgService.cs`
- `JinoSupporter.Web/Services/BmesFcostActualService.cs`
- `JinoSupporter.Web/Services/BmesBomCacheService.cs`

## 3. 핵심 계산식

공정별 실제 NG 수량이 있으면 그 값을 우선 사용한다.

```text
공정 NG률 = 공정 NG 수량 / 공정 Input 수량

공정 NG 수량 = 공정 Input 수량 × 공정 NG률

자재 폐기수량
= 공정 NG 수량
× 완제품 1대당 누적 자재 사용량
× 폐기계수

공정 예상 F-COST
= Σ(자재 폐기수량 × 해당 기간 자재 단가)
```

예시:

```text
Input             = 100,000개
NG률              = 2%
NG 수량           = 2,000개
자재 사용량       = 2 EA/제품
자재 단가         = 150 VND/EA
폐기계수          = 1.0

자재 폐기수량     = 2,000 × 2 × 1.0 = 4,000 EA
예상 F-COST       = 4,000 × 150 = 600,000 VND
```

재작업이나 자재 회수로 30%를 살릴 수 있으면 폐기계수는 `0.7`로 적용하고 예상 F-COST는 `420,000 VND`가 된다.

## 4. 공정별 원가 범위

기본 계산은 `누적 WIP` 방식으로 한다.

예를 들어 공정 순서가 다음과 같다고 가정한다.

```text
S1: FRAME 투입
S2: YOKE, MAGNET 투입
S3: COIL, DOME 투입
```

- S1 불량: FRAME 원가만 포함
- S2 불량: FRAME + YOKE + MAGNET 원가 포함
- S3 불량: FRAME + YOKE + MAGNET + COIL + DOME 원가 포함

공정별 예외 처리를 위해 다음 계산범위를 지원한다.

- `Cumulative`: 해당 공정까지 누적된 모든 자재
- `DirectOnly`: 해당 공정에서 새로 투입된 자재만
- `Rework`: 폐기하지 않고 재작업 비용 또는 지정 폐기계수만 적용

같은 불량 한 건을 여러 공정에 중복 귀속하지 않도록 NG가 실제 발생한 공정에서 한 번만 계산한다.

## 5. 현재 계산 구조의 보완점

현재 `ProcessMaterialNgService`는 한 공정에 자재가 여러 개 매핑되면 동일한 공정 NG/Input을 각 자재 행에 복사한다. `UsageQty`는 결과에 포함되지만 자재수량 계산에 곱하지 않는다.

새 계산에서는 다음 값들을 명확히 분리해야 한다.

- 제품 불량수량: 공정에서 발생한 NG 제품 수량
- 자재 폐기수량: `제품 불량수량 × 자재 사용량 × 폐기계수`
- 예상 F-COST: `자재 폐기수량 × 단가`
- 실제 F-COST: BMES E008의 실제투입액과 표준투입액 차이

`Material NG`라는 하나의 이름으로 제품 불량수량과 자재 폐기수량을 함께 표현하지 않는다.

## 6. 예상 F-COST와 실제 F-COST의 관계

예상 F-COST와 기존 F-COST는 같은 값이 아니다.

```text
예상 F-COST
= 공정 NG × 자재 사용량 × 자재 단가 × 폐기계수

실제 F-COST
= BMES E008의 양수 초과투입액
= MAX(실제 투입액 - 표준 투입액, 0)
```

실제 F-COST에는 다음 항목이 함께 들어갈 수 있다.

- 공정 불량으로 인한 추가 투입
- 설비 셋업과 초도 손실
- 재작업 투입
- 계수와 재고 차이
- 불량시스템에 기록되지 않은 폐기
- BOM 또는 표준소요량 차이

따라서 화면에서는 두 금액을 합치지 않고 다음과 같이 비교한다.

| 항목 | 내용 |
|---|---|
| 예상 F-COST | 공정 NG 데이터로 계산한 이론 손실액 |
| 실제 F-COST | E008에서 확인된 실제 초과투입 금액 |
| 차이 | 실제 F-COST - 예상 F-COST |
| 설명 필요 | 단가·BOM·NG·재작업·재고 차이 점검 |

## 7. 통합 화면 구성

기존 테스트 메뉴를 최종적으로 하나의 `공정 F-COST` 메뉴로 통합한다. 기존 URL과 권한키는 호환을 위해 유지하거나 새 화면으로 연결한다.

### 7.1 BOM·도면 탭

- 모델 검색과 BOM 캐시/서버 조회
- 제품 L/R 및 공통·편측 자재 구분
- BOM 경로와 원본 사용량 표시
- PDMNO, 도면 버전, 도면 다운로드
- BOM 기준일과 캐시 기준일 표시

### 7.2 공정 자재 설정 탭

- Routing 기준 공정 목록 표시
- 드래그로 공정 순서 설정
- BOM 자재를 공정으로 선택·추가
- 자재 사용량과 단위 편집
- 참조공정과 L/R 공정 연계
- 계산범위, 폐기계수, 재작업 여부 설정
- 미배치 BOM 자재와 중복 배치 경고

### 7.3 F-COST 산정 탭

- 기간, 모델, 주차/월 선택
- 공정별 Input, NG, NG률 표시
- 공정별 직접 자재와 누적 자재 표시
- 자재별 폐기수량, 단가, 예상금액 표시
- 공정 소계, 모델 소계, 전체 합계 표시
- 단가 또는 단위가 없는 행은 금액을 0으로 만들지 않고 오류/경고 처리

### 7.4 실제 비교 탭

- 예상 F-COST
- 실제 F-COST
- 차이 금액과 차이율
- 실제 환산수량과 예상 폐기수량 비교
- 자재·공정·모델 단위 Drill-down
- Excel/HTML 내보내기

## 8. 권장 저장 구조

수동 마스터와 계산 결과는 재수집 가능한 `fcost_raw.db`와 분리한다. 권장 파일은 `02. FCOST/process_fcost.db`다.

### ProcessFcostProfiles

- ProfileId
- Plant
- ModelCode
- ModelName
- BomWorkDate
- EffectiveFrom / EffectiveTo
- Version
- Status: Draft / Approved / Retired
- CreatedAt / UpdatedAt

### ProcessFcostProcesses

- ProcessId
- ProfileId
- ProcessCode
- ProcessName
- ProcessNo
- Sequence
- ReferenceProcessNo
- CostScope: Cumulative / DirectOnly / Rework
- DefaultLossFactor

### ProcessFcostMaterials

- ProcessMaterialId
- ProcessId
- MaterialCode
- MaterialName
- UsageQty
- UsageUnit
- UnitConversionFactor
- LossFactor
- IncludeInCost
- BomPath
- PdmNo / PdmVersion
- Source: BOM / FCOST / Manual
- EffectiveFrom / EffectiveTo

### ProcessFcostRuns

- RunId
- StartDate / EndDate
- Plant
- ModelFilter
- ProfileVersion
- GeneratedAt
- CalculationVersion
- TotalEstimatedFcostVnd
- TotalActualFcostVnd

### ProcessFcostRunDetails

- RunId
- PeriodKey
- ModelCode / ModelName
- ProcessCode / ProcessName / ProcessNo
- MaterialCode / MaterialName
- InputQty
- NgQty
- NgRate
- DirectUsageQty
- CumulativeUsageQty
- LossFactor
- CalculatedScrapQty
- UnitPrice
- Currency / PriceUnit
- UnitPriceVnd
- EstimatedFcostVnd
- ActualFcostVnd
- DifferenceVnd
- BomSource / PriceSource / NgSource
- WarningCode / WarningMessage

계산 상세에 당시의 BOM, 단가, 환율, NG 원천을 보존해야 나중에 같은 보고서를 재현할 수 있다.

## 9. BOM 수량과 단위 처리

현재 BOM 트리의 `UsageQty`는 각 `BOMC.MENGE` 행의 원본 수량이다. 다단계 BOM의 완제품 기준 누적수량은 보장되지 않는다.

정확한 계산 전에 다음을 확정해야 한다.

1. BOM 헤더 또는 대체안의 기준수량
2. 하위 BOM 사용량의 누적 계산 방법
3. EA/PC/KG/G/M 등의 단위 환산
4. 가격의 `PriceUnit`이 1EA, 1KG, 100EA 중 무엇인지
5. BOM 대체안과 유효기간 선택 규칙

단위가 일치하지 않거나 기준수량을 확인하지 못한 자재는 자동 계산에서 제외하고 경고로 표시한다. 임의로 1 또는 0을 적용하지 않는다.

## 10. 구현 서비스 구성

### ProcessFcostRepository

- SQLite 스키마 생성과 마이그레이션
- 공정·자재 마스터 버전 관리
- 계산 Run/Header/Detail 저장과 조회

### ProcessMaterialEffectiveMappingService

- Test 3과 Test 4에 흩어진 직접 매핑·참조공정·L/R 연계 로직 통합
- 공정 순서별 직접 자재와 누적 WIP 자재 계산

### ProcessMaterialFCostCalculationService

입력:

- 기간과 모델
- 공정 Input/NG
- 승인된 공정·자재 마스터
- 기간별 자재 단가와 환율
- 공정별 계산범위와 폐기계수

출력:

- 공정·자재·기간별 예상 F-COST 상세
- 공정/모델/전체 합계
- 실제 F-COST 비교
- 누락·단위·단가·BOM 버전 경고

## 11. 구현 순서

### 1단계: 마스터 저장 기반

- `process_fcost.db` 생성
- 기존 JSON 매핑을 SQLite로 이관하는 마이그레이션 추가
- 공정·자재·유효기간·버전·폐기계수 저장
- Test 3의 읽기/쓰기를 Repository로 전환

### 2단계: BOM·도면 통합

- BOM & Drawing의 캐시와 PDM 정보를 설정 화면에서 재사용
- BOM 기준일을 캐시키와 마스터 출처에 반영
- 사용량, 단위, BOM 경로, 도면번호를 공정 자재 설정에 연결
- 미배치/중복/사용량 누락 검사

### 3단계: 이론 F-COST 계산

- 공정 Input/NG/NG률 로딩
- 공정순서에 따른 누적 WIP 자재 계산
- `NG × 사용량 × 폐기계수 × 단가` 계산
- 단위 환산과 누락 경고
- 계산 스냅샷 저장

### 4단계: 실제 F-COST 비교

- 기존 raw breakdown의 단가와 실제 F-COST 재사용
- 예상·실제·차이·차이율 표시
- 자재/공정/모델별 상세 추적
- Excel 및 HTML 내보내기

### 5단계: 검증

- 수작업 계산 가능한 모델 1개로 행별 금액 대조
- 공정 NG 0건, Input 0건, 단가 누락, 단위 불일치 검사
- 참조공정 및 L/R 공통·전용 자재 검사
- 같은 조건 재실행 시 계산 스냅샷 재현성 검사
- 예상 F-COST와 실제 F-COST 차이에 대한 원인 분류 검사

## 12. 권장 기본 정책

- 계산범위: `Cumulative`
- 완전 폐기 공정의 폐기계수: `1.0`
- 재작업 공정: 실제 폐기 비율을 별도 설정
- 단가 기준: 조회 기간의 실제 자재단가를 우선 사용
- 단가 누락: 0원 처리 금지, 경고와 합계 제외
- 단위 불일치: 자동 환산 규칙이 있을 때만 계산
- BOM 기준일: 보고서 기간 또는 승인된 마스터의 유효일 사용
- 실제 F-COST: 예상 F-COST에 더하지 않고 별도 비교
- 계산 결과: 사용한 마스터 버전과 원천 데이터를 스냅샷으로 저장

## 13. 현재 데이터 상태

2026-08-15 읽기 전용 확인 기준:

- 공정·자재 매핑 JSON: 0건
- RoutingTable: 11,580행
- BOM 캐시: 7개 헤더, 301행
- F-COST raw: 177,009행
- F-COST raw breakdown cache: 31건

데이터 원천은 이미 준비되어 있으나 공정별 자재 마스터가 비어 있으므로, 실제 구현은 계산 화면보다 공정·자재 설정과 검증 기능부터 시작한다.

## 14. 구현 전 확정할 업무 규칙

1. 불량 공정에서 해당 공정 신규 자재만 잃는지, 이전 공정 누적 자재까지 잃는지
2. 공정별 완전 폐기·부분 회수·재작업 비율
3. BOM 기준수량과 다단계 누적 사용량 계산 방법
4. 자재 사용단위와 가격단위 환산표
5. 실제 F-COST와 예상 F-COST 차이의 허용범위
6. 설정 승인권자와 마스터 변경 이력 관리 방식

권장 시작값은 `누적 WIP`, `완전 폐기 1.0`, `단가 누락 시 계산 제외`, `실제 F-COST와 별도 비교`다.

## 15. 2026-08-15 Test 3 React 편집기 구현 상태

공정·자재 마스터를 먼저 쉽게 입력할 수 있도록 Test 3의 편집 영역만 React island로 전환했다. 기존 Blazor 메뉴·인증·서비스는 유지하고 React 19와 dnd-kit이 3열 드래그 UI를 담당한다.

- 1열: L/R별 저장 공정 순서와 `MAIN`, `SUB1`~`SUB99` 라인
- 2열: 검색 모델의 미배치 Routing 공정
- 3열: BOM 자재와 최초 투입 공정·사용량·단위
- 공정 카드 드래그: MAIN/SUB 배치, 라인 간 이동, 상하 순서 변경, 미배치 목록으로 회수
- 자재 카드 드래그: 공정 카드에 놓아 최초 투입 공정 저장
- SUB 출력 드래그: 같은 L/R의 MAIN 공정에 놓아 합류 지점 저장
- 공정 카드 자재 표시: 최초 투입 자재는 `+`, 이전 공정부터 이어지는 자재는 `↳`로 표시
- 기준 모델 검색: 직접 연결된 L/R Routing 모델을 한 화면에 동시에 표시
- 자재 범위: `M-P-`는 제외하고 COIL 같은 C-S/R-S 중간 BOM 자재는 포함

현재 JSON 공정 설정에는 기존 필드와 함께 `LaneCode`, `MergeProcessNo`를 저장한다. Test 4 유효 자재 상속은 이전 공정뿐 아니라 해당 MAIN 공정으로 합류하는 SUB 마지막 공정도 함께 따라가도록 확장했다.

React 소스는 `JinoSupporter.Web/ClientApp/Test3Editor`에 있고 빌드 결과는 `JinoSupporter.Web/wwwroot/test3-editor`에 생성한다.

```powershell
Set-Location JinoSupporter.Web/ClientApp/Test3Editor
npm install
npm run build
```

이 단계는 공정·자재 입력 UI와 누적 자재 흐름 기반만 구현한 것이다. 금액 산정을 완료하려면 다음 단계에서 `NG 수량 × BOM 누적 사용량 × 기간 단가 × 폐기계수` 계산 서비스와 결과 스냅샷 저장을 추가해야 한다.
