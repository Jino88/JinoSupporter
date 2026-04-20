# JinoSupporter — 검토 데이터 AI 분석 시스템 설계

---

## 최종 목적

이제까지 했던 모든 검토 데이터를 데이터화해서 AI에 데이터를 준 뒤에,
불량 분석, 산포 분석, "이 불량에는 이 공정을 확인하면 개선 효과가 있다" 등을 알고 싶다.

즉, 이 데이터를 바탕으로, **추후에 어떤 문제가 생겼을 때**,
이전 검토 데이터들을 쉽게 취합하여 무엇이 문제인지, 무엇을 검토해야 하는지 빠르게 파악하기 위함.

---

## 현재 문제

1. 모든 검토 데이터가 Excel로 저장되어 있지만 **DRM 암호화** 되어 프로그램에 넣을 수 없음
   - 지정된 프로그램에서만 열림, 다른 프로그램에서는 열리지 않음
2. 검토 DATA마다 **테이블 양식이 제각각** (불량률 집계 항목도 다름)
3. **제품군마다** 불량 유형, 측정 항목이 다름
4. 검토 조건 변경이 다수 있음 (old lot vs new lot, mold 번호별 비교 등)
5. 리포트에 **이미지(사진 증거)** 가 포함된 경우도 있음
6. **1,000개 이상**의 과거 데이터 존재

---

## DRM 문제 및 데이터 입력 전략

### DRM 환경에서 가능한 방법

| 방법 | 가능 여부 | 비고 |
|---|---|---|
| 스크린샷 → 에디터 붙여넣기 | ✅ 가능 | 현재 방식, Claude가 이미지에서 테이블 추출 |
| IT/DRM 관리자 일괄 해제 요청 | 협의 필요 | 마스터 키로 일괄 복호화 가능할 수도 있음 |
| DRM 앱 내에서 PDF 인쇄 | 정책에 따라 다름 | PDF → 텍스트 추출 가능 |
| 텍스트 복사 (Ctrl+C) | 정책에 따라 다름 | 시트 보호만인 경우 가능 |

### 비용 문제 — 1,000개를 AI로 처리하면?

```
계산:
  이미지 1장 ≈ 1,600 토큰
  리포트 1개 (이미지 2장 + 프롬프트 + 출력) ≈ $0.02~0.03

  1,000개 × $0.03 = 약 $30 (일회성)
  이후 신규 추가 (주 5~10개) ≈ $10/년
```

**비용보다 노동력(일일이 캡처하는 시간)이 진짜 문제.**

### 현실적 입력 전략 — "필요할 때 넣는 방식"

```
[Phase 1] 앞으로 새로 만드는 리포트
  → Input Data 화면에서 작성 시 바로 저장
  → 추가 비용/노력 없음

[Phase 2] 최근 6개월~1년치 (30~50개)
  → 스크린샷으로 수동 입력
  → 현재와 가장 관련 높은 데이터 우선

[Phase 3] 필요 시 소환 방식
  → 특정 문제 발생 시, 그 문제와 관련된 과거 리포트만 그때그때 입력

[Phase 4] IT 협의 (선택)
  → DRM 일괄 해제 가능하면 나머지 처리
```

**→ 1,000개 전부 일괄처리는 불필요. 핵심 200~300개면 충분.**

---

## 실제 리포트 분석 사례 (4개)

### 사례 1 — TIU-C11-20 (UV Gauss 건조 시간 검토)

- **목적**: SPL Low Gauss NG 원인 규명
- **변수 유형**: 공정 파라미터 (건조 시간 1/2/3/5/10 min)
- **테이블**: 1개
- **이미지**: 없음
- **핵심 발견**:
  - 건조 시간 ↑ → Gauss 값 ↓ → SPL NG 발생
  - 1분: 477G (PASS), 2분~: 367G 이하 (FAIL)
  - Safe zone = 1분 이하
  - Before Dry 전부 PASS → 문제는 공정이 원인, 소재 아님

### 사례 2 — BRS-161016 ALL MOLD 비교

- **목적**: VP 전체 Mold NG율 비교
- **변수 유형**: 금형 번호 (Mold #6, #7, #10, #11, #12)
- **테이블**: 2개 (SUB1 공정검사 + FUNCTION 기능검사)
- **이미지**: 없음
- **핵심 발견**:
  - Mold 번호 높을수록 NG율 감소 (#6: 18.9% → #12: 1.0%)
  - 주요 불량: VP+CD 분리

### 사례 3 — BRS-161016 NEW LOT 비교 (old vs new)

- **목적**: 신규 lot (27.02.2026) 사용 가능 여부 판단
- **변수 유형**: Mold 번호 + Lot 날짜 조합
- **테이블**: 2개
- **이미지**: 있음 (Mold별 VP+CD 분리 불량 패턴 사진)
- **핵심 발견**:
  - VP mold #6 new lot → NG 10% (공정), 32.1% (기능) → **사용 불가**
  - VP mold #3 new lot → 0.5%로 오히려 개선됨
  - 이미지가 특정 Mold의 불량 패턴 증거로 포함

### 사례 4 — TIA-338-L/R FPCB ASS'Y 공정 추적

- **목적**: RA1 공정 Ass'y FPCB Damage 원인 규명
- **변수 유형**: Worker (사람, Worker 1~7)
- **테이블**: 1개 (Before/After 쌍 구조 + 위치별 6개 컬럼)
- **이미지**: 있음 (특정 Position + Before/After + Worker 연결)
- **특이사항**:
  - Before / After Ass'y FPCB 가 같은 테이블에 쌍으로 존재
  - 불량 위치 6개가 별도 컬럼으로 분산 (Wide Format → Long 변환 필요)
  - "Worker 1 (After retrained)" 행 존재 → 재교육이라는 개입(Intervention) 기록
  - 한 리포트에 TIA-338L, TIA-338R 두 제품 동시 포함
  - 이미지가 Position + Before/After + Worker 단위로 연결됨
- **핵심 발견**:
  - Before Ass'y 불량 0.28% (L) / 0.89% (R) → 공정 전 일부 불량 존재
  - After Ass'y에서 불량 감소 → 공정 자체가 문제는 아님
  - Worker 1 재교육 후 0.00% → 교육 효과 확인

---

## 핵심 전제: "Normal" = 항상 기준값

**사용자 규칙:**
> 변경이 없는 기존 공정을 항상 **"Normal"** 이라 부른다.  
> 모든 분석의 핵심은 **"Normal 대비 얼마나 개선/악화되었는가?"** 이다.

```
VariableGroup = "normal"  → 기준값 (변경 없음, 기존 공정)
VariableGroup = "test"    → 개선 조건 (변경 적용)
VariableGroup = "new_lot" → 신규 lot (기준: old_lot)
VariableGroup = "before"  → 개입 전 (기준: 사람 편차 비교용)
VariableGroup = "after"   → 개입 후 (재교육 등)
```

**개선 폭 계산 공식:**
```
개선율 = (normal_ng_rate - test_ng_rate) / normal_ng_rate × 100%

예시:
  AI COIL: Normal 10.0% → Test 33.3% = 악화 (-233%)
  SPOT WELDING: Normal 4.4% → Test 2.9% = 개선 +34%
```

**분석 쿼리 패턴 (Normal이 있을 때):**
```sql
-- 동일 dataset + 동일 variableDetail(테이블명) 내에서 normal vs test 비교
SELECT
    variableDetail,
    defectCategory,
    MAX(CASE WHEN variableGroup='normal' THEN ngRate END) AS normal_ng_rate,
    MAX(CASE WHEN variableGroup='test'   THEN ngRate END) AS test_ng_rate,
    ROUND(
      (MAX(CASE WHEN variableGroup='normal' THEN ngRate END)
       - MAX(CASE WHEN variableGroup='test' THEN ngRate END))
      / MAX(CASE WHEN variableGroup='normal' THEN ngRate END) * 100, 1
    ) AS improvement_pct
FROM NormalizedMeasurements
WHERE DatasetName = ?
  AND variableGroup IN ('normal', 'test')
GROUP BY variableDetail, defectCategory;
```

**Normal이 없는 경우 (Mold 비교, Worker 비교 등):**
- 최솟값/최댓값 기준으로 상대 비교
- 또는 AI가 Summary에서 "Worker 1이 가장 낮음" 형태로 정성 기술

---

## 리포트 포맷 다양성 문제

각 리포트마다 테이블 구조가 달라 자동 추출이 어렵다.
현실적으로 100% 정확한 추출은 불가능하며, **검토 후 수정**하는 흐름이 필요.

### 포맷 유형 분류

| 유형 | 예시 | 비교 축 | variableGroup |
|---|---|---|---|
| Test vs Normal (단순) | C11-20 AWF, SPL Gauss | 조건명 | test / normal |
| Mold 번호 비교 | BRS ALL MOLD | Mold # | (빈값, variable에 Mold#) |
| Lot 비교 | BRS NEW LOT | Lot 날짜 | new_lot / old_lot |
| Worker × Before/After | TIA-338 | Worker + 공정 단계 | before / after |
| 복수 테이블 같은 리포트 | C11-20 AWF (AI COIL + SPOT WELDING) | 테이블명 → variableDetail | test / normal |

### Claude 추출 신뢰도 현실적 평가

```
✅ 높음: 숫자 값 (Input, NG 수, NG%)
△ 중간: variableGroup 레이블 (Normal/Test 판단)
△ 중간: 복수 테이블에서 테이블명 분리
⚠ 낮음: 병합 셀이 많은 복잡한 구조
⚠ 낮음: 리포트마다 컬럼명이 다른 불량 유형
```

**운영 방침: 추출 후 검토 필수**
- Analyze 버튼 → 결과 미리보기 → 이상 있으면 수정 후 Save
- 완전 자동화보다 "80% 자동 + 20% 수동 검토"가 현실적

---

## 리포트 간 구조 비교

| 항목 | TIU-C11-20 | BRS ALL MOLD | BRS NEW LOT | TIA-338 | C11-20 AWF |
|---|---|---|---|---|---|
| 제품군 | TIU-C11-20 | BRS-161016 | BRS-161016 | TIA-338-L/R | C11-20 |
| 변수 유형 | 공정 파라미터 | 금형 번호 | 금형+Lot | 사람(Worker) | **조건(Test/Normal)** |
| 비교 구조 | 단방향 | 단방향 | old vs new | Before/After 쌍 | **test vs normal** |
| Normal 기준값 | ✅ (1분 = safe) | ✅ (Mold #12) | ✅ (old lot) | - | **✅ "Normal (AWF#1)"** |
| 복수 테이블 | X | X | X | X | **✅ (AI COIL + SPOT WELDING)** |
| variableDetail 필요 | X | X | X | X | **✅ (테이블명으로 구분)** |
| Wide→Long | X | X | X | ✅ position | ✅ 불량 유형 컬럼 |
| 개입 기록 | X | X | X | ✅ retrained | X |

**→ 공통 구조**: `제품 + 날짜 + 변수 + 투입/불량수/불량률 + 불량유형`

---

## 핵심 문제: 컬럼명 불일치

같은 불량인데 리포트마다 이름이 다름:

| 리포트 A | 리포트 B | 실제 의미 |
|---|---|---|
| Not enough glue | Glue not enough | 접착 불량 |
| Process AI coil NG wire offset | NG wire offset | wire offset |
| NG AUDIOBUS → Hearing | NG Hearing | 청감 불량 |

→ **Claude 정규화가 저장 시 canonical name으로 통일**

---

## 설계: 3계층 저장 구조

```
┌──────────────────────────────────────────┐
│  Layer 1: 원본 그대로 (현재 구현됨)        │
│  Editor HTML + Tables + Images           │
│  → 절대 손실 없음, 사람이 직접 열람 가능   │
└──────────────────────────────────────────┘
                   ↓ Claude 정규화 (저장 시 자동)
┌──────────────────────────────────────────┐
│  Layer 2: 정규화된 측정값 (신규 구현 필요) │
│  NormalizedMeasurements 테이블           │
│  → 제품/리포트 형식 무관하게 숫자 비교 가능│
└──────────────────────────────────────────┘
                   ↓ Claude 요약 (저장 시 자동)
┌──────────────────────────────────────────┐
│  Layer 3: AI 요약 텍스트 (신규 구현 필요)  │
│  DatasetSummary 테이블                   │
│  → 검색/분석 시 컨텍스트로 활용            │
└──────────────────────────────────────────┘
```

---

## Layer 2 — NormalizedMeasurements 스키마 (최종)

```sql
CREATE TABLE NormalizedMeasurements (
    Id              INTEGER PRIMARY KEY AUTOINCREMENT,
    DatasetName     TEXT NOT NULL,       -- 원본 dataset 참조
    ProductType     TEXT NOT NULL,       -- BRS-161016, TIU-C11-20 등
    TestDate        TEXT NOT NULL,       -- 검사 날짜
    Line            TEXT DEFAULT '',     -- C2-2B, C2-1B 등 라인 정보
    CheckType       TEXT NOT NULL,       -- process / function / visual_inspection
    Variable        TEXT NOT NULL,       -- 비교 대상 (Worker 1, Mold #6, 건조 1min 등)
    VariableDetail  TEXT DEFAULT '',     -- 세부 조건 (lot 날짜, Before/After 등)
    VariableGroup   TEXT DEFAULT '',     -- 비교 그룹 (new_lot/old_lot, before/after 등)
    Intervention    TEXT DEFAULT '',     -- 개입 기록 (retrained, condition_changed 등)
    InputQty        INTEGER DEFAULT 0,
    OkQty           INTEGER DEFAULT 0,
    NgTotal         INTEGER DEFAULT 0,
    NgRate          REAL DEFAULT 0,
    DefectCategory  TEXT DEFAULT '',     -- Claude가 매핑한 canonical 카테고리
    DefectType      TEXT DEFAULT '',     -- 원본 컬럼명 그대로
    DefectCount     INTEGER DEFAULT 0,
    CreatedAt       TEXT NOT NULL
);
```

### Wide → Long 변환 (TIA-338 같은 경우)

테이블에 불량 위치가 컬럼으로 분산된 경우, Claude가 저장 시 자동 변환:

```
원본 (Wide):
Worker 5 | Before | pos1=0 | pos2=0 | pos3=0 | pos4=0 | pos5=1 | pos6=0

저장 (Long) — 6개 행으로:
Worker 5 | Before | NG Rear damage position 1 | 0
Worker 5 | Before | NG Rear damage position 2 | 0
Worker 5 | Before | NG Rear damage position 3 | 0
Worker 5 | Before | NG Rear damage position 4 | 0
Worker 5 | Before | NG Rear damage position 5 | 1  ← NG
Worker 5 | Before | NG Rear damage position 6 | 0
```

### DefectCategory 예시 (canonical 카테고리)

| 원본 컬럼명 | DefectCategory |
|---|---|
| VP+CD separate / Not enough glue / Glue not enough | assembly_defect |
| Dome damage / Particle | cosmetic_defect |
| SPL / NG AUDIOBUS | function_spl |
| THD | function_thd |
| Noise / Touch / Hearing | function_hearing |
| wire offset / wire forming | wire_defect |
| Gauss low | magnetic_defect |
| NG Rear damage position 1~6 | rear_visual_damage |

---

## Layer 3 — DatasetSummary 스키마

```sql
CREATE TABLE DatasetSummary (
    Id          INTEGER PRIMARY KEY AUTOINCREMENT,
    DatasetName TEXT NOT NULL UNIQUE,
    ProductType TEXT NOT NULL,
    Summary     TEXT NOT NULL,    -- Claude 자동 생성 요약
    KeyFindings TEXT NOT NULL,    -- 핵심 발견 사항
    CreatedAt   TEXT NOT NULL
);
```

**자동 생성 요약 예시:**
```
"TIA-338 FPCB ASS'Y 공정 추적 (2026-04-16, Line C2-2B/C2-1B)
 Worker 1~7 Before/After 비교. Before 불량 0.89% (R라인).
 Worker 1 재교육 후 NG율 0.00%로 개선 → 교육 효과 명확.
 After Ass'y에서 전반적 불량 감소 → 공정 자체 문제 없음.
 핵심: 작업자 편차가 주 원인."
```

---

## 이미지 연결 구조 (확장)

```sql
-- DatasetImages에 컬럼 추가 (migration)
ALTER TABLE DatasetImages ADD COLUMN LinkedVariable TEXT DEFAULT '';
-- 예: "Worker 5"

ALTER TABLE DatasetImages ADD COLUMN LinkedPosition TEXT DEFAULT '';
-- 예: "Position 5"

ALTER TABLE DatasetImages ADD COLUMN LinkedPhase TEXT DEFAULT '';
-- 예: "before_assembly" / "after_assembly"
```

에디터에서 이미지 캡션에 Variable / Position / Phase를 입력하면  
Claude가 저장 시 자동으로 태깅.

---

## 나중에 가능해지는 질문들

### 핵심 질문: Normal 대비 개선 확인

```
"이 공정 변경이 실제로 효과 있었나?"
→ VariableGroup IN ('normal','test') AND DatasetName=X
  → normal_ng_rate vs test_ng_rate 비교 → 개선율 계산

"와이어 불량에서 Normal 대비 개선이 가장 컸던 리포트는?"
→ DefectCategory='wire_defect' AND VariableGroup IN ('normal','test')
  GROUP BY DatasetName → 개선율 내림차순

"조건 변경 후 일부 불량은 줄었지만 다른 불량이 늘지 않았나?"
→ 동일 DatasetName 내 불량 유형별 normal vs test 비교 (trade-off 확인)
```

### 기타 분석

```
"BRS 제품에서 VP+CD 분리 불량이 높았던 mold는?"
→ ProductType='BRS-161016' AND DefectCategory='assembly_defect' 집계

"TIU-C11 SPL 불량 발생 시 어떤 공정 확인해야 하나?"
→ Summary 검색 + Layer1 원본 + Claude 크로스 분석

"신규 lot 적용 시 리스크가 높은 mold는?"
→ VariableGroup='new_lot' AND NgRate > threshold

"작업자별 불량률 차이가 가장 큰 공정은?"
→ Variable LIKE 'Worker%' GROUP BY DefectCategory

"재교육 후 효과가 있었던 사례는?"
→ Intervention='retrained' → VariableGroup='before' vs 'after' 비교

"특정 Position에서 반복적으로 불량 나는 라인은?"
→ DefectType LIKE '%position%' GROUP BY Line, DefectType
```

### 미래 AI 분석 시나리오

```
"지금 wire offset 불량이 나는데, 과거 유사 사례는?"
→ DefectCategory='wire_defect' 인 모든 dataset의 Summary 검색
  + Normal이 있는 경우 어떤 조건 변경이 효과 있었는지

"이 제품에서 불량률이 5% 넘어간 리포트들의 공통점은?"
→ NgRate > 5 인 데이터 필터 + Claude에게 공통 패턴 분석 요청
```

---

## 스키마가 4개 리포트를 모두 커버하는지

| 항목 | TIU-C11-20 | BRS ALL MOLD | BRS NEW LOT | TIA-338 |
|---|---|---|---|---|
| ProductType | ✅ | ✅ | ✅ | ✅ |
| Line | - | - | - | ✅ |
| Variable | ✅ 시간 | ✅ Mold# | ✅ Mold+Lot | ✅ Worker |
| VariableGroup | - | - | ✅ old/new | ✅ before/after |
| Intervention | - | - | - | ✅ retrained |
| Wide→Long | 불필요 | 불필요 | 불필요 | ✅ position 컬럼 |
| 이미지 연결 | - | - | ✅ Mold 레벨 | ✅ Position+Phase 레벨 |
| 복수 제품 | - | - | - | ✅ L/R 분리 |

**→ 모든 케이스 커버 가능. 스키마 확정.**

---

## 구현 로드맵

| 단계 | 내용 | 상태 |
|---|---|---|
| 6 | 유사 사례 검색: 불량 유형 입력 → 관련 과거 데이터 + Summary | ⬜ 예정 |
| 7 | Claude 자동 개선 제안: 불량 발생 시 어떤 공정 확인해야 하는지 | ⬜ 예정 |
| 8 | 키워드 태그 기반 필터/검색 (DB Data 화면) | ⬜ 예정 |

### 현재 운영 방침

```
[사용자 워크플로우]
1. DRM Excel 스크린샷 촬영 (Ctrl+PrtSc 또는 캡처도구)
2. Input Data 화면에서 Ctrl+V 붙여넣기 (여러 장 가능)
3. Dataset Name + Product Type + Date 입력
4. Analyze 클릭 → Claude가 NormalizedMeasurements 추출
5. 결과 검토 (숫자/레이블 확인) → 이상 없으면 Save
6. DB Data 화면에서 저장 내역 확인

[신뢰도 기준]
- 숫자(Input/NG/NG%): 높음 → 빠른 눈 검토로 충분
- variableGroup 레이블: 중간 → Normal/Test 확인 필수
- defectCategory 매핑: 중간 → 주요 불량 유형 확인 권장
```
