# CLI AI Batch — Runbook for future agents

> 목적: `JinoSupporter.Web` 의 **AI Batch** 기능 (`DataInferenceBatchPage.razor`) 과
> 동일한 결과를 **서버 기동 없이, Anthropic API 호출 없이** CLI 로 만들어
> `RawReports` 미처리분을 `NormalizedMeasurements` + `DatasetSummary` 로 저장한다.
>
> **v3 (2026-04-21 오후)** — 앱의 배치 프로세스 현행화 반영:
> ① UI 에서 **Vision / OCR→Text 두 모드** 선택 (Vision 기본),
> ② **`Kind='excel_paste'` 를 authoritative tiebreaker 로 양 모드에 주입**,
> ③ 사후 `MeasurementValidator.FindIssues` 로 checksum 검증 (Issues 탭),
> ④ `ClaudeService.cs` 프롬프트 line 번호 갱신, `NormalizeFromImagesAsync`
> (line 779~1220) 신규 반영.

---

## 0. TL;DR — 에이전트가 해야 할 일 (요약)

> **🔑 핵심 원칙 (v2, 유지):** Vision/Normalize 작업은 **에이전트(Claude) 가
> 직접 수행**한다. Anthropic API 는 **기본 경로에서 호출하지 않음**.
> 이유: ① `workhost-settings.json` Anthropic 크레딧 자주 소진
> (`HTTP 400 credit balance low`), ② 소량 처리는 에이전트 멀티모달 `Read`
> 로 충분, ③ 사용자 지시 ("API 가 할일을 너가 해라").
> Python 은 **DB IO · 이미지 BLOB export · 트랜잭션 커밋**만 담당.

> **🧭 모드 선택 (v3):** 앱은 두 가지 경로를 지원한다.
> - **Vision** (기본, 앱 UI 기본값) — 이미지 + `excel_paste` 를 그대로 normalize
>   프롬프트에 넣어 **한번에 measurements JSON** 생성. `NormalizeFromImagesAsync`.
>   transcript 를 저장하지 않음 → OCR 캐시 미생성.
> - **OCR→Text** — ① `ExtractStructuredTextAsync` 로 markdown transcript 생성·
>   `Kind='ocr'` 로 캐시 → ② `NormalizeFromTextAsync` 로 JSON. 재분석 시 캐시
>   재사용으로 호출 1회 절약. 디버깅 용이.
>
> CLI 경로에서 **기본은 Vision 모드**: 에이전트가 이미지를 직접 보고 → excel
> paste 가 있으면 그것을 tiebreaker 로 사용 → measurements JSON 을 직접 산출.
> OCR→Text 모드는 **원하는 경우 에이전트가 markdown transcript 를 중간 산출물
> 로 만들어 `Kind='ocr'` 에 저장하도록 선택할 수 있음** (품질 확인/디버깅 목적).

1. `workhost-settings.json` 에서 **DB 경로** 확보 (API 키는 §1.2 폴백 전용)
2. `SELECT ... LEFT JOIN NormalizedMeasurements ... WHERE n.DatasetName IS NULL`
   로 **미처리 목록** 조회 (기본적으로 `BatchExcluded=0` 제외 포함)
3. 각 dataset 마다:
   - Python 이 이미지 BLOB 을 `_tmp_img_{i}.png` 로 export
   - `Kind='excel_paste'` 레코드가 있으면 그 텍스트를 **authoritative tiebreaker**
     로 함께 사용 (§2b) — 이미지의 OCR-애매한 숫자는 paste 값으로 덮어쓰기
   - **에이전트가 `Read` 로 이미지 직접 분석** → §5 의 Vision-Normalize 규칙
     (ClaudeService.cs **line 779~1220**) 으로 measurements JSON 산출
     (OCR→Text 모드를 택했다면 §5 Extract + NormalizeFromText 규칙 차례로 적용)
   - 에이전트가 `MeasurementValidator` 규칙 (§2e) 으로 **자체 체크섬 검증**
     후 이상 있으면 재추출
   - Python 이 단일 트랜잭션으로 `NormalizedMeasurements` 를 `DELETE + INSERT`,
     `DatasetSummary` 를 `ON CONFLICT(DatasetName) DO UPDATE` 로 upsert,
     (OCR→Text 모드 한정) `RawReportText(Kind='ocr')` 도 upsert (실패시 롤백)
4. 종료 시 `=== BATCH DONE ===` 요약 출력 + 임시 파일 (`_tmp_commit.py`,
   `_tmp_img_*.png`, `_batch_work/`) 전원 삭제

**영구 보관 런처 없음** (사용자 요구). 매 실행마다 임시 스크립트 작성 → 실행 → 삭제.

**폴백 (API 경로):** 에이전트가 한 번에 다루기 힘든 **대량 배치** (예: >10건,
수십만 토큰) 에서만 §1.2 의 DPAPI API 키로 `_tmp_ai_batch.py` 구성.
실행 전 **반드시 사용자에게 크레딧 충전 상태 확인** — 크레딧 소진 사고 재발 방지.

---

## 1. 환경 해석 (resolution order)

### 1.1 SQLite DB 경로

```
우선순위:
  1. appsettings.json   →  "Database:Path"
  2. workhost-settings.json  →  "DataInference.DatabasePath"      ← 현재 사용
  3. default               →  %LOCALAPPDATA%\JinoWorkHost\process-review.db
```

`settings-bootstrap.json` (`%LOCALAPPDATA%\JinoWorkHost\`) 의 `SettingsFilePath` 가
실제 `workhost-settings.json` 위치를 가리킨다. 현재 해당 파일 경로:
`D:\000. MyWorks\005. Program\새 폴더\workhost-settings.json`

> ⚠️ `%LOCALAPPDATA%\JinoWorkHost\process-review.db` 는 빈 부트스트랩 DB다.
> 실제 데이터는 `workhost-settings.json` 이 가리키는 경로 (현재
> `D:\000. MyWorks\000. 일일업무\04. DB\process-review.db`, ~665 MB) 에 있다.
> 절대 default 로 폴백하지 말 것.

### 1.2 Claude API 키 (폴백 경로 전용)

> **기본 경로에서는 필요 없음.** 아래는 §0 의 "폴백 (API 경로)" 를 실제로
> 가동할 때만 적용. 먼저 사용자에게 Anthropic Console 크레딧 잔액 확인 요청.

```
우선순위 (ClaudeService.cs 의 생성자 로직과 일치):
  1. Settings 테이블 (DB)  →  "Claude:ApiKey"                 ← 현재 비어있음
  2. workhost-settings.json  →  "Claude.EncryptedApiKey"       ← 현재 사용 (DPAPI)
  3. appsettings.json        →  "Claude:ApiKey"
```

DPAPI 복호화는 `CurrentUser` scope 이므로 **반드시 본인 계정 세션에서 실행**.
Python 에서는 `pywin32` (`win32crypt.CryptUnprotectData`) 사용.

```python
import json, base64, win32crypt
cfg  = json.load(open(SETTINGS, encoding='utf-8'))
blob = base64.b64decode(cfg["Claude"]["EncryptedApiKey"])
_, pt = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
api_key = pt.decode("utf-8")   # → sk-ant-api03-...
```

### 1.3 작업 디렉토리

`./_batch_work/<sanitized-dataset-name>/` 에 이미지 export.
정상 처리 후 **디렉토리 전체 삭제** (사용자 기본 방침).
JSON parse 실패 시에만 `error.txt` 를 해당 폴더에 남김.

---

## 2. STEP 1 — 미처리 목록 쿼리

```sql
SELECT r.DatasetName, r.ProductType, r.ReportDate
FROM   RawReports r
LEFT JOIN NormalizedMeasurements n ON n.DatasetName = r.DatasetName
WHERE  r.BatchExcluded = 0
  AND  n.DatasetName IS NULL
GROUP  BY r.DatasetName
ORDER  BY r.CreatedAt;
```

규모 확인을 위해 동시에 아래도 뽑으면 비용 예측에 유용:

```sql
-- 각 dataset 의 이미지 수 + RawReportText 캐시 유무
SELECT  r.DatasetName,
        (SELECT COUNT(*) FROM RawReportImages i WHERE i.DatasetName=r.DatasetName) AS imgN,
        (SELECT 1        FROM RawReportText  t WHERE t.DatasetName=r.DatasetName) AS hasCache
FROM    RawReports r
LEFT JOIN NormalizedMeasurements n ON n.DatasetName=r.DatasetName
WHERE   r.BatchExcluded=0 AND n.DatasetName IS NULL;
```

---

## 3. STEP 2 — 각 dataset 처리 파이프라인

### 2a. 이미지 추출

```sql
SELECT MediaType, ImageData, FileName, SortOrder
FROM   RawReportImages
WHERE  DatasetName = :name
ORDER  BY SortOrder;
```

**기본 경로**: Python 이 BLOB 을 `_tmp_img_{i}.png` 로 저장 → 에이전트가
`Read` 로 직접 열람. Claude Code `Read` 는 PNG/JPEG 이미지를 그대로 인식하므로
**리사이즈 불필요**. 에이전트 측에는 Anthropic API 의 8,000 px 하드리미트가
동일하게 걸리므로 `Read` 호출 시 400 이 나면 §2a-fallback 처리.

**폴백 (API 경로) 에서만**: **✱ Anthropic Vision 하드리미트: 가로/세로 어느 한
변이라도 8,000 px 초과 시 HTTP 400.** 무조건 리사이즈 체크. 바이트 크기(3.5MB)
만 보는 건 불충분 — 작은 파일이라도 픽셀 사이즈는 클 수 있다.

```python
from PIL import Image
MAX_DIM = 7900   # 안전 마진

def resize_if_oversize(raw: bytes, mt: str) -> tuple[bytes, str]:
    img = Image.open(io.BytesIO(raw))
    w, h = img.size
    if max(w, h) <= MAX_DIM: return raw, mt
    scale = MAX_DIM / max(w, h)
    img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG" if mt=="image/png" else "JPEG",
             quality=88 if mt=="image/jpeg" else None, optimize=True)
    return buf.getvalue(), mt
```

`MediaType` 은 선언값이 틀릴 수 있으니 **매직 바이트로 재검출**
(`ClaudeService.cs::DetectMediaType` 와 동일 규칙):
- `89 50 4E 47` → `image/png`
- `FF D8 FF`    → `image/jpeg`
- `47 49 46`    → `image/gif`
- `52 49 46 46 .. 57 45 42 50` → `image/webp`

### 2b. Tiebreaker 확보 + (OCR→Text 모드 한정) Transcript 캐시 조회

`RawReportText` 는 복합 PK `(DatasetName, Kind)` 로 **두 종류**를 공존시킨다:

| Kind | 의미 | Batch 사용 (v3) |
|---|---|---|
| `ocr` | Vision/에이전트 가 생성한 **구조화 markdown transcript** | OCR→Text 모드 한정 — 캐시로 재사용 |
| `excel_paste` | 사용자가 입력 시점에 엑셀에서 붙여넣은 **raw TSV** | **양 모드 모두 authoritative tiebreaker 로 주입** |

**v3 중요 변경:** `excel_paste` 는 더 이상 "보조" 가 아니라 **정답 tiebreaker**.
앱 쪽 `DataInferenceBatchPage.razor:400` 이 두 모드 모두에서 이 값을 읽어
Claude 호출 때 주입한다:
- Vision 모드 → `NormalizeFromImagesAsync(rawText: excelPaste)` — 프롬프트에
  "AUTHORITATIVE RAW EXCEL TEXT (use this for exact cell values; prefer it
  over OCR'd numbers from the screenshot when they disagree)" 블록으로 첨부
- OCR→Text 모드 → `ExtractStructuredTextAsync(rawExcelText: excelPaste)` —
  transcript 단계에서 "only as a tiebreaker for cell values" 로 첨부
  (Normalize-from-text 단계에는 주입 안됨 — 그 시점엔 이미 transcript 에 반영됨)

조회 쿼리:
```sql
-- 양 모드 공통: tiebreaker 로 쓸 excel paste
SELECT ExtractedText FROM RawReportText
WHERE  DatasetName = :name AND Kind = 'excel_paste';

-- OCR→Text 모드 한정: 이전 실행에서 저장한 markdown transcript 캐시
SELECT ExtractedText FROM RawReportText
WHERE  DatasetName = :name AND Kind = 'ocr';
```

**Kind 필터 필수** — 2026-04-21 실수: Kind 필터 없이 조회해 `excel_paste` 의
TSV 를 OCR 결과로 오인함. 반드시 명시적으로 filter.

에이전트 처리 지침:
- `excel_paste` 가 있으면 에이전트는 이미지에서 읽은 값과 paste 값을
  **셀 단위로 대조**, 불일치 시 paste 값을 채택 (프롬프트 규칙과 동일)
- `Kind='ocr'` 캐시가 있고 사용자가 OCR→Text 모드를 지시하면 그대로 재사용,
  transcript 미존재면 에이전트가 §5 Extract 규칙으로 생성
- 기본 Vision 모드에서는 transcript 를 만들지 않음 (앱과 동일)

**폴백 API 호출 파라미터:**
```python
model       = "claude-sonnet-4-6"
max_tokens  = 64000      # ← 기본값 16384 는 대형 리포트에서 잘림
timeout_sec = 420        # ← 180s 는 Sonnet 4.6 장문 출력엔 부족
```

프롬프트 플레이스홀더 (공통): `{{datasetName}}`, `{{productType}}`, `{{testDate}}`
(+ OCR→Text Normalize 시 `{{extractedText}}`).

### 2c. Normalize (→ measurements JSON)

- **기본 경로 (Vision):** 에이전트가 이미지 + (있으면) `excel_paste` 를 함께
  보고 §5 의 Vision-Normalize 규칙 (`ClaudeService.cs` **line 779~1220**) 을
  자기 자신에게 적용 → measurements JSON 직접 출력. Python 의 `_tmp_commit.py`
  에는 그 JSON 을 파이썬 dict 로 하드코딩해서 커밋.
- **기본 경로 (OCR→Text):** 에이전트가 먼저 §5 Extract 규칙 (line 456~542) 으로
  transcript 생성 → 이어서 §5 NormalizeFromText 규칙 (line 583~713) 으로 JSON
  산출. transcript 는 `Kind='ocr'` 로 캐시 저장.
- **폴백 (API) 경로:** 동일 프롬프트를 verbatim 으로 API POST.
  - Vision 모드: `NormalizeFromImagesAsync(images, ..., rawText=excelPaste)`
  - OCR→Text 모드: `ExtractStructuredTextAsync(..., rawExcelText=excelPaste)`
    이어서 `NormalizeFromTextAsync(transcript, ...)`

### 2d. 응답 파싱 (폴백 경로만)

폴백 API 경로에서 모델 응답이 코드펜스/서문을 붙이는 경우가 있으므로 견고하게:

```python
def extract_json_object(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        nl = raw.find("\n")
        if nl >= 0: raw = raw[nl+1:]
        if raw.rstrip().endswith("```"): raw = raw[:raw.rfind("```")]
    o = raw.find("{"); c = raw.rfind("}")
    if o >= 0 and c > o: raw = raw[o:c+1]
    return raw.strip()
```

파싱 실패 시 → **SKIP** + `./_batch_work/<name>/error.txt` 에 원본 응답 저장.
**DB 에는 아무것도 쓰지 말 것 (부분 저장 금지).**

기본 경로 (에이전트 직접) 에서는 파이썬 dict 로 바로 구조화해서 커밋하므로
이 단계 불필요.

### 2e. 자체 체크섬 검증 (`MeasurementValidator` 규칙)

앱은 커밋 후 `MeasurementValidator.FindIssues` (`Services/MeasurementValidator.cs`)
로 사후 검증하고 UI 의 **Issues** 탭에 경고 배지를 띄운다. CLI 에서는 커밋
**이전에** 에이전트가 같은 규칙을 자체 적용해 이상 있으면 다시 뽑는다.

검증 규칙 요약 (실제 파일 참조):
- `DefectType == "__ALIGN_ERROR__"` 마커가 있으면 컬럼 정렬 실패 → 재추출
- `(Variable, VariableGroup, Line, CheckType, InputQty, OkQty, NgTotal)` 로
  그룹핑해서 각 그룹의 **positive defectCount 합** 이 `NgTotal` 이상인지 확인
- `NgTotal > 0` 이면 최소 1개의 per-defect row 존재해야 함 (aggregate-only
  테이블, criterion-level 테이블, picture-sample (Input=OK=0) 은 skip 예외)
- Undercount (`sum < NgTotal`) 은 "missing defect row(s)" 경고
- Overcount (`sum > NgTotal`) 은 silently skip — 두 개의 별도 sub-table 이
  동일 키를 갖는 드문 케이스 (validator 가 그렇게 처리하므로 동일 준수)

### 2f. 단일 트랜잭션으로 DB 커밋

앱의 `WebRepository.SaveNormalizedMeasurements` 는 DELETE + INSERT (트랜잭션),
`SaveDatasetSummaryRecord` 는 `ON CONFLICT(DatasetName) DO UPDATE` 로 upsert.
CLI 에서도 동일한 의미로 쓴다.

```python
cur.execute("BEGIN")

# measurements — DELETE + INSERT (앱 동작과 동일)
cur.execute("DELETE FROM NormalizedMeasurements WHERE DatasetName=?", (name,))
for m in obj["measurements"]:
    cur.execute("""
        INSERT INTO NormalizedMeasurements
          (DatasetName, ProductType, TestDate, Line, CheckType, Variable,
           VariableDetail, VariableGroup, Intervention, InputQty, OkQty,
           NgTotal, NgRate, DefectCategory, DefectType, DefectCount, CreatedAt)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (name, m.get("productType") or product, m.get("testDate") or reportDate,
         m.get("line",""), m.get("checkType",""), m.get("variable",""),
         m.get("variableDetail",""), m.get("variableGroup",""), m.get("intervention",""),
         int(m.get("inputQty",0)), int(m.get("okQty",0)), int(m.get("ngTotal",0)),
         float(m.get("ngRate",0)), m.get("defectCategory",""), m.get("defectType",""),
         int(m.get("defectCount",0)), now))

# summary — PK 가 DatasetName 이므로 ON CONFLICT DO UPDATE 로 upsert
cur.execute("""
    INSERT INTO DatasetSummary
      (DatasetName, ProductType, Summary, KeyFindings, Tags, CreatedAt,
       Purpose, TestConditions, RootCause, Decision, RecommendedAction)
    VALUES (?,?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(DatasetName) DO UPDATE SET
      ProductType=excluded.ProductType, Summary=excluded.Summary,
      KeyFindings=excluded.KeyFindings, Tags=excluded.Tags, CreatedAt=excluded.CreatedAt,
      Purpose=excluded.Purpose, TestConditions=excluded.TestConditions,
      RootCause=excluded.RootCause, Decision=excluded.Decision,
      RecommendedAction=excluded.RecommendedAction""",
    (name, product, obj.get("summary",""), obj.get("keyFindings",""),
     json.dumps(obj.get("tags") or [], ensure_ascii=False), now,
     obj.get("purpose",""), obj.get("testConditions",""),
     obj.get("rootCause",""), obj.get("decision",""),
     obj.get("recommendedAction","")))

# OCR→Text 모드에서만: 새로 만든 markdown transcript 를 Kind='ocr' 로 upsert.
# Vision 모드에서는 transcript 를 만들지 않으므로 이 블록 skip.
if new_transcript is not None:
    cur.execute("""INSERT INTO RawReportText
                     (DatasetName, Kind, ExtractedText, CreatedAt)
                   VALUES (?, 'ocr', ?, ?)
                   ON CONFLICT(DatasetName, Kind)
                   DO UPDATE SET ExtractedText=excluded.ExtractedText,
                                 CreatedAt=excluded.CreatedAt""",
                (name, new_transcript, now))

conn.commit()
```

예외 시 `conn.rollback()` 후 `raise` — 다음 dataset 으로 진행.

`RawReports.BatchedAt` 은 별도 컬럼이 아니다. `GetAllRawReports` 쿼리가
`COALESCE(DatasetSummary.CreatedAt, MAX(NormalizedMeasurements.CreatedAt), '')`
로 유도하므로 summary row 만 upsert 해도 UI 에 "Batched At" 이 표시됨.

---

## 4. 실패 처리 규칙 (실전 검증됨)

### 기본 경로 (에이전트 직접)

| 상황 | 동작 |
|---|---|
| 이미지 0장 | SKIP, `[SKIP no-images]` 로그 |
| 이미지 전부 illegible | SKIP, error.txt 에 이유 기록, DB 미수정 |
| 판독 불확실 cell | 프롬프트 §5 규칙대로 `[?]` 표기, 추측 금지 |
| 에이전트 출력 JSON 이 스키마 위반 | 에이전트가 바로잡아 재생성, DB 는 최종본만 1회 커밋 |
| 대량 배치로 context 한계 | §0 의 API 폴백 경로로 전환 (사용자 승인 후) |

### 폴백 경로 (API)

| 상황 | 동작 | 비고 |
|---|---|---|
| HTTP 400 credit balance low | **즉시 중단, 사용자에게 충전 요청** | 2026-04-21 실사례 |
| 이미지 변수 >8000 px | **리사이즈 후 진행** (SKIP 하지 말 것) | PIL LANCZOS → ≤7900 px |
| JSON parse 실패 | SKIP, error.txt 저장, **DB 미수정** | 부분 저장 금지 |
| HTTP read timeout | dataset SKIP, error.txt 에 stacktrace | `timeout=420s` 로 줄이기 |
| HTTP 400 (기타) | SKIP, error.txt 저장 | 요청 자체가 기각되면 토큰 청구 없음 |
| 출력 토큰 상한 도달 | `max_tokens` 를 올려 재시도 | 기본 64,000; 극단 케이스만 더 상향 |

### 출력 잘림 (output truncation) 대응 (폴백 전용)

`usage.output_tokens == max_tokens` 이면 거의 확실히 잘린 것 →
동일 dataset 에 `max_tokens` 를 2배 올려 즉시 1회 retry. 권장 상한:

- 1차: 64,000 (기본값)
- 2차: 128,000 (beta 헤더 `anthropic-beta: output-128k-2025-02-19` 필요할 수 있음)

2026-04-21 관측: 가장 rows 많은 케이스(137 rows, Normal-rich multi-page) 가
`output_tokens=29,286` 이므로 **64,000 은 충분한 마진**. 128k 까지 가는 케이스는 희귀.

---

## 5. 프롬프트 원문 (single source of truth)

**복사해서 수정하지 말고 원본을 항상 참조.** `ClaudeService.cs` 가 바뀌면
이 문서도 재검토.

| # | 프롬프트 | 원본 위치 (line) | 메서드 | 역할 | max_tok | model |
|---|---|---|---|---|---|---|
| ① | Extract (vision → markdown) | **456~542** | `ExtractStructuredTextAsync` | 이미지 → 구조화 markdown transcript (OCR→Text 모드 1단계) | 64,000 | claude-sonnet-4-6 |
| ② | NormalizeFromText (markdown → JSON) | **583~713** | `NormalizeFromTextAsync` | transcript → measurements JSON (OCR→Text 모드 2단계) | 64,000 | claude-sonnet-4-6 |
| ③ | NormalizeFromImages (vision → JSON) | **779~1220** | `NormalizeFromImagesAsync` | 이미지 직접 → measurements JSON (**Vision 모드, 앱 기본**) | 64,000 | claude-sonnet-4-6 |

③ 은 ② 와 같은 출력 스키마를 공유하지만 **STEP 0 (LAYOUT CLASSIFICATION
A/B/C/D/E/F)** 이 훨씬 상세. Vision 모드에서는 반드시 이 분류부터 적용.

프롬프트 내 플레이스홀더 (C# `$$"""...{{x}}..."""` 구문):
- `{{datasetName}}`, `{{productType}}`, `{{testDate}}` (①②③ 공통)
- `{{extractedText}}` (② 에서만, transcript 주입 자리)

Tiebreaker 블록은 **별도 content block** 으로 프롬프트 텍스트 뒤에 붙음
(`CallWithContentAsync` 의 blocks 배열, type=text). 앞부분에 "AUTHORITATIVE
RAW EXCEL TEXT" 리터럴이 들어가며 코드펜스 안에 `excel_paste` 원문이 그대로
붙는다. 에이전트가 수동 재현할 때도 동일한 방식으로 구성 (§2b 참조).

Python 에서 프롬프트를 string 조작할 때 JSON 예시의 `{` `}` 를 이스케이프
(`.format` 사용 시 `{{` `}}`) 하거나 단순 `str.replace()` 로 치환할 것.

---

## 6. 기대 출력

```
=== STEP 1: N unprocessed dataset(s), mode=Vision ===

[1/N] <DatasetName>
    product=…  reportDate=…
    K image(s) exported
    tiebreaker: excel_paste ✓ (…chars)   # 또는 excel_paste ✗ (none)
    Vision normalize … done — …chars   # OCR→Text 모드면 "OCR extract … → Normalize …"
    self-check: 0 issues               # MeasurementValidator 규칙
    ✔ M measurement rows, tags=T, summary=Sc

    ── progress: 5/N in …s ──          # 5 단위마다

[k/N] <Next dataset>
    ...
    ✗ <reason>                         # 실패 시

=== BATCH DONE ===
Mode:      Vision  (또는 OCR→Text)
Processed: X
Skipped:   Y (reasons: …)
Issues:    Z datasets with self-check warnings (see log)
Elapsed:   …s
```

---

## 7. 비용 가이드

### 기본 경로 (에이전트 직접)

**Anthropic Console 크레딧 소비 없음.** Claude Code 세션의 tool use 토큰만
사용. 대략 dataset 1건당 이미지 읽기 + transcript + JSON 출력 합쳐 컨텍스트
수만 토큰 수준이며, 세션 토큰 예산이 충분하면 체감 비용 없음.

### 폴백 경로 (Sonnet 4.6 표준 요율)

- 입력 $3 / MTok · 출력 $15 / MTok
- 2026-04-21 실측: 6 datasets × 11 images → **~230k tokens ≈ $2.37**
  (재시도 포함; 재시도 제거 시 ~130k ≈ $1.50)
- 리포트 1건 평균: **~20~40k tokens / $0.25~0.50** (페이지·테이블 수에 비례)

> 🧾 차감 계정: workhost-settings.json 의 `sk-ant-api03-…` → **Anthropic Console 크레딧**
> Claude Code 세션(구독 or 별도 API 키)과는 완전 분리. 2026-04-21 에 크레딧
> 소진으로 HTTP 400 발생했던 사례 있음 — 폴백 경로 실행 전 잔액 확인 필수.

---

## 8. Known gotchas — 재발 방지 체크리스트

### 공통
- [ ] DB 경로는 **항상 `workhost-settings.json` 을 먼저** 읽는다 (default 로 가면 빈 DB)
- [ ] `RawReportText` 조회 시 **반드시 `Kind=…` 필터** 명시 — `excel_paste` 와 `ocr` 을 혼동 금지
- [ ] `excel_paste` 는 **tiebreaker 로 반드시 주입** (v3) — Vision/OCR→Text 양 모드 모두
- [ ] RawReportText 커밋은 `ON CONFLICT(DatasetName, Kind) DO UPDATE` (Kind 명시)
- [ ] `DatasetSummary` 커밋은 `ON CONFLICT(DatasetName) DO UPDATE` — PK 단일, upsert
- [ ] `RawReports.BatchedAt` 은 컬럼이 아니라 COALESCE 유도값 — 따로 쓰려 하지 말 것
- [ ] `BatchExcluded=0` 으로 제외 대상 필터 (미처리 쿼리 WHERE 에 반영 — §2)
- [ ] JSON 파싱 실패 → **반드시 DB 롤백**, error.txt 저장 (부분 저장 금지)
- [ ] Run 중 WPF/Blazor 앱이 동시에 돌면 DB busy 가능 → `PRAGMA busy_timeout = 30000`
- [ ] RawReportText 캐시는 **normalize 성공 후에만** 저장 (실패한 transcript 캐시 금지)
- [ ] dataset name 이 파일명으로 쓰일 때 Windows 금지 문자 sanitize (`<>:"/\\|?*\r\n\t`)
- [ ] 종료 후 임시 파일 (`_tmp_commit.py`, `_tmp_ai_batch.py`, `_tmp_img_*.png`,
  `_batch_work/`) **전원 삭제** (사용자 기본 방침)
- [ ] Python `print` 한글 → `sys.stdout.reconfigure(encoding="utf-8")` 필수
  (Windows cp1252 기본값이면 `UnicodeEncodeError`)
- [ ] 커밋 **직전**에 §2e `MeasurementValidator` 규칙으로 self-check,
  align-error / undercount / missing-defect 발견 시 재추출

### 기본 경로 (에이전트 직접) 전용
- [ ] 이미지는 `Read` 로 **직접 확인** 후 처리 — 페이지 수, 테이블 수, 헤더 계층 선(先)파악
- [ ] **Vision 모드 (기본):** transcript 안 만듦 → `Kind='ocr'` 에 쓰지 말 것
- [ ] **OCR→Text 모드:** transcript 를 명시적으로 만들고 `Kind='ocr'` 로 upsert
- [ ] `excel_paste` 있으면 이미지 숫자와 **셀 단위로 대조** 후 paste 값 우선 채택
- [ ] Normalize 결과 JSON 은 `_tmp_commit.py` 에 **파이썬 dict 로 하드코딩** (API 호출 없음)
- [ ] defectCategory 매핑 표 (§5 프롬프트 ③) 준수 — 특히 Audiobus/Hearing/THD 헤더 규칙
- [ ] Before/After · Normal/Test · new_lot 판별 → variableGroup 반드시 채우기
- [ ] Vision 프롬프트의 STEP 0 LAYOUT CLASSIFICATION (A/B/C/D/E/F) 을 각 테이블에
  먼저 적용 — 잘못 분류하면 다운스트림에 잘못된 데이터 유입

### 폴백 경로 (API) 전용
- [ ] **실행 전 사용자에게 Anthropic Console 크레딧 잔액 확인 요청** (400 재발 방지)
- [ ] API 키는 **DPAPI 복호화 필수** — `CurrentUser` scope 라 다른 계정으로 실행 불가
- [ ] 이미지는 **바이트가 아니라 픽셀** 로 한계 체크 (>8000 px → HTTP 400)
- [ ] `excel_paste` 주입은 **별도 text block** 으로 — 프롬프트 템플릿 안에 치환하지 말 것
- [ ] `max_tokens` 은 최소 **64,000** 으로 시작 — ClaudeService.cs 기본 16,384 는 부족 케이스 있음
- [ ] HTTP read `timeout` 은 최소 **420 s** — Sonnet 4.6 장문 응답은 느릴 수 있음
- [ ] `raise_for_status()` 만 쓰지 말 것 — 400 바디가 잘려서 진짜 원인 (예: credit)
  이 안 보임. `r.status_code >= 400` 분기에서 `r.text[:2000]` 로그 남기기

---

## 9. 에이전트 호출 템플릿

다음부터는 이 줄 한 줄로 실행시키면 된다:

> **"CLI_AI_BATCH.md 읽고 AI Batch 실행해"**

### 에이전트 동작 순서 (기본 경로, v3)
1. §1.1 환경 해석으로 DB 경로 획득
2. §2 STEP 1 쿼리로 미처리 목록 확인 (`BatchExcluded=0` 포함)
3. 각 dataset 에 대해:
   - Python 으로 이미지 BLOB export → `_tmp_img_*.png`
   - `Kind='excel_paste'` 쿼리로 tiebreaker 텍스트 확보 (있을 수도/없을 수도)
   - (OCR→Text 모드를 택했고 `Kind='ocr'` 캐시 존재 시) 캐시 재사용, 아니면:
     - 에이전트가 `Read` 로 이미지를 **직접 시각 분석**
     - Vision 모드 → §5 프롬프트 ③ 규칙 (line 779~1220) 적용 → measurements JSON 직접 출력
     - OCR→Text 모드 → §5 프롬프트 ① 로 transcript 생성 → ② 로 JSON
     - excel_paste 가 있으면 이미지 숫자와 **셀 단위 대조 후 paste 값 우선**
   - 커밋 전 §2e 자체 checksum 검증 → 이상 있으면 재추출
   - `_tmp_commit.py` 에 JSON 을 파이썬 dict 로 하드코딩 → 실행 → 트랜잭션 커밋
     (Vision 모드는 `Kind='ocr'` 쓰기 skip, OCR→Text 모드는 upsert)
4. §6 형식 요약 출력 → 임시 파일 전원 삭제

### 기본 모드 선택 기준
- **Vision (권장·기본)** — 이미지 1~몇 장, transcript 불필요, 한번에 끝내고 싶을 때
- **OCR→Text** — ① 재분석이 여러 번 있을 것으로 예상되어 transcript 캐시 이득이 클 때,
  ② 결과 품질 디버깅을 위해 중간 transcript 를 사람이 확인하고 싶을 때,
  ③ 동일 dataset 을 여러 변형된 Normalize 규칙으로 돌려 비교하고 싶을 때

### 폴백 전환 조건 (API)
대량 배치 (>10건 or 예상 토큰 >30만) 이거나 에이전트의 이미지 판독 품질이
명백히 낮은 경우에만. 전환 시 사용자 사전 승인 + 크레딧 확인 필수.

---

*이 runbook 은 v3 (2026-04-21 오후) — JinoSupporter.Web 의 AI Batch 프로세스
현행화를 반영한다:
① Vision / OCR→Text 두 모드 공존 (UI 기본 Vision),
② `excel_paste` 가 양 모드에서 authoritative tiebreaker 로 주입,
③ `NormalizeFromImagesAsync` (line 779~1220) 신규 프롬프트 반영,
④ `MeasurementValidator` 의 사후 체크섬 검증 규칙을 CLI 에서는 **커밋 직전**
에 self-check 로 적용,
⑤ `DatasetSummary` 는 `ON CONFLICT(DatasetName) DO UPDATE` upsert,
⑥ `RawReports.BatchedAt` 은 별도 컬럼이 아니라 COALESCE 유도값임을 명시.

이전 실전 교훈:
- v2: 6 datasets × API 경로 → 307 rows (재시도 포함),
- v2: 1 dataset × 에이전트 직접 → 12 rows (API credit 소진 HTTP 400 회피).
- v3: ClaudeService.cs 변경에 맞춰 프롬프트 line 번호 및 모드 구조 동기화.

이후 새 실패 패턴이 발견되면 §4 와 §8 에 추가할 것.*
