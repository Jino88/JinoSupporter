# CLI Ask AI — Runbook for future agents

> 목적: `JinoSupporter.Web` 의 **Ask AI** 기능 (`DataInferenceAskPage.razor`) 과
> 동일한 결과를 **서버 기동 없이, Anthropic API 호출 없이** CLI 로 만들어
> `AskAiHistory` 테이블에 저장하고 터미널에도 답을 출력한다.
>
> **v1 (2026-04-23)** — 구독(MAX 5x) 경로 전용. `NormalizeFromImages` 류와는
> 달리 vision 불필요 — 텍스트 context 합성 + reasoning 만 수행.

---

## 0. TL;DR — 에이전트가 해야 할 일

1. **요청 읽기**: `tmp/ask_request.json` 의 `question`, `language`, `productTypeFilter` 확보
2. **DB 경로 확인**: `workhost-settings.json` 의 `DataInference.DatabasePath` (기본 경로 폴백 금지 — CLI_AI_BATCH.md §1.1 참조)
3. **Context 수집**: `FilteredReports` 에 해당하는 dataset 마다 summary + measurements 요약 블록 합성 (§2)
4. **Reasoning**: §3 프롬프트 규칙 그대로 적용 — overall + per-dataset 답변 산출
5. **커밋 + 출력**: `AskAiHistory` 에 single-row INSERT (`last_insert_rowid()` 회수), 터미널에 포맷해서 출력
6. **청소**: `tmp/ask_request.json` 과 `_tmp_*.py` 삭제

**핵심 원칙 (CLI_AI_BATCH.md 와 동일)**: 에이전트가 직접 reasoning 수행 — Anthropic API 호출 없음. Python 은 DB IO 만 담당.

---

## 1. 입력 해석

### 1.1 Request 파일
`tmp/ask_request.json` 스키마:
```json
{
  "question":           "BRS-161016 frequently has SPL NG — how should we improve this?",
  "language":           "Korean",        // "English" | "Korean" | "Vietnamese"
  "productTypeFilter":  "",              // "" = all product types
  "createdAt":          "2026-04-23T12:34:56.789Z"
}
```
파일이 없거나 `question` 이 비어있으면 **즉시 중단**, 터미널에 "No question" 출력.

### 1.2 DB 경로
`CLI_AI_BATCH.md §1.1` 과 동일. `workhost-settings.json` 의 `DataInference.DatabasePath` 를 우선. 현재 경로: `D:\000. MyWorks\000. 일일업무\04. DB\process-review.db`.

---

## 2. Context 빌드 (앱의 `BuildDatasetsContext` 미러)

### 2.1 대상 dataset 쿼리
```sql
SELECT DISTINCT r.DatasetName, r.ProductType, r.ReportDate
FROM   RawReports r
WHERE  r.BatchExcluded = 0
  AND  (:pt = '' OR r.ProductType = :pt)
ORDER BY r.DatasetName;
```
(`:pt` 는 `productTypeFilter` — 빈 문자열이면 전체)

### 2.2 dataset 당 block 생성
각 dataset 에 대해 **둘 중 하나라도 있으면** 포함, 전부 비어있으면 skip:
- `DatasetSummary` 에 `Summary`/`KeyFindings`/`Tags` 가 하나라도 있거나
- `NormalizedMeasurements` 에 row ≥ 1개

```sql
-- Summary 필드 (없으면 빈 문자열)
SELECT Summary, KeyFindings, Tags, Purpose, TestConditions,
       RootCause, Decision, RecommendedAction
FROM   DatasetSummary WHERE DatasetName = :name;

-- Measurements — defect 통계 + normal/test 비교에 사용
SELECT Line, CheckType, Variable, VariableGroup, InputQty, OkQty,
       NgTotal, NgRate, DefectType, DefectCount
FROM   NormalizedMeasurements WHERE DatasetName = :name;
```

블록 포맷 (`DataInferenceAskPage.razor:471-484` 와 1:1 매칭):
```
───── [{idx}] Dataset: {DatasetName}
    ProductType:       {ProductType}
    Date:              {ReportDate}
    Tags:              {tag1, tag2, ...}            ← json 파싱 후 join
    Purpose:           {Purpose}
    TestConditions:    {TestConditions}
    RootCause:         {RootCause}
    Decision:          {Decision}
    RecommendedAction: {RecommendedAction}
    Summary:           {Summary}
    KeyFindings:       {KeyFindings}
    TopDefects:        maxNgRate={maxNg:F1}%, topDefects=[{type(count), ...}]
    NormalVsTest:      normal={n:F1}%, test={t:F1}%, improvement={i:F0}%, best/worst=...
```
빈 필드는 그 줄 자체를 생략. `Tags` 는 DB 에 JSON array 문자열로 저장돼 있음 → `json.loads` 후 `', '.join(...)`.

### 2.3 파생 통계

**TopDefects** (앱 `BuildDefectStats` 동일):
- `measurements` 에서 `DefectType≠''` && `DefectCount>0` 만 필터
- `DefectType` 으로 GroupBy → 합산 → 내림차순 Top5
- `maxNgRate` = `max(NgRate)` 전체 행
- 포맷: `"maxNgRate=12.5%, topDefects=[SPL(23), Audiobus(12), THD(8)]"`
- defect row 없으면 `"topDefects=[none]"`

**NormalVsTest** (앱 `BuildNormalVsTestStats` 동일):
- `VariableGroup` 이 `'normal'` 또는 `'test'` 인 aggregate row (`DefectType=''`) 만 대상
- 각 그룹 가중평균 NG rate: `sum(NgTotal) / sum(InputQty) * 100`
- 둘 다 있고 normal > 0 이면 `improvement = (normal-test)/normal * 100`
- `'test'` aggregate 중 `(Line, CheckType, Variable)` 로 GroupBy → NG rate 최저/최고 → best/worst 조건 도출
- 해당 dataset 에 normal/test 구분 없으면 **블록 생략**

---

## 3. Reasoning 프롬프트 (ClaudeService.cs `AskAiAsync` @ line 1343-1374 미러)

에이전트가 자기 자신에게 적용할 규칙:

```
You are a manufacturing quality improvement assistant.

A user has asked a question about a production problem. Answer it USING ONLY the
information found in the registered dataset reports below.

══ STRICT RULES ══
1. Do NOT use external/general knowledge. Only use facts present in the reports below.
2. If no registered report contains relevant information, set "overall" to a short
   {lang} notice that no relevant data was found, and return an empty "perDataset"
   array. Do not invent an answer.
3. Produce ONE entry in "perDataset" for EVERY dataset that genuinely contributes
   to the answer. Copy "datasetName" VERBATIM from the "Dataset:" header.
4. In each per-dataset "answer": explain in {lang} what this SPECIFIC dataset shows
   and how it addresses the user's question. Cite concrete values (NG rate, defect
   type, product type, date, specific findings). 2-5 sentences is ideal.
5. Do NOT include datasets that are irrelevant to the question.
6. In "overall": give a 2-3 sentence {lang} synthesis across the per-dataset findings
   — top recommendations in priority order. If there is only one relevant dataset,
   you may leave "overall" empty.
7. ALL human-readable text MUST be written in {lang}. Keep dataset names, product
   codes, defect type labels, and numeric values as-is.
8. Produce valid JSON structure internally for the AskAiHistory row.
```

→ 출력 schema:
```json
{
  "overall": "2-3 sentence {lang} overall recommendation.",
  "perDataset": [
    { "datasetName": "<verbatim>", "answer": "{lang} answer with concrete numbers." }
  ]
}
```

`lang` 치환 규칙: `request.language` 값을 그대로 (English / Korean / Vietnamese).

---

## 4. DB 커밋 (Python 한 번에 실행)

```python
import json, sqlite3, datetime

DB = r"D:\000. MyWorks\000. 일일업무\04. DB\process-review.db"
REQ = json.load(open("tmp/ask_request.json", encoding="utf-8"))

question = REQ["question"]
pt_filter = REQ.get("productTypeFilter", "") or ""
overall = "..."                          # 에이전트가 reasoning 후 하드코딩
per_dataset = [                          # 에이전트가 reasoning 후 하드코딩
    {"datasetName": "...", "answer": "..."},
    ...
]

now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.%fZ")
per_json = json.dumps(per_dataset, ensure_ascii=False)

conn = sqlite3.connect(DB)
conn.execute("PRAGMA busy_timeout = 30000")
cur = conn.cursor()
cur.execute("""
    INSERT INTO AskAiHistory (Question, ProductTypeFilter, Overall, PerDatasetJson, CreatedAt)
    VALUES (?, ?, ?, ?, ?)
""", (question, pt_filter, overall, per_json, now))
new_id = cur.lastrowid
conn.commit()
conn.close()

print(f"=== ASK DONE === history_id={new_id}")
```

트랜잭션 실패 시 `conn.rollback()` 후 에러 메시지 print 하고 종료. **부분 저장 금지.**

---

## 5. 터미널 출력 포맷

```
=== ASK AI (CLI) ===
Question: {question}
Filter:   ProductType={pt or 'ALL'}
Language: {language}
Datasets in scope: {N}

── Overall ──────────────────────────────────────────────
{overall}

── Per-dataset ──────────────────────────────────────────
[1] {datasetName1}
    {answer1}

[2] {datasetName2}
    {answer2}

=== ASK DONE === history_id={id}
Open "Ask AI" in app → History tab → latest row to reload.
```

터미널 한글 깨짐 방지:
```python
import sys; sys.stdout.reconfigure(encoding="utf-8")
```
맨 위에 반드시.

---

## 6. 실패 처리

| 상황 | 동작 |
|---|---|
| `tmp/ask_request.json` 없음 | "No request file" print, exit |
| `question` 비어있음 | "Empty question" print, exit |
| context block 0개 (필터에 해당 dataset 없음) | overall="해당 filter 에 등록된 리포트가 없습니다.", perDataset=[] 로 그대로 commit — 앱 동작과 동일 |
| DB 쓰기 실패 (busy 등) | rollback + stacktrace print, exit (history 미기록) |
| 에이전트가 관련 dataset 못 찾음 | perDataset=[], overall={lang}로 "No relevant data found" 계열 메시지 (rule #2) |

---

## 7. 청소 (runbook 기본 방침)

성공 여부와 무관하게 종료 직전:
- `tmp/ask_request.json` 삭제
- `_tmp_ask_commit.py` 등 임시 파이썬 스크립트 삭제
- `_tmp_*.py`, `_ask_work/` 등 잔재 삭제

---

## 8. Known gotchas

- [ ] DB 경로 — **항상 `workhost-settings.json`** 우선, default 폴백 금지
- [ ] `RawReports.BatchExcluded=0` 필터 (Ask AI UI 와 동일한 scope)
- [ ] `Tags` 컬럼은 JSON 문자열이므로 `json.loads` 필요
- [ ] `datasetName` 은 context 에 찍힌 문자열 **verbatim** 으로 (앞뒤 공백 유지)
- [ ] Answer 은 **전부 `language` 값 언어로만** 작성 (dataset 이름/숫자는 원형)
- [ ] perDataset 은 **유관한 것만** — irrelevant dataset 포함 금지
- [ ] Python UTF-8 reconfigure 빠뜨리면 cmd 에서 `UnicodeEncodeError`
- [ ] `AskAiHistory.CreatedAt` 은 UTC ISO-8601 문자열 (앱 포맷과 호환)
- [ ] Context 가 너무 크면 (dataset 수십 개) 에이전트 토큰 부담 → 현재 앱도 동일 문제. 필요 시 `productTypeFilter` 활용을 사용자에게 권유.

---

## 9. 에이전트 호출 템플릿

> **"CLI_ASK_AI.md 읽고 실행해"**

### 동작 순서
1. §1.1 `tmp/ask_request.json` 읽고 검증
2. §1.2 DB 경로 획득
3. §2 context 빌드 (`BuildDatasetsContext` 로직 SQL + Python 으로 재현)
4. §3 프롬프트 규칙 적용 → overall + perDataset 산출 (에이전트 직접 reasoning)
5. §4 `_tmp_ask_commit.py` 생성 → 실행 → `AskAiHistory` INSERT
6. §5 포맷으로 터미널 출력
7. §7 청소

---

*v1 (2026-04-23) — JinoSupporter.Web 의 Ask AI 기능을 CLI 구독 경로로 대체.
ClaudeService.AskAiAsync (line 1322-1398) + BuildDatasetsContext (DataInferenceAskPage.razor:442-487) 의
정확한 1:1 재현이 목적. 코드 변경 시 이 문서의 line 레퍼런스 재검토.*
