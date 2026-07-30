# CLI Ask AI — Runbook for future agents

## Current v2 policy (2026-05-13)

Use the current `AI_PROMPTS/data-inference/ai-excel-proc.md` schema first. Read `tmp/ask_request.json`,
resolve the production DB path as `D:\000. MyWorks\002. DB\process-review.db`,
and answer from:

`AiDocuments`, `AiTestConditions`, `AiResults`, `AiNgBreakdowns`,
`AiConclusions`, `AiTroubleshootingHints`, `AiExtractionLogs`, and their
translation tables. Treat `AiDocuments.RawJson` as the raw AI-analysis payload
for each review result and use it as evidence when the normalized child tables
omit detail. Use legacy `DatasetSummary` / `NormalizedMeasurements` only as
fallback context.

## Current v4 policy (2026-06-04)

Use the current INPUT DATA / RESULT analysis basis when answering Ask AI:

- Prefer generated_report_markdown and AiDocuments.RawJson that include
  reportReviewMatrix, visualization decision, and review domain evidence.
- For every contributing dataset, identify what was reviewed, why it was
  reviewed, what result was visible, and whether the evidence is
  `공정 불량 검토`, `기능 불량 검토`, or `공정+기능 연계 검토`.
- The overall answer must be a complete standalone HTML report when relevant
  reports exist, not a Markdown checklist table. The HTML layout is adaptive:
  first decide the analysis direction, then choose the most useful comparison
  and visualization for that direction.
- If the question asks what to do next, prioritize actions by evidence strength:
  same-event Normal comparison, repeated result, sample/spec evidence, then
  weak ranking-only evidence.
- For any NG-rate/defect-rate/yield/PPM comparison, answer with the numeric
  evidence and denominator. Codex CLI chooses whether bars, matrix, heatmap, or
  compact table best communicates the comparison. Do not claim improvement/
  worsening when Normal is missing.
- For continuous raw measurements such as tension, gauss, height, impedance,
  DCR, SPL, or THD, use a scatter or strip plot when raw points are visible and
  it improves readability. Show n per condition. If raw points are unavailable, write
  `raw distribution unavailable` and compare only visible avg/max/min values.
- Visible wording must use `Normal`, `Normal 대비`, `Normal 값`, or
  `Normal 미확인`. Do not display `Local Control`, `Control`, `Baseline`, or
  `대조군` as visible comparison labels.
- Speed: build a shortlist from indexed fields and generated reports first,
  then inspect RawJson only for shortlisted datasets. Avoid DB-wide deep
  reanalysis when the question narrows model/defect/process.

For NG-rate questions, do not judge by absolute NG-rate ranking when a
same-event Normal exists. Pair each Test/After/New/changed row against the
Normal row from the same source sheet/table and the same carried-forward
Date/Model/Line/measurement type. Source labels such as Baseline, Control,
Reference, Before, Old, or OK may be mapped internally to Normal, but visible
answers must say Normal.
Excel merged cells may display blanks in continuation rows; treat blank
Date/Model/Type/Line cells below a visible value as carrying that value forward.
If stored input came from Data Input, merged cells may already be expanded and
prefixed with metadata such as `{merged=A1:A4}` or `〔merged=A1:A4〕`; strip that
metadata and use the following value as the actual cell value. Percentage-only
subrows are not standalone result rows; use them only as rate/breakdown evidence
for the preceding real count row.

Respect the current `AI_PROMPTS/data-inference/ai-excel-proc.md` report types:
`normal_comparison`, `ng_without_baseline`, `before_after_dimension`,
`measurement_spec`, `defect_root_cause`, `lot_supplier_mold_comparison`,
`process_condition_change`, `reliability_spec`, `doe_matrix`,
`image_dependent`, `mixed`. Shape the answer to the type: comparison/process
reports need process-by-process comparison, measurement reports need
spec/min/max/avg or pass/fail, and cause-analysis reports need phenomenon,
checked action, result, and remaining risk.

Compute relative change as:

```text
(test_ng_rate / normal_ng_rate - 1) * 100
```

Positive means worse; negative means improved. If no same-event Normal exists,
answer as `ng_without_baseline`: rank actual NG rates, defect mix, source sheet,
and sample size, but do not say improved/worsened.

When saving the answer, insert exactly one `AskAiHistory` row with the requested
language, product type filter, and `TranslationsJson` containing Korean,
English, and Vietnamese renderings of the same final result.

> 목적: `JinoSupporter.Web` 의 **Ask AI** 기능 (`DataInferenceAskPage.razor`) 과
> 동일한 결과를 **서버 기동 없이, Anthropic API 호출 없이** CLI 로 만들어
> `AskAiHistory` 테이블에 저장하고 터미널에도 답을 출력한다.
>
> **v1 (2026-04-23)** — 구독(MAX 5x) 경로 전용. `NormalizeFromImages` 류와는
> 달리 vision 불필요 — 텍스트 context 합성 + reasoning 만 수행.

---

## 0. TL;DR — 에이전트가 해야 할 일

1. **요청 읽기**: `tmp/ask_request.json` 의 `question`, `language`, `productTypeFilter` 확보
   - 앱 launcher가 `tmp/ask_runs/<runId>/ask_request.json` 같은 run-scoped 경로를 지정하면 그 경로가 우선이다.
2. **DB 경로 확인**: `workhost-settings.json` 의 `DataInference.DatabasePath` (기본 경로 폴백 금지 — AI_PROMPTS/data-inference/ai-excel-proc.md §1.1 참조)
3. **Context 수집**: `FilteredReports` 에 해당하는 dataset 마다 summary + measurements 요약 블록 합성 (§2)
4. **Reasoning**: §3 프롬프트 규칙 그대로 적용 — HTML overall + per-dataset 답변 산출
5. **번역 + 커밋 + 출력**: 최종 결과를 Korean/English/Vietnamese 로 번역 후 `AskAiHistory` 에 single-row INSERT (`last_insert_rowid()` 회수), 터미널에 포맷해서 출력
6. **청소**: run-scoped 실행이면 해당 `tmp/ask_runs/<runId>/` 안의 현재 run 파일만 삭제한다. 다른 run 폴더나 공유 `tmp/ask_request.json`는 건드리지 않는다.

**핵심 원칙 (AI_PROMPTS/data-inference/ai-excel-proc.md 와 동일)**: 에이전트가 직접 reasoning 수행 — Anthropic API 호출 없음. Python 은 DB IO 만 담당.

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
`AI_PROMPTS/data-inference/ai-excel-proc.md §1.1` 과 동일. `workhost-settings.json` 의 `DataInference.DatabasePath` 를 우선. 현재 경로: `D:\000. MyWorks\002. DB\process-review.db`.

### 1.3 MicroSpeaker evidence pack
Newer Ask AI launches may include a deterministic evidence pack directly in
`tmp/ask_request.json`:

```json
{
  "microSpeakerEvidence": {
    "searchTerms": ["<question factor and outcome terms>"],
    "requiredTermGroups": [["<factor axis terms>"], ["<outcome axis terms>"]],
    "questionAnalysis": {
      "factorAxisLabel": "<factor/condition from the user's question>",
      "outcomeAxisLabel": "<result/defect from the user's question>",
      "factorTerms": [],
      "outcomeTerms": [],
      "suggestedReviewSections": []
    },
    "microSpeaker": {
      "modelCoverage": [],
      "verifiedReviewCases": [
        {
          "sourceModels": "<model names from MicroSpeaker files table>",
          "sourceCategories": "<detected source categories>"
        }
      ],
      "termHits": [],
      "pairAggregates": [],
      "pairConditionAggregates": [],
      "pairRows": [],
      "metricRows": [],
      "measurementRows": []
    },
    "jino": {
      "termHits": [],
      "documentRows": [],
      "resultRows": []
    }
  }
}
```

Use this pack before broad DB scanning. First inspect
`microSpeaker.verifiedReviewCases`; these are AI-verified and approved for Ask
AI, so use matching cases before raw extracted rows. Prefer
`microSpeaker.pairConditionAggregates` over individual `pairRows` when the same
file/table/factor/condition has repeated daily/lot rows. Rows with
`matchesAllRequiredTerms=true` matched every required question term group and
are the strongest evidence. If a row has `strictFallbackUsed=true`, it came from
broader OR matching because no strict row existed for that evidence class; state
that limitation instead of over-claiming. Use `termHits`, row counts, and
`pairAggregates` as coverage evidence.

Model boundary rule: keep model/product type separate. `sourceModels` on
verified ReviewCases, `models` on MicroSpeaker rows, and `productType` on Jino
rows are grouping boundaries. Do not sum, average, or draw one conclusion across
different models unless the user explicitly asks for a cross-model summary. If
the user has not selected a model and evidence spans multiple models, show
model-separated sections or label mixed-model rows as fallback context.
Coverage rule: `microSpeaker.modelCoverage` is the checklist of models touched
by the question. If it has multiple rows, include a visible model coverage
section or model-separated sections. Do not stop with one verified model while
other models have fallback rows or candidate files. Label each model as
`verified`, `fallback rows`, or `candidate files only`.

MicroSpeaker linkage guardrails:
- `NG` alone is generic result wording and must not count as proof of a specific
  outcome category. Match the specific outcome axis from `questionAnalysis`,
  such as function/SPL/THD for a function question, or the corresponding defect/
  measurement label for another question.
- Do not treat isolated short tokens as a combined part/process label. Prefer
  the exact phrase variants from `factorTerms` and `requiredTermGroups`.
- A factor-outcome linkage conclusion requires same-source evidence: the same
  dataset/file/table or a row that explicitly contains both the factor axis and
  outcome axis. If this is missing, answer "direct linkage evidence not found"
  and show the separate factor/outcome evidence instead of building a causal
  chart.
- Never use a 120-row cap as the review count. It is only a UI/evidence-pack
  limit. Report actual selected evidence counts and the limits clearly.
- If `pairConditionAggregates` contains the requested comparison, use its
  aggregate `testInput/testNg/testRatePercent` instead of one daily pair row.
  `aggregationMethod=total_row` means the source Total row was used;
  `aggregationMethod=summed_rows` means repeated same-condition rows were
  summed because no Total row was available.

MicroSpeaker answer layout and calculation rules:
- Before writing HTML, decide the analysis direction from the question and the
  strongest evidence. Use one primary direction and only add secondary
  directions when they materially help:
  `Normal/Test comparison`, `condition impact`, `defect ranking`,
  `process-function linkage`, `measurement/spec review`, `next-action review`,
  or `data gap review`.
- After the direction is chosen, Codex CLI should choose the comparison and
  visualization by itself:
  - rate comparisons: bars, matrix, heatmap, or compact table plus numeric
    evidence;
  - repeated same-condition rows: aggregate/Total comparison first;
  - many factors versus one outcome: grouped table, matrix, or heatmap;
  - continuous measurements: scatter or strip plot when raw points exist;
  - no Normal row: ranking view with a clear `Normal missing` limit;
  - weak linkage: separate evidence blocks plus missing-data note.
- The HTML layout is not fixed. Do not reuse one template for every question.
  Choose the report structure from the user's question and the available
  evidence. Cards, compact tables, grouped sections, matrices, charts, and
  short notes are all allowed when they make the evidence easier to read.
- Keep only the evidence contract fixed: the answer must show what was compared,
  what Normal/Test values were used, the denominator, the NG/result rate, the
  source link, the aggregation basis, and the limit. The visual arrangement is
  flexible.
- Start with the user's answer and the strongest evidence. A summary table is
  optional, not mandatory. Do not start with a generic category table when a
  direct comparison, matrix, or grouped evidence view answers the question more
  clearly.
- Arrange evidence by data similarity, not by retrieval order. Classify rows
  internally as `공정 불량 검토`, `기능 불량 검토`, `공정+기능 연계 검토`,
  `측정/Spec 검토`, or `데이터 한계`, but do not force those buckets to
  appear as a fixed column or fixed section. Use only the groups that are
  relevant to the question.
- Group similar evidence by same source file/table, reviewed factor, compared
  outcome metric, Normal condition, and Test condition. If the purpose differs
  across rows, split or label the group so the reader sees the purpose once at
  group/section level instead of repeated in every row.
- For each numeric comparison, include the comparison context somewhere visible:
  reviewed factor/item, compared outcome metric, original file link, Normal
  state/value for that reviewed factor, Normal Input/NG, Normal NG rate, Test
  state/value for that reviewed factor, Test Input/NG, Test NG rate, relative
  change, judgement, and limit. These may be columns, card fields, labels, or
  section metadata; omit repeated columns when the same value is already stated
  in the section header.
- Do not use ambiguous headers such as `Normal condition` and `Test condition`
  by themselves. The visible header or each cell must make clear what the
  condition belongs to, for example `Normal state of reviewed factor` and
  `Test state of reviewed factor`.
- Keep table cells compact. Use only necessary labels and numbers, not full
  sentences. Target 1-5 words per non-numeric cell. Examples: `press method`,
  `Function NG%`, `Normal`, `Changed`, `960/26`, `2.708%`, `worse +151.3%`,
  `합산 불가`. Put explanations below the table, not inside table cells.
- Purpose is a grouping aid, not a mandatory repeated table cell. Put exact
  purpose at the section/group level when possible, for example
  `<factor> versus <outcome metric>` or `Normal/Test NG% check`. Avoid a
  different long purpose sentence in every row.
- If a row has `originalFileUrl`, do not display the long file name in the
  table. Show a hyperlink only: `<a href="{originalFileUrl}" target="_blank"
  rel="noopener">원본 파일 보기</a>`. If no URL exists, show the shortest source
  label available.
- `reviewed factor/item` must be the thing being compared, such as jig
  condition, press method, mold condition, plasma condition, lot/supplier/line,
  material, measurement item, or inspection/retest condition. Do not put only
  the problem/outcome name there. If the source row does not reveal the reviewed
  factor, write `reviewed factor unclear` and do not use that row as primary
  evidence.
- In each Normal/Test state cell, include both the state value and the reviewed
  factor when needed, e.g. `press method = normal dry UC press` versus
  `press method = changed UC press`. Avoid bare values like `Normal` or
  `Total | Test New machine VP/CD` when the factor context is not visible.
- Defect rate is always `NG count / Input count * 100`. MicroSpeaker DB rates
  are usually decimal values; convert 0.00661 to 0.661% for display.
- Do not label `3420/9 vs 3327/22` as Count. Label it as `Input/NG`.
- Do not aggregate unrelated rows. Sum NG counts only when source dataset/file,
  source table/sheet, metric, and date/lot/line/model basis are the same or
  clearly compatible. Otherwise write `합산 불가` and present separate rows.
- Never combine process NG and function NG into one denominator or one total.
  They answer different review categories.
- Charts are optional and secondary. If used, place them after the detailed
  evidence table; do not let charts replace the count/rate explanation.

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
   to the answer. In "datasetName", copy only the actual name after "Dataset:";
   do not include "Dataset:", bracket numbers, bullets, or prefixes.
4. In each per-dataset "answer": avoid long prose. Use a compact Markdown table
   or short bullet list that shows only concrete evidence from that dataset:
   what was reviewed, source/result count, key value, and judgement.
5. Do NOT include datasets that are irrelevant to the question.
6. In "overall": when relevant reports exist, return a complete standalone HTML
   document string, not a Markdown checklist table. Start with
   `<!doctype html><html>`, include CSS/JS/SVG inline, and do not reference
   external assets, fonts, CDNs, or network URLs. The HTML is the final browser
   report that JinoSupporter will display directly.
   First decide the analysis direction from the question and evidence, then
   choose the most useful comparison/visualization for that direction. Do not
   use a fixed HTML template. A dense summary table, candidate action table,
   bar chart, scatter plot, matrix, heatmap, or grouped evidence table is used
   only when it fits the chosen direction.
6a. Preserve the report domain in the reasoning: process review, function
    review, linked process-function review, measurement/spec review, or data
    gap. Display only the domains that matter to the chosen direction. If the
    domain is unclear, write `미확인` and explain the missing evidence briefly.
6b. For rate comparisons, include the Normal value, target/test value, relative
    change, and numeric denominator. Codex CLI chooses whether to show bars,
    matrix, heatmap, or compact table based on readability.
6c. For continuous raw measurements such as tension, gauss, height, impedance,
    DCR, SPL, or THD, use scatter/strip plot when raw points are visible. If
    raw points are unavailable and only avg/max/min are visible, write
    `raw distribution unavailable` and compare avg-to-avg, max-to-max, and
    min-to-min only.
6d. Do not mix process NG and function NG into one denominator or one chart.
    If both matter, show them as separate evidence blocks or as a linked review
    with separate denominators.
6e. Every displayed chart/table/comparison must state the checked item, visible
    result, related dataset/source, interpretation, and limits.
7. ALL human-readable text MUST be written in {lang}. Keep dataset names, product
   codes, defect type labels, and numeric values as-is.
8. Produce valid JSON structure internally for the AskAiHistory row.
9. Also create `translations` with keys `ko`, `en`, `vi`. Each value must have
   the same `overall` + `perDataset` schema. Preserve HTML tags, CSS,
   JavaScript, SVG geometry, chart data arrays, and datasetName values unchanged;
   translate visible human-readable text only. Do not convert HTML to Markdown.
```

→ 출력 schema:
```json
{
  "overall": "Complete standalone HTML document when relevant reports exist; short no-data notice only when no relevant data exists.",
  "perDataset": [
    { "datasetName": "<actual Dataset name only>", "answer": "{lang} answer with concrete numbers." }
  ]
}
```

`lang` 치환 규칙: `request.language` 값을 그대로 (English / Korean / Vietnamese).

---

## 4. DB 커밋 (payload file + short Python script)

Windows command-length rule:
- Do **not** run the full answer through an inline PowerShell command such as
  `@' ... '@ | python -` or `python -c`. Standalone HTML reports and translation
  JSON can exceed the Windows command-line limit and fail with
  `InvalidFilename` / "파일 이름이나 확장명이 너무 깁니다."
- Do **not** embed `overall`, `perDataset`, or `translations` as giant string
  literals inside the shell command or commit script.
- First write the final answer to `tmp/ask_result_payload.json` as a real file.
  Then run a short `_tmp_ask_commit.py` that only reads
  `tmp/ask_request.json` and `tmp/ask_result_payload.json`.
- If the file-edit tool is unavailable, write the payload in small chunks. Keep
  each shell command short; never place full HTML in a command argument or
  PowerShell here-string.

Payload file schema:

```json
{
  "overall": "<complete standalone HTML or no-data notice>",
  "perDataset": [
    { "datasetName": "<actual Dataset name only>", "answer": "<answer>" }
  ],
  "translations": {
    "ko": { "overall": "<same report translated>", "perDataset": [] },
    "en": { "overall": "<same report translated>", "perDataset": [] },
    "vi": { "overall": "<same report translated>", "perDataset": [] }
  }
}
```

```python
import json, sqlite3, datetime, sys

sys.stdout.reconfigure(encoding="utf-8")

DB = r"D:\000. MyWorks\002. DB\process-review.db"
REQ = json.load(open("tmp/ask_request.json", encoding="utf-8"))
PAYLOAD = json.load(open("tmp/ask_result_payload.json", encoding="utf-8"))

question = REQ["question"]
pt_filter = REQ.get("productTypeFilter", "") or ""
lang = REQ.get("language", "") or ""
overall = PAYLOAD.get("overall", "")
per_dataset = PAYLOAD.get("perDataset", [])
translations = PAYLOAD.get("translations", {})

now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.%fZ")
per_json = json.dumps(per_dataset, ensure_ascii=False)
translations_json = json.dumps(translations, ensure_ascii=False)

conn = sqlite3.connect(DB)
conn.execute("PRAGMA busy_timeout = 30000")
cur = conn.cursor()
cols = {row[1] for row in cur.execute("PRAGMA table_info(AskAiHistory)")}
if "TranslationsJson" not in cols:
    cur.execute("ALTER TABLE AskAiHistory ADD COLUMN TranslationsJson TEXT NOT NULL DEFAULT '{}'")
cur.execute("""
    INSERT INTO AskAiHistory (Question, ProductTypeFilter, Overall, PerDatasetJson, TranslationsJson, CreatedAt)
    VALUES (?, ?, ?, ?, ?, ?)
""", (question, pt_filter, overall, per_json, translations_json, now))
new_id = cur.lastrowid
conn.commit()
conn.close()

print("=== ASK AI (CLI) ===")
print(f"Question: {question}")
print(f"Filter:   ProductType={pt_filter or 'ALL'}")
print(f"Language: {lang}")
print(f"Datasets in scope: {len(per_dataset)}")
print("\n-- Overall ------------------------------------------------")
print(overall[:2200] + ("\n...[HTML truncated in terminal]..." if len(overall) > 2200 else ""))
print("\n-- Per-dataset --------------------------------------------")
for i, d in enumerate(per_dataset, 1):
    print(f"[{i}] {d.get('datasetName', '')}")
    print("    " + str(d.get("answer", "")).replace("\n", "\n    "))
    print()
print(f"=== ASK DONE === history_id={new_id}")
print('Open "Ask AI" in app -> History tab -> latest row to reload.')
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
- run-scoped 경로가 지정된 경우 `tmp/ask_runs/<runId>/` 안의 현재 run 파일만 삭제
- 공유 `tmp/ask_request.json`, 다른 `tmp/ask_runs/*` 폴더, 다른 run의 payload/commit script 삭제 금지
- run-scoped 경로가 없는 수동 실행일 때만 `tmp/ask_request.json`, `tmp/ask_result_payload.json`, 현재 실행에서 만든 commit script를 정리
- `_tmp_*.py`, `_ask_work/` 같은 광범위 삭제는 금지. 현재 run에서 직접 만든 파일명만 삭제

---

## 8. Known gotchas

- [ ] DB 경로 — **항상 `workhost-settings.json`** 우선, default 폴백 금지
- [ ] `RawReports.BatchExcluded=0` 필터 (Ask AI UI 와 동일한 scope)
- [ ] `Tags` 컬럼은 JSON 문자열이므로 `json.loads` 필요
- [ ] `datasetName` 은 context 에 찍힌 문자열 **verbatim** 으로 (앞뒤 공백 유지)
- [ ] Answer 은 **전부 `language` 값 언어로만** 작성 (dataset 이름/숫자는 원형)
- [ ] perDataset 은 **유관한 것만** — irrelevant dataset 포함 금지
- [ ] `TranslationsJson` 은 `ko`/`en`/`vi` 세 키를 모두 포함
- [ ] Python UTF-8 reconfigure 빠뜨리면 cmd 에서 `UnicodeEncodeError`
- [ ] `AskAiHistory.CreatedAt` 은 UTC ISO-8601 문자열 (앱 포맷과 호환)
- [ ] Windows command length — full HTML/JSON must go through
      `tmp/ask_result_payload.json`, never `@'...'@ | python -`
- [ ] Context 가 너무 크면 (dataset 수십 개) 에이전트 토큰 부담 → 현재 앱도 동일 문제. 필요 시 `productTypeFilter` 활용을 사용자에게 권유.

---

## 9. 에이전트 호출 템플릿

> **"AI_PROMPTS/data-inference/cli-ask-ai.md 읽고 실행해"**

### 동작 순서
1. §1.1 `tmp/ask_request.json` 또는 launcher가 지정한 run-scoped request path 읽고 검증
2. §1.2 DB 경로 획득
3. §2 context 빌드 (`BuildDatasetsContext` 로직 SQL + Python 으로 재현)
4. §3 프롬프트 규칙 적용 → overall + perDataset 산출 (에이전트 직접 reasoning)
5. 최종 결과를 Korean/English/Vietnamese 로 번역 → `translations` 생성
6. §4 result payload + commit script 생성 → 실행 → `AskAiHistory` INSERT
7. §5 포맷으로 터미널 출력
8. §7 청소

---

*v1 (2026-04-23) — JinoSupporter.Web 의 Ask AI 기능을 CLI 구독 경로로 대체.
ClaudeService.AskAiAsync (line 1322-1398) + BuildDatasetsContext (DataInferenceAskPage.razor:442-487) 의
정확한 1:1 재현이 목적. 코드 변경 시 이 문서의 line 레퍼런스 재검토.*

