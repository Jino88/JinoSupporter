# InferenceDataAIService

## Current contextual AI question path

The current production path freezes `table-first-builder-v8` and
`table-first-analysis-prompt-v4`. Completed request, analysis, and projection
JSON files are resumable and are not regenerated when search wording or the
WPF presentation changes.

Workbooks containing deterministic SPL, THD, or IMP frequency-response RAW
data are excluded as a whole from table-first semantic AI and recipe learning.
Their lossless source capture is retained for provenance, but none of the
workbook's tables, text blocks, or numeric values are sent to semantic AI.

After a complete table-first batch, build the searchable history database:

```powershell
python inference_data_ai_cli.py table-first-history-index `
  --batch-dir outputs/table-first/full-989-v8-prompt-v4 `
  --db outputs/table-first-history/history.sqlite
```

Verify that broad candidate retrieval still finds the representative primary
workbooks and preserves the review gate:

```powershell
python inference_data_ai_cli.py table-first-history-acceptance `
  --db outputs/table-first-history/history.sqlite `
  --manifest pilot/representative-pilot-v1.json `
  --out outputs/table-first-history/history-acceptance-report.json
```

`table-first-history-query` remains available as a deterministic retrieval
diagnostic. It does not understand the complete question relation and is not
the WPF user-answer path.

Ask a contextual question with query-time AI relevance judgment:

```powershell
python inference_data_ai_cli.py table-first-contextual-query `
  --db outputs/table-first-history/history.sqlite `
  --question "VP CD 조립에 따른 Hearing 불량률 추이" `
  --candidate-limit 30 `
  --detail-candidate-limit 18
```

Keyword scoring is used only to collect candidates. The query-time AI separates
the requested subject, conditions, metrics, comparison, and time axis, rejects
candidates that merely share words, and can use only registered source-cell
facts. A numeric trend requires at least two comparable dated observations;
otherwise the answer explicitly reports partial or insufficient evidence.
The WPF renders the direct answer, interpreted intent, findings, comparable
observations, limitations, and at most ten core `TF-EVD-*` citations instead of
dumping every retrieved source.

Table-first projections remain review-gated. Observed source values may be
described, but the contextual path never creates an approved effect or causal
conclusion. The WPF uses `outputs/table-first-history/history.sqlite`
automatically when present and falls back to the canonical review DB otherwise.

The full-corpus static HTML report is intentionally not part of this path.
The user-facing deliverable is the history question/answer flow and its source
drill-down.

Standalone CLI workspace for turning mixed Excel report files into source-backed SQLite/JSON outputs before any JinoSupporter Web integration.

## Default Input

`D:\000. MyWorks\test\result\InputDataFinish`

## Output Location

All generated outputs are kept under this folder:

`D:\000. MyWorks\005. Program\Repository\JinoSupporter\InferenceDataAIService`

## Source-of-Truth Storage

For hundreds of mixed Excel layouts, use the **Capture v2 tables inside the universal-grid SQLite DB** as the raw tabular source of truth.

- Capture v2 reads DRM-free `.xlsx` files through Open XML/openpyxl and stores SHA-256 revisions, sparse source coordinates, raw/formula/cached/display values, number formats, styles, dimensions, hidden state, and merges.
- It is resumable, keeps per-file terminal states, and bridges each Capture v2 revision to the canonical source revision.
- Embedded images are intentionally outside scope and are neither extracted nor analyzed.
- The older COM fixed-grid tables remain a compatibility/evidence adapter for already imported records; new historical-corpus work uses Capture v2.
- It does not infer a business table shape, so different report layouts can coexist in the same dataset.
- It also supports a reusable **analysis layer**: report purpose/type, review types, Test/Control (or other) cohorts, metric values, numerator/denominator, ppm/rate differences, conclusions, and exact Excel evidence ranges.
- Analysis is derived data, not a replacement for the raw grid. Every stored comparison and conclusion must point back to a verified source sheet/range, and it becomes stale when its source workbook is re-imported.
- The quick-index DB remains an optional sidecar for candidate metrics, statistics, and comparison hints. It is not a replacement for the fixed grid because it stores only non-empty cells and no merge map.

## Commands

Show help:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py --help
```

Create the universal grid DB schema:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py init-db --dataset InputDataFinish
```

Run the fast existing candidate indexer into this service folder:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py quick-index --dataset InputDataFinish
```

Run only a small smoke index:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py quick-index --dataset InputDataFinishSmoke --limit 5 --force
```

Run Excel COM extraction and import to the universal grid DB:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py com-index --dataset InputDataFinish --limit 25 --covered-cell-mode blank --verify-after-import
```

Run the same command again to continue with the next new/changed batch. Existing workbooks with an unchanged source fingerprint and raw JSON artifact are skipped automatically. Do not use `--sparse` for the audit/source DB because it omits ordinary blank coordinates.

Re-extract every selected workbook only when intentionally needed:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py com-index --dataset InputDataFinish --force --covered-cell-mode blank --verify-after-import
```

Verify every stored universal workbook against its raw COM JSON artifact:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py verify-universal-db --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --dataset InputDataFinish
```

Capture the 30-workbook Open XML pilot and verify stored counts, canonical
bridges, and current source hashes:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py openxml-index --pilot-manifest .\JinoSupporter\InferenceDataAIService\pilot\representative-pilot-v1.json --dataset InputDataFinish --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py capture-v2-verify --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --source-sha256
```

Import an evidence-linked reusable analysis summary from a layout-independent manifest:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-import --input .\JinoSupporter\InferenceDataAIService\outputs\analysis-manifests\BRS2015_G06_0003_analysis.json
```

Validate every stored analysis summary (source freshness, evidence ranges, ppm arithmetic, comparison deltas, and conclusion evidence):

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-verify --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite
```

Install the additive canonical Study/Comparison/Evidence schema and migrate
existing `analysis_*` records without using unstable SQLite row ids as public
data numbers:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py knowledge-migrate --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py knowledge-inspect --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite
```

The canonical layer uses stable `DATA-*`, `CMP-*`, `EFF-*`, and `EVD-*`
identifiers. Legacy comparisons are migrated conservatively as
`NEEDS_REVIEW` and are not aggregation-eligible until comparison validity,
confounding, evidence, and calculation checks explicitly pass.

Build lossless source chunks for one captured revision, run resumable parallel
read-only AI locator passes, and consolidate them into a non-self-approved
canonical Study draft:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py semantic-packets --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --revision-id 18 --out outputs\semantic-source-packets\pilot-revision-18.json
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py semantic-locate --packet outputs\semantic-source-packets\pilot-revision-18.json --workers 3 --out-dir outputs\semantic-locators\pilot-revision-18
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py semantic-draft --packet outputs\semantic-source-packets\pilot-revision-18.json --locator-dir outputs\semantic-locators\pilot-revision-18 --db outputs\universal-grid\InputDataFinish.sqlite --out outputs\semantic-study-drafts\pilot-revision-18.study-draft.json
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py study-import --input outputs\semantic-study-drafts\pilot-revision-18.study-draft.json --db outputs\universal-grid\InputDataFinish.sqlite
```

Every non-empty Capture v2 source cell belongs to exactly one primary chunk.
Chunk processing never silently truncates cells, and all locator results are
required before workbook consolidation. The factor/outcome domain is open-ended:
VP+CD and FUNCTION NG are acceptance examples, not extraction rules or a query
whitelist. AI drafts remain `NEEDS_REVIEW`, cannot calculate effects, cannot
claim causality, and cannot make comparisons aggregation-eligible.

Ask any domain-neutral relationship question, create a deterministic Korean
answer, validate that the answer has not changed, and open one exact EVD source
range:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py evidence-query --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --question "waiting 2 day가 FUNCTION NG와 관계가 있나?" --out outputs\evidence-packs\question.json
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py evidence-answer --pack outputs\evidence-packs\question.json --out-json outputs\evidence-answers\question.answer.json --out-markdown outputs\evidence-answers\question.answer.md
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py evidence-answer-validate --pack outputs\evidence-packs\question.json --answer outputs\evidence-answers\question.answer.json
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py evidence-detail --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --evidence-id EVD-61F6173DCB92 --out outputs\evidence-details\example.json
```

`evidence-answer` can also use `--db` plus `--question` directly. Numeric
relationship wording can only come from answer-eligible verified effects with
a direct current-revision EVD citation. Descriptive observations and excluded
records remain visible, but the answer never invents their difference or
relative change. Question outcome terms filter displayed descriptive details
on token boundaries; when no stored outcome matches, the complete descriptive
set remains visible instead of being silently discarded.

Incrementally ingest one completely new DRM-free `.xlsx` with a durable,
resumable journal:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py ingest-workbook --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --input "D:\review\new-review.xlsx" --artifact-root outputs\incremental-ingest
```

This runs Capture, semantic packet construction, batched locator, conservative
Study draft, source/numeric validation, idempotent import, and canonical
integrity verification. AI drafts remain `NEEDS_REVIEW`; terminal empty or
non-tabular workbooks are `EXCLUDED` without an AI call. Images remain outside
the contract.

Freeze or process the deterministic 989-workbook corpus inventory through the
same workflow:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py ingest-corpus --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --input "D:\000. MyWorks\test\result\InputDataFinish" --artifact-root outputs\corpus-ingest\full-989-v1 --inventory-only
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py ingest-corpus --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --input "D:\000. MyWorks\test\result\InputDataFinish" --source-manifest pilot\representative-pilot-v1.json --artifact-root outputs\corpus-ingest\full-989-v1 --offset 0 --limit 1
```

The corpus journal retains path, SHA-256, size, timestamps, attempts, terminal
status, failures, and source-change history. It accounts for every source even
when a selected slice is small. The SQLite writer defaults to one workbook at
a time.

Find exact-content duplicates and lexically related current Studies after
ingestion:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py related-studies --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --target DATA-EB5ABF306FA7 --limit 12 --out outputs\related-studies\example.json
```

This discovery ranking uses open-domain factor, outcome, context, and title
terms. It is explicitly not relationship or causal evidence. WPF shows the
same warning after a newly selected workbook is ingested.

Inspect the fail-closed human review queue, open one comparison with its paired
values and exact EVD ranges, and record an explicit decision:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py review-queue --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --out outputs\human-review\queue.json
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py review-detail --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --comparison-id CMP-EXAMPLE --out outputs\human-review\CMP-EXAMPLE.json
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py review-decide --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --comparison-id CMP-EXAMPLE --decision APPROVE --reviewer "reviewer-id" --reason "Checked the current source rows and pairing." --study-comparability VALID --study-confounding NONE --comparison-validity VALID --comparison-confounding NONE --matching-basis "same unit, period, and measurement method"
```

Approval is never automatic. It fails unless the source is the current
SHA-256 revision, the comparison and both paired observations have direct
VERIFIED evidence, the four explicit assessment states are `VALID/NONE`, and
the matching basis is nonempty. `REJECT`, `EXCLUDE`, and
`RETURN_TO_REVIEW` are also supported and disable aggregation.

Run all ten representative questions through the real query, deterministic
answer, answer validation, and exact primary-source coverage checks:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py golden-acceptance --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --manifest .\JinoSupporter\InferenceDataAIService\pilot\representative-pilot-v1.json --out-dir outputs\golden-acceptance\pilot-current
```

The report distinguishes `PENDING_INGEST` from an actual `RETRIEVAL_MISS`.
Natural-language required behaviors remain explicit manual-review items rather
than being silently marked correct by the runner.

### Current semantic pilot checkpoint (2026-07-18)

- Corpus coverage is 9/989 completed, 980 pending, and 0 failed. Completed
  records are 6 `NEEDS_REVIEW` and 3 `EXCLUDED`.
- P12
  (`014.MSU-20S15-07 Result test AWF cooling time 4s,8s,10s_clean.xlsx`)
  is stored as one current analysis with 2 Studies, 7 Arms, 24 Outcomes,
  70 Observations, 4 measurement series, 2,697 measurement points,
  3 review-gated Comparisons, and 0 Effects.
- Wide measurement matrices remain in SQLite. Evidence packs and answers carry
  only deterministic summaries and exact EVD ranges, never thousands of raw
  points. Axis and replicate counts use source coordinates so duplicate labels
  do not collapse Excel geometry.
- The actual frequency-response answer is written to
  `outputs/qa/p12-frequency-series.answer.json` and `.md`. It retrieves only
  the P12 Study, reports Spec/20000V/1800V/1600V descriptively, attaches
  12 exact citations, and returns `INSUFFICIENT_COMPARISON` with zero eligible
  Effects. A same-model XRAY-only source is excluded as irrelevant.
- Current golden acceptance is `BLOCKED_PENDING_INGEST`: 2 pass, 8 pending,
  0 fail, and 0 retrieval miss. Images analyzed: false.
- Current verification: 100 scoped Python tests pass, WPF builds with
  0 warnings/errors, SQLite integrity is `ok`, and canonical FK errors are 0.
  WPF was not launched.

Inspect or export one analysis for another dashboard/AI workflow:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-inspect --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --report-id 1
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-export --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --report-id 1
```

The manifest contract is intentionally layout-independent: each review can declare any number of cohorts, scoped metrics, pairwise comparisons, conclusions, and source ranges. This allows Normal/Test, supplier variants, before/after, and multi-sheet reports to reuse the same DB model.

The versioned 30-workbook coverage pilot and ten golden questions are in
`pilot/representative-pilot-v1.json`. They are coverage fixtures, not approved
meaning labels. Embedded images are not extracted or analyzed; files without
usable tabular evidence terminate explicitly as `NO_TABULAR_EVIDENCE`,
`DESCRIPTIVE_ONLY`, or an explained exclusion.

## Common Pipeline UI

Use the local Windows UI to drag Excel files or folders into a visible list,
then run the same shared CLI pipeline for the existing batch and later added
Excel files:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_ui.py
```

The packaged drag-and-drop GUI launcher is available as:

`InferenceDataAIServiceUI_DragDrop.exe`

Keep this EXE in the `InferenceDataAIService` folder beside
`inference_data_ai_cli.py`. The EXE packages the GUI itself, but intentionally
uses the existing local Python CLI and installed Microsoft Excel for COM
extraction. If Python is not found automatically, set
`INFERENCE_DATA_AI_PYTHON` to the full `python.exe` path before launching it.

The UI runs each listed source through `com-index` with fixed-grid verification
and, when selected, `quick-index` to create the candidate dashboard HTML. With
**AI draft + verified HTML** selected, it then invokes the locally logged-in
Codex CLI to produce a source-backed `universal-analysis-v1` draft, imports and
verifies it, and writes an analysis dashboard under
`outputs/analysis-rendered`. Drafts remain `NEEDS_REVIEW` unless their source
supports a verified conclusion. The quick-index HTML is still only a candidate
dashboard, not an approved conclusion.

The GUI requires the `codex` command to be installed and logged in. It uses the
original absolute paths without copying files. Unchanged workbooks with a
current analysis dashboard are skipped; changed source workbooks are marked
stale and analysed again.

The extractor always opens source workbooks read-only. Failed extractions are recorded per run and retain the last successful workbook grid instead of deleting it.

The WPF application now also exposes:

- `검토 DB 질문`: unrestricted free-text questions through the Python
  canonical evidence/answer contract, with EVD citation rows;
- `선택 근거 표 보기`: revision-safe Capture v2 cells rendered as an
  Excel-coordinate grid with blank positions, formulas, number formats, hidden
  dimensions, and merge spans;
- `Excel 정확한 범위 열기`: user-triggered read-only Excel COM navigation to
  the cited sheet and A1 range;
- `사람 검토 승인`: current comparison queue, paired values, direct EVD
  ranges, explicit comparability/confounding/validity assessment, and
  confirmed approve/reject/exclude/re-review decisions;
- `신규 Excel DB 적재`: the same durable `ingest-workbook` path, with
  `NEEDS_REVIEW`/`EXCLUDED` status and no image analysis.

WPF does not reimplement evidence eligibility or effect math in C#. It displays
the deterministic Python answer and exact EVD detail contract.

Search indexed files:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py search --db .\JinoSupporter\InferenceDataAIService\outputs\quick-index\InputDataFinish.sqlite --q "VP CD"
```

Build an AI ReviewCase packet for one indexed file:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py build-packet --db .\JinoSupporter\InferenceDataAIService\outputs\quick-index\InputDataFinish.sqlite --file-id 21
```

## Design

Use `GENERALIZATION_PLAN.md` as the implementation intent.

For the current target architecture, confirmed role boundaries, actual DB state,
session history, and the next implementation sequence, use
`WORKING_CONTEXT.md`.

## Current P13 semantic checkpoint (2026-07-18)

- Current P13 analysis: `ANALYSIS-B6586C6CDFB9`.
- Existing Capture v2, 48 complete chunks, and 48 locator results were reused;
  only the prompt-v6 Study draft ran again. Images were not analyzed.
- P13 contains 5 Studies, 27 Arms, 33 Outcomes, 197 scalar Observations, 27
  measurement series, 13,311 points, 24 review-gated Comparisons, and 0
  Effects. The ten MASK values are scalar observations, so no numeric source
  values were dropped.
- Measurement points are 10,962 `RAW` plus 2,349 `AGGREGATE`. AVG values stay
  source-backed but are excluded from default RAW descriptive statistics and
  replicate counts. All 2,349 AVG values equal their source RAW-row average.
- Percent scalars use displayed `%` units. All 20 P13 rate observations have
  exact numerator/denominator pairs; all 124 non-Input count observations are
  denominator-backed; Input uses denominator-only `sample_size`.
- Nine same-mold 180-versus-190 comparisons are now aligned. Fixed cure/dry,
  agent percentage, line, lot, and mold conditions remain explicit. Fourteen
  comparisons are confounded and ten unassessed; none is approved or
  aggregation-eligible.
- The approval planner rejects raw count subtraction and avoids duplicating a
  count-derived rate effect when the same aligned explicit rate outcome is
  available.
- Corpus accounting is 10 completed / 979 pending / 0 failed. Golden
  acceptance is 3 structural pass / 7 pending / 0 fail, with 6 of 15 required
  source appearances represented and 0 retrieval misses.
- Verification: 115 scoped Python tests pass, `verify-universal-db` is 16/16,
  SQLite/canonical integrity passes, and the WPF project builds with 0
  warnings and 0 errors. WPF was not launched.

## Current GQ06/P14 safety checkpoint (2026-07-18)

- Effectless/confounded comparisons now keep their factor values, arm
  conditions, matching basis, gate statuses, and direct EVD ranges in the
  deterministic answer.
- `CONFOUNDED_MULTI_FACTOR` requires at least two real recorded value
  differences. GQ06 lists the four stored differences for Function
  comparisons (`VP mold`, `2nd Cure`, `Test type`, and `1st Molding`) and
  calculates no temperature-only Effect.
- GQ06 now passes all three declarative required-behavior assertions. It has
  0 eligible Effects, 0 unrelated SPL/THD/IMP descriptive series, and 28
  direct citations rather than the former 91 broad citations.
- Prompt v8 and the DRAFT pre-import gate reject unsupported arm roles,
  inferred Before/After blocks, blank cells inside dense measurement ranges,
  duplicate empty replicate identities, blank REF sibling columns, and
  treating MIN/MAX/AVG as raw specimens.
- Any general draft validation error can now be repaired from the rejected
  JSON plus the exact validator message and focused source packet. Pure
  reference repairs remain unable to change unrelated content.
- Full Python regression is 199 tests passing. P14 is still in its v8
  fail-closed draft/audit cycle; it is not counted as completed until the
  source checklist passes. No application was launched and images remain out
  of scope.

## Current P14 completed checkpoint (2026-07-18)

This section supersedes the interim P14 status above. The earlier conclusion
that Before/After was invented is retracted for this workbook.

- P14
  `38. MSU-L20S1507 DOE Air preasure-air leak_clean.xlsx` is imported as
  `ANALYSIS-B35468F7B510`.
- Its header cells store `1..10`, while source-authored custom number formats
  display pressure, replicate, and Before/After identities. Capture retained
  `number_format`, but the rendered HTML/WPF path displayed the raw value.
  Import now derives formatted identities only from directly cited captured
  cells. This is a confirmed Excel-versus-renderer semantic-loss cause.
- The source-backed 18/100/200 kPa Before/After blocks produce nine
  review-gated comparisons. The value-empty 300 kPa After headers remain
  header-only. No comparison is approved or aggregation-eligible, and no
  Effect is calculated.
- Standalone `AGGREGATE` measurement series now require `AVERAGE`, explicit
  RAW source-series references, matching Study/Outcome, and exact same-axis
  arithmetic. All seven P14 AVG series pass.
- Final preservation: 5 Studies, 39 Arms, 9 Outcomes, 17 scalar
  Observations, 46 series (39 RAW + 7 AGGREGATE), 18,916 points
  (18,651 RAW + 265 AGGREGATE), 9 Comparisons, 0 Effects.
- Capture, packet, 71 locator results, and draft artifacts were reused; there
  was no recapture, relocation, image processing, or locator AI call. Failed
  interim drafts were not imported.
- Corpus status is 11 completed / 978 pending / 0 failed. Golden acceptance
  is 4 pass / 6 pending-ingest / 0 fail, with GQ05 represented.
- Full Python regression is 218 tests passing, universal DB verification is
  16/16, canonical/SQLite integrity passes, and WPF builds with 0 warnings
  and 0 errors. No application was launched.
- Existing rendered HTML is still not accepted as an Excel-faithful view.
  The database now preserves the missing format semantics, but renderer
  correction and visual acceptance remain separate work.

The next gate is P15, still one workbook at a time before corpus throughput is
increased.

## Current P15 completed checkpoint (2026-07-18)

- P15
  `1. BRS-161014 Report test DOE sub 2 manual_clean.xlsx` is imported as
  `ANALYSIS-8CAED455EB68`.
- Two sheets were separated into four source-coherent Studies: CM+B-PT
  Visual by bonding amount, Tension by bonding amount, CM+CP
  bonding/drying DOE, and CM+B-PT bonding-line DOE. `In Spec` remains a
  `TEST` arm and is not relabeled as a control.
- Final preservation is 4 Studies, 18 Arms, 21 Outcomes, 90 scalar
  Observations, 8 RAW measurement series, 64 points, 10 Comparisons, and 0
  Effects. The 64 points represent 32 physical specimen rows measured for
  both actual bonding amount and tension; they are not 64 specimens.
- All four bonding-amount series use `mg`; all four tension series use
  `kgf`. Every series has eight dense source points, all 64 stored values
  match Capture v2, and source-coordinate duplication within a series is 0.
- The four Visual rate formulas have no cached values. They are retained as
  formula-derived observations with `valueNumber=null`, exact
  numerator/denominator, and deterministic `ratePpm`. The 12 uncached
  Tension MAX/MIN/AVG cells remain formula-lineage evidence only; no
  aggregate numeric value was invented.
- CM+B-PT row 46 remains exactly Input 8, OK 0, Total NG 5, with three
  unclassified/unreconciled specimens. No complement, correction, or
  imputation was applied.
- Ten review-gated comparisons are retained: three adjacent Visual, three
  adjacent independent Tension, two CM+CP amount-only, and two amount-only
  comparisons inside the explicitly changed bonding-line group. The two
  tempting comparisons involving blank bonding-line cells were excluded
  because blank does not prove a shared baseline or unchanged line.
- Every Comparison is `NEEDS_REVIEW`/`UNASSESSED`, non-eligible for
  aggregation, and effect-free. Human review decisions remain 0.
- Capture revision, the complete 345-cell packet, nine locator results, and
  the prompt-v17 draft artifact were reused during import. Locator AI calls,
  draft AI calls, recapture, and image analysis were all 0 for the import
  run.
- AI output can no longer decide deterministic packet coverage: the runner
  restores `source.contentComplete` from the packet before strict validation,
  while the direct validator still rejects a coverage mismatch.
- Corpus status is 12 completed / 977 pending / 0 failed. Golden acceptance
  remains 4 pass / 6 pending-ingest / 0 fail; represented required primary
  sources increased to 7 of 15, with 0 retrieval misses.
- Full Python regression is 222 tests passing, universal DB verification is
  16/16, P15 numeric/evidence/SQL audits pass, and WPF builds with 0 warnings
  and 0 errors. No WPF, Excel, server, or desktop application was launched.

The next gate is P16 through the same one-workbook, image-free source audit.
Throughput must remain gated until the remaining representative pilots prove
the same completeness and comparison safety.

## Current P16 completed checkpoint (2026-07-18)

- P16 is
  `017.MSU-20S15-07 Result test sample waitting 2 day and check function_clean.xlsx`,
  SHA-256
  `97ee9e16690a7b7818313ce6d6c0644779878ade5513459102aef933e6257085`,
  imported as `ANALYSIS-F0756E286A00`.
- The obsolete reconciled analysis `ANALYSIS-0D3FE4FD3695` was superseded.
  It had two unlabeled `OTHER` arms, mixed counts and percentages, missing
  sample sizes, human percentages stored as raw Excel fractions, and no
  source-backed comparison.
- The current source-backed representation has 1 Study, 2 Arms, 22 Outcomes,
  44 scalar Observations, 1 Comparison, and 0 Effects. Waiting is a `TEST`
  arm with n=299; Normal is a `CONTROL` arm with n=920.
- Main Function NG is exactly 66/299 = 22.07357859531772% for Waiting and
  236/920 = 25.65217391304348% for Normal. The descriptive difference is
  -3.57859531772576 percentage points, but the source provides no
  randomization, matching, lot equivalence, or causal identification.
- The one Normal-to-Waiting comparison is therefore
  `NEEDS_REVIEW`/`UNASSESSED`, aggregation-ineligible, and effect-free.
  Component counts and composition percentages remain separate; +0V is not
  added to main NG because their inclusion relationship is not documented.
- Prompt v18 and the import gate now require percentage values to equal the
  exact human-scale `raw × 100` value from a percent-formatted source cell.
  Raw fractions, rounded display percentages, and ordinary count cells cited
  beside a percentage are rejected.
- Capture revision `capture_revision_334844564a86a310e4ceccdc` and its
  complete 67/67 packet were reused. No recapture, regrouping, image
  extraction, or image analysis occurred.
- Focused semantic/import/workflow verification passes 62 tests,
  `verify-universal-db` passes 16/16, and the post-import SQL/source audit
  passes. No application, Excel process, server, or preview was launched.

After P16, a deterministic 30-workbook small-tier benchmark was selected to
measure v18 throughput and quarantine types with three workbook workers and
two locator workers. Its durable journal and final result are under
`outputs/corpus-ingest/full-989-v1`; images remain excluded.
