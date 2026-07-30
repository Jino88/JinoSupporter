Read AI_PROMPTS/data-inference/cli-ask-ai.md and execute it.

First read tmp/ask_request.json. If it contains `microSpeakerEvidence`, use that
object as the deterministic pre-filtered evidence pack for the user's question.
It is produced by JinoSupporter before launching Codex CLI from the current
MicroSpeaker SQLite and Jino imported AI tables. Treat it as the first source to
inspect before doing broader database scans.

If the prompt begins with a run-scoped override such as
`tmp/ask_runs/<runId>/ask_request.json`, use that path instead of the shared
`tmp/ask_request.json`. Write payload and commit helper files only under the
same run folder, and never delete another run's files.

Windows execution safety:
- Never pass the completed HTML, translations JSON, or Python commit code through
  an inline PowerShell command such as `@' ... '@ | python -` or `python -c`.
  Ask AI HTML can exceed Windows command-line length limits and fail before
  Python starts.
- Persist the generated answer to a real temporary payload file first, such as
  `tmp/ask_result_payload.json`, then run a short `_tmp_ask_commit.py` script
  that reads that file and writes `AskAiHistory`.
- If a file-edit tool is unavailable, write large payload files in small chunks;
  keep each shell command short. Do not put full standalone HTML in a command
  argument or here-string.

For MicroSpeaker linkage questions, especially condition/factor versus
result/defect questions, answer from
`microSpeakerEvidence.microSpeaker.verifiedReviewCases`,
`pairConditionAggregates`, `pairRows`, `metricRows`, `measurementRows`,
`pairAggregates`, and
`microSpeakerEvidence.jino.resultRows` first. Use `pairConditionAggregates`
before individual `pairRows` when the same file/table/factor/condition appears
in repeated daily/lot rows. Use rows where `matchesAllRequiredTerms=true` as strongest evidence. If
`strictFallbackUsed=true`, state that the row matched the broader query but not
every required term group. Use `termHits` and row counts as coverage evidence.
Use `microSpeakerEvidence.questionAnalysis.factorAxisLabel`,
`outcomeAxisLabel`, `factorTerms`, and `outcomeTerms` to name the report
sections. Do not hard-code VP+CD, function, SPL, or any other domain term unless
those terms are actually present in the user's question or evidence pack.
Model boundary rule: model/product type is a grouping boundary. Do not merge
or total rows across different `sourceModels`, `models`, or `productType`
values. If the evidence pack contains multiple models and the user did not pick
one, split the answer by model or clearly label mixed-model fallback evidence.
Coverage rule: if `microSpeakerEvidence.microSpeaker.modelCoverage` lists more
than one model, do not stop after the first verified ReviewCase. Show every
covered model as `verified`, `fallback rows`, or `candidate files only`. It is
acceptable to give a strong conclusion for one model and a "needs more data" or
"fallback only" note for the others, but do not hide the other models.

ReviewCase usage:
- Verified ReviewCases are the primary evidence layer for Ask AI. Use them
  before raw pair rows, metric rows, measurement rows, and broad DB scans.
- Do not rebuild a verified ReviewCase from raw rows during ordinary Ask AI
  answering. Use the stored `changedFactors`, `outcomes`, `evidenceRows`,
  limitations, and verification status.
- Use cited source rows/cells to explain or audit a ReviewCase, especially when
  showing numeric basis or source evidence to the user.
- If the evidence pack contains only heuristic ReviewCase candidates or raw rows
  and no verified ReviewCase, then reconstruct a temporary ReviewCase from those
  extracted rows for the answer and clearly label it as unverified/fallback.
- If a verified ReviewCase is missing, insufficient for the user's question, or
  has verification issues, fall back to raw evidence rows and state the coverage
  gap.
- Every numeric conclusion must cite the source row basis and show numerator /
  denominator when it is a defect rate.
- If grouping is ambiguous, state the ambiguity and keep competing groupings
  separate instead of forcing one answer.
- User-verified calibration examples may be used as reference patterns for how
  to split grouping keys, DOE parameters, secondary conditions, outcomes, and
  exclusions. Do not answer by matching exact file names or example-specific
  part/process names. Re-check the current evidence rows every time.

Important MicroSpeaker linkage rules:
- `NG` by itself is generic result wording. It is not enough to prove a specific
  outcome category. Specific outcome evidence must match the user's outcome
  axis, for example function/functional/SPL/THD/Rub/Buzz/Noise/Hearing/Sound/
  Audio/Acoustic/DCR/IMP for a function question, or another explicit defect/
  measurement label for a different question.
- Do not treat isolated short tokens as a combined part/process label. Prefer
  the exact phrase variants from `factorTerms` and `requiredTermGroups`.
- Do not invent a factor-outcome linkage chart from unrelated factor rows and
  unrelated outcome rows. A direct linkage requires the same source dataset/file/
  table or a row that explicitly contains both the factor axis and outcome axis.
- If direct linkage rows are missing, say that the DB has related process rows
  and related outcome rows but no direct same-source linkage evidence. Then list
  the exact additional data needed instead of forcing a conclusion.
- Do not use the fixed 120-row UI cap as a review count. Use the actual selected
  rows, `termHits`, and `pairAggregates` only as coverage context.
- If `pairConditionAggregates` contains the same factor/outcome comparison,
  prefer its `testInput/testNg/testRatePercent` over a single daily `pairRows`
  item. If `aggregationMethod=total_row`, state that the source Total row was
  used. If `aggregationMethod=summed_rows`, state that repeated same-condition
  rows were summed because no Total row was available.

MicroSpeaker report shape:
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
  internally as `process defect review`, `function defect review`,
  `linked process-function review`, `measurement/spec review`, or `data limit`,
  but do not force those buckets to appear as a fixed column or fixed section.
  Use only the groups that are relevant to the question.
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
  Use charts only after this evidence table and only when the numeric comparison
  is clear.

MicroSpeaker count/rate rules:
- Always show how defect counts are used. Use `control_input`, `control_ng`,
  `test_input`, and `test_ng` from pairRows when available.
- Defect rate formula is `NG count / Input count * 100`. If the DB field is a
  decimal rate such as 0.00661, display it as 0.661%. Do not display the raw
  decimal as a percent.
- Label counts explicitly as `Input/NG`, never as an ambiguous `Count` column.
- Do not sum NG counts across unrelated rows. Aggregate only when rows share
  the same source dataset/file/table, same metric, same date/lot/line/model
  basis, and compatible denominator. If those keys differ, keep rows separate
  and write "합산 불가".
- Do not mix process NG count/rate and function NG count/rate into one total.
  Process review and function review must have separate denominators and
  separate conclusions.
- If Normal is missing, do not create a Normal comparison. Mark it as "Normal
  미확인" and explain which baseline row is missing.

Use the current AI_PROMPTS/data-inference/ai-excel-proc.md schema first:
AiDocuments, AiResults, AiNgBreakdowns, AiConclusions,
AiTroubleshootingHints, AiExtractionLogs, and translation tables.

When useful, inspect AiDocuments.RawJson as the raw AI analysis payload for
each review result, then synthesize across matching datasets to answer the
user's question.

Apply the 2026-06-04 Ask AI HTML policy:

- Use reportReviewMatrix/report domain evidence when RawJson or
  generated_report_markdown contains it.
- Decide the analysis direction first, then separate evidence only when it helps
  the chosen direction: process defect review, function defect review,
  linked process-function review, measurement/spec review, or data gap review.
- For every relevant dataset, state what was reviewed, why it was reviewed,
  what result appeared, and whether the next action should focus on process,
  function, or both.
- If the answer includes an NG-rate, defect-rate, yield, PPM, OK/NG ratio, or
  below-spec comparison, include Normal value, target/test value, relative
  change, and denominator. Codex CLI chooses whether bars, matrix, heatmap, or
  compact table best communicates the comparison.
- Use visible wording `Normal`, `Normal 대비`, `Normal 값`, or `Normal 미확인`.
  Do not display `Local Control`, `Control`, `Baseline`, or `대조군`.
- For continuous raw measurements such as tension, gauss, height, impedance,
  DCR, SPL, or THD, use a scatter or strip plot when raw points are visible and
  it improves readability. Show n per condition. If raw points are unavailable, write
  `raw distribution unavailable` and compare only visible avg/max/min values.
- Speed: filter to relevant datasets first, inspect RawJson only for those
  datasets, and avoid full DB-wide reanalysis.

Use workbook-derived merged-cell handling when available. Otherwise treat stored
excel_paste metadata tags such as `{merged=A1:A4}` as evidence that the following
value came from the merged range. Treat blank Date/Model/Type cells below a
visible value as carrying the visible value forward before pairing rows.

Do not treat percentage-only subrows as standalone result rows.

When answering, compare NG rates against same-event Normal using
`(test / normal - 1) * 100`. Source labels such as Baseline, Control, Reference,
Before, Old, or OK may be mapped internally, but visible output must say Normal.
If no Normal exists, do not claim improvement/worsening; use
ng_without_baseline ranking.

Respect the current AI_PROMPTS/data-inference/ai-excel-proc.md report types:
normal_comparison, ng_without_baseline, before_after_dimension,
measurement_spec, defect_root_cause, lot_supplier_mold_comparison,
process_condition_change, reliability_spec, doe_matrix, image_dependent, mixed.

Set `overall` to a complete standalone HTML document string when relevant
reports exist. Start with `<!doctype html><html>`, include CSS/JS/SVG inline,
and do not reference external assets or CDNs. Do not return a Markdown checklist
table for `overall`.

The HTML report is adaptive, not template-based. Include only the blocks needed
for the chosen analysis direction. Candidate action tables, bar charts, numeric
tables, scatter plots, process/function sections, and linkage sections are
available tools, not mandatory blocks. Every displayed comparison still needs a
checked item, visible result, source, interpretation, and limit.

If no registered report contains relevant information, set `overall` to a short
notice in the requested language and `perDataset` to an empty array.

Also translate the final Ask AI result into Korean, English, and Vietnamese.
Save TranslationsJson in AskAiHistory as a JSON object with keys ko, en, vi;
each value must use the same schema as the main answer: overall and perDataset.
Preserve HTML tags, CSS, JavaScript, SVG geometry, chart data, and datasetName
values unchanged; translate visible human-readable text only. Do not convert
HTML to Markdown.

If the AskAiHistory table is missing TranslationsJson, add it with:
ALTER TABLE AskAiHistory ADD COLUMN TranslationsJson TEXT NOT NULL DEFAULT '{}'

Persist the final answer to AskAiHistory.
