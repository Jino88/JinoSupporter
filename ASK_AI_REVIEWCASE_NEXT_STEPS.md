# ASK AI ReviewCase Next Steps

Date: 2026-07-04

## Latest Session Status

Status as of 2026-07-06:

- User confirmed that calibration examples must be used for generalization, not
  hard-coded special cases.
- User requested that each calibration question should open the original Excel
  workbook first, then ask only one question about that file.
- User confirmed ReviewCase creation should be AI-led: AI analyzes extracted
  workbook rows/cells, creates `changedFactors`, `outcomes`, and
  `evidenceRows`, then verifies the cited evidence.
- User-verified keep/exclude decisions are recorded in
  `REVIEWCASE_AI_AUDIT_DECISIONS.md`.
- Implemented a ReviewCase AI input packet endpoint:
  - `GET /microspeaker/review-cases/ai-packet/{fileId}.json`
  - optional limits: `rowLimit=1200`, `candidateLimit=300`
  - reads the MicroSpeaker SQLite DB in read-only mode
  - returns source file metadata, sheets, extracted rows/cells, context rows,
    pair/metric/measurement candidates, term hints, and any user audit decision
  - does not create or modify MicroSpeaker DB tables
- Generated batch ReviewCase pre-analysis drafts for all 989 extracted Excel
  files:
  - script: `tools/generate_reviewcase_batch.py`
  - manifest: `REVIEWCASE_AI_DRAFTS/batch/reviewcase_batch_manifest.json`
  - summary: `REVIEWCASE_AI_DRAFTS/batch/reviewcase_batch_summary.md`
  - file drafts: `REVIEWCASE_AI_DRAFTS/batch/files/*.reviewcase-draft.json`
  - counts: 922 `needs_ai_verification`, 60 `needs_review`, 7 `excluded`
  - evidence row existence check passed with 0 missing evidence references
- Ran AI verification for all 989 files:
  - script: `tools/verify_reviewcase_ai_batch.py`
  - manifest: `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_manifest.json`
  - summary: `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_summary.md`
  - file verifications: `REVIEWCASE_AI_DRAFTS/verified/files/*.reviewcase-ai-verification.json`
  - final counts: 13 `verified`, 917 `needs_review`, 59 `excluded`
  - file coverage check passed: 989 expected, 989 actual, 0 missing, 0 invalid JSON
  - rerun check returned `processed=0`, so there are no remaining unprocessed entries
- Continued on 2026-07-07 by wiring the 13 approved `verified` ReviewCases
  into Ask AI's deterministic MicroSpeaker evidence pack:
  - service: `JinoSupporter.Web/Services/MicroSpeakerAskEvidenceService.cs`
  - evidence field: `microSpeakerEvidence.microSpeaker.verifiedReviewCases`
  - reads `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_manifest.json`
    and approved per-file verification/draft JSON without modifying the DB
  - Ask AI prompts now name `verifiedReviewCases` as the primary evidence layer
- Next likely continuation point:
  - Validate Ask AI behavior against the initial target questions and confirm
    that matching `verifiedReviewCases` are cited before raw pair/metric rows.
  - Triage the 917 `needs_review` cases by the AI `issues`,
    `requiredUserQuestions`, and `correctionPlan` fields.
  - Promote corrected `needs_review` cases into verified ReviewCases only after
    user or AI verification confirms grouping and cited evidence.

Do not proceed by hard-coding VP/CD, Coil/CD, Suspension D, TIN, or any other
specific part/process names. The goal is a generic extractor that detects
changed factors and outcome groups from source file/table/header/row text.

## Current Decision

Current MicroSpeaker/JinoSupporter DB parsing is useful for searching related
evidence and doing some Normal/Test numeric comparisons, but it is not yet a
reliable answer layer for this real working question:

> When a process, method, raw material, jig, supplier, lot, or condition changed,
> what happened in the previous review history, and what were the results?

The current state is closer to:

- related file/search candidate extraction
- term matching
- metric candidate extraction
- Normal/Test pair comparison when obvious
- raw row/sheet evidence lookup

The desired state is:

- extracted workbook data stays broad and source-backed
- AI reconstructs review-history cases from extracted rows/cells
- AI identifies changed factors, outcomes, and evidence rows
- AI verifies that every ReviewCase field is supported by extracted source rows
- verified ReviewCases become the primary evidence layer for Ask AI
- Ask AI uses ReviewCases first, and only opens cited source rows when it needs
  to explain, audit, or handle missing/ambiguous ReviewCase coverage

## User-Defined Initial Problem Scope

The first ReviewCase implementation should be driven by these real target
questions:

1. Does VP+CD assembly condition affect function defect?
2. Does VP FILM LOT affect function defect?
3. Does FP+VP assembly bonding amount affect function defect?
4. Does Magnet Lot affect adhesion strength?

Required outcome evidence:

- matching process NG when the question is process-related
- function defect rate
- SPL / Hearing related function evidence
- adhesion strength or related measurement evidence for Magnet Lot questions

Current comparison rule from the user:

- In most useful files, Test and Normal are in the same file and usually the
  same table.
- Same-file/same-table Normal-Test comparison should be treated as the primary
  high-confidence evidence.
- If dates differ, the comparison may be weak. Keep it separate and explicitly
  show that date mismatch is a limitation instead of hiding it.
- Judgement is not the primary goal. The answer does not need to force
  improved/worse/no-change. It should first show the compared conditions,
  result values, denominator, source, and limitation clearly.

Important implication:

- Do not over-focus only on the currently listed changed factors. The system
  needs a generic changed-factor discovery layer because there will be many more
  change types than VP+CD, VP FILM LOT, FP+VP bonding amount, and Magnet Lot.
- The initial implementation can still use these four target questions as
  regression examples to tune extraction quality.
- The extractor must not be the final ReviewCase classifier. It should preserve
  workbook structure, row/cell text, numeric values, table context, and existing
  candidate rows. AI must create and verify the final ReviewCase structure.

## AI ReviewCase Analysis Contract

ReviewCase creation must be AI-led:

1. Input to AI is extracted workbook data, not only pre-made pair rows.
2. AI decides whether a workbook contains one ReviewCase, multiple ReviewCases,
   or no ReviewCase.
3. AI outputs `changedFactors`, `outcomes`, and `evidenceRows`.
4. Every output field must cite extracted row IDs, cell ranges, or source row
   references.
5. AI must run a verification pass before saving or using the ReviewCase.

Verification requirements:

- Every `evidenceRows` reference must exist in the extracted input packet.
- Every numeric value in an outcome must be copied from or calculated from cited
  extracted rows.
- Any calculated rate must show numerator, denominator, and cited rows.
- AI must flag missing baseline, missing changed condition, ambiguous grouping,
  mixed outcomes, or insufficient evidence instead of forcing a ReviewCase.
- Ask AI must treat verified ReviewCases as the source-backed answer layer.
  It should not rebuild ReviewCases from raw rows during ordinary questions.
  It should inspect cited evidence rows only for audit, display, or fallback
  when the ReviewCase is missing, unverified, or insufficient for the question.

Extractor responsibility:

- Preserve sheet names, row numbers, cells, row text, table titles, nearby
  purpose/content rows, candidate metric rows, candidate pair rows, measurement
  rows, and source file links.
- Produce candidate structure only as hints for AI. Hints are not final
  `changedFactor`, `outcome`, or judgement.

AI responsibility:

- Decide changed-factor grouping.
- Decide outcome grouping.
- Select evidence rows.
- Verify evidence support.
- Produce a structured ReviewCase with confidence and limitations.

## AI Analysis Guidelines From User Calibration

Use these examples as ground truth for AI ReviewCase analysis and verification.
They are not hard-coded extractor rules.
They should be used as calibration references: when AI sees a similar workbook
structure, it should reason by analogy from the verified pattern while still
checking the actual rows/cells in the current workbook.

Calibration reference handling:

- Keep the user's verified judgements as examples for AI analysis.
- Do not convert a verified example into an exact filename, part-name, or
  keyword rule.
- AI must always re-read the current extracted rows and cite evidence from the
  current workbook.
- A calibration example can guide what to look for, such as grouping keys,
  secondary conditions, DOE parameters, measurement-only outcomes, or exclusion
  cases.
- If the current workbook differs from the reference pattern, the current
  evidence wins.

Calibration workflow:

- Before asking the user about a source file, open the original Excel workbook.
- Ask only one source-file question at a time.
- Record the user's interpretation as ground truth.
- Do not infer the final classification from DB text alone when the user is
  available to confirm it from the opened workbook.

### Guideline 0: Exclude Non-Comparable Analysis Reports

- Do not force every report/test workbook into ReviewCase.
- If a workbook is primarily post-failure analysis, root-cause analysis, or NG
  sample analysis and does not provide a clear changed factor plus comparable
  before/after or Normal/Test result rows, route it outside ReviewCase.
- Such files may later belong to a separate `NG Analysis Case` or issue-analysis
  layer, but they should not pollute changed-factor review history.

### Guideline 1: Compare Within Row-Level Grouping Keys

- Do not collapse a whole table into one Normal/Test pair when the rows contain
  nested grouping dimensions such as equipment number, cavity, side, line,
  station, fixture, or position.
- When a table has repair/change state wording such as before/after, old/new, or
  normal/test, compare rows only within matching grouping keys.
- Treat descriptive columns such as content, type, item, sample, lot, material,
  line movement, station movement, or condition description as grouping keys
  when they distinguish separate test conditions.
- Carry row-level grouping keys into the ReviewCase outcome, so Ask AI can show
  which subgroup improved, worsened, or did not change.
- Preserve process, vision, function, and measurement outcomes separately under
  the same changed-factor context.

### Guideline 2: Use DOE Parameter Columns As Change Keys

- In DOE-style result tables, do not treat the test number alone as the changed
  factor.
- Detect controlled parameter columns such as condition name, on/off state,
  setting value, gap, pressure, speed, temperature, time, distance, amount, or
  other numeric/spec settings.
- Compare each DOE row against the baseline/standard row using the parameter
  values as the change key.
- Keep the test number as a row label only; the ReviewCase changed factor should
  be the actual parameter state and value being tested.

### Guideline 3: Keep Multiple Outcome Domains Under One Change Context

- When a report validates one material, surface treatment, supplier, process
  method, or condition change across several result sections, keep it as one
  ReviewCase change context.
- Do not split process defect, functional defect, line/module result, and
  continuous measurement sections into unrelated ReviewCases when they belong to
  the same validation.
- Store each result section as a separate outcome under the shared changed
  factor, preserving its own metric, denominator, judgement, and source rows.
- If one outcome improves while another worsens, keep the mixed result instead
  of forcing a single overall judgement.

### Guideline 4: Treat Equipment Validation As One Review With Adjustment Subconditions

- When a report validates a new or repaired machine, equipment, fixture, or
  tool, keep the validation as one ReviewCase context.
- Treat repair states, adjustment states, added support conditions, and setup
  changes as subconditions under that equipment validation instead of splitting
  them into unrelated ReviewCases.
- Compare each subcondition against the relevant baseline row within the same
  table and grouping keys.
- Keep pickup, press, vision, process, and function results as separate outcomes
  under the same equipment validation context.

### Guideline 5: Preserve Measurement Comparison Tables Without Forcing Judgement

- For validation reports that include measurement tables, preserve the measured
  values, sample counts, specs, and compared conditions as outcome evidence.
- Do not infer a final accept/reject or usable/not-usable judgement from notes,
  remarks, or measured values unless that judgement is explicitly requested for
  the downstream answer.
- When defect-rate outcomes and measurement outcomes coexist, keep both, but do
  not let a small NG-rate difference override the measurement evidence.
- Measurement outcomes should be available as comparison tables first; summary
  judgement can be added later by a separate rule.

### Guideline 6: Preserve Secondary Conditions Inside A Test Condition

- A row labelled as a test/new/changed condition can still contain secondary
  condition dimensions such as coating use, primer use, cleaning state, packing
  variant, adhesive amount, fixture state, or other treatment options.
- Do not collapse these secondary dimensions into the main changed factor.
- Compare rows only where both the main changed factor and the secondary
  condition are explicit, and keep the secondary condition as a subgroup key.
- When the same main condition has with/without variants, preserve both variants
  so Ask AI can compare them separately.

### Guideline 7: Recover Explicit Comparisons Missed By Pair Extraction

- A workbook can be a valid ReviewCase even when the current DB has no extracted
  comparison pair.
- If the workbook text or table structure explicitly defines two or more
  comparable states of a tool, fixture, material, process, equipment, or method,
  keep it as a ReviewCase candidate.
- For tool/fixture condition checks, treat wear state, repair state, normal
  state, and abnormal state as comparable condition values.
- The extractor should scan purpose/content sections and nearby result tables
  for explicit type definitions, then link them to the outcome table even when
  Normal/Test labels are not present.

### Guideline 8: Root-Cause Titles Can Still Contain Valid Condition Tests

- Do not exclude a workbook only because its title or purpose says root-cause,
  reason check, issue check, or analysis.
- If the content/result table defines a baseline condition and one or more
  changed condition values, keep it as a ReviewCase.
- Treat numeric process settings such as temperature, pressure, speed, delay,
  time, distance, amount, voltage, and force as changed condition values when
  they appear in the compared rows.
- Exclude only when the workbook lacks comparable changed-condition result rows.

### Guideline 9: Preserve Composite Changed Factors And Measurement Subgroups

- A single validation can include a primary changed factor plus secondary
  material, supplier, coating, source, or processing conditions.
- Keep the primary changed factor and the secondary condition dimensions
  together as a composite changed-factor context when the report validates them
  together.
- Preserve each result section as a separate outcome, including process/result
  tables, reliability/drop tests, and measurement tables.
- For measurement tables with item subgroups, such as magnet type, component
  type, material type, supplier group, or measurement location, keep the subgroup
  in the outcome key and compare values only within the same subgroup.
- Do not let one good outcome hide another bad outcome. Mixed section results
  are valid and should remain visible.

### Guideline 10: Keep Measurement-Only Decomposition Reviews

- A workbook can be a valid ReviewCase even when it has no Normal/Test NG-rate
  pair, if it analyzes a problem through measurement tables and explicit
  component or condition combinations.
- For measurement-only decomposition reviews, preserve the measured table,
  spec, sample rows, and combination conditions as outcome evidence.
- Do not force improved/worse judgement or accept/reject judgement unless the
  source explicitly provides it and the downstream question asks for it.
- Treat combination conditions such as max/min material values, component
  pairings, re-assembly conditions, separated part checks, and measurement
  locations as grouping keys.
- The ReviewCase should answer "what was measured under which combination" first.

### Guideline 11: Preserve Treatment Process Condition Combinations

- When a workbook compares a material or lot under multiple treatment/process
  conditions, keep the main material/lot question as one ReviewCase context.
- Preserve treatment method, chemical/primer/alcohol use, plasma state, speed,
  before/after state, and date/lot conditions as subgroup keys when they define
  separate test combinations.
- Do not require an existing `comparison_pairs` row. If the source table clearly
  lays out condition combinations and result columns, AI should build the
  ReviewCase from sheet rows/cells.
- Split outcomes by result domain such as surface/material check, tension or
  other measurement, and function/defect result.
- Compare only rows that share compatible subgroup keys; otherwise show them as
  separate condition rows rather than forcing a pair.

### Guideline 12: Preserve Shift, Date, And Total Aggregation Basis

- When a validation table includes day/night shift, date, lot, line, or other
  repeated production grouping rows, preserve those fields as subgroup keys.
- A `TOTAL` row is an aggregation basis, not the same kind of condition as a
  day/date/shift row.
- Compare TOTAL rows with compatible TOTAL or baseline rows when available, and
  compare shift/date rows only with compatible shift/date or clearly matching
  baseline rows.
- Keep process, separation, function, and measurement sections as separate
  outcomes even when they share the same changed factor.
- If one subgroup improves and another worsens, keep subgroup-level results
  visible instead of collapsing to one judgement.

### Guideline 13: Preserve Repeated Validation Runs Across Dates And Trials

- A single ReviewCase can contain repeated validation runs across multiple
  dates, shifts, machines, lots, or trial numbers.
- Preserve trial labels such as first/second/third run, repeated test, new
  machine run, and normal/baseline rows as condition metadata.
- Do not reduce repeated runs to one row unless the workbook provides a clear
  total/summary row or the downstream task explicitly asks for aggregation.
- When repeated runs contain several outcome sections, keep each section as a
  separate outcome and keep date/shift/trial evidence attached to each row.
- Existing pair extraction may capture only one outcome section; AI should use
  sheet rows/cells to recover the other sections before creating ReviewCase.

### Guideline 14: Preserve Supplier And Coating As Subgroups In Material Reviews

- When a material review changes supplier, coating source, coating location,
  production source, or vendor-related condition, keep those fields as subgroup
  keys under the material ReviewCase.
- Preserve test round or repetition labels as subgroup keys when the workbook
  has first/second or repeated validation sections.
- Keep decap, drop, tension, function, and measurement sections as separate
  outcomes.
- For tension or other measurement sections, compare within the same component
  or magnet/material subgroup.

### Guideline 15: Preserve Dimension/Spec Context In Material Reviews

- When a review is driven by dimension/spec issues, preserve the source spec
  wording and actual-over/under/out-of-spec description as changed-factor
  evidence.
- Keep repeated sheets, dates, runs, or second-check sections as subgroup keys
  instead of overwriting earlier results.
- Split AOI/vision, height/dimension, tension, function, and reliability/drop
  sections into separate outcomes under the same dimension/spec ReviewCase.
- Do not judge the whole ReviewCase from one outcome section when other sections
  show different behavior.

### Calibration 1: New VP+CD Assembly Equipment

Source file:

`01.1.. BRS-161016DT  Report NEW machine Bond Ass'y VP-CD  improve NG separate VP+CD ( after improve base) Date  2.06.2025_1778470943_clean.xlsx`

User interpretation:

- This report is a validation of new VP+CD assembly equipment.
- Do not classify this only as a generic method/process change.
- This is also only a calibration example. Do not hard-code `VP+CD`; the
  generalized rule must detect the assembly target and equipment/process wording
  dynamically from each source file.

ReviewCase mapping:

| Field | Value |
| --- | --- |
| change_domain | `equipment` / `assembly equipment` |
| changed_factor | source-derived new assembly equipment target, in this file: `new VP+CD assembly equipment` |
| reviewed_process | source-derived assembly process, in this file: `VP+CD assembly` |
| comparison_scope | same file, same table Normal/Test where available |
| process outcome | `Vision VP/CD NG rate` |
| function outcome | `Function NG rate` |
| result handling | keep process NG and function NG as separate outcomes |
| confidence | high when using same-table Total/Normal rows |

Generalized extraction rule from this example:

- Do not hard-code `VP+CD`.
- Detect `new machine`, `new equipment`, `new assembly machine`,
  `new ass'y`, `new jig`, or similar wording as an equipment/process validation
  signal.
- Extract the nearby assembly target dynamically from the file title, table
  title, and Normal/Test condition labels.
- Classify as equipment validation when the source wording indicates new
  machine/equipment, even if the outcome tables are process NG or function NG.
- Keep process outcome and function outcome separate when both appear in the
  same report.

### Calibration 2: VP+CD and Coil+CD Bonding Amount

Source file:

`009.MSU-L20S15-07 Report test change bonding amount VP+CD, Coil+CD_1778470550_clean.xlsx`

User interpretation:

- The report reviews both `VP+CD bonding amount` and `Coil+CD bonding amount`.
- In many files, the reviewed/changed items are written in the table and are
  also roughly present in the file title.
- File title and table title should both be used as strong changed-factor
  evidence, but table content should remain the primary source when available.

ReviewCase mapping:

| Field | Value |
| --- | --- |
| change_domain | `bonding amount` / `process condition` |
| changed_factors | source-derived bonding targets in this file: `VP+CD bonding amount`; `Coil+CD bonding amount` |
| has_multiple_changed_factors | true |
| reviewed_process | bonding / assembly bonding |
| outcome candidates | `Air leak`, `Sigma`, `Hearing`, `Function NG rate` |
| changed_factor_evidence | file title + table title + row labels |
| confidence | high when table labels and title agree |

Generalized extraction rule from this example:

- Do not hard-code `VP+CD` or `Coil+CD`.
- Detect `bonding amount`, `bond amount`, `glue amount`, `bonding quantity`,
  or similar process-amount phrases as the change domain.
- Extract the nearby component/assembly target dynamically from the same phrase,
  table title, header, or row label.
- If multiple targets appear in the title/table, create multiple changed-factor
  candidates or one case with `has_multiple_changed_factors=true`, depending on
  whether the result rows can be separated.
- Raise confidence when the same target/domain appears in both file title and
  table/header/row labels.
- Keep the original source wording in evidence even after normalization.

### Calibration 3: Suspension D Material and Tin Plating Method

Source file:

`02. BRS-161016 GMI Report  test MTR Suspension D used Susp- Array 161014 D-1 changing tin lating method - 2025.05.13_clean.xlsx`

User interpretation:

- This report reviews `161014 Suspension D` material.
- The tin plating method was changed from `Polish` to `Non Polish`.
- The report includes SPOT process defect rate, tensile strength, and function
  defect data.
- Do not classify this file as only a SPOT process defect-rate review. The
  ReviewCase must preserve all three relevant outcome groups when they are
  present in the workbook.

ReviewCase mapping:

| Field | Value |
| --- | --- |
| change_domain | `material` + `plating method` |
| changed_factor | source-derived Suspension material and plating method change |
| before_condition | `Polish` tin plating method |
| after_condition | `Non Polish` tin plating method |
| reviewed_process | `SPOT` / spot welding process |
| outcome groups | `SPOT process defect rate`; `tensile strength`; `function defect rate` |
| result handling | keep process NG, tensile measurement, and function NG as separate outcomes under the same changed factor |
| confidence | high when each outcome uses same-table Normal/Test rows or clearly linked same-file sections |

Generalized extraction rule from this example:

- Do not hard-code `Suspension D` or `TIN`.
- Detect material/part identifiers and process-method changes separately when
  both appear in the source title/table.
- If a phrase indicates a method transition such as `A -> B`, `from A to B`,
  `Polish to Non Polish`, or similar, store it as before/after condition.
- Outcome groups should follow the workbook contents, not only the first table.
  If a material/method change report includes process NG, tensile strength, and
  function NG, preserve all of them as separate outcomes under the same changed
  factor. Do not collapse them into one denominator or hide non-primary outcomes.

## Current DB State

MicroSpeaker parsed DB:

`D:\000. MyWorks\005. Program\Repository\MicroSpeaker_ProductTech_DB\db\InputDataFinish.sqlite`

Observed table counts:

| Table | Count |
| --- | ---: |
| files | 989 |
| term_hits | 16,227 |
| metric_candidates | 7,976 |
| comparison_pairs | 1,587 |
| measurement_stats | 22,035 |
| sheet_rows | 88,257 |

Main parsed tables:

- `files`
- `sheet_rows`
- `sheet_cells`
- `term_hits`
- `metric_candidates`
- `comparison_pairs`
- `measurement_stats`

JinoSupporter process review DB:

`D:\000. MyWorks\002. DB\process-review.db`

Observed relevant table counts:

| Table | Count |
| --- | ---: |
| AiDocuments | 989 |
| AiTestConditions | 1,582 |
| AiResults | 31,586 |
| AiNgBreakdowns | 6,460 |
| DatasetSummary | 989 |

Important Jino tables:

- `AiDocuments`
- `AiTestConditions`
- `AiResults`
- `AiNgBreakdowns`
- `AiConclusions`
- `AiTroubleshootingHints`
- `DatasetSummary`

## What Works Now

Current ASK AI can often answer questions like:

- "press method change related NG rate comparison"
- "VP mold change and Function NG result"
- "Normal/Test rows where condition changed"
- "find related review files for a factor/outcome"

It can use:

- `comparison_pairs` for obvious Normal/Test rows
- `metric_candidates` for Input/OK/NG/rate rows
- `measurement_stats` for min/max/avg/n measurement evidence
- `term_hits` for related file discovery
- Jino `AiTestConditions` and `AiResults` when the AI extraction produced usable condition/result links

## Main Gap

The missing layer is a normalized review-case model.

Right now, changed factors are often buried in free text fields such as:

- `Change method UC press VP/CD`
- `dry UC press`
- `new machine`
- `material make press jig`
- `VP #7 (Improve mold)`
- `Test VP change thickness 70 -> 65`

The DB does not consistently know whether each text means:

- process change
- method change
- equipment change
- jig change
- raw material change
- supplier/lot change
- spec/measurement change
- inspection/retest change

Also, the result side is not always linked to the changed factor as a single
review case. For example:

- one row may describe a process condition
- another row may describe function NG
- the direct relationship may be in the same file/table, but not normalized
- ASK AI currently has to infer too much from raw strings

## Required Next Layer: ReviewCase

Build a normalized `ReviewCase` layer on top of the parsed DB.

Generalization rule:

- Do not hard-code specific part names, assemblies, lots, suppliers, or process
  labels such as `VP+CD`, `Coil+CD`, `Magnet`, or `press method` as the only
  supported cases.
- Use those terms only as examples and calibration evidence.
- The extractor must discover the changed factor from source text at runtime:
  file title, sheet name, table title, purpose/content rows, headers, row labels,
  Normal/Test condition labels, and Jino AI condition fields.
- Normalize the discovered source phrase into a generic domain while preserving
  the original phrase as evidence.
- Expected generic domains include assembly condition, process method,
  machine/equipment, jig/fixture, material, supplier, lot, mold, bonding amount,
  dry/UV/plasma, dimension/spec, inspection/retest, and measurement condition.

Recommended conceptual schema:

| Field | Meaning |
| --- | --- |
| case_id | stable generated id |
| source_system | `MicroSpeaker` or `Jino` |
| file_id / document_id | source identity |
| source_file | original file name/path |
| original_file_url | JinoSupporter source-file link when available |
| sheet_name | evidence sheet |
| source_rows / source_cells | row/cell evidence |
| review_title | source table title or report title |
| review_purpose | normalized purpose |
| change_domain | `process`, `method`, `equipment`, `jig`, `material`, `supplier`, `lot`, `spec`, `inspection`, `unknown` |
| changed_factor | normalized changed item, e.g. `press method`, `VP mold`, `bonding amount` |
| before_condition | Normal/before/old condition |
| after_condition | Test/after/new condition |
| outcome_domain | `process defect`, `function defect`, `measurement`, `reliability`, `unknown` |
| outcome_metric | e.g. `Vision VP/CD NG%`, `Function NG%`, `SPL`, `THD`, `tension` |
| normal_input | Normal denominator |
| normal_ng | Normal NG count |
| normal_rate | Normal NG/result rate |
| test_input | Test denominator |
| test_ng | Test NG count |
| test_rate | Test NG/result rate |
| relative_change_percent | `(test_rate / normal_rate - 1) * 100` |
| judgement | `improved`, `worse`, `no_change`, `mixed`, `not_comparable` |
| aggregation_method | `single_row`, `total_row`, `summed_rows`, `not_aggregated` |
| confidence | `high`, `medium`, `low` |
| limit_reason | why the case is weak or not comparable |

## Extraction Approach

Use existing parsed DB first. Do not re-open Excel unless required for a
specific missing field or parser bug.

1. Start from `comparison_pairs`.
   - It already gives Normal/Test-like comparisons and rates.
   - Join to `files` for source info.
   - Use `table_title`, `compare_item`, `control_condition`, and `test_condition`
     to infer changed factor and outcome metric.

2. Add aggregation over repeated same-condition rows.
   - Prefer Total-equivalent row when present.
   - Otherwise sum only rows with compatible file/table/metric/date/lot/line/model basis.
   - Never combine process NG and function NG denominators.

3. Use `metric_candidates` to recover rows not paired into `comparison_pairs`.
   - Especially NG-rate matrix and one-off result rows.

4. Use `measurement_stats` for continuous measurement reviews.
   - Preserve min/max/avg/n.
   - Mark `raw distribution unavailable` when only aggregate stats exist.

5. Use Jino `AiTestConditions` + `AiResults` as a second source.
   - Good target fields: `ChangedFactor`, `BeforeValue`, `AfterValue`, `Process`,
     `Machine`, `Jig`, `MaterialLot`, `Supplier`, `DryTimeSec`, `Temperature`,
     `Pressure`, `BondAmount`, `UvEnergy`.
   - Current AI extraction quality is mixed, so do not trust it blindly.
   - Cross-check against source rows and MicroSpeaker pair data when possible.

6. Add alias/normalization dictionary.
   - Example: `UC press method`, `dry UC press`, `press method change` should map
     to `press method`.
   - Example domains:
     - `dry`, `UV`, `curing`, `press`, `bonding`, `plasma` -> process/method
     - `jig`, `fixture`, `machine`, `AWF`, `MC` -> equipment/jig
     - `MTR`, `material`, `supplier`, `lot`, `vendor` -> material/supplier/lot
     - `SPL`, `THD`, `F0`, `DCR`, `IMP`, `hearing`, `function` -> function/measurement outcome

## ASK AI Behavior Target

When the user asks:

> "이런이런 공정/공법/원자재가 변경됐을 때 결과들이 어땠지?"

ASK AI should:

1. Decide analysis direction first.
   - `condition impact`
   - `Normal/Test comparison`
   - `process-function linkage`
   - `measurement/spec review`
   - `data gap review`

2. Search `ReviewCase`, not raw tables first.

3. Group by normalized changed factor and outcome metric.

4. Show source-backed results:
   - changed factor
   - Normal condition
   - Test condition
   - outcome metric
   - Input/NG or measurement n/min/max/avg
   - rate or value change
   - judgement
   - original file link
   - confidence/limit

5. Let Codex CLI choose visualization.
   - No fixed HTML template.
   - Use bar, matrix, heatmap, compact table, scatter/strip plot, or grouped
     evidence block based on the selected analysis direction.

## Important Existing Changes

Parser fix already made:

`D:\000. MyWorks\005. Program\Repository\MicroSpeaker_ProductTech_DB\tools\incremental_dataset_indexer.py`

- `extract_metric_candidates` scan range was increased from `idx + 18` to
  `idx + 80`.
- Reason: long tables were missing Total rows.
- Example fixed case:
  - file_id `511`
  - `3.1 MSU-L20S15-07 Report test change method Dry VP+CD...`
  - DB now includes `960/26` vs `80928/608`, `2.708%` vs `0.751%`.

JinoSupporter ASK AI evidence changes:

`JinoSupporter.Web\Services\MicroSpeakerAskEvidenceService.cs`

- Builds deterministic `microSpeakerEvidence`.
- Adds `questionAnalysis`.
- Adds `pairConditionAggregates`.
- Uses aggregate/Total row before individual daily pair rows.

Prompt changes:

- `AI_PROMPTS/data-inference/ask-ai-cli.md`
- `AI_PROMPTS/data-inference/cli-ask-ai.md`

Current prompt direction:

- decide analysis direction first
- no fixed HTML template
- Codex CLI chooses comparison/visualization
- keep evidence contract fixed
- do not mix process NG and function NG denominators
- use original file links instead of long file names

## Recommended Next Implementation

Create a small service/table first, not a full UI.

Suggested implementation path:

1. Add a ReviewCase extraction service.
   - Possible file:
     `JinoSupporter.Web\Services\MicroSpeakerReviewCaseService.cs`

2. Add a generated SQLite table or in-memory pack.
   - If persisted, suggested DB/table:
     `MicroSpeaker_ProductTech_DB\db\InputDataFinish.sqlite`
     table `review_cases`
   - Safer first step: generate in memory and inspect output before altering DB.

3. Build a diagnostic export.
   - Generate `tmp/review_cases_sample.json` or a small HTML/CSV preview.
   - Review 30-50 cases manually:
     - press method
     - VP mold
     - bonding amount
     - material/supplier/lot
     - jig/machine

4. Once extraction quality is acceptable, make ASK AI use ReviewCase first.
   - `ReviewCase` evidence first
   - fallback to `pairRows`, `metricRows`, `measurementRows`, raw rows

5. Only after that, tune the HTML output.

## Verification Notes

Per repository instructions:

- Do not launch app/dev server unless explicitly requested.
- Verify code changes with the narrowest relevant command.
- For prompt/docs-only changes, `git diff --check` is enough.
- For C# service changes, run:
  `dotnet build .\JinoSupporter.Web\JinoSupporter.Web.csproj --no-restore -p:UseAppHost=false -p:OutputPath=.\artifacts\codex-check\<name>\`
- Remove generated verification artifacts after build.

## Bottom Line

Current DB is not wasted. It has enough raw material to build the desired answer
system. But the missing piece is the `ReviewCase` layer.

Do not spend more time forcing fixed HTML layouts. The next meaningful work is
to normalize previous review files into comparable review cases, then let Codex
CLI analyze and visualize those cases based on the user's question.
