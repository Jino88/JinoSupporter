# Final Goal Execution Plan

Confirmed by the user on 2026-07-17. This document is the execution contract for the complete InferenceDataAIService goal.

## 2026-07-20 execution-path amendment

The user-confirmed default semantic path is now the table-first pipeline in
[`TABLE_FIRST_SEMANTIC_PIPELINE.md`](TABLE_FIRST_SEMANTIC_PIPELINE.md).
The trust rules and canonical destination below remain valid, but exhaustive
per-cell AI semanticization and workbook-wide blocking validation are no
longer the default ingestion method. Code preserves and calculates raw values;
AI classifies table purpose, comparison structure, metrics, and table grouping
from one compact workbook request. Ambiguity is stored as `NEEDS_REVIEW` or
`PARTIAL` and must not discard unrelated usable tables.

## Goal

Convert all historical and future Excel review workbooks into an evidence-traceable database of comparable studies, produce one consolidated analysis per workbook, answer natural-language questions from valid control/comparison evidence, and support incremental ingestion and source-table drill-down through WPF.

`VP+CD assembly -> FUNCTION NG` is a representative acceptance example, not a fixed domain boundary. The query system must support any factor, context, process, material, equipment, model, lot, condition, or outcome represented by validated database evidence. Seed concepts and aliases may improve initial recall but must never become workbook-specific extraction rules or a whitelist of answerable questions.

The current source corpus is:

- root: `D:\000. MyWorks\test\result\InputDataFinish`
- files: 989
- format: 989 `.xlsx`
- total size at the start of execution: 3,214,102,996 bytes
- DRM constraint: none; Open XML/openpyxl extraction is allowed

## Non-negotiable trust rules

1. The source workbook remains read-only and authoritative.
2. Every numeric or directional claim must cite stable evidence/data ids and an exact workbook, sheet, and cell/table range.
3. An effect may be calculated only when a control/comparison basis and compatible denominator or measurement basis are validated.
4. Different models, lots, periods, lines, units, sample bases, or confounded changes must not be silently pooled.
5. Association and causation must be distinguished. Weak, unmatched, descriptive, or incomplete evidence cannot produce a causal claim.
6. Missing, conflicting, truncated, stale, and unverified evidence must be visible in stored status and user answers.
7. Each workbook receives a consolidated study-level analysis. Raw cells, tables, and candidate regions remain drill-down evidence rather than the main answer.
8. Historical and newly added workbooks use the same canonical schema, validation gates, and query path.

Scope update confirmed by the user on 2026-07-17: embedded workbook images do not need to be extracted or analyzed. Image-only content may be reported as `NO_TABULAR_EVIDENCE` or excluded with a reason; it must not block tabular ingestion.

## Reuse boundary

The existing universal-grid and analysis tables are the migration base because they already preserve workbook fingerprints, fixed coordinates, merge ranges, cohorts, metrics, comparisons, conclusions, and evidence ranges.

The following are reusable only as supporting layers:

- layout groups and renderer keys: extractor-routing hints;
- numeric candidates: non-authoritative discovery hints;
- v5 HTML: audit evidence for the old renderer, not the knowledge model;
- existing verified analysis manifests: pilot fixtures that must be migrated and revalidated.

The current generic HTML renderer will not be refined as the main deliverable.

## Canonical knowledge contract

The normalized chain is:

`Study -> Context -> ChangedFactor/Intervention -> Cohort/Arm -> Comparison -> Outcome -> Observation -> Effect -> Evidence`

It also requires:

- canonical concepts and aliases while retaining original wording;
- normalized model, process, component, material, equipment, lot, line/site, period, unit, and metric fields;
- explicit control, treatment/comparison, before/after, paired, multi-arm, descriptive, and confounded design states;
- stable public study, comparison, effect, and evidence identifiers;
- validation issues and approval history;
- workbook/sheet/range provenance;
- duplicate and related-study links;
- versioned extraction, semantic-analysis, and verification metadata.

## Execution phases and completion gates

### Phase 0 — Freeze, inventory, and baseline

Status: completed on 2026-07-17.

- The rejected v5 render outputs are retained as audit artifacts.
- The Excel/HTML audit and final goal are recorded in the session handoff and Obsidian note.
- The corpus is confirmed as 989 DRM-free XLSX files.
- The existing universal DB contains only 16 workbooks and therefore is a pilot DB, not the historical corpus.
- The relevant baseline Python tests pass.

Gate: inputs, current limitations, immutable trust rules, and reuse boundaries are documented.

### Phase 1 — Canonical schema and migration

Status: completed on 2026-07-17.

Deliverables:

- idempotent SQLite migrations for canonical concepts, contexts, factors, arms, outcomes, observations, effects, evidence, validation issues, relationships, and pipeline versions;
- backward-compatible migration of existing analysis records;
- schema constraints and indexes that prevent unsupported effects and broken evidence references;
- inspect/export commands and focused tests.

Gate: a database can represent the VP+CD example and unrelated factor/outcome domains without layout-specific columns, all existing analysis tests still pass, and repeated migration is safe.

### Phase 2 — Representative pilot and golden questions

Status: completed on 2026-07-17.

Deliverables:

- an automatically selected, diverse pilot set covering layout, size, sheets, merges, formulas, metric types, comparison designs, and graceful handling of files without tabular evidence;
- expected study counts, expected evidence regions, exclusions, and ambiguity notes;
- golden questions, including VP+CD assembly versus FUNCTION NG, with machine-checkable expected retrieval properties.

Gate: the pilot set exercises every major extraction and reasoning risk before full-corpus processing.

### Phase 3 — Provenance-complete Capture v2

Status: implementation and representative-pilot verification completed on 2026-07-17; full-corpus capture remains gated by Phase 5 pilot acceptance.

Deliverables:

- read-only Open XML extraction for all 989 XLSX files;
- workbook/sheet/cell coordinates, displayed/raw/formula values where available, merge structure, style/number format, row heights, column widths, and visibility;
- deterministic fingerprints, resumability, per-file failure isolation, and source-change invalidation;
- capture verification independent of HTML.

Gate result for the representative pilot:

- 30/30 current source revisions captured and SHA-256 matched to the source files;
- 105 sheets, 480,102 sparse source/structural cells, 10,415 formulas, and 4,643 merge ranges stored;
- 27 workbooks are `CAPTURED`, two are `EMPTY_WORKBOOK`, and one is `NO_TABULAR_EVIDENCE`;
- zero failed items, zero unfinished runs, 30/30 canonical source-revision bridges, and SQLite integrity `ok`;
- formulas, cached/display/raw values, number formats, styles, dimensions, hidden state, and merges remain distinguishable;
- embedded images are ignored by contract and do not enter the evidence database.

The full 989-workbook execution is intentionally deferred until semantic extraction passes the representative end-to-end gate, so a flawed semantic contract is not multiplied across the corpus.

### Phase 4 — Semantic study extraction and validation

Status: implementation complete for the versioned packet/locator/draft/validator/import
contract; representative pilot expansion is in progress.

Deliverables:

- deterministic hints plus AI-produced versioned study manifests;
- canonical term/unit mapping with original terms preserved;
- study, context, changed factor, comparator, outcome, observation, conclusion, limitation, and exact evidence extraction;
- validators for evidence existence, sample/denominator compatibility, unit consistency, formula/rate arithmetic, confounding, completeness, and stale source state;
- `VERIFIED`, `NEEDS_REVIEW`, `EXCLUDED`, `FAILED`, and `STALE` states with explicit reasons.

Implemented and verified on 2026-07-17:

- all 30 representative Capture v2 revisions were converted to complete
  `semantic-source-packet-v1` artifacts: 864 lossless row-first chunks own
  exactly 234,287 nonempty semantic cells, with zero missing or duplicate
  primary coordinates;
- 493 numeric/formula continuation chunks are retained deterministically for
  later evidence retrieval without unnecessary AI calls; 371 text-bearing
  chunks require semantic location;
- the locator is open-domain, preserves unfamiliar original terms, excludes
  images, and can now process several independent chunks per read-only AI call
  while validating each result against its own sheet/range boundary;
- AI drafts cannot self-verify, calculate effects, enable aggregation, or claim
  causality; source identity, packet completeness, A1 existence, and every
  numeric observation are checked deterministically before import;
- one real tabular workbook and both terminal-status classes completed the
  source-to-canonical path. The tabular study remains `NEEDS_REVIEW` because no
  source-backed control/matching basis was explicit; the terminal cases are
  `EXCLUDED` with zero invented studies or effects.

Gate: no verified claim lacks source evidence; incomplete packets cannot yield approved or causal conclusions.

### Phase 5 — Pilot end-to-end acceptance

Status: in progress. One tabular workbook plus two terminal workbooks have
completed the full path; the remaining representative workbooks and all ten
golden questions still require reviewed acceptance.

Deliverables:

- pilot ingestion from source XLSX through consolidated study records;
- workbook-level summaries and exact source-table drill-down;
- retrieval and effect calculations for golden questions;
- reviewed false-positive, false-negative, unsupported-claim, and aggregation results.

Gate: every pilot answer is reproducible from stored observations/effects and every cited id opens the exact supporting range.

### Phase 6 — Full historical migration

Status: pending.

Deliverables:

- resumable processing of all 989 workbooks;
- one terminal ingestion status per source fingerprint;
- consolidated analysis coverage report;
- retry and quarantine manifests for failures;
- schema-candidate review without silently changing the canonical contract.

Gate: all 989 files are accounted for as verified, needs-review, excluded, or failed with a concrete reason; no file disappears from coverage statistics.

### Phase 7 — Evidence retrieval, calculation, and answer synthesis

Status: deterministic retrieval, evidence-pack, answer construction, answer
validation, and exact EVD drill-down are implemented; the full
golden-question acceptance set remains in progress.

Deliverables:

- alias-aware concept search;
- comparability partitions for model, lot, process, intervention, outcome, unit, period, and design;
- deterministic absolute, percentage-point, relative, and applicable risk calculations;
- consistency, range, conflicting-evidence, and insufficiency summaries;
- evidence packs that constrain AI answers to cited records;
- stable links to exact source tables and original workbooks.

The generic query path now tokenizes arbitrary Unicode terms, searches
canonical/original labels and aliases without a domain whitelist, ranks studies
by multi-term coverage, exposes descriptive observations when no valid
comparison exists, and keeps unverified/confounded/invalid effects out of the
answer-eligible set. The real `waiting 2 day function NG` smoke query applies a
generic factor/context-plus-outcome relevance gate and retrieves the one Study
matching both sides of that relationship. It contains 14 outcomes, 44
observations, and exact source citations, while correctly returning zero
answer-eligible effects because no verified comparison exists.

`canonical-evidence-answer-v1` now generates Korean answers without allowing AI
to rewrite numbers, direction, stable IDs, or citations. Compatible effects are
grouped only when outcome, metric/unit/denominator, factor transitions,
control/comparison conditions, all contexts, design/matching basis, and
stratum/replicate signatures are identical. Opposing directions are reported
as conflicting rather than averaged. Descriptive records are visible without
new difference calculations, and outcome details are filtered using
token-boundary matching so a term such as `NG` does not falsely match
`hearing`. `canonical-evidence-detail-v1` resolves one `EVD-*` only through its
explicit current Capture v2 bridge and returns exact cells, formulas, cached
values, number formats, styles, merges, and hidden dimensions.

Gate: the VP+CD/FUNCTION NG question, the remaining cross-domain golden questions, and previously unseen concept combinations use the same generic retrieval path, return all relevant studies, separate invalid comparisons, calculate only validated effects, and attach citations to every quantitative statement.

### Phase 8 — WPF integration and new-workbook ingestion

Status: core implementation completed; non-UI build verification passed.
Interactive runtime acceptance and related-study/duplicate review presentation
remain in progress.

Deliverables:

- natural-language evidence query screen;
- study list, consolidated summary, evidence table, and original workbook/range drill-down;
- incremental new-XLSX intake through the same Capture v2, semantic extraction, validation, and database path;
- duplicate/related-study detection, progress, failure, and review status;
- no requirement for the user to manually group every workbook.

The durable `incremental-xlsx-ingest-v1` workflow is source-read-only and
journals Capture, Packet, Locator, Draft, Import, and Verify stages. Repeated
source fingerprints resume valid AI artifacts and re-import idempotently.
Terminal empty/non-tabular workbooks do not invoke AI, and all non-terminal AI
drafts remain `NEEDS_REVIEW`.

WPF now has unrestricted `검토 DB 질문` and `신규 Excel DB 적재` tabs. It
delegates eligibility and answer math to the Python canonical contract, shows
stable EVD citations, renders the selected evidence range as a coordinate- and
merge-aware table, and opens the exact source sheet/range through user-triggered
read-only Excel COM. It was intentionally not launched during verification;
the WPF project builds with zero warnings and zero errors.

Gate: a newly selected workbook can be ingested without modifying it, becomes queryable after validation, and produces the same evidence contract as historical data.

### Phase 9 — Regression, operations, and handoff

Status: pending.

Deliverables:

- narrow automated tests for each layer and end-to-end pilot regression;
- corpus coverage, validation, and query-quality reports;
- backup, schema migration, restart/resume, and failure-recovery instructions;
- final handoff with exact commands and known limitations.

Gate: the complete system is reproducible from documented commands and no required stage remains represented only by a plan.

## Verification policy

- Verify each change with the narrowest relevant unit test, migration test, or project build.
- Run pilot processing before the 989-file migration.
- Run full-corpus work only through resumable commands with checkpoints and per-file status.
- Do not launch WPF or other desktop applications for verification unless the user explicitly asks; use build and non-UI tests.
- Never promote `NEEDS_REVIEW`, `STALE`, truncated, or failed records into primary answer evidence.

## Progress ledger

| Date | Phase | Result |
|---|---|---|
| 2026-07-17 | 0 | Final goal, renderer audit, corpus count, baseline DB state, and test baseline recorded |
| 2026-07-17 | 1 | Additive canonical schema installed; 11 legacy analyses migrated to 22 Studies, 56 Comparisons, and 281 Evidence items with canonical integrity passing |
| 2026-07-17 | 2 | Thirty existing representative workbooks and ten golden questions saved; embedded-image extraction/analysis explicitly disabled |
| 2026-07-17 | 3 | Capture v2 implemented and verified for all 30 pilot sources: 105 sheets, 480,102 sparse cells, 10,415 formulas, 4,643 merges, 30/30 source hashes and canonical bridges valid, zero failed/unfinished items |
| 2026-07-17 | 4 | Thirty complete semantic packet sets built: 864 chunks and 234,287 exact primary cells; strict open-domain locator, conservative draft gate, source-range validation, and numeric-evidence validation implemented |
| 2026-07-17 | 5 | Real E2E completed for one tabular workbook and two terminal workbooks; 1 Study/2 Arms/14 Outcomes/44 Observations imported as review-required with no unsupported Comparison or Effect |
| 2026-07-17 | 7 | Generic Unicode evidence retrieval implemented with answer-eligible/excluded separation, deterministic eligible-effect summaries, stable DATA/CMP/EFF/EVD citations, and descriptive no-comparison drill-down |
| 2026-07-17 | 7 | Deterministic Korean answer/validation and current-revision EVD detail implemented; exact compatibility grouping prevents cross-model/lot/condition averaging, and the actual waiting-2-day query returns one relevant DATA record with zero invented effects |
| 2026-07-17 | 8 | Durable single-XLSX journaled ingestion and WPF free-text query/intake/EVD table/exact Excel-range surfaces implemented; focused Python regressions and the WPF narrow build pass, with no desktop launch |
| 2026-07-18 | 4/5 | P12 was redrafted under prompt v5 after rejecting the lossy v2 result. It now preserves 22 scalar outcomes/66 scalar observations plus four measurement series/2,697 exact points; no unsupported cooling comparison or effect was created |
| 2026-07-18 | 7 | Compact measurement-series retrieval and deterministic answers were added. Coordinate-based geometry counts preserve repeated labels, exact EVDs remain attached, and relationship filtering no longer retrieves an unrelated same-model XRAY source |
| 2026-07-18 | 9 | Scoped regression reached 100 passing tests; WPF builds with 0 warnings/0 errors; SQLite integrity is `ok` with 0 canonical FK errors; golden acceptance has 2 pass, 8 pending-ingest, 0 fail, and 0 retrieval miss |

## Current execution checkpoint — 2026-07-17

This is the latest authoritative state. The final goal is not complete.

- The frozen corpus journal accounts for exactly 989 DRM-free XLSX sources
  (3,214,102,996 bytes). Eight are completed, 981 are pending, and none have
  failed. Current results are five `NEEDS_REVIEW` and three `EXCLUDED`.
- P11 (`SUS-D cutting 40%/50%/60%`) validates the v2 baseline behavior:
  4 Studies, 16 Arms, 4 Outcomes, 16 Observations, 12 Comparisons, and zero
  Effects. Every `Normal` row is preserved as a control/baseline, and each
  40/50/60 condition has a draft comparison to Normal.
- P11 remains correctly fail-closed. Study comparability/confounding and
  Comparison validity/confounding require human judgment; 9 comparisons also
  need a matching basis. Different sample sizes and other possible changes
  must be reviewed from the exact source ranges before any approval.
- P03 was produced under the earlier v1 draft prompt and omitted explicit
  Normal comparisons. It remains blocked and must be redrafted/reviewed under
  the corrected contract rather than being treated as accepted evidence.
- P10's label-only draft was deterministically converted to
  `NO_TABULAR_EVIDENCE` without recapture, regrouping, or a second AI call.
- Terminal workbook analyses now participate in retrieval as source-level
  exclusions. Legacy descriptive values without direct current-revision EVD
  evidence are withheld rather than displayed as trustworthy observations.
- The human review contract exposes a read-only queue/detail view and explicit
  `APPROVE`, `REJECT`, `EXCLUDE`, and `RETURN_TO_REVIEW` decisions. Approval
  requires current SHA-256 source identity, direct verified Comparison and
  paired Observation evidence, explicit `VALID/NONE` judgments, a nonempty
  matching basis, and deterministic effect calculation. No real Comparison
  has been auto-approved; `review_decisions` remains empty.
- WPF now includes the human review queue, paired values, exact EVD
  sheet/ranges, source-table preview, exact Excel-range navigation, explicit
  assessment controls, and a confirmation step. The WPF project builds with
  zero warnings and zero errors and was not launched.
- `related-studies` separates exact SHA-256 duplicates from transparent
  lexical similarity; similarity is discovery-only and is never relationship
  or causal evidence.
- The ten-question acceptance runner executes every golden question through
  the real query/answer/validation path. Current status is
  `BLOCKED_PENDING_INGEST`: 1 question passes structurally, 9 are blocked by
  pending primary sources, 0 fail, and 0 have a retrieval miss. Eleven of 15
  required primary-source appearances remain pending. There are currently
  zero human-approved eligible effects, so the runner reports zero eligible
  effect/data/evidence counts instead of inventing results.
- Future workflow journals record locator/draft prompt versions, whether AI
  actually ran in that attempt, and whether a prior artifact was reused. P11's
  verified v2 provenance was backfilled into its journal.
- Images remain excluded from extraction and analysis everywhere.

Next execution gate: continue the remaining representative pilot sources
through the resumable corpus journal, inspect v2 comparison quality before
scaling, perform explicit human reviews where the source supports them, rerun
all ten acceptance questions, and only then expand the same validated workflow
across the remaining corpus.

## P12 semantic-preservation checkpoint — 2026-07-18

This checkpoint supersedes the counts in the 2026-07-17 current execution
checkpoint. The final goal is still in progress.

- Corpus journal: 9 `COMPLETED`, 980 `PENDING`, 0 `FAILED`; completed results
  are 6 `NEEDS_REVIEW` and 3 `EXCLUDED`.
- P12 is
  `014.MSU-20S15-07 Result test AWF cooling time 4s,8s,10s_clean.xlsx`,
  current analysis `ANALYSIS-7BCA4D3F017F`.
- The v2 draft was rejected because it collapsed separately labelled scalar
  metrics, omitted source sample sizes, and could not represent the wide
  frequency-response matrix. It was never approved.
- Prompt compression reduced the deterministic source payload from roughly
  1.6 million to 172,480 characters without dropping any of the 5,080 source
  cells. Prompt v5 records provenance and validates all reference keys. The
  one rejected-draft repair changed only an unknown Arm reference
  (`cool_time_10s` to `cooling_time_10s`); a projection guard now rejects any
  repair that changes non-reference content.
- The cooling Study preserves 3 Arms with source sample sizes 100/156/122,
  22 distinct scalar Outcomes, and 66 Observations. It has 0 Comparisons and
  0 Effects because the workbook does not explicitly identify a valid control
  for those conditions.
- The frequency Study preserves 4 Arms, 4 measurement series, and exactly
  2,697 points: Spec 87 and 20000V/1800V/1600V 870 each. All 2,697 stored
  values match Capture v2, all source coordinates are distinct, axis values
  are complete, and the exact header/value/row-identity ranges are cited by
  12 EVD records.
- The three Spec comparisons remain `NEEDS_REVIEW`/confounded and have no
  Effects. No real Comparison has been human-approved; `review_decisions`
  remains empty.
- Re-import safely superseded the earlier unreviewed canonical analysis for
  the same revision, leaving one current analysis and stable public source
  identities. A reviewed or verified prior analysis would instead be
  preserved as `STALE`.
- Query/answer returns one relevant Study record, compact summaries for all
  four series, 12 exact citations, and no answer-eligible Effect. Repeated
  semantic labels are counted by source coordinates, so the 1600V series is
  correctly reported as 87 axes and 10 source rows. The unrelated
  same-model XRAY-only workbook is not returned for the frequency question.
- Golden acceptance is `BLOCKED_PENDING_INGEST`, not failed: 10 questions =
  2 structural pass, 8 pending-ingest, 0 fail; 15 primary-source appearances =
  5 represented, 10 pending, 0 retrieval miss. Eligible effect/data/evidence
  counts remain zero until explicit human approval.
- Verification: 100 scoped Python tests pass; Python compilation passes; WPF
  builds with 0 warnings and 0 errors; `verify-universal-db` passes 16/16
  stored universal-grid workbooks; SQLite integrity is `ok`; canonical FK
  errors are 0. The 66 pre-existing FK warnings are confined to legacy
  `analysis_*` compatibility tables.
- Images remain outside extraction and analysis. No WPF/Desktop application
  was launched, and the user's open Excel/WPF process was not touched.

Next gate: process the next representative source through the same resumable
workflow, inspect its exact scalar/comparison/series preservation, and expand
only after that per-workbook quality gate passes.

## P13 semantic-safety checkpoint — 2026-07-18

This checkpoint supersedes the P12 corpus and acceptance counts. The final
goal remains active.

- Corpus journal: 10 `COMPLETED`, 979 `PENDING`, 0 `FAILED`; completed
  results are 7 `NEEDS_REVIEW` and 3 `EXCLUDED`.
- P13 is
  `51. BRS-161014 Report test VP mold #4,#7,#8 change mold temperature and
  valcunizing agent 10% date 17.7.2024_clean.xlsx`, current analysis
  `ANALYSIS-B6586C6CDFB9`.
- The first P13 draft preserved source values but exposed four semantic
  defects: Excel percent fractions were stored on the wrong human `%` scale,
  raw Input/OK/NG counts could have been treated as continuous effects, fixed
  dry/cure/agent conditions and aligned 180-versus-190 mold comparisons were
  incomplete, and AVG columns inflated replicate counts.
- The pipeline now normalizes directly cited percent-formatted scalar cells
  to displayed percent units, requires count effects to have an explicit
  numerator/denominator, treats `sample_size` as denominator-only, and
  suppresses duplicate count-derived effects when the same source also
  supplies an explicit aligned rate outcome.
- Measurement points now carry `RAW` or `AGGREGATE` replicate roles.
  Aggregate values remain fully source-cited, but default range/average/axis/
  replicate summaries use RAW points only. Prompt v6 requires exact
  `aggregateReplicateRanges`, explicit count denominators, displayed percent
  scale, aligned same-entity comparisons, and all fixed conditions.
- P13 v6 reused the current Capture v2 revision, all 48 complete semantic
  chunks, and all 48 locator results. Locator AI calls were 0; only the Study
  draft was regenerated. Images were not extracted or analyzed.
- The current draft has 5 Studies, 27 Arms, 33 Outcomes, 197 scalar
  Observations, 27 measurement series, 13,311 measurement points, 24
  Comparisons, and 0 Effects. The former 10-point MASK series is not lost; its
  ten source values are represented as scalar observations. Total preserved
  numeric values remain 13,508.
- All 13,311 point values match their exact Capture v2 source cells. They are
  split into 10,962 RAW and 2,349 AGGREGATE points. Every one of the 2,349 AVG
  values matches the corresponding RAW row average; each ordinary series has
  87 axes, five RAW replicates and one AVG, while ST has two RAW replicates
  and one AVG.
- All 20 rate observations have explicit numerator/denominator pairs and
  correct percentage arithmetic. All 124 non-Input count observations have a
  denominator, while the 14 Input observations use `sample_size`. Numeric
  evidence validation reports no mismatch.
- Nine aligned same-mold 180-versus-190 draft Comparisons now exist across
  vision, function, and SPL/THD/IMP. Dry 150 C/1 hour, vulcanizing agent 10%,
  second cure, line, lot, and mold identities are preserved. Fourteen
  reference comparisons remain explicitly confounded and ten remain
  unassessed. Every Comparison is `NEEDS_REVIEW`, aggregation eligibility is
  0, Effects are 0, and `review_decisions` is still empty.
- Verification: 115 scoped Python tests pass; P13 semantic validation and DB
  integrity pass; canonical FK errors, orphan evidence links, invalid
  aggregation effects, point mismatches, rate arithmetic mismatches, and AVG
  arithmetic mismatches are all 0. `verify-universal-db` passes 16/16. WPF
  builds with 0 warnings and 0 errors and was not launched.
- Golden acceptance remains `BLOCKED_PENDING_INGEST`, but coverage improved:
  10 questions = 3 structural pass, 7 pending-ingest, 0 fail; 15 required
  primary-source appearances = 6 represented, 9 pending, 0 retrieval miss.
  Eligible effect/data/evidence counts remain 0 until explicit human approval.

Next gate: process P14 through the same resumable, image-free workflow and
apply the same per-workbook scalar, comparison, series, source-coordinate,
and review-safety audit before increasing throughput.

## GQ06 confounding-answer and P14 pre-import gates — 2026-07-18

- GQ06 previously returned only generic `NO_VALID_COMPARISON` text and
  attached unrelated SPL/THD/IMP series to a Function-NG relationship
  question. The cause was loss of Comparison context in `excludedRecords`,
  fallback to all measurement series when no outcome matched, and citation
  merging from the whole candidate instead of the selected series.
- Effectless comparisons now retain their validity, confounding,
  aggregation, design, matching, arm condition, direct EVD, and exact
  factor-difference gates. `CONFOUNDED_MULTI_FACTOR` is emitted only when two
  or more recorded factor values actually differ; baseline/held-constant
  metadata alone cannot create a false difference.
- The deterministic answer lists every control-to-compared factor value and
  states why no single-factor effect was calculated. Current GQ06 has 0
  eligible Effects, 14 confounded comparison explanations, 12 multi-factor
  classifications, 0 unrelated descriptive series, and 28 direct comparison
  citations instead of 91 broad citations.
- GQ06's three required behaviors now use declarative, non-AI acceptance
  assertions: complete confounded-factor preservation, required answer code,
  and maximum eligible Effect count 0. All three pass; the question itself is
  `PASS`, while overall acceptance remains `BLOCKED_PENDING_INGEST` because
  seven other questions still await pilot ingestion.
- P14 attempts remained fail-closed and imported no bad data. They exposed
  reusable draft defects: blank REF columns inside dense ranges, unsupported
  `BASELINE` arm role, invented Before/After roles for repeated unlabeled
  blocks, duplicate empty replicate identities, and omitted Fo/AVG values.
- Prompt v8 explicitly restricts arm roles, forbids inferred Before/After,
  preserves repeated blocks as source-backed strata with unchanged sample
  size, permits only data-bearing REF columns, requires dense numeric series,
  and requires unique raw replicate identities or valid series mappings for
  Air-leak/Fo values.
- Dense measurement series are now expanded read-only during DRAFT
  validation, so blank, malformed, and error cells fail before IMPORT.
  General contract/numeric failures now feed the rejected draft, exact
  validator error, and focused source packet into a constrained repair pass;
  unknown-reference-only repairs retain their stricter projection guard.
- The full service Python suite passes 199 tests. Images remain excluded and
  no WPF, Excel, server, or desktop application was launched.

Next gate: complete the P14 v8 draft using the existing Capture revision and
71 locator results, audit it against every source invariant, and import only
if every gate passes.

## P14 custom-format semantic correction and completed import — 2026-07-18

This section supersedes the interim P14 statements above. In particular, the
claim that the workbook's Before/After labels were invented is retracted for
this source.

- P14 is
  `38. MSU-L20S1507 DOE Air preasure-air leak_clean.xlsx`, current analysis
  `ANALYSIS-B35468F7B510`.
- The cells' stored values are `1..10`, but their source-authored Excel custom
  number formats display identities such as `18kPa #1_Before` and
  `18kPa #1_After`. Capture v2 had preserved `number_format` while exposing
  the raw number as `display_value_json`; the HTML/WPF renderer then displayed
  the raw value without applying the format. This format loss was a concrete
  cause of the Excel-versus-rendered-HTML semantic mismatch.
- Import now restores an evidence-linked formatted measurement identity only
  from the actual captured header cell and its custom number format. It does
  not authorize AI-written labels or locator summaries. The format can
  establish Before/After stage identity, but cannot create a CONTROL or
  BASELINE role or prove causality.
- The workbook supplies aligned formatted replicate identities for
  18/100/200 kPa Before and After. Nine paired draft Comparisons cover
  SPL/THD/IMP/Fo as available. They all remain
  `NEEDS_REVIEW`/`UNASSESSED`, aggregation eligibility is 0, Effects are 0,
  and there are no human review decisions.
- The 300 kPa After headers are source-authored but have no values. Their
  header-only arms are preserved without inventing a series or comparison.
  The malformed THD cell, `#REF!` IMP cells, and blank REF siblings are
  excluded; data-bearing REF values remain reference-only and never become a
  control.
- A standalone aggregate-series contract now preserves pooled AVG curves
  without duplicating them as specimens. Aggregate series must use
  `AVERAGE`, identify their RAW source series, share the same Study/Outcome,
  and match the arithmetic mean at every axis. All seven P14 AVG series pass
  this validation.
- P14 contains 5 Studies, 39 Arms, 9 Outcomes, 17 scalar Observations, 46
  measurement series (39 RAW and 7 AGGREGATE), 18,916 points (18,651 RAW and
  265 AGGREGATE), 9 Comparisons, and 0 Effects. Source sheet-coordinate
  uniqueness and source-value matching both pass.
- The v14 draft reused Capture revision
  `capture_revision_c1e1c775deaa07463430899c`, 71 locator results, the packet,
  and existing artifacts. Locator AI calls were 0; no recapture, relocation,
  image extraction, or image analysis occurred. Failed interim drafts were
  never imported.
- Corpus accounting is now 11 `COMPLETED`, 978 `PENDING`, 0 `FAILED`.
  Golden acceptance is 4 pass / 6 pending-ingest / 0 fail; GQ05 is now
  represented by P14, and all three automated required behaviors pass.
- Verification: the complete service Python suite passes 218 tests,
  `verify-universal-db` passes 16/16, SQLite integrity and canonical
  integrity pass, and the WPF project builds with 0 warnings and 0 errors.
  No application, Excel process, server, or preview was launched.
- Rendered HTML remains a supporting artifact and is not accepted as an
  Excel-faithful view. This gate fixes database semantics and identifies the
  number-format rendering loss; it does not claim that the existing HTML has
  been visually corrected.

Next gate: process P15 through the same one-workbook, image-free quality gate.
Do not increase throughput until its scalar values, comparison groups,
measurement series, formatted identities, aggregate arithmetic, evidence
coordinates, and review safety all pass.

## P15 source-audited DOE import completed — 2026-07-18

- P15 is
  `1. BRS-161014 Report test DOE sub 2 manual_clean.xlsx`, current analysis
  `ANALYSIS-8CAED455EB68`.
- Capture revision `capture_revision_53a1a1e20535ff386f3e7a54` contains two
  sheets, 345 non-empty captured cells, 16 formulas, and 148 merges. The
  packet is complete at 345/345 cells across nine chunks, with no missing or
  duplicate cell keys.
- The source is represented as four Studies rather than one mixed DOE:
  CM+B-PT Visual, Tension, CM+CP bonding/drying, and CM+B-PT bonding-line.
  Final counts are 18 Arms, 21 Outcomes, 90 scalar Observations, 8 RAW
  series, 64 points, 10 Comparisons, and 0 Effects.
- The original 86 source scalar observations are complete. Four additional
  Visual formula-derived rates retain `valueNumber=null`, the exact 0/8 or
  8/8 counts, sample size 8, formula-cell evidence, and deterministic
  `ratePpm`.
- Tension contains 32 physical specimens and two measured variables, so 64
  points are retained without inflating the sample size. Four actual bonding
  amount series use `mg`; four tension series use `kgf`. Every source value,
  axis identity, and coordinate matches Capture v2.
- The 12 uncached Tension MAX/MIN/AVG formulas are evidence lineage only.
  Their raw values permit later descriptive calculation, but the import does
  not claim a cached source result.
- CM+B-PT row 46 explicitly records Input 8, OK 0, Total NG 5, and three
  unclassified/unreconciled specimens. The source inconsistency remains a
  limitation and cannot be repaired by taking complements or imputing a
  category.
- The exact safe Comparison set is three adjacent Visual, three adjacent
  independent Tension, two CM+CP amount-only, and two amount-only within the
  explicitly changed bonding-line group. Comparisons spanning blank
  bonding-line cells are excluded because blank cannot establish a baseline
  or unchanged line. `In Spec` remains `TEST`.
- Every Comparison remains `NEEDS_REVIEW`/`UNASSESSED`, aggregation
  eligibility is 0, Effects are 0, and `review_decisions` is empty.
- The first quality-repair response was fail-closed because it changed
  deterministic `contentComplete=true` to false. The packet boolean and four
  numerator/denominator-derived `ratePpm` values were restored
  deterministically; all semantic content then passed the canonical,
  numeric-evidence, and strengthened P15 source gates. The general runner now
  restores packet coverage before validation while the direct validator
  continues to reject mismatches.
- Import reused the existing Capture, packet, all nine locator results, and
  the prompt-v17 draft. Import-stage locator AI and draft AI calls were both
  0; recapture and image work were 0.
- Corpus accounting is 12 `COMPLETED`, 977 `PENDING`, 0 `FAILED`. Golden
  acceptance is 4 pass / 6 pending-ingest / 0 fail; 7 of 15 required primary
  sources are represented and retrieval misses remain 0.
- Verification: 61 focused semantic/workflow/import tests and the full 222
  test service suite pass. `verify-universal-db` passes 16/16, P15 SQL
  integrity and source-value audits pass, and WPF builds with 0 warnings and
  0 errors. No application, Excel process, server, or preview was launched.

Next gate: process P16 through the same single-workbook, image-free audit.
Continue the representative pilots before increasing corpus concurrency.

## P16 source correction and first chunk benchmark (2026-07-18)

- P16
  `017.MSU-20S15-07 Result test sample waitting 2 day and check function_clean.xlsx`
  is current as `ANALYSIS-F0756E286A00`; the weak stale analysis
  `ANALYSIS-0D3FE4FD3695` was superseded.
- Its canonical preservation is 1 Study, 2 Arms, 22 Outcomes, 44 scalar
  Observations, 1 review-gated Comparison, and 0 Effects.
- Waiting has n=299 and exact main Function NG 66/299 =
  22.07357859531772%; Normal has n=920 and 236/920 =
  25.65217391304348%. The -3.57859531772576 percentage-point difference is
  descriptive only because matching, randomization, lot equivalence, and
  causal identification are not established.
- Prompt/import contract v18 rejects Excel raw fractions and rounded display
  percentages as canonical human-percent values. It accepts only the exact
  percent-formatted source-cell raw value multiplied by 100, and count cells
  cannot satisfy percentage evidence merely because they share a citation
  range.
- The existing complete Capture/packet/locator artifacts were reused. There
  was no recapture, regrouping, image extraction, or image analysis.
- Focused tests pass 62/62, universal DB verification passes 16/16, and the
  post-import source/SQL audit passes. Human approvals and eligible Effects
  remain 0.
- The next scaling gate is
  `pilot/corpus-benchmark-small-30-v1.json`: 30 deterministic pending files,
  each at most 1 MB, processed with three workbook workers and two locator
  workers. The benchmark must report completion, quarantine classes,
  integrity, and measured throughput before the remaining corpus is
  scheduled in size tiers.

## v23 semantic-safety correction and resumed 30-workbook gate (2026-07-18)

- The first 30-workbook run was intentionally stopped after source audit
  found semantic false passes in B03 and B05. Both analyses were quarantined
  as `STALE`, removed from answer-visible current results, and requeued.
- The shared contract is now provenance v23. It distinguishes source-authored
  conclusions from AI-derived descriptions, preserves factor quantity plus
  unit, separates quantitative and categorical outcomes, maps unsupported
  Normal labels to `REFERENCE`, and rejects representation-misaligned
  comparisons.
- A transactional quarantine command and corpus-journal downgrade prevent a
  quarantined analysis from remaining current or journal-complete. Correct
  same-source reimport reactivates the analysis and resolves the quarantine.
- B03 and B05 were reprocessed without recapture or image work. Independent
  source audit passed all 18/18 and 46/46 raw points respectively. Both remain
  `NEEDS_REVIEW`, with zero eligible Effects.
- Full Python verification passes 246/246 tests and universal DB verification
  passes 16/16.
- At 2026-07-18 14:52 +07, the v23 benchmark resumed with three workbook
  workers and two locator workers. Four corrected/passed workbooks are
  skipped; the other 26 failed or interrupted records are being retried.
- The honest full-corpus estimate remains approximately 30–45 hours for one
  pass and 36–72 hours including semantic audits and repair retries. Bulk
  expansion remains blocked on this 30-workbook quality gate.

## v24 repair run and staged drafting gate (2026-07-18)

- The v23 30-workbook run ended with 13 completed and 17 fail-closed records.
  All completed records have zero answer-eligible Effects.
- Current-contract audit quarantined B06 and B20 because embedded quantity
  tokens with unresolved units bypassed exact whole-cell validation.
- v24 now covers exact merge-anchor identity/header provenance, strict numeric
  text repair, ordered multi-cell conclusions, unit-catalog-independent
  quantity syntax, compound factors, horizontal/vertical series-axis policy,
  unsupported count-pair removal, categorical status rows, and source-grounded
  grouped reference roles.
- B24/B25 proved that monolithic drafting is not scalable: 9.6k captured cells
  and 32/33 chunks produced 65k–72k-token outputs and transport loss. Stable
  per-call artifact output paths are implemented.
- A staged-draft gate is now required before full-corpus expansion: include
  numeric continuation chunks in candidate-bearing sections, create bounded
  sheet/section fragments, persist and validate each fragment independently,
  consolidate deterministically without invented cross-fragment comparisons,
  prove exact chunk/source-cell ownership, and import only the fully validated
  final manifest. B24 and B25 are the acceptance fixtures.
- The non-large 17-record v24 retry started at 15:48 +07 while staged drafting
  is implemented in parallel. Images remain excluded.

## 2026-07-18 v24 source-content gate checkpoint

- The 17-record v24 retry ended at 12 pipeline completions and 5 fail-closed
  failures. Independent source audit accepted 7 completions and quarantined
  5 false passes that omitted quantitative panels or source conclusions.
- A new inverse source-content gate now requires every owned numeric result,
  aggregate/formula result, exact numeric design value, and locator-identified
  source conclusion to have a canonical representation before import.
- The gate was applied to 12 real v24 artifacts, not only synthetic fixtures:
  all 5 known omissions failed and all 7 source-audited safe controls passed.
  The only initial safe-control mismatch was an error-only SPL axis tail,
  independently verified as two isolated `numeric/#REF!` rows beyond the
  valid series and excluded by a narrow adjacent-row rule.
- Staged draft v1 is not an accepted scaling path. Review found that it
  structurally appended part studies instead of merging source-identical
  logical studies, could lose cross-part comparisons, and lacked a complete
  continuation registry, evidence allowlist, and resume provenance contract.
- Phase 3 therefore remains gated on staged draft v2: exact whole-request
  budget selection, source-identity registry, append-only fragments,
  owned/shared cell enforcement, deterministic conflict-failing merge,
  evidence-backed comparison intents, and exact resume hashes. B24/B25 must
  pass this gate before the 30-workbook retry and 989-workbook migration.

## 2026-07-18 staged-v2 completion and bounded semantic gate

- Staged draft v2 now enforces the exact prompt text/hash, bounded parallel
  fragments, source-order deterministic merging, exact owned/shared evidence
  ranges, bidirectional record/disposition equality, merged-disposition
  provenance, and full part/final validation on resume.
- Independent focused verification passed 88/88 tests: staged v2 9, workflow
  9, semantic AI 45, and content coverage 25. Python compilation passed for
  the six changed runtime modules.
- Full 989-workbook execution has not started. Before v25 and the 30-workbook
  quality gate, five bounded adversarial barriers must pass:
  semantic-label inverse coverage; source Result/Axis/Factor/Aggregate roles;
  field-specific count/range/comparison binding plus primary-cell ownership;
  locator-independent conclusion discovery; and legend versus real status-row
  geometry.
- Scaling remains fail-closed until those fixtures pass together with the
  real v24 matrix: 5/5 known omissions blocked and 7/7 audited controls passed.
  No image analysis, recapture, regrouping, WPF, Excel, or server launch is
  authorized for this checkpoint.

## 2026-07-18 bounded semantic gate completed

- The five semantic/content blind spots are now implemented and covered by
  adversarial fixtures, including actual-cell-only axis authorization and
  empty-range laundering rejection.
- Focused regression passes 212/212. The real v24 matrix passes its exact
  acceptance boundary: all five known omissions are blocked and all seven
  audited controls pass with zero unresolved quantitative, categorical,
  narrative, formula, semantic-label, or field-binding items.
- Phase 3 may now proceed to the v25 ten-workbook retry and then the
  30-workbook gate including the oversized staged-v2 fixtures B24/B25.
- The v25 ten-workbook retry started at 18:08 +07 with three workbook workers
  and two locator workers. All ten selected records are journal-RUNNING;
  images remain disabled. The full 989 migration has not started.

## 2026-07-18 full source capture completed; semantic migration remains gated

- The v25 ten-workbook retry ended at 2 completed and 8 fail-closed. No failed
  draft was imported.
- The deterministic source-preservation layer was decoupled from semantic AI.
  Four parallel OpenXML readers now extract workbooks while SQLite imports
  remain serialized in deterministic source order.
- The complete 989-workbook source corpus was captured from 18:32:33 through
  18:41:50 +07: 930 imported, 59 exact-SHA skips, zero failures, and no image
  extraction or analysis.
- Verification passed for all 989 current Capture v2 revisions with all 989
  source SHA-256 values checked. There are zero unfinished Capture runs,
  failed Capture items, missing canonical bridges, or duplicate current
  Capture revisions. SQLite quick-check is `ok`.
- This is completion of the lossless source/cell/merge/formula/style layer,
  not completion of 989 semantic Study analyses. The corpus semantic journal
  is currently 33 completed, 9 failed, and 947 pending.
- B15 was safely recovered without another AI call by splitting six
  comma-separated A1 unions into individual evidence objects. All 6,546
  required quantitative cells are covered and `ANALYSIS-683E8804F3FB` is
  current `NEEDS_REVIEW`.
- B07, B09, and B22 exposed subsequent fail-closed defects and remain blocked:
  an unsupported composite reference Arm, five uncovered F6:J6 values after
  comparison omission, and one uncovered `NG function!L7` value.
- B14 remains a genuine large-draft omission; B16 remains a genuine
  queryability failure because 163 numeric leaves are embedded in composite
  text rather than normalized numeric records.
- Monolithic selection now checks both exact prompt bytes and source-cell
  count; more than 2,000 cells selects staged-v2.
- Retry journals now snapshot prior attempts, reset current stages/results,
  invalidate downstream stages on re-execution, and clear terminal results on
  failure. When an old draft fails the current contract, unverified canonical
  analyses for that revision are fail-closed to `STALE`; verified or
  human-decided analyses remain protected.
- Next gate: resolve and re-audit the remaining v26 failures, rerun the
  representative 30-workbook gate including B24/B25, then start resumable
  semantic migration of the 947 pending workbooks.

## 2026-07-18 v27 deterministic recovery

- Three source-proven repairs now run before any new Study AI call when the
  rejected artifact is newer than a stale target: exact split-cell Arm
  identity restoration for B07, a normalized numeric height-axis Outcome plus
  three aligned RAW count series for B09, and one exact missing L7 percentage
  Outcome for B22.
- AI provenance is now marked only immediately before a real Codex subprocess,
  rather than when entering the monolithic draft stage.
- Focused semantic/workflow/repair regression passed 77/77; the final combined
  semantic/content/staged/workflow/repair/CLI regression passed 130/130 and
  all seven related Python modules compiled. The real
  three-workbook v27 retry completed 3/3 in 33.8 seconds with zero failures,
  zero image work, and `DRAFT.aiExecuted=false` in every workbook journal.
- All three canonical analyses were imported and verified as NEEDS_REVIEW;
  SQLite quick-check remains `ok`.
- The semantic corpus is now 36 completed, 6 failed, and 947 pending. The six
  fail-closed records are B04, B08, B14, B16, B24, and B25. These must pass
  before rerunning the full representative gate and scaling the pending 947.
- Remaining-gate audit:
  - B04 and B08 are bounded evidence repairs: remove an unsupported numeric
    claim embedded only in prose, preserve four missing semantic headers, and
    widen the short `Can use` conclusion to its exact two-line source.
  - B16 requires deterministic normalization of 154 matrix values and nine
    Sigma rates plus exclusion of three explicit IR identifiers.
  - B14 has 861 quantitative, 123 semantic, 20 categorical, 1,234 unresolved
    formula, and two binding gaps across 2,947 selected cells, so it requires
    a staged-v2 reconstruction rather than a patch.
  - B24/B25 select 9,638/9,656 cells. Default staged plans still contain
    finalized prompts over 400 KB. One-chunk parts produce safe 29/31-part
    plans with maxima of 353,267/389,814 bytes. No reusable fragments exist,
    so 60 new fragment AI calls are required.
  - The planner must permanently budget the finalized prompt including
    registry/shared anchors and preflight every part before starting workers.

## 2026-07-18 v28-v30 recovery and operational E2E audit

- B04 and B08 were recovered with source-bounded deterministic repairs. B04
  removes the unsupported numeric interpretation of prose-only `50pcs` while
  preserving exact context cells; B08 binds the conclusion to exact B25:B26.
  The real retry completed 2/2 with zero AI calls.
- B16 now expands four composite matrix outcomes into 14 scalar outcomes and
  154 observations, adds three Sigma-rate outcomes and nine observations, and
  excludes exactly the repeated IR identifiers E17/E19/E21 while retaining the
  first queryable IR field. The real retry completed 1/1 with zero AI calls.
- The semantic corpus is therefore 39 completed, 3 failed, and 947 pending.
  The only failures are B14, B24, and B25. SQLite quick-check remains `ok`.
- Staged-v2 now packs by the exact finalized prompt including the registry and
  shared anchors, preflights every part before any worker starts, uses stable
  UUID sibling transport files, and records AI execution only immediately
  before a real subprocess. The focused combined regression passed 124/124.
  Current exact plans are B24=28 parts (max 353,267 bytes) and B25=29 parts
  (max 389,814 bytes), with exact source ownership and no 400 KB violations.
- A pre-staged backup was created and verified at
  `outputs/universal-grid/backups/InputDataFinish.pre-staged-b24-b25-v30-20260718.sqlite`.
- The first real B24/B25 staged-v2 attempt ended before fragment generation or
  import. Codex rejected the transport response schema because nested object
  schemas were not strict (`additionalProperties:false`). Both workbooks
  remain fail-closed, zero fragment artifacts were accepted, and the canonical
  DB remains unchanged. The strict transport schema is being corrected before
  retry.
- Read-only end-to-end audit confirms that WPF Ask -> CLI evidence-answer ->
  query/answer -> EVD detail -> exact Excel range and WPF new-XLSX ingest ->
  resumable backend ingest are wired in code. Runtime UI execution remains
  intentionally unperformed.
- The operational database is not yet capable of the requested quantitative
  relationship answer: 165 comparisons and 73 effects exist, but zero are
  currently verified and answer-eligible; 120 items remain in the review
  queue. The VP+CD assembly/FUNCTION NG example therefore returns
  `INSUFFICIENT_COMPARISON`, as required by the unsupported-claim gate.
- Arbitrary-domain retrieval is still partial rather than complete semantic
  generalization. The database has only 7 canonical concepts and 28 aliases,
  while 601 schema candidates remain open. The next retrieval vertical slice
  must add fail-closed human approval/merge of candidate terms into canonical
  concepts and aliases, make related-workbook discovery use the same aliases,
  and expose actual factor changes in eligible effect answers.

## 2026-07-18 formula-safe B14 and concept-curation vertical slices

- A deterministic formula overlay now evaluates a deliberately restricted A1
  arithmetic grammar without writing derived values back to Capture v2. It
  checks source revision/content, exact dependencies, cycles, unsupported
  syntax, non-finite values, and overlay checksums fail-closed.
- Real B14 read-only projection proves 1,234 uncached formulas = 1,149 numeric
  derivations + 85 `#DIV/0!` results. Unresolved formulas fall from 1,234 to
  zero, required numeric cells rise from 935 to 2,084, and all 1,234 Capture
  formula caches remain NULL. The projected staged plan is 24 parts with a
  maximum finalized prompt of 357,901 bytes.
- Workflow/CLI integration is opt-in through `derive_formula_values` /
  `--derive-formula-values`. It atomically stores and revalidates the overlay,
  projects only a deep-copy semantic packet, binds overlay provenance through
  draft validation and import, and records `captureMutated=false`. The focused
  formula/import/workflow/CLI regression passed 100/100. Actual B14 AI/import
  remains pending.
- A new fail-closed concept-curation backend and additive migration now support
  candidate/concept listing plus atomic human CREATE, MERGE, and REJECT
  decisions. Immutable resolution and alias-approval history record the
  request hash, candidate/action/concept/alias snapshots, reviewer, note, and
  timestamp. Exact replay is idempotent; conflicting replay, UNIT misuse, kind
  mismatch, empty values, inactive targets, and alias ownership conflicts are
  rejected.
- Four JSON CLI commands expose this boundary: `concept-candidates`,
  `concept-list`, `concept-resolve`, and `concept-alias-upsert`. Related-study
  profiles now consume the same aliases as evidence query while explicitly
  retaining that similarity is neither relationship nor causal evidence.
  Sol re-ran the expanded schema/curation/related/query/CLI/import suites:
  95/95 passed.
- Eligible effect answers now carry each exact recorded factor difference and
  control/compared condition into the structured result and Korean rendering,
  for example `Bonding amount: A -> B`, before the stored effect estimate.
  Missing values are displayed as unrecorded rather than invented. Combined
  query/answer regression passed 36/36.
- The real operational DB has not yet received
  `canonical-concept-curation-v1`; migration must be a separate backed-up,
  verified operator step and must never be triggered implicitly by the WPF
  review screen. No real candidate was approved or rejected in this
  checkpoint.

## 2026-07-18 representative-30 gate correction

- The WPF concept/alias normalization tab is now implemented between human
  review and new-workbook ingest. It lists up to 10,000 OPEN candidates,
  filters same-kind ACTIVE concepts, requires reviewer and note, confirms
  irreversible CREATE/MERGE/REJECT decisions with exact IDs and values, never
  mutates the list optimistically, distinguishes save success from refresh
  failure, and surfaces idempotent replay. The focused WPF project build passes
  with zero warnings and zero errors; the app was not launched.
- A fresh read-only audit proved that journal `COMPLETED` is not equivalent to
  the current quality contract. Of the representative B01-B30 set, only 18
  currently pass end-to-end source/manifest/content/evidence validation:
  B01/B02/B03/B04/B07/B08/B09/B10/B13/B15/B16/B18/B20/B21/B22/B26/B28/B29.
- Nine answer-visible NEEDS_REVIEW analyses are false passes and must be
  quarantined/reprocessed:
  - B05/B19/B27: insufficient source wording for a declared REFERENCE role.
  - B06: missing quantitative `Test (2)!C23`.
  - B11: 298 quantitative omissions plus a single selected chunk whose
    finalized prompt exceeds 400 KB before AI.
  - B12: missing semantic labels `201507!J10:J11`.
  - B17: 586 quantitative omissions, conclusions at `Report!C109,C111`, and
    two semantic omissions.
  - B23: missing categorical status `201507!F28`.
  - B30: missing quantitative `Test!C17,C19,C21,C23`.
- B14 remains failed pending formula-safe actual execution; B24/B25 are
  currently RUNNING in staged-v31. Therefore the representative gate is not
  30/30 and the pending 947 migration remains prohibited.
- All representative sources still match their journal SHA and current
  Capture v2 revision; SQLite quick-check is `ok`, canonical knowledge FK
  errors are zero, invalid aggregation effects are zero, and orphan evidence
  links are zero. The whole legacy database has an existing exact
  `foreign_key_check` baseline of 66 rows (48 analysis_evidence,
  12 conclusions, 6 review_items); future checkpoints must prove this baseline
  is unchanged rather than incorrectly claiming global FK zero.
- Before the 947 corpus launch: finish and re-audit B24/B25; recover the nine
  false passes and B14; prove the current contract 30/30; freeze exact runtime
  file hashes; create and verify an immutable backup; generate an exact
  PENDING-only 947 manifest hash; then run a 25-50-workbook canary followed by
  size-tiered resumable batches. Formula-free, supported-formula, and
  unsupported-formula workbooks must be separated so formula mode cannot
  accidentally reprocess completed records or fail an entire bulk tier.

## 2026-07-18 21:24 +07 usage-pause checkpoint

- The user explicitly requested a pause because Codex usage is unavailable.
  Both active implementation workers were interrupted. The staged-v31 retry
  process tree (supervisor PID 42984, Python PID 11892, console PID 43920) was
  stopped; none of those PIDs remained afterward. No app, Excel instance,
  WPF window, server, or image-analysis path was launched.
- The generic B11 scaling blocker is implemented. A source-contiguous
  within-chunk segment retains the original chunk/locator identity, exact
  ordered and unique `sourceCellKey` ownership, bounded shared
  merged/context anchors, logical-study registry scope, segment-bound
  part/plan/provenance/resume identity, and an atomic-cell fail-closed path.
  The real B11 read-only plan has 3,323 owned cells in 17 parts, 15 segmented
  parts, and a maximum finalized prompt of 373,061/400,000 bytes. Exact
  reconstruction, ownership, locator identity, and ordering passed. The
  staged/workflow regression passed 30/30 and compile passed.
- Generic content-coverage corrections are implemented for:
  - B05: a farther merged `Total` family header no longer converts its
    nearest `Position 1` raw value into an aggregate.
  - B06: a complete increasing ordinal sequence remains excluded across
    merged record-row gaps.
  - B12: `Normal` under a captured merged factor matrix such as
    `S-MG -> Spec/Supplier` is a factor level without making bare
    `Spec`/`Supplier` global factor requirements.
  - B23: a vertically merged sparse `OK` column heading with no data below
    is not a categorical observation.
  - B30: exact multi-value HEADER-series row identities are covered as
    replicate identities rather than dropped.
- The content-coverage module passed 47/47 tests; content/workflow/import
  passed 111/111. Real saved-artifact revalidation now passes:
  B05 quantitative 72/72, semantic 4/4, categorical 4/4;
  B06 quantitative 112/112;
  B12 quantitative 69/69 and semantic 14/14;
  B23 quantitative 124/124, semantic 9/9, categorical 0;
  B30 quantitative 64/64. These figures prove content coverage only. B05
  still needs exact REFERENCE Arm normalization before full strict-contract
  acceptance.
- The completed read-only audits establish the remaining real repairs:
  B05/B19 require exact Test/Normal Arm identities while retaining every
  other identity component as factor values; B27 requires an exact Type
  factor for identities 1-4 plus safe factor-level handling; B17 genuinely
  omitted the `Report!C15:O107` 45-lot by 12-outcome table (540
  observations/586 previously uncovered quantitative cells) and conclusion
  context at C109:C112. B17 has an in-memory deterministic no-AI proof for
  quantitative 1117/1117, semantic 2/2, categorical 44/44, narrative 6/6,
  binding errors 0. The two implementation workers were interrupted before
  adding their Arm/B17 repair modules, so these remain pending.
- staged-v31 first pass preserved B24 23/28 accepted parts and B25 24/29.
  The exact missing part indexes at that checkpoint were B24
  `[2,3,20,27,28]` and B25 `[2,3,11,21,29]`. The retry was started and then
  stopped on the user's pause request. A post-stop inventory confirmed the
  counts and missing indexes were unchanged and no in-flight artifact
  remained. The corpus journal still says RUNNING for these two sources
  because the process was intentionally stopped; resume must reconcile that
  state rather than treating it as a live writer. Re-run
  the same `run_corpus_staged_b24_b25_v31.ps1` only after proving that no
  writer is live. It will reuse validated fragments and call only missing
  parts.
- Resume order: inventory B24/B25 fragments and safely resume v31; finish
  exact Arm and B17 deterministic repair modules with focused regression;
  reprocess/re-audit the nine false passes; run formula-safe B14; prove the
  representative set 30/30; only then freeze runtime hashes, back up and
  inspect the DB, migrate concept curation, generate the exact pending-947
  manifest, and start canary/tiered corpus migration.
