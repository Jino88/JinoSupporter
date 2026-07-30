# InferenceDataAIService session handoff — 2026-07-16

## User goal

New Excel files must be handled inside WPF without manual grouping:

1. capture the complete workbook structure/content (images are out of scope),
2. have AI group genuinely similar forms,
3. define group-specific extraction and dashboard rules,
4. render only after grouping is sufficiently fine and validated.

The user explicitly rejects a large `fallback`/`needs-review` bucket and rejects rendering from coarse top-level categories.

## Current batch and persisted artifacts

- Batch: `outputs/batches/wpf-list-preview`
- Input root: `D:\000. MyWorks\test\result\InputDataFinish`
- 990 items total: 988 `SCANNED`, 1 truncated/unreadable, 1 non-workbook.
- Complete capture is persisted in `numeric-capture.sqlite` (text, numeric/date/formula values, section candidates, merges).
- Inventory artifacts persist: `document-inventory.jsonl`, `layout-signatures.json`, `layout-clusters.json`, `layout-cluster-summary.json`, `layout-semantic-summary.json`.
- Source Excel files were never modified.

## What was completed

- WPF argument batch runs populate the left file list and show pipeline columns.
- WPF supports source Excel open (double click/context menu) and external result HTML open.
- WPF now shows a pipeline state banner; current stages are capture, AI grouping, and rendering.
- Numeric capture was expanded to full-document grid data rather than only simple facts.
- Group renderer has progress logging every 10 workbooks and caps raw structural tables at 250 rows / 32 columns to avoid giant HTML stalls.
- Codex invocation now sends the prompt through standard input (`codex exec ... -`) instead of a command-line prompt, avoiding Windows command-length failures.
- Build command repeatedly passed after changes:
  `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore`

## Important failures and lessons

1. Old strict layout signatures included absolute coordinates, merge count, dimensions and raw headers. This produced 969 signatures for 989 captures, so AI put almost everything in fallback.
2. The first broad semantic pass produced only 6 top-level buckets. It is NOT a final renderer grouping and must not be rendered:
   - acoustic 634, quality 159, measurement 113, general 66, process 13, tension 4.
3. Earlier `MaterializeFileAssignments` silently sent unmatched files to fallback. This was changed to fail rather than silently hide unassigned normal files.
4. A temporary fallback invariant was made contradictory with the AI prompt. WPF now injects an empty exception fallback if an otherwise complete plan does not contain one.
5. Codex CLI cannot reliably be asked to discover large local JSON artifacts: it attempted shell, web, GitHub, and printed large JSON. Do not return to that design.
6. Passing hundreds of cluster ids as a command-line prompt exceeded Windows argument length. Standard input fixes the transport limit, but the current semantic-summary prompt is still too coarse for final grouping.
7. Rendering started once from the six top-level groups; it was stopped and generated `group-reports`/render indexes were deleted. Do not render until second-level grouping validation passes.

## Current code state

Key modified files:

- `InferenceDataAIService.Wpf/MainWindow.xaml`
- `InferenceDataAIService.Wpf/MainWindow.xaml.cs`
- `InferenceDataAIService.Wpf/DocumentInventoryEngine.cs`
- `InferenceDataAIService.Wpf/GroupCatalogRendererEngine.cs`
- `InferenceDataAIService.Wpf/DashboardHtmlRenderer.cs`
- `group-plan.schema.json`

The latest run was started after changing `SemanticCategory` to append a section-count/table-candidate pattern. Its result must be inspected before any rendering. It may still be too coarse; the final approach needs category-internal AI refinement.

## Required next implementation (do not skip)

Implement a real two-level grouping flow before launching renderer:

1. Use top-level semantic categories only to partition work.
2. For each top-level category, construct compact batches of cluster summaries containing:
   cluster id, section candidate sequence, normalized title/header facets, table count/shape, 1–3 representative paths.
3. Invoke Sol xhigh per compact batch through stdin. The response must assign every cluster in that batch to a named second-level group. Do not rely on the model reading local files.
4. Merge the batch results. Validate:
   - no overlap;
   - every `ASSIGNABLE` file assigned;
   - fallback only `EMPTY_LAYOUT` / `CAPTURE_INCOMPLETE`;
   - fallback <= 1.5% of scanned items;
   - registered renderer key only;
   - groups sharing one renderer still have compatible component recipes.
5. If validation fails, create a batch-local validation report and feed only failing/mixed clusters into a bounded retry. Never render a failed plan.
6. Only then materialize `group-catalog.json` and start `GroupCatalogRendererEngine`.

## Restart checkpoint (2026-07-16)

- The WPF process launched with `--ai-group-analysis-batch wpf-list-preview` and its direct Codex child were deliberately stopped for the session restart. Older unrelated Codex processes were not touched.
- Keep and reuse the persisted capture database and inventory artifacts. A restart does **not** need to re-open all source workbooks unless the source input changes or the capture schema changes.
- The stopped run still used one coarse semantic-category prompt rather than the required category-internal cluster-batch refinement. It did not produce a trusted plan and no renderer was started from it.
- Current state is **not complete**: no validated fine-grained catalog exists and no trustworthy rendered report set exists.

## User-facing constraints

- Do not ask the user to manually group files.
- Do not claim completion merely because a plan or HTML exists.
- Actively monitor WPF/CLI after launch: verify plan, catalog, validation counts, renderer start, first output, and final index.
- Tell the user only when the validated grouping and renderer completion conditions are actually met.
- Update Obsidian upon meaningful completed milestones.

## Implementation milestone (2026-07-16, two-level pipeline complete)

- Implemented `TwoLevelGroupAnalysisEngine`. It reuses `numeric-capture.sqlite` and never re-opens source Excel workbooks during grouping.
- `DocumentInventoryEngine` now emits the compact second-level evidence contract: top-level category, cluster id, section-candidate sequence, normalized title/header facets, table count/shape, and at most two representative paths.
- The engine partitions only by top-level category, submits bounded cluster batches to `gpt-5.6-sol` with `model_reasoning_effort=xhigh` through standard input, and uses a JSON response schema (`group-refinement.schema.json`). The model is not asked to discover local files.
- Every compact response is validated for complete one-time cluster coverage, stable ids, exact component recipes, and a registered category-specific renderer. A failed compact batch writes `group-refinement/*.validation.json` and retries only that batch once with the validation errors.
- The merged plan validates no overlap, all `ASSIGNABLE` files covered, exception-only fallback, fallback <= 1.5% of scanned files, registered renderer keys, and renderer/component-recipe compatibility. Scanner-proven empty/incomplete layouts are emitted as named structural exception groups rather than one large fallback bucket (the current inventory has 19 `EMPTY_LAYOUT` items, which would otherwise exceed 1.5%). It writes `group-validation-report.json`; failed validation blocks both catalog materialization and rendering.
- WPF now invalidates old coarse plans while preserving the capture DB, runs the two-level engine, and starts `GroupCatalogRendererEngine` only after validation passes. The renderer independently requires the validated two-level plan/report and clears old report pages before a new render.
- Verification passed: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` (0 warnings, 0 errors).
- Runtime grouping/rendering was intentionally not launched in this implementation session, so the current persisted batch still has no newly validated final catalog or completed report set. On the next explicit WPF/CLI run, monitor the validation report, catalog, renderer start, first report, and final render index.

## Renderer contract correction (2026-07-16, code complete; rerender intentionally withheld)

- The completed two-level grouping run produced a valid catalog for 988 `SCANNED` workbooks: 969 assignable files, 19 named `EMPTY_LAYOUT` exceptions, 256 second-level groups, 16 compact AI batches, zero fallback assignments, and zero validation errors. The persisted catalog and capture DB must be reused; never recapture Excel just to render.
- A subsequent renderer inspection proved that the previous `GroupCatalogRendererEngine` discarded `componentRecipe`, `extractionRule`, and `htmlRule` while loading the catalog. It selected every `numeric_table_candidates` row only from `rendererKey`, which made the dashboard a raw source-table list instead of the AI group's component-defined report.
- The renderer now preserves the full catalog group contract and materializes actual `numeric-capture.sqlite` candidates by ordered component recipe. Each primary dashboard contains the validated recipe, important-data/extraction/HTML contracts, component-specific table and numeric-cell counts, source sheet/range evidence, and merge-aware captured value grids. Uncontracted or missing candidates are explicit `CONTRACT_MISMATCH`/review findings; they are never silently omitted.
- A non-destructive preflight now compares the ordered distinct candidate type sequence for every assigned workbook against its catalog recipe before deleting any previous report. The exception-only `EMPTY_LAYOUT`/`CAPTURE_INCOMPLETE` groups remain structural review pages.
- Read-only verification against `wpf-list-preview` passed: all 969 assignable workbooks match their catalog recipe in exact order, with zero missing capture rows and zero mismatches. Representative checks confirmed both a three-component measurement/defect workbook (2/1/2 captured tables; 8/24/54 numeric cells) and a repeated-defect workbook (1/1 captured tables; 53/54 numeric cells).
- Verification passed after the correction: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` (0 warnings, 0 errors). WPF was stopped and **no corrected reports were rendered**. The existing `group-reports` and render indexes are old raw-list output and must not be treated as acceptance evidence.

## Component-contract render completion (2026-07-16)

- The corrected WPF renderer was run with only `--ai-group-analysis-batch wpf-list-preview`. It detected the already validated catalog and **did not recapture or regroup** any Excel workbook.
- Runtime preflight passed for all 969 assignable workbooks before report cleanup. It then rendered 989 captured workbook records: 969 `COMPONENTS_RENDERED`, 20 `STRUCTURE_READY` scanner exceptions, and 0 `CONTRACT_MISMATCH`.
- Final artifacts are `group-render-index.json` and `group-render-index.html`, both reporting `group-catalog-renderer-v2-component-contract`. The index has 989 primary report rows, all primary report paths exist, and `group-reports` contains 5,956 HTML pages including component review pages.
- The representative three-component report `1. BRS-161014 Report test reduce weight glue process ass'y VP-CD 2024.1.5_clean.xlsx` records `NUMERIC_TABLE_UNCLASSIFIED=2`, `MEASUREMENT_SUMMARY_NUMERIC_TABLE=1`, and `DEFECT_RATE_NUMERIC_TABLE=2`. Its dashboard exposes recipe/extraction contracts and actual capture evidence: `Report!D15:D18`, `Report!H15:I18`, `Report!D23:O26`, `Report!D33:M36`, and `Report!E53:N56`.
- The renderer log records: component-contract preflight passed, 989-workbook render started, then 969 component-contract renders completed. No startup failure was written after the old 20:13 renderer error. The WPF process is closed after verification.

## Context-grid rendering correction and verified completion (2026-07-17)

- Root cause: a numeric table candidate bounded only the numeric island. The prior renderer emitted that rectangle as a generic table, discarding the nearby original title rows, left row-label gutter, multi-row/group headers, and merges that extended outside the candidate. This produced generic headers such as `Column N` and detached `Gap NG rate` values.
- Added the shared persisted-capture contract `InferenceDataAIService.Wpf/CapturedTableContext.cs`. It restores each candidate's nearby title band, label gutter, fully contained merge ranges, header labels, and inherited logical-row facets from `numeric-capture.sqlite`. It is used by both the HTML renderer and `DocumentInventoryEngine`, so future compact AI evidence contains logical row/merge context as well.
- `GroupCatalogRendererEngine` now uses the global `group-catalog-renderer-v3-context-grid` path for every catalog profile and structural exception page. `DashboardHtmlRenderer` and the detail-page shell both style title/header/label grid cells consistently. The renderer does not recapture Excel or regroup the already validated batch.
- Rates are rendered as `1.16%` rather than raw decimal values. Since the persisted existing capture does not store Excel number-format/style metadata, this uses only captured rate headers and textual/numeric values. The display rule is conservative: an immediately following unlabeled row is hidden only when it has at least two 0..1 percentage values and its preceding numeric row has at least two values including one above 1. Merged Gap metrics and independently labeled/rate-only tables are retained.
- Build verification passed: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` (0 warnings, 0 errors).
- Full WPF render was run with `--ai-group-analysis-batch wpf-list-preview`, reusing the existing capture and validated catalog. Preflight completed and the final index reports 989 workbooks: 969 `COMPONENTS_RENDERED`, 20 `STRUCTURE_READY`, 0 `CONTRACT_MISMATCH`; `group-reports` contains 5,956 HTML files. The 18 generated HTML pages carrying the auxiliary-percentage notice demonstrate conservative suppression in affected source sections.
- Representative evidence: `outputs/batches/wpf-list-preview/group-reports/234c909c0174d0f0b76f.html` for `015.MSU-20S15-07 GMI -Result check NG function type no reason 201507 improve TF_clean.xlsx`. It preserves `RESULT CHECK FUNCTION OF LOT TEST USE VP +CD C2`, the merged `Gap NG rate` header (`rowspan=2`, `colspan=2`), both merged Gap pairs, and `1.16%` / `2.60%`; the raw values `0.01161198418160932` / `0.02595061971629173` and generic `Column 16` are absent.

## Context-grid blank-axis compaction (2026-07-17, v4 verified)

- Root cause: v3 correctly restored the broader source context, but then rendered every coordinate of `TopRow..BottomRow × LeftColumn..RightColumn`. Fully empty spacer rows/columns and title/header trailing cells became repeated empty `<th>`/`<td>` slots.
- `GroupCatalogRendererEngine` now globally compacts only wholly empty rows and columns. It retains ordinary blank cells inside a retained data grid, so meaningful sparse measurements are not erased. A non-empty merged cell keeps all visible axes it covers; its `rowspan`/`colspan` are recomputed from those visible axes, preventing repeated merge-covered blanks or invalid zero spans. For title/header bands, only trailing cells after the final meaningful cell are omitted; interior alignment cells and the actual hierarchy remain.
- Read-only representative assertions passed before rerender. In `005.MSU-L20S15-07 REPORT CHECK VP CD Method to Improve Tension_250119_1778470541_clean.xlsx`, `Sheet1!F10:Z22` has content on rows `10-14,19-22` and no capture content on rows `15-18`, while data extends through `Z`. In `006.MSU - 20S15-07 REPORT TEST DOE LED UV_1778470544_clean.xlsx`, the title band merges are `B25:G25` and `H25:M25`, whereas data legitimately extends through merged `U:V` cells; this distinguishes header-tail removal from data loss.
- Verification passed: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` (0 warnings, 0 errors). Full WPF render reused the same capture DB and validated catalog only. Final `group-catalog-renderer-v4-compacted-context-grid` index: 989 workbooks, 969 `COMPONENTS_RENDERED`, 20 `STRUCTURE_READY`, 0 `CONTRACT_MISMATCH`.
- Final HTML assertions: the `Sheet1!F10:Z22` representative is 9 rendered rows (not 13), has no spacer-only row, and keeps the three `TEST` merged headers. The `Sheet1!B25:V32` representative has exactly two cells in its first title row (`TEST LED UV`, `NORMAL LED UV`) while retaining the merged total-rate data tail. The earlier GAP representative still retains its merged header and `1.16%` / `2.60%`. No HTML report contains `rowspan='0'` or `colspan='0'`. WPF was closed after verification.

## Serial/index support-column suppression (2026-07-17, v5 verified)

- Root cause: a one-column `No` candidate was simultaneously rendered as `NUMERIC_TABLE_UNCLASSIFIED` even when it was only the serial-number gutter of a nearby primary table. This created a duplicate standalone `1, 2, ...` panel and an inaccurate captured-table summary.
- `GroupCatalogRendererEngine` now marks a candidate as support-only only when every conservative signal agrees: it has a recognized index header (`No`, `No#`, `S/N`, `STT`, `Index`, or `Seq`), five or more non-negative integer values, at least 75% sequential/reset transitions, and a nearby overlapping multi-column primary candidate to its right. Independent single-column measurements and logs remain renderable.
- Support-only candidates are retained in the context grid of their primary table but are excluded from standalone cards, component counts, review detail pages, and dashboard captured-data totals. The validated component recipe/preflight is unchanged; no Excel capture or AI grouping was rerun.
- Representative proof: `01. MSU-L20S15-07DT Report test check lot material CD  -11.4.2025_1778470667_clean.xlsx`, report `outputs/batches/wpf-list-preview/group-reports/110c248aa31bb51e3b1d.html`. Its `11.04!C14:C113` `No` gutter is no longer a standalone panel (`NUMERIC_TABLE_UNCLASSIFIED=0`), while the primary `G14:L113` defect table preserves the `No` context column and has `DEFECT_RATE_NUMERIC_TABLE=1`.
- Engine-level full traversal found 182 support-only candidates under the v5 rule. All 182 corresponding primary reports exist and none contains its suppressed `numeric candidate ...` standalone marker; the representative assertion also confirmed no `Captured 1 table(s), 50 numeric cell(s)` support summary leak.
- Verification passed: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` (0 warnings, 0 errors). Full WPF rerender reused the same `numeric-capture.sqlite` and validated catalog: `group-catalog-renderer-v5-support-index-suppression`, 989 workbooks, 969 `COMPONENTS_RENDERED`, 20 `STRUCTURE_READY`, 0 `CONTRACT_MISMATCH`, and 5,774 report HTML files. WPF PID 40536 was closed after verification.

## Session restart checkpoint (2026-07-17, saved state)

- The trusted persisted batch remains `outputs/batches/wpf-list-preview`. It has a valid two-level catalog, `numeric-capture.sqlite`, `group-validation-report.json`, `group-render-index.json`, and `group-render-index.html`. Do **not** recapture Excel or rerun AI grouping merely to inspect or rerender this batch.
- The current accepted renderer is `group-catalog-renderer-v5-support-index-suppression`: 989 workbook reports, 969 `COMPONENTS_RENDERED`, 20 `STRUCTURE_READY`, 0 `CONTRACT_MISMATCH`, and 5,774 HTML files. It includes v3 merge/title/label context, v4 blank-axis compaction, and v5 serial-column suppression.
- The WPF now has a left-side `기존 학습 결과` tab. It discovers validated batches, lists every stored rendered report, opens the selected report in the right pane, and supplies right-click actions to open the source Excel or report HTML in a new window. This is read-only browsing and must not start capture/grouping.
- Current manually opened WPF viewer PID is `38088`; it has no AI/Codex child process and does not mutate the batch. It can be closed or reopened normally after restarting the session.
- Merge semantics required for later Ask AI are persisted: `captured_merge_ranges` retains range and anchor values, while `CapturedTableContext` derives logical inherited row facets used by both renderer and compact AI evidence. Example GAP pair anchors display as 1.16%/2.60% rather than raw decimals. Limitation: the existing capture DB does not retain native Excel number formats/styles; percentage display is inferred conservatively from stored header/value evidence.
- The worktree contains extensive pre-existing user changes outside this service. Preserve them; only inspect or modify files required for the current task.
- Last scoped verification: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` passed with 0 warnings and 0 errors.

## Full Excel-to-HTML audit (2026-07-17, analysis only; v5 output rejected)

- The user rejected the current rendered HTML quality and requested root-cause analysis before any further correction. No renderer, capture, grouping, catalog, or HTML artifact was changed during this audit.
- The user confirmed that the current `.xlsx` files are not DRM-protected. All 989 indexed Excel/primary-HTML pairs were therefore read directly with `openpyxl` and compared in three read-only chunks (330/330/329 pairs). No recapture, AI regrouping, or rerender was run.
- Overall source coverage: 1,502,519 non-empty Excel cells were examined; 525,444 (34.97%) lie outside every rendered evidence range. This is a structural omission metric, not merely a CSS difference.
- Media loss: 831/989 workbooks contain 9,937 embedded images, while all 989 primary HTML reports contain zero `img`, `svg`, or `canvas` elements. The corpus contains no Excel chart objects, so embedded photographs/screenshots are the current media-loss issue.
- Visual-layout loss: 982 workbooks use multiple cell styles, 973 use explicit row heights, and 978 use explicit column widths. The source corpus contains 79,574 cell-style definitions (median 74 per workbook), but all component-rendered reports use the same generic dashboard/table CSS and do not carry cell-level font, fill, border, width, or height mappings.
- Candidate quality: `numeric-capture.sqlite` contains 4,967 numeric candidates; 3,735 (75.2%) are LOW-confidence `NUMERIC_TABLE_UNCLASSIFIED`. Numeric regions are split/merged by numeric row/column gaps, not by the visual table boundary, photos, borders, or report semantics.
- Repetition/fragmentation: the 989 primary reports contain 4,778 source cards (4.83 per workbook on average) plus 4,785 detail-review HTML files. Across the three chunks, roughly half or more of the workbooks showed highly overlapping evidence ranges, repeated table text, or one cohesive source table fragmented into candidate-specific cards. A single workbook reached 65 cards.
- Hard cropping: the context resolver caps data at 250 rows, 32 candidate columns, and 40 rendered context columns. There are 177 over-limit candidates in 74 workbooks. Confirmed examples include candidates ending at rows 392/626/1000 rendered only through row 251, and an `A:ZK` raw-data surface reduced to approximately `A:AF`.
- Non-numeric-sheet loss: 195 captured sheets in 151 workbooks have no numeric candidate. Seventy of those sheets are non-empty and collectively contain 1,472 text cells, 42 dates, 46 formulas, and 1,292 merge ranges, yet the candidate-driven renderer has no normal path for them. Image/text-driven defect reports can therefore become empty `STRUCTURE_READY` shells.
- Formula loss: 98 workbooks contain 14,087 formulas. The persisted formula rows have no numeric cached values in the current capture, and the source-grid renderer does not read `formula_cells`. Formula-based totals/rates/analysis therefore disappear unless they also exist as independent captured constants.
- Merge/coordinate loss: merge handling is limited to merges fully contained in a selected context. The chunk audits found widespread merge anchors outside evidence and, in the middle chunk, 6,068 merge structures that were inside evidence but did not match the resulting HTML spans. Blank-axis and support-index compaction also change source coordinates; strict source-cell-to-HTML position comparison in that chunk disagreed for 52.6% of in-evidence cells.
- Percentage distortion: source number formats are not persisted. Any rendered column whose inferred header contains `RATE` is formatted by multiplying its value by 100. Confirmed ordinary quantities such as `26` and `20` became `2600.00%` and `2000.00%`. Pattern-based auxiliary-percentage-row suppression also removes source rows in six current primary reports.
- Group validation is structural, not semantic or visual. `isValid=true`, `COMPONENTS_RENDERED`, and zero contract mismatches prove assignment coverage and ordered-distinct candidate-type compatibility only. They do not prove that the form group is correct, that Excel meaning is preserved, or that the HTML resembles the workbook.
- The catalog is also over-fragmented: 969 assignable files produced 256 second-level groups; among 259 assigned group ids including exceptions, 108 are singletons and 190 have at most three members. Nevertheless, all normal groups route through only six top-level renderer keys and ultimately the same `ComponentDashboard`/shared CSS shell.
- `importantData`, `extractionRule`, and `htmlRule` are displayed as `REVIEW_REQUIRED` prose. Only the ordered-distinct candidate-type recipe affects materialization; the detailed extraction/HTML rules are not executable layouts. Repeated/interleaved candidate types are collapsed by `Distinct`, which can reorder the workbook narrative by component type.

### First-pair proof

- Source: `00. Report TEST new dry machine with material make press JIG 2024.03.26_1778470442_clean.xlsx`.
- Render: `outputs/batches/wpf-list-preview/group-reports/c643735b0fcf8216afe2.html`.
- Excel contains one cohesive report: the `B2:K22` material/time table, 15 sample photographs in the Picture column, six method/setup photographs on the right, and `M4 = Method test and dry machine setting at 230°C`. Excel also uses yellow fills and red bold text to mark the NG sample and highlights MIN/MAX.
- Numeric detection split the same table into `B6:B22` (`Sample`) and `F6:H22` (`1min/3min/5min`), both LOW-confidence unclassified candidates. The AI catalog then invented `Test Minimum Measurement — Two Panels`, an adjacent minimum-value table, and `TEST; MIN` important data.
- HTML consequently repeats the same `B2:K22` context twice under two `Note` cards, omits all 21 images, omits the `230°C` method condition, and loses the NG/MIN/MAX visual emphasis. The incorrect AI group is downstream of incorrect candidate evidence, not an isolated wording error.

### Representative failure classes

- `02. TIU L5S3 Result Test Frame (PV1 Frame, MP T2 Frame)_251112_clean.xlsx`: the main `12.11` sheet (80 values, 60 images, 130 merges) is omitted while one minor `Sheet1` candidate is rendered.
- `81. 161014 SPot NG Anylize_231208_clean.xlsx`: 35 meaningful text cells, 37 images, and the defect-analysis table are reduced to a zero-card `STRUCTURE_READY` page because there is no qualifying numeric candidate.
- `63.1. TIU L5S3-01 LR REPORT DATA SPL - 2026.01.20_clean.xlsx`: 60,116/64,203 non-empty cells (93.63%) are outside rendered evidence; a source surface reaching `ZK` is cropped to 32 columns.
- `19. BRS-201506 Report checking and test problem NG function high date 24.2.2024 -_clean.xlsx`: one source range is repeated up to 15 times across 36 cards, with percentage and coordinate distortions.

### Required interpretation for the next session

- The current v5 artifacts are technically complete but are **not accepted visual or semantic reports**. Do not describe `COMPONENTS_RENDERED=969` as user acceptance or faithful Excel rendering.
- The primary failure chain is: lossy capture contract -> numeric-island candidate errors -> limited AI evidence and recipe constraints -> coverage-only validation -> generic candidate-card renderer -> cropping/compaction/format heuristics.
- The next work should begin from this audit and redesign the representation/rendering contract. Do not apply another isolated CSS or support-column patch and do not recapture/regroup/rerender until the user explicitly asks for implementation.

## AUTHORITATIVE FINAL GOAL (user-confirmed 2026-07-17)

This section is the authoritative product goal and supersedes renderer-centric interpretations elsewhere in this handoff. HTML rendering, form grouping, and WPF browsing are supporting capabilities, not the final product.

### One-sentence goal

Build a WPF knowledge system that converts all historical and future Excel review documents into an evidence-traceable database of comparable review experiments, then answers natural-language questions by retrieving every relevant study, calculating and aggregating only valid control-versus-comparison evidence, and linking every conclusion back to its source table and workbook.

### 1. Convert the existing Excel corpus into a review-evidence database

- Ingest all currently held review-data Excel files (approximately 900; the current indexed corpus has 989 primary workbook records).
- Do not model a workbook as a list of numeric islands. Model it as one or more semantically complete review studies/experiments.
- Apply all three forms of harmonization:
  - **Database normalization:** separate reusable entities such as study, condition, intervention, comparator, outcome, result, and evidence.
  - **Standardization:** normalize names, units, rates, dates, model identifiers, process names, and equivalent labels such as `NG rate`, `NG율`, and `Function 불량률`.
  - **Semantic generalization:** map heterogeneous terms such as `VP+CD 본드량`, `접착제 도포량`, and `Bonding amount` into canonical concepts while retaining the original wording.
- Each study must be able to retain:
  - review purpose and hypothesis;
  - product, model, process, equipment, material, lot, site/line, and review period;
  - changed factor/intervention such as bond amount, pressing amount, assembly method, pressure, time, material, or equipment setting;
  - control group and comparison/treatment group;
  - conditions and sample size for every group;
  - input, OK, NG count, NG rate, FUNCTION NG, and detailed outcome metrics;
  - calculable absolute difference, percentage-point change, relative change, or risk measure;
  - source conclusion, limitations, and unresolved questions;
  - source workbook, worksheet, cell/table range, and stable evidence/data id;
  - embedded workbook images are outside the required extraction and analysis scope as confirmed by the user on 2026-07-17.

### 2. Produce one consistent, consolidated analysis per workbook

- Do not expose a raw list of every detected table as the main analysis.
- Consolidate all relevant workbook content into a study-level summary:
  - what was reviewed;
  - what was compared;
  - whether non-tested conditions were held constant;
  - which outcome changed and by how much;
  - whether the result direction is consistent;
  - whether the evidence is strong enough for a conclusion;
  - which exact source tables support each statement.
- A workbook should be represented as meaningful units such as `two comparison experiments plus one supporting measurement`, not `five numeric candidates`.

### 3. Answer natural-language questions from the evidence database

Example user question:

> VP+CD 조립 공정이 FUNCTION NG에 영향을 끼치나? 이 관계에 대해 검토 된 데이터 모두를 보여주고 요약해줘.

The program must:

1. retrieve every review related to VP+CD assembly and FUNCTION NG;
2. classify the retrieved reviews by changed factor, including bond amount, pressing amount, assembly method, equipment, material, pressure, time, and other conditions;
3. identify which reviews contain a valid control/comparison basis;
4. keep unlike models, lots, sample bases, or confounded multi-factor changes separate instead of averaging them blindly;
5. extract control and comparison sample sizes, input/OK/NG counts, rates, and detailed outcomes;
6. calculate appropriate effect measures such as absolute percentage-point change and relative change;
7. aggregate comparable results, report direction/range/consistency, and surface conflicting evidence;
8. distinguish association from causation and state when the evidence is insufficient;
9. attach a stable data/evidence number and direct source-table link to every quantitative claim;
10. let the user open the exact comparison table and, when needed, the original Excel workbook.

An expected answer pattern is:

> VP+CD 조립 관련 검토 18건을 찾았으며, 이 중 비교 조건이 확인된 자료는 11건입니다. 본드량 증가 검토 4건 중 3건에서 FUNCTION NG가 상승했습니다. 비교 가능한 검토의 상승 범위는 약 0.8~2.4%p였습니다. 다만 1건은 누름량과 경화시간도 동시에 변경되어 본드량만의 영향으로 단정할 수 없습니다. 근거: 데이터 #102, #284, #511, #728.

The numbers above are an answer-format example only; the final program must calculate real values from validated database evidence.

The VP+CD/FUNCTION NG question is an example, not a product-specific boundary. The final query path must answer any relationship supported by the canonical evidence database without adding a workbook-specific parser or a hard-coded whitelist for each new factor/outcome pair.

### 4. Incrementally add completely new Excel reviews through WPF

After the historical corpus is established, WPF must support new workbooks by:

1. reading the new workbook without changing the source;
2. mapping its terms, units, study design, intervention, comparator, outcomes, and evidence into the canonical schema;
3. detecting duplicate or closely related historical reviews;
4. validating comparison structure and evidence traceability;
5. updating the same evidence database;
6. making the new study immediately available to later AI questions.

The user must not manually group every new workbook. Automation may use AI, deterministic extraction, and bounded review, but accepted records must satisfy the same evidence contract as the historical corpus.

### Core data model

The central schema must be organized around:

`Review/Study -> Context -> Intervention/Changed factor -> Control group -> Comparison group -> Outcome metric -> Result/Effect -> Evidence source`

Form/layout groups may help choose an ingestion extractor, but they are not the knowledge model. Numeric candidates may be low-level evidence, but they are not studies or conclusions.

### Completion and trust requirements

- Every numeric or directional AI claim must cite stable evidence/data ids and link to the exact source table.
- Effect calculations must use only records with a valid control/comparison basis.
- Different models, lots, conditions, sample bases, or confounded interventions must not be combined silently.
- The system must not claim causation from weak, unmatched, or observational evidence.
- Insufficient or conflicting evidence must be reported explicitly.
- Original values, units, comparison conditions, and source locations must remain auditable.
- Each workbook must have a consolidated analysis, while raw tables remain available as drill-down evidence rather than the primary output.
- Newly added Excel files must enter the same canonical schema and validation path as the historical corpus.
- HTML is an evidence/detail viewer and optional presentation layer. Pixel-faithful Excel reproduction and embedded-image analysis are not final-goal requirements; source tabular meaning and comparison structure must remain accessible.

### Consequence for current implementation

- The current `layout group -> numeric candidate -> generic HTML card` pipeline is not the final architecture.
- Its validated grouping, renderer keys, and v5 HTML may be reused only as low-level inventory or migration evidence after their limitations are accounted for.
- The next major implementation must start with the canonical study/comparison/evidence schema and a traceable ingestion contract, then build consolidated workbook analysis and retrieval/aggregation on top.
- Do not resume renderer-only refinement as if it were the main deliverable.

## FULL EXECUTION AUTHORIZATION (2026-07-17)

The user instructed the project to save the plan and proceed through every stage of the authoritative final goal. The persistent phase order, deliverables, trust rules, and completion gates are recorded in:

- `FINAL_GOAL_EXECUTION_PLAN.md`

Execution must continue from the earliest incomplete gate. A phase is not complete merely because it has been planned; its artifacts and scoped verification must exist.

## IMPLEMENTATION CHECKPOINT (2026-07-17)

This checkpoint records implemented evidence-database work after the final-goal
authorization above.

### Scope that must not narrow

- `VP+CD assembly -> FUNCTION NG` is only one acceptance example.
- A user may ask about any factor, context, process, material, equipment,
  model, lot, condition, outcome, or relationship represented by validated
  database evidence.
- Extraction and retrieval must remain open-domain: there is no fixed concept
  whitelist and no workbook-specific parser may be added for each new question.
- Embedded images remain outside extraction and analysis scope.
- Do not restart the rejected renderer grouping as the main architecture, and
  do not recapture already current sources unless the source fingerprint or
  capture contract actually changed.

### Implemented and verified

- Canonical Study/Context/Factor/Arm/Outcome/Observation/Comparison/Effect/
  Evidence schema is installed in `outputs/universal-grid/InputDataFinish.sqlite`.
- Capture v2 completed for all 30 representative sources: 30/30 hashes and
  canonical bridges match, 105 sheets, 480,102 sparse/structural cells, 10,415
  formulas, 4,643 merges, zero failed or unfinished items.
- `outputs/semantic-source-packets/pilot-30-v2` contains 30 complete packet
  sets, 864 chunks, and exactly 234,287 primary semantic cells with no missing
  or duplicate ownership.
- The semantic locator and Study draft prompts are domain-neutral and explicitly
  state that VP+CD/FUNCTION NG are examples only.
- AI output cannot self-verify, create answer-eligible effects, enable
  aggregation, or assert causality. Exact source identity, A1 ranges, packet
  completeness, and every numeric observation are deterministically checked.
- Real end-to-end tabular pilot revision 18 imported one review-required Study,
  two Arms, 14 Outcomes, and 44 Observations. It deliberately imported no
  Comparison or Effect because the source did not explicitly prove a matched
  control. Two terminal workbooks imported as `EXCLUDED` with zero Studies.
- The generic evidence query supports arbitrary Unicode terms and searches
  canonical names, aliases, and original labels. It returns stable DATA/CMP/
  EFF/EVD identifiers, exact workbook/sheet/range citations, descriptive
  observations when no valid comparison exists, and a separate excluded set
  for unverified/confounded/invalid candidates.
- The real `waiting 2 day function NG` evidence pack applies a generic
  factor/context-plus-outcome relevance gate and retrieves the one Study
  matching both sides of that relationship. It exposes 14 Outcomes, 44
  Observations, and exact citations, but zero answer-eligible effects, which is
  the intended conservative result until a valid comparison is verified.
- Locator batching is implemented. A real two-chunk smoke run completed both
  chunks in one read-only AI call with zero failures while preserving
  per-chunk range validation.
- The current focused regression is 77 tests passing; the new locator-batching
  slice adds 27 passing CLI/semantic tests and Python compilation passes.

### Current phase

Phase 4 semantic extraction/validation and Phase 5 representative E2E
acceptance remain in progress. Full 989-workbook migration stays gated until
the representative semantic contract and golden questions pass, so errors are
not multiplied across the corpus. WPF incremental ingestion and final answer
synthesis remain later phases; no desktop application was launched for this
checkpoint.

## CANONICAL ANSWER + WPF INTAKE CHECKPOINT (2026-07-17)

This checkpoint supersedes the final sentence above: the core Phase 7 answer
path and Phase 8 WPF/intake surfaces are now implemented, although the full
pilot/golden-question gate and 989-workbook migration are still incomplete.

- `inference_data_ai_answer.py` implements deterministic
  `canonical-evidence-answer-v1`. Quantitative wording can only use a verified,
  answer-eligible effect with direct current-revision EFFECT evidence. Exact
  compatibility signatures prevent combining different models, lots,
  contexts, factor transitions, control/comparison conditions, metric bases,
  or designs. Conflicting directions are not averaged.
- `inference_data_ai_evidence_detail.py` resolves one stable `EVD-*` through
  its explicit current Capture v2 bridge. It returns exact cells, formulas,
  cached/display/raw values, number formats, styles, merge context, row/column
  hidden state, and `imagesAnalyzed=false`. Stale or mismatched evidence is
  rejected rather than redirected to another revision.
- `inference_data_ai_workflow.py` implements journaled single-`.xlsx`
  incremental ingestion through Capture → Packet → Locator → Draft → Import →
  Verify. It preserves the source SHA-256, resumes valid artifacts, skips AI
  for `EMPTY_WORKBOOK`/`NO_TABULAR_EVIDENCE`, and never auto-promotes an AI
  draft beyond `NEEDS_REVIEW`.
- CLI commands now include `evidence-answer`,
  `evidence-answer-validate`, `evidence-detail`, and `ingest-workbook`.
- Existing bridged canonical revisions now preserve terminal source semantics
  in `source_content_status` while canonical `capture_status` remains the
  revision lifecycle. The actual DB backfill reports 27 current `CAPTURED`, 2
  `EMPTY_WORKBOOK`, and 1 `NO_TABULAR_EVIDENCE` Capture v2 revisions.
- The actual `waiting 2 day function NG` query retrieves exactly one relevant
  `DATA-680D2C3E4D83` record and zero answer-eligible effects. The generated
  answer is `INSUFFICIENT_COMPARISON`; it displays source-backed Total NG and
  NG Rate observations but does not calculate a difference or claim an effect.
  Token-boundary filtering prevents `NG` from falsely matching `hearing`.
- WPF now contains unrestricted `검토 DB 질문` and
  `신규 Excel DB 적재` tabs. It uses the Python canonical contract rather than
  duplicating eligibility/effect math in C#, lists EVD citations, renders a
  merge/coordinate-aware source table, and opens the exact Excel sheet/range
  only after a user action. Images remain excluded.
- Verification:
  - relevant Python regression: 57 tests passed before the WPF integration;
  - query/answer/CLI slice after outcome filtering: 34 tests passed;
  - answer-focused slice after token-boundary correction: 7 tests passed;
  - `capture-v2-verify`: 30/30 valid, zero failed/unfinished;
  - canonical integrity: `ok=true`, zero orphan evidence links, zero invalid
    aggregation effects; previously recorded legacy FK warnings remain;
  - WPF narrow build: zero warnings, zero errors.
- The WPF application was not launched. Full Phase 5 acceptance, the rest of
  the pilot semantic drafts/golden questions, duplicate/related-study review
  presentation, and full Phase 6 migration remain the next gates.

## CORPUS / HUMAN REVIEW / ACCEPTANCE CHECKPOINT (2026-07-17)

The goal is domain-neutral. VP+CD is only an example; users may ask about any
factor, context, process, material, equipment, model, lot, condition, outcome,
or relationship represented by validated DB evidence.

### Exact current state

- Frozen corpus: 989 sources, 3,214,102,996 bytes.
- Corpus journal: 8 `COMPLETED`, 981 `PENDING`, 0 `FAILED`.
- Completed result states: 5 `NEEDS_REVIEW`, 3 `EXCLUDED`.
- Canonical DB: 38 source documents, 39 revisions, 19 workbook analyses,
  35 Studies, 113 Arms, 75 Outcomes, 251 Observations, 68 Comparisons,
  73 Effects, 627 Evidence items, 892 evidence links, and 0 review decisions.
- SQLite integrity check: `ok`.
- Images extracted/analyzed: no.

### Representative findings

- P01 is an accurate descriptive multi-material/time result and correctly has
  no invented control comparison.
- P03 is a known v1 quality defect: the source contains Normal rows, but the
  earlier draft omitted comparisons. Keep it blocked; do not approve it.
- P10 was corrected deterministically from a label-only pseudo-Study to
  `NO_TABULAR_EVIDENCE`, without recapture, regrouping, or another AI call.
- P11 proves the v2 correction works: 4 Studies, 16 Arms, 4 Outcomes,
  16 Observations, and 12 explicit 40/50/60-versus-Normal Comparisons.
  All four Normal rows are control/baseline Arms. Effects remain zero because
  every comparison is review-gated. Vision/laser/function have unequal sample
  bases and no stored matching basis; tension has n=5 per Arm but still
  requires review for other differences.

### New fail-closed review and retrieval behavior

- `review-queue`, `review-detail`, and `review-decide` expose explicit human
  decisions. Approval accepts human assessment fields in the same transaction,
  but succeeds only with Study comparability `VALID`, Study confounding
  `NONE`, Comparison validity `VALID`, Comparison confounding `NONE`, a
  nonempty matching basis, current SHA-256 source identity, and direct
  current-revision VERIFIED evidence for the Comparison and both paired
  Observations.
- Approval calculates Effects deterministically and links each Effect to the
  exact approval EVDs. Reject/exclude/re-review disables aggregation. No real
  record was approved during implementation.
- WPF includes the review queue, values and exact EVD ranges, source-table
  preview, user-triggered exact Excel range open, explicit assessment fields,
  and a confirmation dialog. It was built only, not launched.
- Terminal sources without Studies are returned as source-level exclusions.
  A generic single-distinctive-exact-term fallback fixes P08 retrieval without
  adding a product/domain whitelist.
- Uncited legacy descriptive observations are withheld and reported as
  limitations instead of being shown without current-revision EVD support.
- Exact-content duplicates and related Studies are separate. Related ranking
  is discovery-only lexical similarity and never relationship/causal evidence.

### Golden-question baseline

`outputs/golden-acceptance/pilot-current/acceptance-report.json` is the current
machine-readable baseline:

- overall: `BLOCKED_PENDING_INGEST`;
- 10 questions: 1 structural pass, 9 blocked pending ingestion, 0 failures;
- 15 primary-source appearances: 4 represented, 11 pending, 0 retrieval miss;
- 0 eligible Effects because no new Comparison has received human approval.

This is not a failed final system claim and not a pass: it is a precise
checkpoint showing which required sources have not yet entered the canonical
path.

### Provenance and verification

- Future incremental journals record locator/batch-locator/draft prompt
  versions, AI execution in the current attempt, and artifact reuse.
- P11's verified `canonical-study-draft-prompt-v2` provenance is present in its
  journal.
- Review/acceptance/CLI focused tests, query/answer/acceptance tests, workflow
  and corpus tests pass.
- WPF narrow build: 0 warnings, 0 errors.
- Do not launch WPF for automated verification.

### Next action

Continue the remaining representative pilot workbooks through the existing
resumable corpus journal. Inspect v2 comparison quality before increasing
throughput, perform only explicit source-backed human approvals, rerun the ten
golden questions, then expand the validated workflow to the remaining corpus.

## P12 정밀 보존 및 조회 체크포인트 — 2026-07-18

이 항목이 위 2026-07-17 수치보다 최신이다. 최종 GOAL은 아직 진행 중이다.

- 전체 journal: 989개 중 9개 `COMPLETED`, 980개 `PENDING`, 실패 0개.
  완료 결과는 `NEEDS_REVIEW` 6개, `EXCLUDED` 3개다.
- 대상 P12:
  `014.MSU-20S15-07 Result test AWF cooling time 4s,8s,10s_clean.xlsx`.
  현재 분석 ID는 `ANALYSIS-7BCA4D3F017F`이며 같은 revision의 이전 미검토
  분석은 안전하게 supersede되어 현재 분석은 하나만 남았다.
- v2 초안은 개별 scalar 지표를 합치고 표본 수와 넓은 주파수 행렬을 놓쳐
  폐기했다. 승인하거나 효과 계산에 사용하지 않았다.
- source prompt는 약 160만 자에서 172,480자로 압축했지만 5,080개 셀은
  하나도 버리지 않았다. prompt v5와 provenance를 저장했고, 잘못된 Arm
  참조 한 곳(`cool_time_10s` → `cooling_time_10s`)만 고친 repair 결과도
  전체 diff로 확인했다. 이후에는 reference 이외 내용이 바뀌면 repair를
  거부하는 projection guard가 있다.
- 냉각시간 Study: Arm 3개, 표본 수 100/156/122, scalar Outcome 22개,
  Observation 66개. 명시적인 유효 대조군이 없어 Comparison/Effect는 0개다.
- 주파수 Study: Arm 4개, measurement series 4개, point 2,697개
  (Spec 87, 20000V 870, 1800V 870, 1600V 870). 2,697개 값은 Capture 원본과
  전수 대조하여 불일치 0개였고 좌표 중복 0개, 축 누락 0개다. header/value/
  row-identity 범위는 series당 3개, 총 EVD 12개로 연결돼 있다.
- Spec 대비 Comparison 3개는 모두 `NEEDS_REVIEW`/confounded이고 Effect는
  0개다. 실제 human approval는 하지 않았으며 `review_decisions=0`이다.
- 반복 라벨을 의미 문자열이 아니라 원본 좌표로 세도록 수정했다.
  따라서 1600V도 87축/10개 원본 행으로 표시된다.
- 실제 frequency 질문 결과는 관련 Study 1개, series 4개, citation 12개,
  eligible Effect 0개다. 같은 모델명만 겹친 XRAY 이미지 전용 파일은 더
  이상 이 질문의 관련 자료로 나오지 않는다.
- 골든 수용성: `BLOCKED_PENDING_INGEST`; 10문항 중 2 pass, 8 pending,
  fail 0. 기대 원본 15회 중 5 represented, 10 pending, retrieval miss 0.
- 검증: 관련 Python 테스트 100개 통과, py_compile 통과, WPF build 경고 0/
  오류 0, universal-grid 검증 16/16, SQLite `integrity_check=ok`, canonical
  FK 오류 0. 남은 FK 경고 66개는 기존 `analysis_*` 호환 테이블에만 있다.
- 이미지는 계속 추출·분석하지 않는다. WPF/데스크톱 앱은 실행하지 않았고,
  사용자가 열어 둔 Excel/WPF 프로세스도 건드리지 않았다.

다음 순서: 다음 대표 파일을 동일 journal로 한 건 처리한 뒤 scalar,
대조군/비교군, 측정 series, EVD 좌표 보존을 검사한다. 이 품질 게이트를
통과한 후에만 처리 폭을 늘린다.

## P13 의미 안전성 체크포인트 — 2026-07-18

- 대상:
  `51. BRS-161014 Report test VP mold #4,#7,#8 change mold temperature and
  valcunizing agent 10% date 17.7.2024_clean.xlsx`
- 현재 분석 ID: `ANALYSIS-B6586C6CDFB9`
- 기존 Capture v2, 48개 완전 청크, 48개 locator 결과를 그대로
  재사용했다. locator AI 호출은 0회이고 v6 Study 초안만 새로
  생성했다. 이미지 추출/분석은 하지 않았다.
- 최초 P13 대조에서 원본 값 불일치는 없었지만 의미 오류 네 가지가
  발견됐다: Excel 퍼센트 저장값/표시값 혼동, Input/OK/NG 원시 건수
  차이의 잘못된 효과화 가능성, dry/cure/agent 조건과 동일 몰드
  180↔190 비교 누락, AVG 열의 반복수 포함.
- 공통 코드에서 퍼센트 표시 단위 정규화, `sample_size` 분모 전용
  처리, 분모 없는 count 효과 차단, 명시적 rate와 중복되는 count
  파생 효과 억제, RAW/AGGREGATE point 역할을 구현했다.
- 현재 P13은 Study 5, Arm 27, Outcome 33, scalar Observation 197,
  measurement series 27, point 13,311, Comparison 24, Effect 0이다.
  MASK 10개 값은 series point가 아니라 scalar observation으로
  보존되어 전체 숫자 13,508개는 그대로다.
- point는 RAW 10,962개 + AVG 파생값 AGGREGATE 2,349개다. AVG
  2,349개 모두 같은 행 RAW 평균과 일치하고 기본 통계/반복수에서는
  제외된다. 원본 sheet+좌표와 point 값 불일치는 0개다.
- rate observation 20개는 모두 분자/분모 및 퍼센트 산술이 맞고,
  Input 이외 count observation 124개는 모두 분모가 있다. Input
  14개는 `sample_size`다.
- Vision/Function/SPL·THD·IMP에서 같은 몰드의 180↔190 비교 9개가
  정렬되어 추가됐다. dry 150도/1시간, vulcanizing agent 10%,
  second cure, line, lot, mold 조건을 보존했다.
- Comparison은 14개 confounded, 10개 unassessed이고 전부
  `NEEDS_REVIEW`, `aggregationEligible=0`, Effect 0,
  `review_decisions=0`이다. 자동 승인하지 않았다.
- corpus: 10 completed / 979 pending / 0 failed.
- 골든 검증: 10문항 중 3 pass, 7 pending, 0 fail. 필수 원본 등장
  15건 중 6 represented, 9 pending, retrieval miss 0.
- 검증: Python 115개 통과, P13 manifest/숫자/좌표/AVG/비율 산술
  오류 0, canonical FK/orphan/invalid aggregation 오류 0,
  universal-grid 16/16 통과, WPF build 경고 0/오류 0. WPF/Excel
  앱은 실행하거나 건드리지 않았다.

다음 순서: P14를 동일한 재개 가능·이미지 제외 workflow로 처리하고,
파일별 scalar/비교군/series/원본좌표/승인 안전성 gate를 통과한 뒤
처리량을 늘린다.

## GQ06 혼입 설명 및 P14 사전 차단 체크포인트 — 2026-07-18

- GQ06의 기존 문제는 Function NG 질문에 P13 전체 SPL/THD/IMP series를
  붙이고, 혼입 비교를 `NO_VALID_COMPARISON` 한 줄로만 설명한 것이었다.
- effect가 없는 Comparison도 이제 validity/confounding/aggregation,
  양쪽 arm 조건, matching basis, factor 차이, 직접 EVD를 보존한다.
  factor 값 차이가 실제 두 개 이상일 때만
  `CONFOUNDED_MULTI_FACTOR`를 반환한다.
- 현재 GQ06은 eligible Effect 0, confounded 설명 14개,
  multi-factor 12개, 관련 없는 descriptive series 0개다. 직접 인용은
  91개에서 28개로 줄었고, Function 비교의 `VP mold`, `2nd Cure`,
  `Test type`, `1st Molding` 양쪽 값을 열거한다.
- GQ06 required behavior 세 가지는 선언형 자동 검증으로 3/3 PASS다:
  factor 전수 보존, `CONFOUNDED_MULTI_FACTOR` 코드, eligible Effect
  최대 0. 전체 골든 상태는 다른 7개 미적재 질문 때문에 계속
  `BLOCKED_PENDING_INGEST`다.
- P14 실패 초안은 DB에 반영되지 않았다. 반복된 일반 오류는 빈 REF
  열을 dense range에 포함, 허용되지 않은 `BASELINE`, 원본에 없는
  Before/After 창작, 빈 replicate identity 중복, Fo/AVG 누락이었다.
- prompt v8에 정확한 arm enum, Before/After 추정 금지, 반복 블록
  stratum 보존과 sample size 유지, data-bearing REF 열만 사용, dense
  series 빈셀 금지, Air-leak/Fo raw replicate identity 및 AVG aggregate
  보존을 명시했다.
- measurementSeries는 이제 DRAFT 단계에서 원본 셀을 read-only로 전개해
  blank/malformed/error 셀을 IMPORT 전에 차단한다. 일반 validator 오류도
  rejected JSON+오류+focused source packet으로 repair하며, 순수 reference
  repair의 projection guard는 유지한다.
- 전체 Python 테스트 199개 통과. 이미지 분석과 WPF/Excel/서버 실행은
  하지 않았다.

다음: 기존 P14 Capture revision과 locator 71개를 그대로 재사용한 v8
초안을 원본 체크리스트로 전수 감사하고, 모두 통과할 때만 import한다.

## P14 custom number format 의미 정정 및 적재 완료 — 2026-07-18

이 항목이 바로 위 P14 임시 결론보다 최신이다. 특히 이 파일의
Before/After가 원본에 없다는 결론은 철회한다.

- 대상:
  `38. MSU-L20S1507 DOE Air preasure-air leak_clean.xlsx`
- 현재 분석 ID: `ANALYSIS-B35468F7B510`
- header의 저장값은 `1..10`이지만 Excel custom `number_format`에는
  `18kPa #1_Before`, `18kPa #1_After`처럼 압력·순번·단계 의미가
  원본 작성자에 의해 들어 있다. Capture는 number format을 보존했지만
  `display_value_json`과 기존 HTML/WPF 렌더러가 raw 숫자만 표시했다.
  이것이 Excel과 렌더링 결과의 의미가 달라진 구체적인 원인 하나다.
- 이제 실제 근거 header 셀의 custom format만 제한적으로 해석해
  measurement identity를 복원한다. AI의 sourceText나 locator 요약으로
  의미를 만들 수 없으며, Before/After 표시는 CONTROL/BASELINE 또는
  인과관계를 뜻하지 않는다.
- 18/100/200 kPa는 동일 순번의 Before/After 정렬 근거가 있어
  SPL/THD/IMP/Fo에서 가능한 비교 9개를 만들었다. 모두
  `NEEDS_REVIEW`/`UNASSESSED`, aggregation eligible 0, Effect 0,
  사람 승인 0이다.
- 300 kPa After는 header만 있고 값이 모두 비어 있어 header-only arm으로
  보존하고 series/comparison은 만들지 않았다. malformed THD 1셀,
  IMP `#REF!` 셀, 값 없는 REF sibling은 제외했다. REF 값은 reference일
  뿐 control로 쓰지 않았다.
- pooled AVG를 표본으로 중복하지 않는 standalone aggregate series
  계약을 추가했다. `AVERAGE`, RAW 원본 series 목록, 같은 Study/Outcome,
  같은 축 평균 일치가 모두 필수다. P14 AVG series 7개는 전 축 산술이
  정확히 일치한다.
- 최종 보존량: Study 5, Arm 39, Outcome 9, scalar Observation 17,
  series 46(RAW 39 + AGGREGATE 7), point 18,916
  (RAW 18,651 + AGGREGATE 265), Comparison 9, Effect 0.
  원본 sheet+좌표 중복과 값 불일치는 0이다.
- v14는 기존 Capture revision
  `capture_revision_c1e1c775deaa07463430899c`, packet, locator 71개,
  기존 artifact를 재사용했다. locator AI 호출 0, 재캡처 0,
  재그룹화/재배치 0, 이미지 추출·분석 0이다. 실패 초안은 적재하지
  않았다.
- corpus는 11 completed / 978 pending / 0 failed. 골든 검증은
  4 pass / 6 pending-ingest / 0 fail이고 GQ05가 P14로 표현됐다.
- 전체 Python 218 tests PASS, universal DB 16/16, canonical/SQLite
  무결성 PASS, WPF build 경고 0/오류 0. 앱·Excel·서버는 실행하지 않았다.
- 현재 렌더링 HTML은 여전히 Excel 충실 뷰로 승인하지 않는다. 이번
  단계는 DB 의미 보존과 number-format 손실 원인 규명까지이며,
  HTML 시각 수정 완료를 뜻하지 않는다.

다음: P15를 동일한 이미지 제외·한 파일 품질 gate로 처리한다. scalar,
대조군/비교군, wide series, formatted identity, AVG 산술, EVD 좌표,
검토 안전성을 모두 확인하기 전에는 병렬 처리량을 늘리지 않는다.

## P15 원본 대조 완료 및 DB 적재 — 2026-07-18

- 대상:
  `1. BRS-161014 Report test DOE sub 2 manual_clean.xlsx`
- 현재 분석 ID: `ANALYSIS-8CAED455EB68`
- 두 sheet의 서로 다른 실험 집단을 섞지 않고 CM+B-PT Visual,
  Tension, CM+CP bonding/drying DOE, CM+B-PT bonding-line DOE의
  Study 4개로 분리했다. `In Spec`은 기준군으로 바꾸지 않고 `TEST`로
  보존했다.
- 최종 적재량은 Study 4, Arm 18, Outcome 21, scalar Observation 90,
  RAW series 8, point 64, Comparison 10, Effect 0이다. point 64개는
  실제 표본 32개에서 본드 실측량과 tension 두 변수를 측정한 것이며
  표본 64개를 뜻하지 않는다.
- 본드량 series 4개는 모두 `mg`, tension series 4개는 모두 `kgf`다.
  각 series는 원본 8점을 빠짐없이 보존하며 64개 값·축·좌표의 Capture
  불일치는 0이다.
- Visual의 수식 4개는 cached value가 없으므로 `valueNumber=null`을
  유지하고, 원본 Input/NG의 0/8 또는 8/8과 수식 셀 근거로만
  `ratePpm`을 결정론적으로 계산했다.
- Tension MAX/MIN/AVG 수식 12개도 cached value가 없다. 원본 수식
  계보 범위 `31.7!D46:O49`만 보존하고 수치 집계값을 창작하지 않았다.
- CM+B-PT row 46은 Input 8, OK 0, Total NG 5로 세 표본이
  미분류/미조정 상태다. OK나 NG를 보수로 고치거나 누락 분류를
  대치하지 않았고 limitation에 정확히 기록했다.
- 안전한 비교 10개만 보존했다: Visual 인접 3, 독립 Tension 인접 3,
  CM+CP 본드량 단일 차이 2, 명시적으로 bonding line을 변경한 집단
  내부의 본드량 단일 차이 2. bonding-line 셀이 빈 행과 연결되는
  유혹적인 비교 2개는 빈칸이 동일 기준선/미변경을 증명하지 않으므로
  제외했다.
- 모든 Comparison은 `NEEDS_REVIEW`/`UNASSESSED`, aggregation eligible
  0, Effect 0, 사람 승인 0이다.
- 최초 품질 보정 응답은 AI가 결정론적 `contentComplete=true`를
  false로 바꿔 import 전에 차단됐다. packet의 값과 4개 비율 산술만
  결정론적으로 복원한 뒤 canonical contract, numeric evidence,
  강화된 P15 원본 gate를 모두 통과했다. 앞으로 runner는 AI 판단이
  아닌 packet coverage를 복원하고, 직접 validator는 불일치를 계속
  거부한다.
- import 시 기존 Capture revision
  `capture_revision_53a1a1e20535ff386f3e7a54`, 345/345 셀 packet,
  locator 9개, prompt-v17 draft를 재사용했다. import 단계 locator AI
  0, draft AI 0, 재캡처 0, 이미지 추출·분석 0이다.
- corpus는 12 completed / 977 pending / 0 failed. 골든 검증은
  4 pass / 6 pending-ingest / 0 fail이고 필수 원본 표현은 7/15,
  retrieval miss 0이다.
- semantic/workflow/import 집중 테스트 61개와 전체 Python 222 tests가
  통과했다. universal DB 16/16, P15 SQL/원본값 감사, WPF build
  경고 0/오류 0이다. WPF·Excel·서버·데스크톱 앱은 실행하지 않았다.

다음: P16을 같은 이미지 제외·한 파일 원본 gate로 처리한다. 남은
대표 파일에서 동일한 완전성과 대조군/비교군 안전성을 확인하기 전에는
전체 corpus 병렬 처리량을 늘리지 않는다.

## P16 교정 완료 / 30개 병렬 벤치마크 중단 체크포인트 — 2026-07-18

- P16
  `017.MSU-20S15-07 Result test sample waitting 2 day and check function_clean.xlsx`
  의 현재 분석은 `ANALYSIS-F0756E286A00`이다. 약한 기존 분석
  `ANALYSIS-0D3FE4FD3695`는 supersede했다.
- P16은 Study 1, Arm 2, Outcome 22, Observation 44, Comparison 1,
  Effect 0이다. Waiting은 TEST n=299, Normal은 CONTROL n=920이다.
  main Function NG는 각각 66/299=22.07357859531772%,
  236/920=25.65217391304348%다.
- 차이 -3.57859531772576%p는 기술 통계일 뿐이다. 무작위화, matching,
  lot 동등성, 인과 식별 근거가 없으므로 Comparison은
  `NEEDS_REVIEW`/`UNASSESSED`, aggregation eligible 0, Effect 0이다.
- prompt/import v18부터 Excel percent-format 셀의 정확한 `raw × 100`
  값만 human percent로 허용한다. raw fraction, 반올림 display 값,
  같이 인용된 일반 count 값은 percent 근거로 통과할 수 없다.
- 기존 Capture revision
  `capture_revision_334844564a86a310e4ceccdc`와 67/67 packet을
  재사용했다. 재캡처·재그룹화·이미지 추출·이미지 분석은 0이다.
- P16 이후 `pilot/corpus-benchmark-small-30-v1.json`의 30개를
  workbook 3병렬, locator 2병렬로 실행했다. 30개는 중복 SHA/path/
  layout cluster가 0이며 SIMPLE 17, COMPLEX 7, EXTREME 6이다.
- 실측 병목은 Excel/DB가 아니라 AI draft였다. Capture/Packet/
  Import/Verify는 대부분 1초 미만이나 작은 파일 draft도 1.5~3.5분,
  큰 파일은 7~13분이다. 현재 3병렬 기준 977개 one-pass는 30~45시간,
  재시도 포함 잠정 36~72시간이다.
- v18 벤치마크는 4 pipeline 완료, 6 계약 실패, 중단 시점에 실행 중이던
  3건의 중단성 실패, 17 예약 상태에서 중단했다. DB에 잘못된 자동
  의미 분석을 더 넣지 않기 위한 의도적 중단이다.
- pipeline 완료 4건의 원본 의미 감사 결과:
  - B01 `ANALYSIS-606B44C2FC52`: semantic PASS.
  - B03 `ANALYSIS-3C55E1FB61F8`: 숫자/series는 정확하지만 AI가 만든
    요약을 `SOURCE_CONCLUSION`으로 저장해 semantic FAIL.
  - B05 `ANALYSIS-80970E38FC9C`: 숫자/series는 정확하지만 `1.56mg`
    numeric/unit 손실, numeric tension outcome에 `Pass` 혼합,
    비정렬 LED 비교, CONTROL 역할 과장 위험으로 HOLD/FAIL.
  - B06 `ANALYSIS-D299E22731EC`: PASS_WITH_WARNINGS. 모든 count/rate
    산술과 중복 가능한 NG category limitation이 정확하다.
- 모든 완료 건은 Effects 0, 전 Comparison `NEEDS_REVIEW`,
  aggregation eligible 0이라 잘못된 정량 효과나 자동 승인은 없었다.
  하지만 pipeline PASS와 semantic PASS를 분리하지 않으면 거짓 성공이
  생기므로 전체 병렬 적재는 재개하지 않는다.
- 공통 실패 4종은 v20에서 보강했다: cross-chunk locator 인용은
  singleton 1회 재실행, 비JSON draft는 schema 강제 1회 재생성,
  빈 rowIdentityRange는 rejected-draft source repair,
  단독 numerator는 denominator를 창작하지 않고 pair만 비우며 raw
  count/evidence를 보존한다. 최신 semantic/import/workflow 집중 테스트
  67개가 통과했다.
- 다음 gate는 source-authored conclusion provenance, factor numeric/unit,
  outcome homogeneity, comparison alignment, Normal role 정책을 공통
  validator로 구현하고 B03/B05를 교정한 뒤 동일 30개를 retry하는 것이다.
  이미지 분석과 WPF/Excel/서버 실행은 계속 금지한다.

## v23 의미 안전성 교정 및 30개 벤치마크 재개 — 2026-07-18

- B03과 B05는 `STALE`로 격리하고 journal도 다시 처리 가능 상태로 내린 뒤,
  공통 계약을 prompt/provenance v23으로 강화했다.
- v23은 출처가 직접 쓴 결론과 AI 기술 통계를 구분하고, `1.56mg` 같은 수량의
  원문·숫자·단위를 함께 보존한다. 정량 tension과 정성 `Pass`를 분리하며,
  직접적인 Control 근거가 없는 `Normal`은 `REFERENCE`로 둔다.
- RAW 대 scalar, 축·shape·stratum이 다른 RAW series처럼 표현이 정렬되지 않은
  비교는 import 전에 차단한다. 격리된 분석은 answer-visible current 분석이나
  `COMPLETED` journal 상태에 남을 수 없고, 같은 원본의 교정 재수입이 성공하면
  격리 이슈를 `RESOLVED`로 바꾼다.
- B03 `ANALYSIS-3C55E1FB61F8`은 잘못된 `SOURCE_CONCLUSION`을
  `AI_DERIVED_DESCRIPTIVE`로 교정했다. 3초/Normal arm은
  `TEST`/`REFERENCE`, 길이가 10 대 8인 RAW series는 비교하지 않는다.
  원시 point 18/18이 Capture v2와 일치하고 Comparison 0, Effect 0,
  `NEEDS_REVIEW`다.
- B05 `ANALYSIS-80970E38FC9C`은 양 arm의 `1.56mg`를 원문 `1.56mg`,
  숫자 `1.56`, 단위 `mg`로 보존한다. `Pass` 4개는 별도
  `tension_note` categorical Outcome이다. 모든 Normal은 `REFERENCE`,
  Old bond는 `COMPARATOR`, Control은 0건이다.
- B05의 3-position LED RAW와 Normal merged scalar 사이 잘못된 비교 3개를
  제거했다. 남은 비교 3개도 전부 `NEEDS_REVIEW`, aggregation eligible 0,
  Effect 0이다. 원시 point 46/46과 scalar 값이 Capture v2와 일치한다.
- 두 수동 격리 이슈는 `RESOLVED`이고 answer-eligible Effect는 0이다.
  전체 Python 회귀시험은 246/246 PASS, universal DB 검증은 16/16 PASS다.
  이미지 추출·분석 및 WPF/Excel/서버/데스크톱 앱 실행은 하지 않았다.
- 14:52(+07)에 `run_corpus_benchmark_small_30_v23.ps1`로 30개 벤치마크를
  재개했다. 완료된 B01/B03/B05/B06 4개는 건너뛰고 나머지 26개를
  workbook worker 3, locator worker 2로 처리 중이다. 결과 파일은
  `outputs/corpus-ingest/full-989-v1/benchmark-small-30-v23.result.json`이다.
- 989개 전체의 정직한 예상은 one-pass 30~45시간, 의미 감사와 재시도 포함
  약 36~72시간이다. 30개 품질 gate 통과 전에는 전체 corpus로 확장하지 않는다.

## v23 30개 종료 / v24 재처리 및 staged draft 구현 — 2026-07-18

- v23 30개 벤치마크는 최종 13 COMPLETED / 17 FAILED로 종료했다.
  모든 실패는 import 전 fail-closed이며 이미지 처리는 0이다.
- 완료 13개를 원본 감사한 결과 answer-eligible Effect는 모두 0이었다.
  다만 이전 prompt 결과인 B06 `ANALYSIS-D299E22731EC`과
  B20 `ANALYSIS-83B828CC2B8B`은 복합 문장에서 각각
  `5.6kg/3.5kg`, `8V/7V/2 day` 토큰을 source-exact factor처럼 보존해
  최신 계약 false pass로 판정했고 둘 다 `STALE` 격리했다.
- 17개 실패 분류:
  - 병합 identity/header 6
  - 복합 조건 quantity 3
  - 결론/상태 provenance 2
  - 대용량 output 전송 2
  - numeric text, aggregate axis, 합성 numerator, Normal 그룹 label 각 1
- 공통 계약/실행은 prompt v24로 보강했다.
  - 병합 covered row identity와 1열 header는 정확한 same-sheet merge anchor만
    상속하고 실제 anchor 좌표를 provenance로 저장한다.
  - 일반 공백, 숫자 valueRange, 여러 header가 한 anchor로 겹치는 경우는 거부한다.
  - strict whole-cell numeric text만 정확한 Decimal 동등 `valueNumber`로 1회
    제한 복구하고 Pass/Fail 또는 다른 필드 변경은 거부한다.
  - 인접 다중 셀 SOURCE_CONCLUSION은 원본 순서와 제한된 구분자를 보존할 때만
    허용하며 역순은 거부한다.
  - 단위 카탈로그에 없는 V/kg/day/min도 quantity-like syntax를 먼저 감지해
    전체 셀 exact match 없이는 factorValue로 통과하지 못한다.
  - 복합 조건은 token을 잘라 정량화하지 않고 원문 whole-cell compound factor로
    보존하거나 component factorValue를 생략한다.
  - 가로 frequency×replicate matrix는 `ROW_IDENTITY`, 세로 replicate+AVG는
    shared header `HEADER` axis로 구분한다.
  - 출처에 없는 numerator/denominator는 그 쌍만 null로 제거하고 category 합산은
    금지한다.
  - 반복 `PASSED`는 결론이 아니라 replicate별 categorical Outcome이다.
  - bare `ST`는 REFERENCE가 아니며, `Normal #1…#N` 그룹은 모든 근거 셀이
    순수하고 순서가 고유한 reference replicate일 때만 REFERENCE로 허용한다.
  - canonical `replicateKey`/`stratumKey`가 DB에 그대로 저장되도록 importer
    legacy-key 불일치도 수정했다.
- 대형 B24/B25는 각각 9,646/9,660 cells와 32/33 chunks에서 단일 draft가
  약 71,950/65,637 output tokens까지 증가했고 system temp last-message 파일
  생성이 실패했다. 호출별 UUID output/schema 파일을 workbook artifact
  폴더에 두고 정확한 파일만 정리하는 안정 전송 경로를 구현했다.
- 더 중요한 누락 위험은 locator 후보가 있는 SPL/THD/IMP section의 숫자-only
  continuation chunks가 단일 focused draft에서 제외되던 점이다. 전체 corpus
  확장 전에 workbook→sheet→section→bounded contiguous chunk fragment,
  조각별 resume/provenance/validation, deterministic consolidation,
  complete coverage와 final validation 후에만 import하는 staged draft를
  구현 중이다. B24/B25가 acceptance fixture다.
- v24 통합 후 전체 Python 회귀시험은 254/254 PASS였다. 이후 grouped role,
  stable output 등 추가 집중 시험도 PASS했으며 최종 전체 suite는 staged
  통합 후 다시 실행한다.
- 15:48(+07)에 대형 B24/B25를 제외한 17개 v24 재처리를 시작했다.
  selection은 `pilot/corpus-benchmark-retry-17-v24.json`, 결과는
  `outputs/corpus-ingest/full-989-v1/benchmark-retry-17-v24.result.json`이다.
  시작 시 corpus 상태는 23 COMPLETED / 17 RUNNING / 2 FAILED / 947 PENDING이다.

## v24 종료 감사 / source-content coverage gate — 2026-07-18

- v24 비대형 17건 재처리는 12 COMPLETED / 5 FAILED로 종료했다. 시스템 성공
  12건을 다시 Excel/Capture v2/manifest와 대조한 결과 7건만 의미적으로
  안전했고, B08/B09/B13/B14/B22 5건은 정량 표 또는 원문 결론을 누락한
  false pass라 격리했다. 나머지 5건은 import 전에 fail-closed했다.
- 안전 사용 가능: B10, B18, B20, B21, B26, B29, B30.
  false-pass 격리: B08, B09, B13, B14, B22.
  기존/현재 답변에 자동 사용 가능한 Effect는 계속 0이다.
- B04 numeric-text repair는 두 번째 AI 응답이 설명문까지 바꾸지 못하도록
  exact whole-cell 숫자에 한해 `valueNumber`만 결정적으로 복구한다.
  B16의 출처 없는 numerator/denominator는 발견된 observation의 그 쌍만
  제거하며, unsafe repair artifact보다 원본 rejected artifact를 우선한다.
  B06의 `1st/2nd/Total`은 복합 셀 안의 순번 토큰으로서 독립 수량이 아니다.
- `study-content-coverage-v1` 역방향 gate를 추가했다. canonical claim의
  근거 존재 여부만 검사하는 기존 정방향 validator에 더해, candidate-bearing
  source의 모든 숫자 결과·집계·수식 cached result·정확한 조건값·원문 결론이
  Observation/RAW 또는 AGGREGATE series/design evidence/conclusion에 실제로
  보존됐는지 검사하며 누락 시 최종 import 전에 중단한다.
- 실제 v24 artifact 12건 acceptance:
  - 불량 B08/B09/B13/B14/B22: 5/5 차단
  - 안전 B10/B18/B20/B21/B26/B29/B30: 7/7 통과
  - B08은 정량 7셀과 B26 결론, B09는 41셀, B13은 47셀,
    B14는 851셀, B22는 2,288셀 누락으로 차단됐다.
- B21/B26의 유일한 초기 오탐 `SPL DATA_(NTI!A175:A176`은 독립 원문
  감사했다. 유효 series A3:A89 뒤 85행 공백을 지나
  `19000/#REF!`, `20000/#REF!`만 남은 고립 오류 축 꼬리다.
  인접 2행 이상이며 같은 행의 다른 비어 있지 않은 셀이 오른쪽 Excel
  error뿐인 경우에만 `ERROR_ONLY_AXIS_TAIL`로 제외한다. 원문 range와
  제외 limitation은 manifest에 보존돼 있다.
- focused content/workflow 테스트 20/20 및 실제 12건 matrix가 통과했다.
  이미지 분석, 재캡처, 재그룹화, WPF/Excel/서버/앱 실행은 하지 않았다.
- staged draft v1은 독립 검토에서 안전하지 않은 것으로 판정했다. part
  Study를 단순 append해 논리 Study 병합과 cross-part 비교를 잃고,
  continuation anchor/registry, evidence allowlist, resume contract hash가
  불충분하다. B24/B25와 전체 corpus에는 사용하지 않는다.
- 다음 gate는 staged draft v2다. 실제 full-request budget을 넘을 때만
  staging하고, source-identity `logicalStudyId`, append-only fragment,
  owned/shared evidence allowlist, deterministic conflict-failing merge,
  evidence-backed comparison intent, exact resume provenance를 구현한 뒤
  B24/B25와 30건을 재검증한다.
- 현재 corpus journal은 35 COMPLETED / 7 FAILED / 947 PENDING이다.
  이는 실행 상태 수치이며, 위 의미 감사의 격리 판정과 동일한 개념이 아니다.

## staged draft v2 완료 / 의미 안전 게이트 고정 — 2026-07-18

- staged draft v2의 bounded fragment 처리, 최대 3-worker 병렬 실행,
  source-order 결정적 병합, exact prompt text/SHA-256, source-cell
  owned/shared allowlist, 양방향 `coverageDispositions` exact equality,
  merged disposition 보존/해시, part/final provenance 재검증과 안전한
  resume를 구현했다.
- Sol 독립 검증에서 staged v2 9, workflow 9, semantic AI 45,
  content coverage 25로 관련 시험 88/88이 통과했고 여섯 Python 모듈의
  compile도 통과했다. staged v1은 계속 사용 금지다.
- 989개 전체 실행은 아직 시작하지 않았다. v25/30 gate 전에 아래 다섯
  재현 가능한 의미 누락/오탐을 모두 차단해야 한다.
  1. outcome/factor/level/arm/unit/count-ratio 원문 라벨 역방향 coverage
  2. Result/Axis/Factor/Aggregate 역할 결속과 series header laundering 차단
  3. numerator/denominator, min/max, baseline/changed field 결속 및 source
     primary-cell 중복 소비 차단
  4. locator `NO_CANDIDATE`와 독립적인 conclusion heading/narrative 탐지
  5. legend/header/instruction과 실제 categorical status row 구분
- 위 다섯 항목은 최소 adversarial fixture, 실제 v24 불량 5/5 차단,
  안전 대조 7/7 통과를 모두 만족한 뒤에만 v25 실패 10개와 30개 품질
  gate로 진행한다. 이미지 분석, 재캡처, 재그룹화, WPF/Excel 실행은
  하지 않는다.

## 의미 안전 게이트 완료 — 2026-07-18

- 위 다섯 blind spot의 구현과 실제 artifact 회귀를 완료했다.
  semantic label/compound header, Result·Axis·Factor·Aggregate 역할,
  numerator/denominator·min/max·baseline/changed 결속, source-cell
  primary scalar 중복, locator-independent conclusion, legend/status
  geometry가 fail-closed 검증에 포함된다.
- 실제 오탐 원인이던 세로 result의 AXIS 오분류, `Change ...` 행 전체의
  CHANGED 오염, `Result` 섹션 제목과 `Improve`의 `imp` 부분일치,
  `OK` 열 헤더, `No sound`와 `No` 순번열 혼동, 동일 numeric alias의
  중복 slot 요구를 좁은 source geometry 규칙으로 교정했다.
- 숫자 header/row-identity axis는 선언 range가 아니라 Capture v2에
  실제 존재하는 정렬된 value cell 행이 2개 이상일 때만 허용한다.
  빈/invented valueRange로 numeric header를 숨기는 fixture는 실패한다.
- 최종 관련 회귀는 212/212 PASS다. 실제 v24 matrix도 불량
  B08/B09/B13/B14/B22 5/5 FAIL, 안전
  B10/B18/B20/B21/B26/B29/B30 7/7 PASS다. 안전 7건의 quantitative,
  categorical, narrative, formula, semantic, field binding 미해결 수는
  모두 0이다.
- 이 checkpoint로 의미 안전 단계는 완료됐고 다음 실행 단계는
  v25 실패 10건 재처리, 이어서 B24/B25를 포함한 30건 품질 gate다.
- 18:08(+07)에 `run_corpus_benchmark_retry_10_v25.ps1`을 숨김
  백그라운드로 시작했다. supervisor PID 42180, Python worker PID
  39368이며 selection 10건은 journal에서 모두 RUNNING이다. 로그는
  `outputs/corpus-ingest/full-989-v1/benchmark-retry-10-v25.stdout.log`
  및 `.stderr.log`, 결과는 `benchmark-retry-10-v25.result.json`이다.
  이미지 분석은 false이며 WPF/Excel/서버는 실행하지 않았다.

## 989개 원본 캡처 완료 / v25-v26 의미 분석 상태 — 2026-07-18

- v25 10건 재처리는 2 COMPLETED / 8 fail-closed로 종료됐다. 통과는
  B06/B13이며, 실패 artifact는 canonical DB에 반영되지 않았다.
- 느린 의미 AI와 빠른 원본 보존을 분리했다. OpenXML 병렬 reader 4개,
  source-order 직렬 DB import를 구현하고
  `run_corpus_capture_989_v26.ps1`로 전체 989건을 처리했다.
- 실행 시간은 18:32:33~18:41:50(+07), 총 9분 17초다.
  930 IMPORTED + 59 SHA 일치 SKIPPED + 0 FAILED이며 이미지 추출·분석은
  모두 false다. current Capture v2 revision은 정확히 989건이다.
- 전체 원본 SHA-256 재검증 989/989, Capture v2 989/989 valid,
  미완료 run 0, failed item 0, canonical bridge 누락 0,
  current duplicate 0, SQLite `quick_check=ok`를 확인했다.
- 실행 전 403,582,976-byte DB 백업:
  `outputs/universal-grid/backups/InputDataFinish.pre-capture-989-v26-20260718-1831.sqlite`
  현재 DB는 약 3.24 GB다.
- 이 완료는 원본 셀·병합·수식·스타일·좌표 DB화 완료를 뜻한다.
  989건 의미 분석 완료를 뜻하지 않는다. corpus journal은 현재
  33 COMPLETED / 9 FAILED / 947 PENDING이다.
- v26 결정 보정 중 B15는 쉼표로 합친 비연속 A1 근거 6개를 개별
  evidence로만 분리해 6 studies, required quantitative 6,546/6,546,
  semantic/binding gap 0으로 통과했고
  `ANALYSIS-683E8804F3FB`로 NEEDS_REVIEW 반영됐다.
- B07/B09/B22는 재처리 중 새 결함이 드러나 계속 차단했다.
  B07은 unsupported 합성 REFERENCE Arm, B09는 비교 제거 후 F6:J6
  수치 5개 누락, B22는 최신 초안의 `NG function!L7` 1개 누락이다.
- B14는 22시트 대형 초안이 수치 861, semantic 123, status 20,
  formula 1,234개를 빠뜨려 소규모 보정 불가다. B16은 실제 숫자
  163개를 queryable field가 아닌 복합 valueText로 뭉쳐 계속 차단한다.
- monolithic/staged 결정은 prompt bytes뿐 아니라 source cell 수도
  검사하며 2,000셀 초과 시 staged-v2로 전환한다.
- 재시도 journal은 attemptHistory에 이전 snapshot을 보존하고 현재
  attempt의 stages/result를 초기화한다. stage 실패 시 downstream은
  PENDING, result는 null이다. 현재 contract에서 기존 draft가 무효면
  같은 revision의 미검증 canonical analysis를 자동 STALE 격리하고,
  VERIFIED/review-decision 분석은 보호한다.

## v27 결정 보정 3건 완료 — 2026-07-18

- B07/B09/B22의 최신 실패를 source-proven 결정 보정으로 구현했다.
  B07은 split-cell 합성 Arm의 label/condition만 원문 Test/Normal로
  축소하고 role/factorValues/evidence는 보존한다. B09는 F6:J6 높이축을
  공통 Outcome 1개와 정렬된 RAW count series 3개로 정규화한다. B22는
  `NG function!L7`의 Hearing (+1V) Noise percentage Outcome 1개만
  복구한다.
- stale target보다 최신 rejected artifact의 결정 보정을 먼저 실행하고,
  실제 Codex subprocess 직전에만 `aiExecutedThisAttempt=true`가 되도록
  provenance 관찰 지점을 수정했다.
- 독립/통합 회귀는 semantic+workflow+B09+B22 77/77 PASS이며,
  content/staged-v2/CLI까지 합친 최종 집중 회귀는 130/130 PASS다.
  관련 Python 모듈 7개의 compile도 통과했다.
  `corpus-benchmark-deterministic-3-v27.json` 실제 실행은 33.8초,
  3 COMPLETED / 0 FAILED였으며 세 journal 모두
  `DRAFT.aiExecuted=false`, IMPORT/VERIFY COMPLETED다.
- DB 반영:
  B07 `ANALYSIS-88FCAF94EFA6` 3 studies,
  B09 `ANALYSIS-9E54710C735D` 4 studies/42 outcomes/11 series/165 points,
  B22 `ANALYSIS-7095FD904623` 6 studies/26 outcomes/25 series/3,328 points.
  모두 NEEDS_REVIEW이며 SQLite `quick_check=ok`다.
- corpus journal은 36 COMPLETED / 6 FAILED / 947 PENDING,
  결과는 3 EXCLUDED / 33 NEEDS_REVIEW / 953 NONE이다.
  남은 실패는 B04/B08/B14/B16과 대형 staged 대상 B24/B25다.
- 남은 6건 감사 결과:
  - B04는 D23:M23 서술 속 `50pcs`를 numeric valueNumber=50으로
    잘못 선언한 문제와 제목/헤더 semantic 4칸 누락이다.
  - B08은 B26 `=> Can use` 한 줄만으로는 결론 서술성이 부족하며
    B25:B26 두 줄의 정확한 결합 인용으로 전체 gate가 통과한다.
  - B14는 2,947 selected cells/24 chunks에서 quantitative 861,
    semantic 123, categorical 20, unresolved formula 1,234, binding 2
    누락이라 소규모 repair가 불가능하고 staged-v2 재구성이 필요하다.
  - B16은 composite valueText 속 154개, Sigma rate 9개, IR 식별자
    오분류 3개가 원인이다. 14 scalar outcomes/154 observations,
    3 rate outcomes/9 observations 추가와 IR identifier 제외가
    안전 경계이며 예상 358/358이다.
  - B24/B25는 각각 9,638/9,656 selected cells이고 default staged
    plan은 17/25 parts지만 finalized prompt가 400 KB를 넘는 part가
    12/4개다. `max-chunks=1`이면 29/31 parts, 최대
    353,267/389,814 bytes로 전 part가 budget 안에 든다.
  - B24/B25에는 재사용 가능한 fragment가 0개라 총 60개 신규
    fragment AI call이 필요하다. planner는 finalized prompt bytes
    기준 preflight로 영구 보강해야 한다.

## v28-v30 복구 및 운영 E2E 감사 — 2026-07-18

- B04/B08은 원본 범위 안에서 결정 보정했고 실제 재처리 2/2를 AI 호출
  없이 통과했다. B04의 서술 전용 `50pcs` 숫자 주장을 제거하면서
  문맥 셀을 보존했고, B08 결론 근거는 정확히 B25:B26으로 묶었다.
- B16은 복합 matrix 4개를 scalar outcome 14개/observation 154개로,
  Sigma rate를 outcome 3개/observation 9개로 정규화했다. 반복 IR
  식별자 E17/E19/E21만 제외하고 최초 IR 값은 queryable로 유지했다.
  실제 재처리 1/1도 AI 호출 없이 통과했다.
- 현재 corpus 의미 상태는 39 COMPLETED / 3 FAILED / 947 PENDING이다.
  남은 실패는 B14/B24/B25뿐이며 SQLite `quick_check=ok`다.
- staged-v2는 registry/shared anchor를 포함한 최종 prompt bytes로
  분할하고, worker 시작 전에 전 part를 preflight하며, UUID sibling
  transport와 실제 subprocess 직전 AI provenance를 사용한다. 집중
  회귀는 124/124 PASS다. 정확한 plan은 B24 28 parts/최대 353,267
  bytes, B25 29 parts/최대 389,814 bytes이고 ownership은 exact다.
- 실행 전 백업은
  `outputs/universal-grid/backups/InputDataFinish.pre-staged-b24-b25-v30-20260718.sqlite`
  이며 quick-check를 통과했다.
- B24/B25 첫 실제 staged-v2 실행은 fragment 생성·import 전에 종료됐다.
  nested object가 strict response-schema 규칙의
  `additionalProperties:false`를 만족하지 못해 API가 거부했다. 승인된
  fragment는 0개이며 원본과 canonical DB는 변경되지 않았다. strict
  transport schema 수정·검증 후 재시도한다.
- WPF Ask→CLI evidence-answer→query/answer→EVD 상세→정확한 Excel 범위,
  신규 XLSX→resumable ingest 경로는 코드상 연결돼 있다. 다만 앱 실행
  금지에 따라 WPF runtime 검증은 하지 않았다.
- 실제 DB에는 comparison 165개/effect 73개가 있어도 verified+eligible은
  0개이고 review queue가 120개다. 따라서 현재 VP+CD 조립/FUNCTION NG
  예시는 정량 관계를 만들지 않고 `INSUFFICIENT_COMPARISON`으로
  fail-closed한다.
- 임의 질문 일반화도 아직 부분 완료다. canonical concept 7개/alias
  28개뿐이고 미승인 schema candidate가 601개다. 다음 검색 vertical
  slice는 candidate를 사람이 기존/신규 concept로 승인·병합하는
  fail-closed API/CLI, related 검색의 동일 alias 사용, eligible effect
  답변의 실제 factor 변화 표시다.

## B14 수식 안전 경로와 개념 정규화 backend — 2026-07-18

- Capture v2를 변경하지 않는 제한 문법 수식 overlay를 구현했다.
  source revision/content, dependency, cycle, 미지원 문법, 비유한 값,
  checksum을 fail-closed 검증한다.
- 실제 B14 read-only projection 결과는 미캐시 수식 1,234 =
  숫자 1,149 + `#DIV/0!` 85다. unresolved는 1,234→0, required numeric은
  935→2,084이고 Capture formula cache NULL은 1,234/1,234 그대로다.
  staged plan은 24 parts, 최대 finalized prompt 357,901 bytes다.
- `derive_formula_values`/`--derive-formula-values` opt-in 경로가 overlay
  저장·재유도 검증, deep-copy packet projection, draft/content/import의
  동일 provenance 검증, `captureMutated=false` journal 기록을 연결한다.
  formula/import/workflow/CLI 집중 회귀는 Sol 재실행 기준 100/100 PASS다.
  실제 B14 AI/import는 아직 실행 전이다.
- 사람 결정형 concept/alias curation backend와 additive migration을
  구현했다. candidate/concept 조회, 원자적 CREATE/MERGE/REJECT, 불변
  resolution/alias approval 이력, exact idempotent replay를 지원한다.
  UNIT 오용, kind 불일치, 빈 값, inactive target, 다른 ACTIVE concept의
  alias 소유 충돌, conflicting replay는 모두 차단한다.
- JSON CLI는 `concept-candidates`, `concept-list`, `concept-resolve`,
  `concept-alias-upsert` 4개다. related 검색도 query와 같은 alias를
  사용하지만 similarity가 관계·인과 근거가 아니라는 계약은 유지한다.
  schema/curation/related/query/CLI/import 통합 회귀는 95/95 PASS다.
- eligible effect 답변은 이제 실제 factor 차이와 control→compared
  조건을 구조화 결과와 한국어 문장에 표시한다. 예를 들어
  `본드량: A → B` 뒤에 저장된 effect를 제시하며 없는 값은 만들지 않고
  `(미기록)`으로 둔다. query+answer 회귀 36/36 PASS다.
- 운영 DB에는 아직 `canonical-concept-curation-v1` migration이 없다.
  WPF가 자동 migration하면 안 되며, v31 종료 후 별도 백업과
  `knowledge-migrate`/`knowledge-inspect` 검증이 필요하다. 실제 후보
  승인·병합·거절은 수행하지 않았다.

## 대표30 gate 정직화와 WPF 정규화 탭 — 2026-07-18

- WPF의 사람 검토와 신규 적재 사이에 `개념·별칭 정규화` 탭을
  구현했다. OPEN 후보 최대 10,000개, 동일 kind ACTIVE concept,
  reviewer/note 필수, ID·값을 포함한 CREATE/MERGE/REJECT 확인,
  낙관적 UI 변경 금지, 저장 성공/새로고침 실패 구분,
  IdempotentReplay 표시를 포함한다. 프로젝트 build는 경고 0/오류 0,
  앱 실행은 하지 않았다.
- 최신 계약으로 대표 B01~B30을 읽기 전용 재검증하니 journal 완료와
  품질 통과가 달랐다. 실제 PASS는 18건:
  B01/B02/B03/B04/B07/B08/B09/B10/B13/B15/B16/B18/B20/B21/B22/B26/B28/B29.
- answer-visible NEEDS_REVIEW false-pass 9건:
  - B05/B19/B27: REFERENCE role 원문 근거 부족.
  - B06: `Test (2)!C23` 수치 누락.
  - B11: 수치 298개 누락 + 단일 selected chunk finalized prompt 400KB 초과.
  - B12: `201507!J10,J11` 의미 라벨 누락.
  - B17: 수치 586개, `Report!C109,C111` 결론 2개, 의미 2개 누락.
  - B23: `201507!F28` categorical status 누락.
  - B30: `Test!C17,C19,C21,C23` 수치 4개 누락.
- B14는 formula-safe 실제 실행 전, B24/B25는 staged-v31 RUNNING이다.
  따라서 대표 gate는 아직 30/30이 아니며 947 전체 처리는 금지한다.
- 대표30 원본 SHA/current Capture는 모두 정상, SQLite quick-check OK,
  canonical knowledge FK/invalid aggregation/orphan evidence는 0이다.
  단 전체 legacy `foreign_key_check`에는 기존 66건
  (analysis_evidence 48/conclusions 12/review_items 6)이 있어 이후에는
  전역 FK 0이라고 보고하지 말고 exact baseline 불변을 확인한다.
- 947 전 필수:
  B24/B25 종료·재감사 → false-pass 9건+B14 복구 → 최신 계약 30/30 →
  runtime exact hash 고정 → immutable backup/quick-check → PENDING 전용
  exact947 manifest hash → 25~50 canary → 크기 tier별 resume다.
  formula-free/지원 formula/미지원 formula는 별도 queue로 분리한다.

## 사용량 소진에 따른 정지 체크포인트 — 2026-07-18 21:24 +07

- 사용자 요청에 따라 추가 Codex/API 사용을 즉시 멈췄다. 실행 중이던
  구현 worker 2개는 interrupt했고, staged-v31 재시도 process tree
  (supervisor 42984 / Python 11892 / conhost 43920)도 종료했다. 종료 후
  세 PID가 남아 있지 않음을 확인했다.
- 앱, Excel, WPF, 서버, 이미지 분석은 실행하지 않았다.

완료된 코드:

- B11의 단일 source chunk가 registry/context 결합 후 400KB를 넘는
  문제를 범용 source-contiguous within-chunk segment로 해결했다.
  원본 chunkId/locator, exact sourceCellKey ordered union·유일 소유,
  merged/context shared allowlist, logical Study anchor, part/plan/provenance/
  resume identity를 보존한다. 단일 cell+context도 한도를 넘으면
  fail-closed한다.
- 실제 B11 read-only proof:
  source 3,323 cells, 17 parts, segmented parts 15, max prompt
  373,061/400,000B. ownership/order/reconstruction/locator identity 모두
  true. staged+workflow 30/30 및 compile PASS.
- `inference_data_ai_content_coverage.py`에 B05/B06/B12/B23/B30 공통
  판정 보정을 반영했다.
  - B05: `Total -> Position 1/2/3`에서 가까운 leaf를 우선해 J22를
    RAW result로 보존.
  - B06: merged record gap을 건너는 완전 증가 순번열 1/2/3 제외.
  - B12: merged factor matrix `S-MG -> Spec/Supplier` 아래 Normal을
    FACTOR_LEVEL로 처리하되 bare Spec/Supplier를 전역 factor로 만들지
    않음.
  - B23: 아래 값이 없는 vertically merged `OK` column header 제외.
  - B30: HEADER series의 exact multi-value rowIdentity를
    SERIES_REPLICATE_IDENTITY로 인정.
- 회귀는 content 47/47, content+workflow+import 111/111 PASS.
- 실제 저장 artifact 재검증:
  B05 q72/72 semantic4/4 categorical4/4,
  B06 q112/112,
  B12 q69/69 semantic14/14,
  B23 q124/124 semantic9/9 categorical0,
  B30 q64/64. 단 B05는 content coverage만 통과한 것이며 full strict
  통과 전 exact REFERENCE Arm 보정이 남아 있다.

감사 완료·구현 대기:

- B05/B19: Line/공정 등 나머지 구성요소는 factorValue로 보존하고
  Arm identity만 exact source Test/Normal로 원자적으로 축약해야 한다.
- B27: Type 1~4를 exact factor/factorValue로 추가하고 factor-level
  Normal과 빈 요약행 Normal을 구분해야 한다.
- B17: `Report!C15:O107` 45 LOT×12 outcomes 전체가 실제 누락됐다.
  deterministic no-AI projector의 메모리 증명은 observation 540개,
  q1117/1117, semantic2/2, categorical44/44, narrative6/6, binding0이다.
  C109:C110과 C111:C112 결론 문맥도 같이 보존한다.
- 위 Arm/B17 구현 worker는 정지 요청으로 중단했고 새 모듈 파일을
  만들기 전이었다. 다음 세션에서 다시 맡긴다.

B24/B25 중단 지점:

- v31 first pass 승인 조각은 B24 23/28, B25 24/29.
- 당시 missing index:
  B24 `[2,3,20,27,28]`,
  B25 `[2,3,11,21,29]`.
- 같은 v31 resume를 21:23에 시작했으나 사용자 정지 요청으로 process
  tree를 종료했다. 종료 후 다시 계산한 결과 accepted/missing은 위
  숫자와 같고 in-flight artifact는 0개다.
- 강제 정지 때문에 corpus journal에는 B24/B25가 RUNNING으로 남아
  있다(전체 39 COMPLETED / 2 RUNNING / 1 FAILED / 947 PENDING). 실제
  writer가 있다는 뜻이 아니므로 다음 resume가 journal 상태를
  reconcile하게 한다.
- live writer가 없음을 먼저 확인한 뒤 같은
  `run_corpus_staged_b24_b25_v31.ps1`만 재실행한다. accepted fragment는
  검증 후 재사용되고 missing part만 호출되어야 한다.

다음 재개 순서:

1. B24/B25 part inventory·live PID 확인 후 v31 resume.
2. exact Arm 보정과 B17 deterministic projector 구현·집중 회귀.
3. false-pass 9건 재처리와 최신 계약 재감사.
4. B14 formula-safe 실제 실행 후 대표30 30/30 증명.
5. runtime hash freeze, DB backup/quick-check/legacy FK66 불변 확인,
   concept migration.
6. exact PENDING947 manifest → 25~50 canary → size/formula tier별 resume.
