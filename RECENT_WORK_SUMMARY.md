# JinoSupporter Recent Work Summary

Checked at: 2026-07-06 06:29:54 +07:00

## Basis

- Repository: `D:\000. MyWorks\005. Program\Repository\JinoSupporter`
- Current Git HEAD: `13e88b0` (`123`, committed 2026-05-19 16:47:49 +0700)
- Working tree status at check time:
  - Tracked changed/deleted/renamed files, excluding `bin` and `obj`: 103
  - Untracked files, excluding `bin` and `obj`: 104
  - Main changed areas: `JinoSupporter.Web/Components/Pages`, `JinoSupporter.Web/Services`, `BmesNgRateStandalone/Components/Pages`, `BmesNgRateStandalone/Services`, and `AI_PROMPTS`
- Existing detailed handoff MD found:
  - `ASK_AI_REVIEWCASE_NEXT_STEPS.md` last modified 2026-07-04 16:53:17
  - `BMES_FCOST_DB_NOTES.md` last modified 2026-06-17 15:29:35
  - `CHANGELOG.md` exists but is only updated through 2026-05-25 release notes

## Short Answer

The latest documented stop point is the Ask AI / MicroSpeaker `ReviewCase` work. The system can already search related evidence, extract Normal/Test style comparison evidence, aggregate repeated comparison rows, and pass structured evidence into Ask AI, but it does not yet have the normalized `ReviewCase` layer that answers "when a process/material/jig/supplier/condition changed, what happened?" as a first-class model.

The next meaningful implementation is to build `MicroSpeakerReviewCaseService` or equivalent, generate inspectable ReviewCase samples, then make Ask AI search that layer first and fall back to raw pair/metric/measurement rows.

## Progress After This Summary

### 2026-07-06

- Added first diagnostic ReviewCase extractor:
  - `JinoSupporter.Web/Services/MicroSpeakerReviewCaseService.cs`
- Added DI registration and JSON sample endpoint:
  - `GET /microspeaker/review-cases/sample.json?limit=80`
  - optional query filter: `q=...`
- Added a `ReviewCase JSON` link to the MicroSpeaker Result page.
- Current behavior:
  - Reads MicroSpeaker SQLite in read-only mode.
  - Emits `reviewCases` as a hierarchy: source report -> changed factors -> linked outcomes.
  - Keeps `flatCandidates` as the raw diagnostic extraction list for rule tuning.
  - Converts `comparison_pairs` into normalized pair outcomes.
  - Converts paired `measurement_stats` Normal/Test rows into measurement outcomes.
  - Preserves process NG, function NG, and measurement outcomes as separate outcome records under the same source review.
  - Uses generic evidence structure: source title, table headers, Normal/Test condition labels, and before/after contrast; no specific workbook title or example part/process name is hard-coded.
  - Does not create or modify MicroSpeaker DB tables yet.
- Verification:
  - `dotnet build .\JinoSupporter.Web\JinoSupporter.Web.csproj --no-restore -p:UseAppHost=false -p:OutputPath=.\artifacts\codex-check\reviewcase\`
  - Build succeeded with existing unrelated warnings only; no new warnings from `MicroSpeakerReviewCaseService.cs`.

### 2026-07-06 ReviewCase AI Packet Update

- Reframed ReviewCase creation as AI-led analysis:
  - AI receives extracted workbook rows/cells plus candidate hints.
  - AI creates `changedFactors`, `outcomes`, `evidenceRows`, and verification.
  - Candidate pairs/metrics/measurements are hints only, not final labels.
- Added prompt:
  - `AI_PROMPTS/data-inference/reviewcase-ai-analysis.md`
- Added user audit decision log:
  - `REVIEWCASE_AI_AUDIT_DECISIONS.md`
- Added file-level AI input packet endpoint:
  - `GET /microspeaker/review-cases/ai-packet/{fileId}.json`
  - optional limits: `rowLimit=1200`, `candidateLimit=300`
  - reads MicroSpeaker SQLite in read-only mode and does not modify DB tables.
- Verification:
  - `dotnet build .\JinoSupporter.Web\JinoSupporter.Web.csproj --no-restore -p:UseAppHost=false -p:OutputPath=.\artifacts\codex-check\reviewcase-ai-packet\`
  - Build succeeded with 26 existing unrelated warnings and 0 errors.

### 2026-07-06 Full Excel ReviewCase Batch

- Added batch generator:
  - `tools/generate_reviewcase_batch.py`
- Generated ReviewCase pre-analysis drafts for all 989 MicroSpeaker Excel files:
  - manifest: `REVIEWCASE_AI_DRAFTS/batch/reviewcase_batch_manifest.json`
  - summary: `REVIEWCASE_AI_DRAFTS/batch/reviewcase_batch_summary.md`
  - file drafts: `REVIEWCASE_AI_DRAFTS/batch/files/*.reviewcase-draft.json`
- Batch result:
  - 922 files: `needs_ai_verification`
  - 60 files: `needs_review`
  - 7 files: `excluded`
  - 1 file has a manual confirmed draft: `REVIEWCASE_AI_DRAFTS/721.reviewcase-draft.json`
- Verification:
  - `python -m py_compile .\tools\generate_reviewcase_batch.py`
  - Full batch JSON parse check passed for 989 draft files.
  - Evidence row existence check passed with 0 missing evidence references.
- Scope:
  - The batch reads MicroSpeaker SQLite in read-only mode.
  - It does not modify MicroSpeaker DB tables or source Excel files.
  - Batch drafts are not final verified ReviewCases; they are the next input for
    AI/user verification.

### 2026-07-06 Full AI Verification

- Added AI verification runner:
  - `tools/verify_reviewcase_ai_batch.py`
- Ran AI verification over all 989 MicroSpeaker Excel file drafts:
  - manifest: `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_manifest.json`
  - summary: `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_summary.md`
  - file results: `REVIEWCASE_AI_DRAFTS/verified/files/*.reviewcase-ai-verification.json`
- Final verification counts:
  - 13 files: `verified`
  - 917 files: `needs_review`
  - 59 files: `excluded`
- Verified file IDs:
  - `219`, `330`, `442`, `477`, `553`, `711`, `721`, `775`, `802`, `854`, `864`, `893`, `936`
- Verification checks:
  - 989 expected file IDs, 989 verification files present.
  - 0 missing file IDs.
  - 0 invalid verification JSON files.
  - Re-run check returned `processed=0`, meaning no unprocessed entries remain.
- Token usage recorded in the verification manifest:
  - input tokens: 5,850,135
  - output tokens: 729,008
  - total tokens: 6,579,143
- Important interpretation:
  - `verified` cases can be the initial Ask AI ReviewCase evidence layer.
  - `needs_review` means the AI rejected automatic approval and left issues,
    questions, or correction steps in the per-file verification JSON.
  - No MicroSpeaker SQLite tables or source Excel files were modified.

### 2026-07-07 Ask AI Verified ReviewCase Layer

- Wired approved verified ReviewCases into the deterministic Ask AI evidence
  pack:
  - service: `JinoSupporter.Web/Services/MicroSpeakerAskEvidenceService.cs`
  - JSON field: `microSpeakerEvidence.microSpeaker.verifiedReviewCases`
  - source: `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_manifest.json`
    plus approved per-file verification and draft JSON
- The loader supports both batch draft schema and the manually corrected file
  721 draft schema, preserves changed factors, outcomes, evidence rows,
  limitations, source decision text, verification status, and source links.
- Updated Ask AI prompt contracts:
  - `AI_PROMPTS/data-inference/ask-ai-cli.md`
  - `AI_PROMPTS/data-inference/cli-ask-ai.md`
- Verification:
  - `dotnet build .\JinoSupporter.Web\JinoSupporter.Web.csproj --no-restore -p:UseAppHost=false -p:OutputPath=.\artifacts\codex-check\reviewcase-ask-evidence\`
  - Build succeeded with 26 existing unrelated warnings and 0 errors.

## Recent Work By Area

### Ask AI / MicroSpeaker ReviewCase

- Saved a detailed handoff in `ASK_AI_REVIEWCASE_NEXT_STEPS.md`.
- Captured the current user decision: do not hard-code examples like VP+CD, Coil+CD, Suspension D, TIN, Magnet, or specific process names.
- Recorded calibration examples:
  - New VP+CD assembly equipment: classify as equipment/process validation and keep process NG and function NG separately.
  - VP+CD and Coil+CD bonding amount: detect multiple changed factors from title/table/row evidence.
  - Suspension material plus tin plating method: keep SPOT process defect, tensile strength, and function defect as separate outcome groups.
- Current pause point:
  - Continue from calibration and ReviewCase extraction design.
  - Re-open or continue reviewing `00.BRS-161016 Report test VP all mold from supplier_1778463079_1778470471_clean.xlsx`.
  - Ask one source-file question at a time after opening the original Excel workbook.
- `MicroSpeakerAskEvidenceService.cs` now builds deterministic `microSpeakerEvidence`, including:
  - `questionAnalysis`
  - `pairConditionAggregates`
  - source links such as `/microspeaker/source-file/{fileId}`
  - aggregate/Total-row preference before individual daily rows
- Ask AI prompts were updated to use structured evidence first:
  - `AI_PROMPTS/data-inference/ask-ai-cli.md`
  - `AI_PROMPTS/data-inference/cli-ask-ai.md`

### Data Inference / Input Data

- Added or expanded MicroSpeaker and input-data pages:
  - `JinoSupporter.Web/Components/Pages/InputDataBatchPage.razor`
  - `JinoSupporter.Web/Components/Pages/MicroSpeakerResultPage.razor`
  - `JinoSupporter.Web/Components/Pages/DataInferenceAskPage.razor`
  - `JinoSupporter.Web/Components/Pages/DataInferenceDbPage.razor`
  - `JinoSupporter.Web/Components/Pages/DataInferenceModelAnalysisPage.razor`
- Added service layer for MicroSpeaker DB and Ask evidence:
  - `JinoSupporter.Web/Services/MicroSpeakerInputDataService.cs`
  - `JinoSupporter.Web/Services/MicroSpeakerAskEvidenceService.cs`
  - `JinoSupporter.Web/Services/InputDataTestBatchExtractor.cs`
  - `JinoSupporter.Web/Services/InputDataTestAnalysisSupport.cs`
- Added Program endpoints for MicroSpeaker dashboards and source-file downloads:
  - `/microspeaker/dashboard/{kind}`
  - `/microspeaker/source-file/{fileId}`
  - `/data-inference/ask-history/{id}/html`
- Added many prompt files under `AI_PROMPTS` and included them in `JinoSupporter.Web.csproj` so they copy to output/publish.

### Daily Test Data / Current Problem Workflow

- Added Daily Test Data input workflow:
  - `JinoSupporter.Web/Components/Pages/DailyTestDataInputPage.razor`
  - `JinoSupporter.Web/Services/DailyTestExtractionSettingsService.cs`
  - `JinoSupporter.Web/Services/DailyTestDataCliRecoveryService.cs`
  - `JinoSupporter.Web/Services/DailyTestReportBuilder.cs`
  - `daily-test-extraction-settings.json`
- Added Current Problem analysis workflow:
  - `JinoSupporter.Web/Components/Pages/CurrentProblemAnalysisPage.razor`
  - `JinoSupporter.Web/Components/Shared/CurrentProblemWorkflowStrip.razor`
  - `JinoSupporter.Web/Services/CurrentProblemAnalysisService.cs`
  - `ai_current_problem_analyze.py`
  - `ai_first_pass_classify.py`
  - `create_current_problem_search_html.py`
- Added prompt categories for daily test data and input-data analysis.

### AI Provider / Prompt / Translation Support

- Added or expanded:
  - `JinoSupporter.Web/Services/CodexApiService.cs`
  - `JinoSupporter.Web/Services/CodexUsageScraper.cs`
  - `JinoSupporter.Web/Services/AiProviderSettingsService.cs`
  - `JinoSupporter.Web/Services/AiPromptRegistry.cs`
  - `JinoSupporter.Web/Components/Pages/AiPromptPage.razor`
- Added AI usage/admin visibility and prompt folders:
  - `AI_PROMPTS/claude`
  - `AI_PROMPTS/data-inference`
  - `AI_PROMPTS/daily-test-data`
  - `AI_PROMPTS/input-data`
  - `AI_PROMPTS/translation`
- Expanded translation page/service work with Codex/OpenAI prompt files and OCR translation prompt variants.

### BMES F-COST / Material / Report

- Added detailed BMES DB exploration notes in `BMES_FCOST_DB_NOTES.md`.
- Expanded F-COST page and services:
  - `JinoSupporter.Web/Components/Pages/BmesFCostPage.razor`
  - `JinoSupporter.Web/Services/BmesFcostActualService.cs`
  - `JinoSupporter.Web/Services/FCostReportService.cs`
  - `JinoSupporter.Web/Services/FCostService.cs`
  - `JinoSupporter.Web/Services/FCostRawBreakdownExcelExporter.cs`
- Added BMES reporting pages:
  - `JinoSupporter.Web/Components/Pages/BmesReportPage.razor`
  - `JinoSupporter.Web/Components/Pages/BmesCauseMonthlyReportPage.razor`
  - `JinoSupporter.Web/Components/Pages/BmesTest3Page.razor`
  - `JinoSupporter.Web/Components/Pages/BmesTest4Page.razor`
- Added process/material mapping and NG services:
  - `JinoSupporter.Web/Services/ProcessMaterialMappingService.cs`
  - `JinoSupporter.Web/Services/ProcessMaterialNgService.cs`

### NG Rate / Reporting

- Split or shared NG Rate UI components:
  - `NgRateModelGroupPicker.razor`
  - `NgRateReportStyles.razor`
  - `NgRateSetupPanel.razor`
  - `NgRateSimpleGroupPicker.razor`
  - `NgRateViewNav.razor`
- Added/expanded daily and weekly report pages:
  - `JinoSupporter.Web/Components/Pages/NgRateForDailyReportPage.razor`
  - `JinoSupporter.Web/Components/Pages/NgRateForWeeklyReportPage.razor`
  - `JinoSupporter.Web/Components/Pages/NgRateAllPage.razor`
  - `JinoSupporter.Web/Components/Pages/NgRateByGroupPage.razor`
  - `JinoSupporter.Web/Components/Pages/NgRatePage.razor`
- Added CSV/export support:
  - `JinoSupporter.Web/Services/NgRateCsvExporter.cs`
  - `JinoSupporter.Web/Services/CsvExportUtility.cs`
  - changes in `NgRateExcelExporter.cs`
  - `NgRateModeSupport.cs`

### Standalone App Sync / Update

- Synced many Web-side BMES/NG Rate changes into `BmesNgRateStandalone`.
- Added standalone compatibility/support files:
  - `BmesNgRateStandalone/Services/StandaloneCompatibilityModels.cs`
  - `BmesNgRateStandalone/StandaloneErrorLog.cs`
  - standalone copies of shared NG Rate components and services
- Added or adjusted standalone update/build workflows:
  - `.vscode/launch.json`
  - `.vscode/tasks.json`
  - `BmesNgRateStandalone/installer.iss`
  - `BmesNgRateStandalone/tools/SyncBmesFromWeb.ps1`
- Added standalone sync endpoints in `Program.cs`, including BMES materials sync.

### Admin / Operations

- Expanded app startup and request logging through `AppActivityLogger`.
- Added or expanded admin/tooling pages:
  - `AdminAiUsagesPage.razor`
  - `AdminDbQueryPage.razor`
  - `AdminPathsPage.razor`
  - `TestExcelConverterPage.razor`
- Updated `.gitignore` for local temporary AI prompt/launch scripts, pid files, temp folders, and artifacts.
- Removed template/demo pages from the Web app:
  - `Counter.razor`
  - `Weather.razor`

## Current Risks / Notes

- The working tree is very large and not committed after `13e88b0`; treat it as active work in progress.
- There are many generated `bin`/`obj` files from prior builds, but they were excluded from this summary.
- Some navigation icon text is visibly mojibake in the current file content and should be cleaned separately if UI polish is needed.
- `CHANGELOG.md` does not reflect the July work; this file and `ASK_AI_REVIEWCASE_NEXT_STEPS.md` are the current handoff documents.
- No app/dev server was launched during this check.

## Suggested Next Steps

1. Implement `MicroSpeakerReviewCaseService.cs` as a small diagnostic service first.
2. Generate a review-case sample export, for example `tmp/review_cases_sample.json` or CSV/HTML, and manually inspect 30-50 cases.
3. Use the three saved calibration examples as regression checks, without hard-coding the example part/process names.
4. Wire Ask AI to search ReviewCase evidence first, then fall back to `pairRows`, `metricRows`, `measurementRows`, and raw rows.
5. Only after the evidence layer is acceptable, tune the Ask AI HTML/visual output.
