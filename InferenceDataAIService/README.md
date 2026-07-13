# InferenceDataAIService

Standalone CLI workspace for turning mixed Excel report files into source-backed SQLite/JSON outputs before any JinoSupporter Web integration.

## Default Input

`D:\000. MyWorks\test\result\InputDataFinish`

## Output Location

All generated outputs are kept under this folder:

`D:\000. MyWorks\005. Program\Repository\JinoSupporter\InferenceDataAIService`

## Source-of-Truth Storage

For hundreds of mixed Excel layouts, use the **universal-grid** SQLite DB as the raw source of truth.

- It stores each workbook as a fixed UsedRange grid: workbook, worksheet, row, column, address, displayed/raw value, and merge anchor/covered-cell metadata.
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

Import an evidence-linked reusable analysis summary from a layout-independent manifest:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-import --input .\JinoSupporter\InferenceDataAIService\outputs\analysis-manifests\BRS2015_G06_0003_analysis.json
```

Validate every stored analysis summary (source freshness, evidence ranges, ppm arithmetic, comparison deltas, and conclusion evidence):

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-verify --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite
```

Inspect or export one analysis for another dashboard/AI workflow:

```powershell
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-inspect --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --report-id 1
python .\JinoSupporter\InferenceDataAIService\inference_data_ai_cli.py analysis-export --db .\JinoSupporter\InferenceDataAIService\outputs\universal-grid\InputDataFinish.sqlite --report-id 1
```

The manifest contract is intentionally layout-independent: each review can declare any number of cohorts, scoped metrics, pairwise comparisons, conclusions, and source ranges. This allows Normal/Test, supplier variants, before/after, and multi-sheet reports to reuse the same DB model.

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
