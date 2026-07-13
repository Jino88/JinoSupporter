# InferenceDataAIService session handoff — 2026-07-13

## Current deliverable

The active desktop application is the WPF operator UI:

`InferenceDataAIService.Wpf/publish-async-utf8/InferenceDataAIService.Wpf.exe`

It is a WPF shell over the existing Python CLI and AI runner. Keep the Python
files beside the `InferenceDataAIService` folder; the WPF executable discovers
that folder by walking upward from its own location (or uses
`INFERENCE_DATA_AI_SERVICE_DIR`).

## Implemented WPF behaviour

- **Folder add** recursively lists Excel files.
- The left grid shows file name, universal DB state, analysis state, and
  current progress.
- Right-click a selected file to start `com-index` followed by the AI runner,
  or display its HTML result.
- The right pane renders the stored HTML analysis.
- The bottom pane shows the Python/Codex log.
- Folder scanning, subprocess execution, and log capture are asynchronous.
  Logs are UTF-8 and batched every 120 ms so Codex output does not freeze the
  WPF UI.

## Analysis behaviour

- Curated CLI manifests in `outputs/analysis-manifests/*_analysis.json` are
  the source of truth for their matching Excel source paths. The AI runner
  preserves or reuses them rather than replacing them with an AI guess.
- New Excel files use the curated manifests as calibration examples. The
  prompt requires source-backed cohorts, matched denominators, ppm/delta
  calculations, evidence ranges, and `NEEDS_REVIEW` for unresolved grouping.
- The HTML renderer now displays each Test/Normal cohort label and its detailed
  condition, not only repeated metric names and values.
- Program-generated draft reports 8 and 9, with their packet/manifest/HTML
  artifacts, were deleted. Curated reports 1–4 remain and their HTML files
  were regenerated in `outputs/analysis-rendered`.

## Verification completed

- `python -m unittest discover -s tests -v`: 12 tests passed.
- `python -m py_compile inference_data_ai_cli.py inference_data_ai_ui.py inference_data_ai_analysis_runner.py`: passed.
- `python inference_data_ai_cli.py analysis-verify --db outputs/universal-grid/InputDataFinish.sqlite`: 4 reports valid, 0 invalid.
- `dotnet build InferenceDataAIService.Wpf/InferenceDataAIService.Wpf.csproj -c Release --no-restore`: passed with 0 warnings and 0 errors.

## Important files

- `inference_data_ai_analysis_runner.py`: calibrated runner, reusable HTML renderer.
- `inference_data_ai_ui.py`: prior Python Tk UI; retained for reference only.
- `InferenceDataAIService.Wpf/MainWindow.xaml(.cs)`: active WPF screen.
- `InferenceDataAIService.Wpf/App.xaml`: starts `MainWindow`.

## Do not do

- Do not delete or overwrite curated CLI manifests/reports 1–4.
- Do not launch a server; the product is a local WPF + Python CLI workflow.
