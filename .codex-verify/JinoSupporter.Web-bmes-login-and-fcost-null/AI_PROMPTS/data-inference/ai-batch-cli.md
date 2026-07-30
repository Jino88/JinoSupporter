Read AI_PROMPTS/data-inference/ai-excel-proc.md and run the AI Batch CLI workflow.
DB={{dbPath}}.
Targets={{fileName}} ({{targetCount}} datasets).
Mode={{modeHint}}.
Current mandatory policy:
- Analyze each dataset/file independently with the same reportReviewMatrix,
  visualization-decision, and result-report rules used by INPUT DATA.
- For each workbook/report block, state what was reviewed, why it was reviewed,
  what result is visible, and whether it is `공정 불량 검토`, `기능 불량 검토`,
  `공정+기능 연계 검토`, or `기타/미확인`.
- Every NG-rate, defect-rate, yield, PPM, OK/NG ratio, or below-spec comparison
  must have both a numeric table and a `report-bars` visual block. In the UI
  this block is rendered as a vertical bar chart, so keep exact numeric values
  and numeric `amount` fields.
- Use `Normal`, `Normal 대비`, `Normal 값`, or `Normal 미확인` in visible report
  text. Do not display `Local Control`, `Control`, `Baseline`, or `대조군` as
  visible labels.
- Speed: classify workbook/report blocks first, then re-read only the relevant
  sheet/table blocks needed for the final judgement and visual block.
The launcher already set AI_BATCH_DB_PATH and AI_BATCH_TARGETS_FILE.
Use _ai_batch_helper.load_targets() and verify it returns exactly {{targetCount}} non-empty names.
Process the target list sequentially in this one CLI session. Do not open extra terminals.
For each dataset, open the DB, call _ai_batch_helper.get_excel_text(con, dataset), and analyze every workbook and every worksheet in workbook order.
Do not analyze only the first visible sheet. If the rendered text is too large, split by '=== SHEET:' and analyze sheet by sheet, then commit one combined result.
Use _ai_batch_helper.get_excel_files(con, dataset) only when workbook file context is required. Use get_excel_paste only as a last-resort fallback.
Do not run or import _batch_auto.py, _batch_build.py, _tmp_ai_excel/auto_normalize.py, or any heuristic fallback/auto-normalizer.
Do not commit placeholders such as primary_defect='NG (auto-extracted)', changed_factor='see workbook title/purpose', 'Workbook stored but extraction surfaced narrative only', 'batch inventory', or measurement_type='inventory'.
If real extraction is not possible for a dataset, call _ai_batch_helper.log_failed(dataset, reason) and leave existing DB rows untouched.
Store AiResults and AiNgBreakdowns from all relevant sheets and set sheet_name/source_cells on every extracted row so Dataset Results can prove sheet coverage.
Merged-cell rule: if Date/Model/Type/Line/Process cells are merged or visually carried across rows, all covered rows inherit that value before pairing Test/Normal rows.
Do not create standalone AiResults from percentage-only subrows; attach them as breakdown/rate evidence to the preceding count row.
Set document.report_type to exactly one allowed AI_EXCEL_PROC report_type.
Purpose priority: first, create an AI-authored analysis report that lets the user quickly decide what to check/change next using an action-board table; second, store structured DB parameters so future AI ASK can quickly find related evidence.
Do not infer document.title, document.purpose, primary_defect, related_defects, parts, processes, or other search labels from dataset/file names alone. Use workbook Purpose/Objective, section headers, and result evidence first; use filename only as product/model fallback.
Create result.generated_report_markdown as the primary user-facing Korean Markdown report. This report is the product. It must be written by AI from workbook contents, not assembled by the UI from parameters.
Follow AI_PROMPTS/data-inference/ai-excel-proc.md Target Report Style and Required Report Shape. The first section must be a practical action board table: Priority | Check/change item | Evidence/result | Judgement | Next action.
The report should look like a manufacturing review sheet: action board first, then defect phenomenon, review items, result table, lot/mold/line comparison, heatmap/comparison matrix, judgement, action, and evidence cells.
Do not write a parameter dump or long narrative. When result rows exist, each report needs compact Markdown sections, a compact result/evidence table, source worksheet/cell references, and only enough prose to support the action board.
When result rows exist, result.generated_report_markdown and every translated report must include at least one AI-authored visual block using fenced code blocks named report-heatmap, report-bars, or report-heatmap-matrix. For NG-rate/defect-rate/yield/PPM comparisons, use report-bars plus the result table. For two-factor matrices, use report-heatmap-matrix. For non-paired rankings without a rate comparison, report-heatmap is allowed. status must be bad, warn, or good.
Also create tr_ko.document.generated_report_markdown, tr_en.document.generated_report_markdown, and tr_vi.document.generated_report_markdown. These are required for the Korean, English, and Vietnamese UI language buttons. Keep equivalent visual blocks in all three translations; translate label/detail text but preserve numeric values, amount, and status.
Do not stop after extracting parameters. commit_dataset rejects parameter-only payloads when generated_report_markdown is empty.
commit_dataset also rejects missing translations, reports that are too short, reports without enough sections, result-row reports without Markdown tables/visual blocks, reports without worksheet/cell evidence, and broken encoding/mojibake output.
Use test_conditions/results/conclusions/troubleshooting as DB/search parameters only; the user-facing report is generated_report_markdown and its translated generated_report_markdown values.
Extract enough fields for a per-workbook mini report: basic info, problem/phenomenon, purpose, changed/tested conditions, result data, Normal/Test or Before/After comparison, writer decision/note, AI judgement/warnings, and evidence location.
If no same-event baseline exists, use ng_without_baseline ranking and do not claim improvement/worsening.
Do not put large result/tr_ko/tr_en/tr_vi payloads inside python -, PowerShell here-strings, Set-Content command text, or any shell command; Windows command length error 206 will fail the run.
Commit by writing a JSON payload file under _batch_tmp with shape {name,result,translations:{ko,en,vi}}, then run python _ai_batch_helper.py commit-json <that-file>. Create the JSON file with file-writing/editing capability, not by embedding the full JSON in a shell command.
Before finishing, verify success_count + failed_count == {{targetCount}}. Delete {{fileName}} only after that verification passes.
