# AI Excel Process Report Contract

This file is the source of truth for AI Batch, Excel report analysis, and
Dataset Results rendering.

## Current Mandatory Policy (2026-06-03)

Use the same analysis basis as INPUT DATA current problem analysis:

- Build a reportReviewMatrix for every workbook/report block before writing the
  final report. The matrix must identify reviewed item, review purpose, visible
  result, review domain, process evidence, function evidence, recommended use,
  weak evidence, and confidence.
- Review domain must be exactly one of `공정 불량 검토`, `기능 불량 검토`,
  `공정+기능 연계 검토`, or `기타/미확인`.
- The report must make it obvious whether the next work is process defect
  review, function defect review, or both.
- Decide the visualization for each file from workbook evidence before writing
  the report: vertical bar, scatter/dot, heatmap/matrix, Pareto/ranking, trend,
  or `확정 불가`.
- Every NG-rate, defect-rate, yield, PPM, OK/NG ratio, or below-spec comparison
  must include both a numeric table and a `report-bars` visual block.
- Measurement/spec sample values must use scatter/dot logic in the written
  analysis: preserve raw points when available and report spec/avg/min/max.
- Two independent variables with one result metric and at least four
  combinations must use a heatmap/matrix.
- Visible labels must say `Normal`, `Normal 대비`, `Normal 값`, or `Normal 미확인`.
  Do not display `Local Control`, `Control`, `Baseline`, or `대조군`.
- Speed: after the matrix and visualization decision, inspect only the
  worksheet/table blocks needed for evidence and commit. Do not repeatedly
  summarize unrelated sheets.

## Purpose

The analysis has two jobs, in this order:

1. Create an AI-authored analysis report that lets the user understand the
   workbook review result quickly. The report must be easy to read and must use
   tables, vertical bar visuals, heatmaps, comparison matrices, and concise
   engineering judgement.
2. Understand the same report deeply enough to store structured DB parameters
   for later AI ASK retrieval. These DB fields exist so future questions can
   quickly find related evidence, defects, conditions, and conclusions.

The visible product is the AI-authored report in `generated_report_markdown`.
Structured fields are secondary DB/search data. The UI must show the AI-authored
report first and must not assemble the visible report from DB parameters.

If there is a tradeoff, optimize for a useful, readable engineering report first
and then fill the DB/search fields from that understanding.

## Target Report Style

Write reports in the style of a good manufacturing review sheet:

- Keep the original technical-report flow: `Purpose`, `Content`, `Result`.
- Start with a practical action board table so the reader immediately sees
  what to check/change, what evidence supports it, and the next action.
- Start with a clear judgement, but still preserve the original report's
  investigation story.
- Use section bands and concise headings: purpose, current phenomenon, analysis
  and review items, result, comparison, judgement, action.
- Convert important workbook tables into readable Markdown tables.
- Rebuild workbook charts as AI-authored heatmaps or comparison matrices when
  the underlying values are available.
- Rebuild defect-rate/yield/PPM comparisons as `report-bars` visuals plus
  numeric tables.
- Describe embedded photos/charts when their captions or surrounding text
  explain the defect. If the image itself is essential and cannot be read from
  extracted text, classify the report as `image_dependent` and state the image
  review limitation.
- For lot, mold, supplier, line, or module comparisons, show the reader which
  item is worst, which is acceptable, and why.
- For defect/root-cause reports, show:
  1. defect phenomenon
  2. suspected cause
  3. checked evidence
  4. management/spec criterion
  5. final judgement and action
- Avoid generic prose. Every judgement must point to a number, row, chart
  label, note, or worksheet/cell evidence.

## Non-Negotiable Output Contract

For every dataset that can be analyzed, create:

1. `result.generated_report_markdown`: Korean user-facing Markdown report.
2. `tr_ko.document.generated_report_markdown`: Korean report for the UI button.
3. `tr_en.document.generated_report_markdown`: English report for the UI button.
4. `tr_vi.document.generated_report_markdown`: Vietnamese report for the UI
   button.
5. Structured `document`, `test_conditions`, `results`, `conclusions`,
   `troubleshooting_index`, and `ai_extraction_log` fields for DB/search only.

Do not commit a payload that only extracts parameters. Do not commit a short
summary that reads like a database note. The Markdown report must be a complete
analysis report that a process engineer can read without opening the DB fields.

## Required Report Shape

Write every user-facing report with these sections, in this order. Use the
target language for headings and prose.

1. Practical action board
   - Start with a compact Markdown table.
   - Required semantic columns: Priority, Check/change item, Evidence/result,
     Judgement, Next action.
   - The check/change item must be the factor to inspect or adjust, not the
     defect name itself.
   - Keep each row short enough to scan.

2. Decision summary
   - Give the final judgement first.
   - State pass/fail/use/hold/improve/retest when the workbook supports it.
   - Include the most important numeric reason.

3. Review target
   - Report/workbook name, model, date or period, line, process, lot, supplier,
     mold/jig/material when available.
   - Mention all worksheets used when the workbook has multiple sheets.

4. Purpose and test design
   - Explain what was being tested and what changed.
   - Preserve the workbook's own `Purpose` and `Content` intent when present.
   - Identify Normal rows when they exist. Source labels such as Control or
     Baseline may be mapped internally, but visible text must say Normal.
   - If no Normal exists, explicitly say the report is a ranking or absolute
     result review, not an improvement claim.

5. Current phenomenon and review items
   - Summarize the defect phenomenon, symptoms, suspect process/part, and
     review items.
   - For image-heavy reports, summarize what the surrounding labels/captions say
     and mark image-dependent uncertainty when needed.

6. Key result table
   - Include at least one compact Markdown table when result rows exist.
   - Keep cell text short. Do not pad table cells with spaces.
   - Show raw counts and rates together when both exist.
   - Include Normal/Test, Before/After, lot/mold/supplier/line labels exactly
     enough to understand the comparison.

7. Report review matrix
   - Include a compact Markdown table when the workbook has multiple sheets,
     multiple reports, or mixed evidence.
   - Required columns: Report/sheet, Reviewed item, Review purpose, Visible
     result, Domain, Process evidence, Function evidence, Use/limit.
   - Domain must show whether the row is process review, function review, or a
     linked process-function review.

8. Visual block
   - When result rows exist, include at least one AI-authored visual block.
   - Place the visual block directly after the key result table.
   - The visual block must be based on workbook evidence, not DB parameters.
   - For NG-rate/defect-rate/yield/PPM/below-spec comparisons, this visual
     block must be `report-bars` and the key result table must stay visible.

9. Interpretation
   - Explain what the numbers mean.
   - Use multiplicative relative change for NG-rate comparison:
     `(test_ng_rate / normal_ng_rate - 1) * 100`.
   - Say "worse by X%" when the result is higher than Normal and
     "improved by X%" when lower.
   - Never compare rows from different events as if they were one Normal/Test
     pair.

10. Recommended action
   - Give concrete next actions such as use, hold, improve mold, align line
     parameter, retest, first-article check, inspect specific process/part.
   - Tie each action to evidence strength.

11. Evidence location
   - List source worksheet names and cell ranges used for the judgement.
   - Include enough cell references to audit the report.

12. Limitations
   - Mention missing Normal, image-dependent evidence, partial sheet coverage,
     or weak sample size when relevant.

## Visual Block Formats

Use fenced code blocks. Do not output HTML.

Use `report-heatmap-matrix` for paired comparisons, condition-vs-Normal,
line A vs line B, or any two-factor matrix comparison.

```report-heatmap-matrix
{
  "columns": ["NG Rate", "Main Defect", "Judgement"],
  "rows": [
    {
      "label": "Normal",
      "cells": [
        {"value": "3.2%", "count": "26 / 800", "detail": "Normal value", "status": "good"},
        {"value": "Noise", "count": "20", "detail": "Dominant NG", "status": "warn"},
        {"value": "Normal", "count": "", "detail": "Normal row", "status": "good"}
      ]
    },
    {
      "label": "Test",
      "cells": [
        {"value": "5.8%", "count": "46 / 790", "detail": "81.3% worse than Normal", "status": "bad"},
        {"value": "Touch", "count": "31", "detail": "Dominant NG changed", "status": "bad"},
        {"value": "Hold", "count": "", "detail": "Retest after parameter change", "status": "bad"}
      ]
    }
  ]
}
```

Use `report-heatmap` for non-paired rankings: mold list, material list, lot
ranking, condition ranking, defect mix, or absolute result review.

```report-heatmap
[
  {"label": "VP #8", "value": "45.14%", "detail": "Worst lot-average bending NG", "status": "bad"},
  {"label": "VP #4", "value": "23.99%", "detail": "Second worst bending NG", "status": "bad"},
  {"label": "VP #7", "value": "2.86%", "detail": "Best lot-average bending NG", "status": "good"}
]
```

Use `report-bars` for NG-rate, defect-rate, yield, PPM, OK/NG ratio,
below-spec count, and other rate/count comparisons where magnitude matters.
The Dataset Results UI renders this as a vertical bar chart, so provide
`amount` as a numeric value whenever possible and keep exact values in `value`.

```report-bars
[
  {"label": "Normal", "value": "20.0%", "amount": 20.0, "detail": "NG 10 / Input 50", "status": "good"},
  {"label": "Bond 30s", "value": "50.0%", "amount": 50.0, "detail": "NG 25 / Input 50, Normal 대비 +150%", "status": "bad"}
]
```

When one result metric is compared across two independent variables and there
are at least four variable combinations, use `report-heatmap-matrix` as the
primary visual. Example: VP LOT x VP IR LOT -> function NG rate. Put one
variable on rows, the other on columns, and show the exact rate plus NG/Input
count in each cell.

When comparing multiple Types, conditions, lots, or lines and the workbook
also provides defect-type counts/rates under each group, the report must show
both the total NG rate for each group and that group's internal defect mix.
Compute defect share as defect_count / group_NG_total * 100 when counts are
available, and display defect name, count, and share in the visual block detail
or a compact companion table. Do not show only the total NG rate.

Allowed `status` values are `bad`, `warn`, and `good`.

All three translated reports must keep an equivalent visual block. Translate
`label` and `detail`; preserve numbers, `value`, `count`, `amount`, and
`status`.

## Report Quality Gate

The commit helper rejects low-quality reports. Avoid these failure cases:

- no `generated_report_markdown`
- missing Korean, English, or Vietnamese report
- no Markdown table when result rows exist
- no valid visual block when result rows exist
- fewer than six report sections when result rows exist
- no source worksheet/cell evidence
- report too short to stand alone
- parameter dump instead of analysis
- placeholder text such as `NG (auto-extracted)`, `batch inventory`, or
  `see workbook title/purpose`
- mojibake or broken encoding text

## Batch Execution Workflow

When launched from Data Inference Batch:

1. Import `_ai_batch_helper.py`.
2. Call `_ai_batch_helper.load_targets()` and verify the count matches the
   launcher prompt.
3. For each dataset, open the DB and call
   `_ai_batch_helper.get_excel_text(con, dataset)`.
4. Analyze every workbook and every worksheet in workbook order.
5. If the rendered text is too large, split by `=== SHEET:` and analyze sheet
   by sheet. Commit one combined result for the dataset.
6. Use `_ai_batch_helper.get_excel_files(con, dataset)` only when file context
   is required. Use `get_excel_paste` only as a last resort.
7. Build `result`, `tr_ko`, `tr_en`, and `tr_vi`.
8. Commit through a JSON payload file, not an inline Python command.
   Save `{ "name": dataset, "result": result, "translations": {"ko": tr_ko,
   "en": tr_en, "vi": tr_vi} }` to `_batch_tmp/<safe_dataset>.json`, then run
   `python _ai_batch_helper.py commit-json _batch_tmp/<safe_dataset>.json`.
9. If real extraction is impossible, call `_ai_batch_helper.log_failed(...)` and
   leave existing DB rows untouched.

Never run or import `_batch_auto.py`, `_batch_build.py`,
`_tmp_ai_excel/auto_normalize.py`, or any heuristic fallback normalizer.

## Workbook Analysis Rules

- Process every worksheet. Do not stop after the first visible sheet.
- Preserve workbook order and sheet names.
- Store `sheet_name` and `source_cells` on every important extracted row.
- Use rendered workbook text as the first source of truth.
- Use workbook files only when rendered text is insufficient.
- Preserve merged-cell meaning. If Date/Model/Type/Line/Process is shown once
  in a merged region, carry that visible value to rows under that region.
- Percentage-only subrows are not independent result rows. Attach them as
  breakdown/rate evidence to the preceding count row.
- Do not fabricate values. Unknown values must be `null` in structured fields
  and described as missing in the report when important.
- Separate original workbook statements from AI interpretation.
- Preserve manufacturing terms such as NG, VP+CD, SPL, THD, F0, Gauss, Tension,
  Jig, Mold, Lot, Line.

## Report Type Classification

Set `document.report_type` to exactly one of:

```text
normal_comparison
ng_without_baseline
before_after_dimension
measurement_spec
defect_root_cause
lot_supplier_mold_comparison
process_condition_change
reliability_spec
doe_matrix
image_dependent
mixed
```

Use these meanings:

- `normal_comparison`: NG or metric rows can be compared with a same-event
  Normal row. Source labels such as Baseline, Control, Reference, Before, Old,
  or OK may map to Normal internally, but visible report text must say Normal.
- `ng_without_baseline`: NG rows exist but no same-event Normal exists. Rank
  actual NG rate, defect mix, process, lot, supplier, mold, or line. Do not say
  improved or worse.
- `before_after_dimension`: dimension, gap, offset, thickness, height, or other
  measurement before/after comparison is the main evidence.
- `measurement_spec`: SPL, THD, F0, impedance, tension, Gauss, force, Min/Max,
  Avg, or spec/pass/fail measurement is the main evidence.
- `defect_root_cause`: workbook investigates symptoms, possible causes, checks,
  and remaining risks.
- `lot_supplier_mold_comparison`: lots, suppliers, molds, machines, labs, or
  lines are compared.
- `process_condition_change`: jig, machine, dry, UV, plasma, bonding, pressure,
  material, laser, or other process condition change is the main test.
- `reliability_spec`: temperature, humidity, aging, reliability, SPL/THD spec,
  or other spec-gate reliability result.
- `doe_matrix`: multiple condition/factor combinations are compared.
- `image_dependent`: important evidence appears in embedded photos/charts/OCR
  unavailable to text extraction.
- `mixed`: two or more categories are truly co-primary.

## Normal Pairing Rules

Pair a test row only with a Normal row from the same event:

- same worksheet or clearly same test block
- same date/period when available
- same model/part/lot/line context unless the test explicitly swaps that factor
- same measurement type
- compatible input/count basis

Valid source labels for internal pairing include Normal, Baseline, Control,
Reference, Before, Old, OK, and equivalent local-language labels. In visible
report text, normalize the comparison label to Normal.

If Normal is missing or ambiguous:

- classify as `ng_without_baseline` or another non-baseline type
- rank actual values
- state that improvement/worsening cannot be claimed

## Structured JSON Shape

Return Python dictionaries or JSON-compatible data in this shape before writing
the commit payload file.

```json
{
  "schema_version": "0.1",
  "generated_report_markdown": "## 결론 요약\n...\n",
  "document": {
    "document_id": "",
    "source_file": "",
    "source_sheet": "",
    "title": "",
    "model": "",
    "report_date": "",
    "department": "",
    "marker": "",
    "line": "",
    "report_type": "",
    "primary_defect": {
      "canonical_name": "",
      "aliases_in_document": []
    },
    "related_defects": [],
    "parts": [],
    "processes": [],
    "purpose": "",
    "content": [],
    "source_cells": {
      "title": [],
      "date": [],
      "purpose": [],
      "content": []
    }
  },
  "test_conditions": [
    {
      "condition_id": "",
      "condition_group": "",
      "line": "",
      "process": "",
      "changed_factor": "",
      "before_value": null,
      "after_value": null,
      "unit": null,
      "machine": null,
      "jig": null,
      "material_lot": null,
      "supplier": null,
      "dry_time_sec": null,
      "temperature": null,
      "pressure": null,
      "bond_amount": null,
      "uv_energy": null,
      "source_file": "",
      "sheet_name": "",
      "source_cells": []
    }
  ],
  "results": [
    {
      "result_id": "",
      "condition_id": "",
      "measurement_type": "",
      "condition_group": "",
      "date": "",
      "line": "",
      "input_count": null,
      "ok_count": null,
      "ng_count": null,
      "ng_rate_decimal": null,
      "ng_rate_percent": null,
      "metric_name": "",
      "metric_value": null,
      "unit": null,
      "judgement": null,
      "ng_breakdown": {},
      "source_file": "",
      "sheet_name": "",
      "source_cells": []
    }
  ],
  "conclusions": [
    {
      "conclusion_id": "",
      "topic": "",
      "statement_from_report": "",
      "normalized_interpretation": "",
      "source_file": "",
      "sheet_name": "",
      "source_cells": []
    }
  ],
  "troubleshooting_index": {
    "defect_name": "",
    "when_user_asks": [],
    "suggested_checks": [
      {
        "hint_id": "",
        "check_item": "",
        "reason": "",
        "evidence_strength": "",
        "related_process": "",
        "related_part": "",
        "source_file": "",
        "sheet_name": "",
        "source_cells": []
      }
    ],
    "limitations": []
  },
  "ai_extraction_log": {
    "confidence": 0.0,
    "assumptions": [],
    "warnings": [],
    "decision_rationale": ""
  }
}
```

Translations use this shape:

```json
{
  "document": {
    "title": "",
    "purpose": "",
    "content": [],
    "generated_report_markdown": ""
  },
  "conclusions": {
    "conclusion_id": {
      "topic": "",
      "statement_from_report": "",
      "normalized_interpretation": ""
    }
  },
  "hints": {
    "hint_id": {
      "check_item": "",
      "reason": ""
    }
  },
  "log": {
    "assumptions": [],
    "warnings": [],
    "decision_rationale": ""
  }
}
```

## Final Commit Rule

Windows command lines are short. Do not put the full `result`, `tr_ko`,
`tr_en`, or `tr_vi` dictionaries inside `python -`, PowerShell here-strings,
`Set-Content` command text, or any shell command. Large inline commands can fail
with Windows error 206: "The filename or extension is too long."

Use a payload file and the helper CLI:

```powershell
python _ai_batch_helper.py commit-json _batch_tmp/<safe_dataset>.json
```

The JSON file must have this top-level shape:

```json
{
  "name": "DatasetName",
  "result": {},
  "translations": {
    "ko": {},
    "en": {},
    "vi": {}
  }
}
```

Create the JSON file with the agent's file-writing/editing capability, not by
embedding the entire JSON in a shell command.

If the helper rejects the payload, revise the report and try again. Do not
bypass validation and do not write directly to SQLite.
