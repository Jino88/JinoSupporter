# Inference Data AI Service - Excel DB Generalization Plan

## Core Answer

Excel forms are not generalized by forcing one extraction rule.

They are generalized by splitting the system into layers:

1. **Raw grid layer**
   - Store every workbook as fixed Excel coordinates.
   - Preserve workbook, sheet, row, column, cell address, value, raw value, and merged-cell metadata.
   - Never delete covered merged cells or compact rows/columns.
   - This layer is format-agnostic and becomes the audit source.

2. **Candidate hint layer**
   - Detect common patterns such as Input/OK/NG/NG rate, Min/Max/Avg, Normal/Test, Before/After, model names, material/supplier/lot/process terms.
   - These rows are only hints. They are not final truth.

3. **Reusable analysis layer**
   - Store the human/AI/rule analysis as a structured object linked to one raw workbook fingerprint.
   - Record the report purpose/type, each review type, comparison/test/control cohorts, metric values, numerator/denominator, ppm or measurement deltas, conclusions, limitations, and exact A1 evidence ranges.
   - Support arbitrary cohort attributes and metric scopes so supplier, line, material, lot, process, round, and component layouts do not require new DB tables.
   - Verify every evidence range against the fixed source grid and mark the analysis stale when the source workbook is re-imported.

4. **ReviewCase AI layer**
   - AI reads source-backed rows/cells and candidate hints.
   - AI can reuse a verified structured analysis rather than re-deriving the same Test/Control comparisons from HTML text.
   - AI decides changed factors, outcomes, evidence rows, and limitations.
   - Every claim must cite source rows/cells.

5. **Verification layer**
   - Reject or mark needs_review when cited evidence does not exist, condition/outcome is ambiguous, denominators are mixed, or the model is uncertain.
   - Only verified ReviewCases are allowed as primary ASK AI evidence.

## Why This Works For Mixed Excel Formats

Different Excel files may have different titles, tables, merged headers, row positions, and measurement layouts.

The raw grid layer does not care about those differences. It only stores:

- `workbooks`
- `worksheets`
- `grid_sheet_rows`
- `grid_sheet_cells`
- `merge_ranges`

That is the common denominator for every Excel workbook.

Meaningful business extraction is moved upward into candidate detection and AI ReviewCase generation, where ambiguity can be preserved instead of hidden.

The reusable analysis layer is the persistent bridge between those two: it stores a dashboard's selected comparison logic without copying the raw grid or hard-coding one Excel layout.

## Two CLI Paths

### Quick Index Path

Use the existing MicroSpeaker indexer through CLI.

Purpose:

- Reproduce current candidate tables and dashboards quickly.
- Produce `files`, `sheets`, `sheet_rows`, `sheet_cells`, `metric_candidates`, `measurement_stats`, `comparison_pairs`, and `term_hits`.
- Useful for comparing with the existing accurate right-side table behavior.

Limitation:

- This path uses `openpyxl`, so it is best for normal `.xlsx/.xlsm` files.
- It does not preserve Excel COM-only behavior for DRM or policy-protected files.

### COM Universal Grid Path

Use Excel COM extraction and import the output JSON into a universal SQLite schema.

Purpose:

- Preserve the exact Excel coordinate grid and merge structure.
- Support files that require local Excel to open.
- Build a reliable source DB even when workbook formats differ completely.

Limitation:

- Slower because it opens Excel through COM.
- It should be run in batches and resumed rather than treated as a one-shot web request.

### Resumable Universal Ingestion

The universal path is designed for hundreds of mixed layouts.

- A source fingerprint identifies an unchanged workbook, so repeat runs skip the existing verified grid by default.
- Raw JSON names are deterministic from the source fingerprint instead of a run ordinal; optional reuse is accepted only after source and extraction-option validation.
- `runs` and `ingest_items` retain per-file `SKIPPED`, `IMPORTED`, or `FAILED` status.
- A failed extraction records the error but preserves the last successful workbook grid.
- Every COM JSON payload is validated as a fixed coordinate grid before import. The importer validates UsedRange dimensions, row sequence, duplicate/missing coordinates for full-grid output, merge ranges, and anchor/covered metadata.
- `verify-universal-db` compares stored rows/cells/merges with the retained raw JSON artifacts.

Use `--covered-cell-mode blank` and do not use `--sparse` for the raw source DB. Blank covered cells remain coordinates with explicit merge metadata, which preserves the original layout without duplicating labels.

## Output Rule

All outputs for this CLI service go under:

`D:\000. MyWorks\005. Program\Repository\JinoSupporter\InferenceDataAIService`

Default output layout:

- `outputs/quick-index/*.sqlite`
- `outputs/quick-index/*_dashboard.html`
- `outputs/universal-grid/*.sqlite`
- `outputs/universal-grid/raw-json/**/*.json`
- `outputs/analysis-manifests/*.json`
- `outputs/analysis-exports/*.json`
- `outputs/packets/*.json`
- `outputs/logs/*.log`

## ReviewCase Contract

The DB is not the final answer. The final AI evidence object should look like:

```json
{
  "reviewCaseId": "",
  "sourceWorkbook": {},
  "modelReview": {
    "selectedModels": [],
    "mappingStatus": "confirmed | needs_user_mapping | missing"
  },
  "changedFactors": [
    {
      "changedFactorId": "cf-1",
      "changeDomain": [],
      "baselineCondition": "",
      "changedCondition": "",
      "evidenceRows": []
    }
  ],
  "outcomes": [
    {
      "outcomeId": "out-1",
      "changedFactorId": "cf-1",
      "outcomeDomain": "",
      "outcomeMetric": "",
      "judgement": "improved | worse | mixed | no_change | not_judged | needs_review",
      "subResults": [],
      "comparisonRows": [],
      "limitations": []
    }
  ],
  "verification": {
    "status": "verified | needs_review | excluded",
    "issues": []
  }
}
```

## Implementation Direction

Start in CLI only:

1. Build raw/candidate DB outputs in this folder.
2. Generate one-file AI packets from DB records.
3. Review packets and AI drafts from CLI.
4. Only after this stabilizes, wire verified ReviewCases into JinoSupporter Web.
