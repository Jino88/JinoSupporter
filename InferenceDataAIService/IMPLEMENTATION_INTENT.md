# InferenceDataAIService - Implementation Intent

## One-Line Goal

Build a CLI-first service that converts mixed-format Excel report files into a source-backed data layer, then generates AI ReviewCase evidence that can later be used by ASK AI.

This must not be wired directly into JinoSupporter Web yet.

## What The User Appears To Want

The target is not a generic Excel parser that magically understands every sheet with one fixed rule.

The target is:

1. Put all Excel files into a common DB without losing structure.
2. Preserve enough source evidence so every extracted claim can be audited.
3. Let AI infer the actual review meaning per workbook:
   - what changed,
   - compared against what,
   - what outcome changed,
   - what rows/cells prove it,
   - what is ambiguous.
4. Verify AI output before it becomes ASK AI evidence.

## Why `report_compare_dashboard.html` Looked Accurate

The dashboard was not doing live AI analysis in the browser.

It was accurate because the source data had already been normalized into candidate tables and then filtered:

- workbook/sheet/row/cell extraction already existed,
- pair candidates were generated from Normal/Test, Before/After, or adjacent condition rows,
- invalid pairs were rejected by validation logic,
- only valid visible candidates were rendered in the right-side extracted table.

So the accuracy came from:

- preserving source rows,
- candidate extraction,
- validation filters,
- manual representative logic for some high-value report types.

## Generalization Strategy

Because every Excel format is different, the stable abstraction is not `NG rate table` or `Normal/Test table`.

The stable abstraction is Excel itself:

- workbook
- sheet
- row
- column
- cell address
- value
- raw value
- merged range
- merge anchor/covered cell role

After this raw layer exists, higher layers can add meaning.

## Service Layers

### 1. RawWorkbookExtraction

Reads Excel files and stores their source structure.

Required output:

- workbooks
- worksheets
- grid_sheet_rows
- grid_sheet_cells
- merge_ranges

Important rule:

- Do not remove blank covered merge cells.
- Do not compact rows or columns.
- Do not shift coordinates.

### 2. CandidateHintBuilder

Builds non-final hints:

- Input/OK/NG/NG rate rows
- Min/Max/Avg measurements
- Normal/Test or Before/After pairs
- model names
- material/supplier/lot/process terms
- title/purpose/result/decision/context rows

These are hints only.

### 3. ReviewCaseAiPacketBuilder

Creates one JSON packet per workbook.

The packet should include:

- source workbook metadata
- sheets
- source rows/cells
- context rows
- candidate hints
- required ReviewCase contract
- notes explaining that candidates are hints, not truth

### 4. ReviewCaseAiGenerator

AI reads the packet and outputs ReviewCases.

AI should decide:

- zero, one, or multiple ReviewCases per workbook,
- changed factors,
- baseline condition,
- changed/test condition,
- outcomes,
- numeric comparison values,
- limitations,
- cited evidence rows/cells.

### 5. ReviewCaseVerifier

Verifies generated ReviewCases before use.

Reject or mark `needs_review` when:

- cited rows/cells do not exist,
- changed factor is unsupported,
- outcome is unsupported,
- numerator/denominator is mixed,
- model mapping is ambiguous,
- process NG and function NG are combined incorrectly,
- date or condition meaning is unclear.

### 6. HumanCalibrationQueue

CLI should support one-workbook-at-a-time confirmation.

The user may need to confirm:

- model name,
- whether this file should be excluded,
- whether a ReviewCase draft is actually correct,
- whether the AI misunderstood the workbook purpose.

### 7. AskAiEvidencePack

Only verified ReviewCases should later become primary ASK AI evidence.

Raw rows and candidate hints remain available as fallback/audit evidence.

## Initial Questions The System Should Eventually Answer

1. Does VP+CD assembly condition affect function defect?
2. Does VP FILM LOT affect function defect?
3. Does FP+VP assembly bonding amount affect function defect?
4. Does Magnet Lot affect adhesion strength?

Expected evidence types:

- process NG when process-related,
- function defect rate,
- SPL/Hearing/function evidence,
- adhesion strength or measurement evidence for Magnet Lot.

## What Not To Do Yet

- Do not directly attach this to JinoSupporter Web.
- Do not treat candidate pairs as final evidence.
- Do not hard-code only VP/CD/Magnet-specific rules.
- Do not use one Excel layout rule for all files.
- Do not aggregate unrelated models or unrelated denominators.
- Do not answer from filenames alone.

## Current CLI Direction

The first CLI should support:

- `init-db`
- `quick-index`
- `com-index`
- `inspect-db`
- `search`
- `build-packet`

All output must stay under:

`D:\000. MyWorks\005. Program\Repository\JinoSupporter\InferenceDataAIService`

