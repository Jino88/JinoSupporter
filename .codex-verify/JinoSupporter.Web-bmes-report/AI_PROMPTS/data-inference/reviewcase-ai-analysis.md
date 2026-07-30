# ReviewCase AI Analysis And Verification

RequestJson={{requestPath}}

Open RequestJson and analyze the extracted workbook evidence. The extractor is
not the final ReviewCase classifier. It only provides source-backed rows, cells,
candidate tables, candidate pairs, candidate measurements, and hints.

## Goal

Create verified ReviewCase JSON from extracted data:

- `changedFactors`
- `outcomes`
- `evidenceRows`
- `verification`

The verified ReviewCase output is the primary evidence layer that Ask AI will
use later. Ask AI should not need to reinterpret the whole workbook when this
ReviewCase is complete and verified.

## Core Rules

- Decide whether the workbook contains one ReviewCase, multiple ReviewCases, or
  no ReviewCase.
- Use source rows/cells as the authority. Candidate pair rows and heuristic
  labels are hints only.
- Do not hard-code product, part, process, material, supplier, jig, equipment,
  or defect names.
- Do not force a final improved/worse judgement when the source only supports a
  measurement comparison or mixed outcomes.
- If the workbook is only post-failure analysis with no comparable changed
  condition result rows, return `reviewCaseStatus="excluded"` with the reason.
- If the workbook has no extracted row/cell evidence that can be cited, and the
  available source is image-only, empty, or reference-only, return
  `reviewCaseStatus="excluded"` unless OCR/re-extraction evidence is provided.

## Calibration References

RequestJson may include user-verified calibration notes or reference patterns.
Use them as analysis guidance, not as hard-coded matching rules.

- A reference pattern can tell you what kind of structure to look for, such as
  row-level grouping keys, DOE parameter columns, secondary test conditions,
  mixed outcome domains, measurement-only evidence, or non-comparable analysis
  reports.
- Do not classify the current workbook by matching a specific file name, part
  name, product name, supplier name, process name, or defect name from a
  calibration note.
- Always inspect the current workbook's extracted rows/cells and cite current
  evidence rows.
- If the current workbook contradicts a reference pattern, trust the current
  extracted evidence and put the mismatch in `verification.issues`.
- Treat calibration references as few-shot examples for reasoning style: what to
  preserve, what to split, what to exclude, and what not to judge.

## Evidence Requirements

Every `changedFactor` must cite rows or cells that identify the changed
condition, tested state, baseline state, or grouping context.

Every `outcome` must cite rows or cells that identify:

- outcome metric or result section
- compared condition rows
- input quantity and NG/result quantity when applicable
- rate value when applicable
- measurement values/spec/sample count when applicable

Every `evidenceRows` item must refer to an extracted row/cell id from
RequestJson. Do not invent row numbers, values, or file names.

## Verification Pass

Before returning the final JSON, run a verification pass:

- Check that every cited evidence row exists in RequestJson.
- Check that every numeric value appears in cited rows or is calculated from
  cited numerator/denominator.
- Check that row grouping is consistent. If rows contain grouping dimensions
  such as content, type, item, sample, lot, material, line movement, station,
  equipment, side, cavity, fixture, position, or secondary condition, compare
  only within matching groups.
- Check that multiple outcomes under one changed factor remain separate.
- Check that missing baseline, missing test condition, ambiguous grouping, or
  insufficient evidence is reported in `verification.issues`.

## Output JSON Shape

Return only JSON:

```json
{
  "reviewCaseStatus": "verified | needs_review | excluded",
  "sourceFileId": 0,
  "sourceFile": "",
  "reviewCases": [
    {
      "reviewCaseId": "",
      "reviewTitle": "",
      "reviewPurpose": "",
      "changedFactors": [
        {
          "changedFactorId": "",
          "changeDomain": "",
          "changedFactor": "",
          "baselineCondition": "",
          "changedCondition": "",
          "subgroupKeys": [],
          "evidenceRows": []
        }
      ],
      "outcomes": [
        {
          "outcomeId": "",
          "changedFactorId": "",
          "outcomeDomain": "",
          "outcomeMetric": "",
          "comparisonRows": [],
          "normal": {
            "condition": "",
            "input": null,
            "ng": null,
            "ratePercent": null,
            "measurementValues": []
          },
          "test": {
            "condition": "",
            "input": null,
            "ng": null,
            "ratePercent": null,
            "measurementValues": []
          },
          "judgement": "improved | worse | no_change | mixed | not_judged",
          "limitations": []
        }
      ],
      "evidenceRows": [],
      "limitations": []
    }
  ],
  "verification": {
    "status": "passed | failed | needs_review",
    "checkedEvidenceRows": [],
    "issues": []
  }
}
```
