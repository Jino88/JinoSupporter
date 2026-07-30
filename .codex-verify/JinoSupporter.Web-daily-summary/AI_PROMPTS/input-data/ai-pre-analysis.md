Mandatory AI pre-analysis phase:

- Read RequestJson.workbook.aiStructure.textPath. If missing, use
  workbook.extractedTextPath as fallback.
- Classify the workbook from Excel text only:
  - review purpose
  - purpose code/category used by review_index.html when clear
  - target defect(s)
  - reviewed item(s)
  - model
  - date/period
  - confidence
- Never derive those fields from workbook.fileName, sourceDataset, source path,
  sequence number, or date suffix. File names may only help identify product or
  model after workbook text supports it.
- If RequestJson.reviewIndex has a matching row, use it only after the text
  classification to align wording or report a conflict.

Build an internal reportReviewMatrix before detailed analysis. Each workbook or
distinct report block must have one row with these fields:

- report/sheet/block name
- reviewed item: what factor, process, lot, condition, material, jig, line, or
  metric was reviewed
- review purpose: why that item was reviewed
- visible result: what the workbook actually shows
- review domain: exactly one of `공정 불량 검토`, `기능 불량 검토`,
  `공정+기능 연계 검토`, `기타/미확인`
- process evidence
- function evidence
- recommended use in this analysis
- do-not-use or weak-evidence note
- confidence

Definitions:

- `공정 불량 검토`: process, line, jig, mold, material, VP/CD/bonding/dry/UV,
  pressure, timing, laser, plasma, lot, supplier, or manufacturing condition is
  the main factor.
- `기능 불량 검토`: final function NG, VP+CD separate NG, SPL/THD/F0,
  tension, spec, measurement distribution, reliability, pass/fail, or customer
  functional symptom is the main result.
- `공정+기능 연계 검토`: process condition is tested and the effect is judged by
  function NG/spec/measurement.

The visible report must include this matrix near the top so the reader can see
what each report was for, what result it showed, and whether to pursue process
or function review.
