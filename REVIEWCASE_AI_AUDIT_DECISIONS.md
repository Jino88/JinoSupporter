# ReviewCase AI Audit Decisions

Date: 2026-07-06

Purpose:

- Record user-verified inclusion/exclusion decisions for ReviewCase AI generation.
- These decisions are data-quality decisions, not hard-coded extraction rules.
- Original Excel files and MicroSpeaker SQLite rows are not deleted by this file.

## Exclude From ReviewCase AI Candidates

Reason category: no usable extracted row/cell evidence for ReviewCase generation.

When a workbook is image-only, empty, reference-only, or otherwise lacks
source rows/cells that AI can cite as evidence, exclude it from ReviewCase AI
generation unless a later OCR/re-extraction step creates citeable rows.

| file_id | decision | reason | file_name |
| --- | --- | --- | --- |
| 247 | exclude | user confirmed delete from candidates | `12. BRS-201506 report checking Wire AOI standard_clean.xlsx` |
| 313 | exclude | user confirmed delete from candidates | `16. BRS-201506 report AOI mask_clean.xlsx` |
| 457 | exclude | user said this file is not needed | `26. BRS-161014 Report compare Coil C2 & E2 (Machine #5)_clean.xlsx` |
| 628 | exclude | user said image-only; delete from candidates | `41. BRS-161016 report compare MTR VP #2, #4, #8 2024.06.03_clean.xlsx` |
| 783 | exclude | user said content is empty; delete from candidates | `60. TIU C11-20 Report check high dimension Frame + CD 2026.2.07_clean.xlsx` |
| 803 | exclude | user said delete; image-only/no extracted evidence class | `63. 2023.11.08. Function NG_clean.xlsx` |
| 190 | exclude | user said delete salt spray from ReviewCase candidates | `1. BRS-161016 TF Salt spray TEST 2024.10.31_clean.xlsx` |

## General Handling

- Do not send excluded files to ReviewCase AI generation.
- Do not use excluded files as Ask AI evidence unless they are reprocessed into
  citeable text/table rows later.
- Keep the exclusion reason visible so future re-extraction can decide whether
  OCR or manual transcription is worth doing.

## Keep As ReviewCase AI Candidates

| file_id | decision | handling | file_name |
| --- | --- | --- | --- |
| 721 | keep | material/supplier/coating review; preserve supplier, coating, test round, and outcome sections separately | `52. BRS-161016 Report test material YK change Supplier (Producft of press line - Glonics - Coating MJ , Doojin VN ) 2024.7.13_clean.xlsx` |
| 266 | keep | material dimension/spec review; preserve spec wording, sheet/date/run, and outcome sections separately | `13. BRS-161016 Report Test PT 161014-S of Press line (Doojin coating) happen  NG dimension 30.9.2025_clean.xlsx` |

## ReviewCase Drafts Pending User Confirmation

| file_id | draft | status | note |
| --- | --- | --- | --- |
| 721 | `REVIEWCASE_AI_DRAFTS/721.reviewcase-draft.json` | primary change factor confirmed | user confirmed primary changed factor is Normal vs Test YK MJ vs Test YK Doojin; keep Baotou/Boutou as subgroup context only |
