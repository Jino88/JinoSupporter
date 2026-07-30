Run JinoSupporter INPUT DATA analysis.
RequestJson={{requestPath}}

Open RequestJson and analyze only the workbook/session described in that JSON.
Step 1 deterministic Excel text extraction is already complete. Your job is AI
interpretation from the extracted workbook text, not filename parsing and not
DB-field filling.

Mandatory execution order:
1. Read RequestJson.workbook.aiStructure.textPath first.
2. Build the AI pre-analysis and reportReviewMatrix from workbook text.
3. Decide the visualization for this file before writing the result.
4. Re-read only the worksheet/table blocks needed to support the judgement.
5. Produce analysisText and a complete self-contained analysisHtml report.

For INPUT DATA (BATCH), each added workbook is still analyzed as one file with
the exact same rules. RequestJson.reviewIndex is only a taxonomy/reference table
after workbook-text classification. Do not use it as a filename-only substitute.

If reviewPurpose is not empty, answer that objective directly. If it conflicts
with workbook evidence, say so in the judgement and keep workbook evidence as
the source of truth.

Speed rule: avoid repeated whole-workbook scans. Once you know the relevant
sheets/tables, focus on those blocks and cite the source. Ignore images unless
the workbook text says the image/caption is essential.

{{aiPreAnalysisPrompt}}

{{defaultAnalysisPrompt}}

{{visualizationSelectionPrompt}}

{{reviewIndexPrompt}}

{{workbookEvidencePrompt}}

{{analysisOutputPrompt}}
