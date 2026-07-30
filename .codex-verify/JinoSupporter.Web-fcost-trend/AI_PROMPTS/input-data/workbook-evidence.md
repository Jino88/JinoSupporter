Workbook evidence rules:

- Read RequestJson.workbook.aiStructure.textPath first as the primary workbook
  structure evidence.
- Use extractedTextPath only as fallback or supplemental raw text.
- When Step 1 extraction includes layout.headerRows, mergedCells, normalizedRows,
  sheetMergedCells, or excelMergeSemantics, trust those fields over fallback
  flat text.
- For merged Date/No/Note/Model/Type/Line/Process cells, use the inherited
  values recorded in normalizedRows. Do not treat blank follower cells as
  missing context.
- For later steps, use the pre-analysis matrix and selected evidence blocks
  first. Re-open full workbook text only when the selected evidence is
  insufficient.
- Read every worksheet only to classify coverage and find relevant blocks. Once
  the relevant sheet/table is found, analyze that block deeply instead of
  repeatedly scanning unrelated sheets.
- Do not use embedded pictures unless explicitly asked or workbook text makes
  the picture essential. Nearby captions/placeholders can be mentioned as
  context when text supports them.
- Always preserve raw workbook numbers, units, labels, and source sheet/table
  names needed to audit the result.
