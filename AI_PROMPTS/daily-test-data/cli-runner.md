Run the Daily Test Data report workflow for JinoSupporter.

RequestJson={{requestPath}}

Open and read RequestJson. It contains the user prompt, extracted Excel text, prior history context, current workflowPhase, and request.analysisPrompt.
If request.programExtractSummary is present and available, read it before raw data and use it as the workbook map/coverage guide.

Use request.analysisPrompt as the complete analysis specification, with this override:
- Return one valid JSON object, not a bare HTML fragment.
- The JSON must include parameters and analysisHtml.
- analysisMarkdown is optional and should be concise if included.
- Do not wrap the JSON in a markdown fence.

Data rules:
- Prefer request.programExtractSummary for workbook metadata, sheet names, header rows, row spans, and extracted-cell coverage. Use raw EXCEL COM EXTRACT text only for metric calculations and evidence details that are not already summarized.
- Prefer EXCEL COM EXTRACT blocks in request.currentInput.dataText and request.cumulativeDataText. Those blocks are the source extracted from DRM Excel.
- Do not reopen DRM workbook paths when EXCEL COM EXTRACT text is present.
- Lines starting with XLSX_ATTACHMENT are legacy references only. If no EXCEL COM EXTRACT text is available for those files, say in the HTML that Excel COM extraction is required.
- Preserve product codes, model names, line names, numbers, dates, units, item-specific dimension labels, defect labels, and source labels verbatim.
- Do not invent facts. If a requested metric or dimension is missing, include a clearly named not-available section in the HTML.
- Use request.existingParametersJson as the canonical parsing/report contract when it is not empty. For new data, map fields into that contract before adding new fields.
- parameters.analysisData must be compact reusable aggregate data, not a copy of all raw rows.
- parameters.reportBuilderSpec must store the data-reading method and HTML Maker contract so the application can later rebuild the same report style from parameters plus newly added aggregates.

Tool / shell rules:
- You may use shell commands only to read RequestJson or run compact local calculations.
- Never put large Python, PowerShell, JSON, HTML, or data text inside a `powershell -Command`, `cmd /c`, or one-line inline command.
- Keep each command line under 2,000 characters.
- If a calculation needs more than a few lines of code, create a temporary script file under `tmp/daily-test-data-ai/`, run it with a short command such as `python tmp/daily-test-data-ai/analyze.py`, then delete or ignore the temp file.
- Do not use base64/EncodedCommand to hide large scripts in the command line.
- Windows error 206 means the command line is too long; avoid that by using file-based scripts.

Scope rules:
- If workflowPhase is current_input_analysis, create a current-input-only HTML report.
- If workflowPhase is cumulative_update, create the final cumulative HTML report for all unique data in request.cumulativeDataText plus currentAdditionAnalysis when present.
- If promptOnlyReanalysis is true, create a new HTML view from existing analysis context and extracted cumulative data.

HTML rules:
- If request.outputMode is "standalone_html_document" or request.allowStandaloneHtml is true, analysisHtml must be a complete standalone HTML document. Include doctype, html, head, meta charset UTF-8, viewport, body, CSS, JavaScript, filters, search controls, canvas/charts, and embedded analysis data when the user asks for them. Do not use external servers, CDNs, or network assets. Keep all CSS, JavaScript, and data inside analysisHtml.
- If request.standaloneOutputPath is present, still return the complete HTML document in analysisHtml. The application will write the file; do not skip JSON output.
- Otherwise, analysisHtml must be an HTML fragment only.
- In fragment mode, do not include doctype, html, head, body, script, iframe, style, link, meta, SVG, canvas, input, select, button, or inline event handlers.
- In fragment mode, allowed tags are: h2, h3, p, ul, ol, li, table, thead, tbody, tr, th, td, strong, em, code, pre, br, span.
- In fragment mode, allowed attributes are class and style only.
- In fragment mode, use inline style for KPI cards, action boards, heatmap cells, risk colors, and compact layout.
- Put the action board / key conclusion first, then KPI summary, heatmap/matrix, evidence table, and notes.
- If a heatmap/status table cell shows a defect rate, display the rate on top and the count line as NG/Total. Never display Total/NG.
- If the requested report includes a heatmap or matrix, render all observed combinations for the item-requested dimensions in the matrix itself, unless the item prompt explicitly asks for a limited/top-N view.
- In matrix cells, use "-" only for combinations that have no source row. Do not write "obs.", "observed", or "seen" as a placeholder. For observed combinations with enough source data, compute and show the requested metric plus its count basis; if the metric cannot be computed, write "metric unavailable" and add a note naming the missing field.

Output JSON schema:
{
  "parameters": {
    "version": 1,
    "projectName": "item name",
    "scope": "current_input or cumulative",
    "dataSummary": {
      "rowCount": 0,
      "dateRangeStart": "",
      "dateRangeEnd": "",
      "periodRange": "",
      "sourceFiles": [],
      "coverage": ""
    },
    "sourceSchema": {
      "fields": [
        {
          "name": "canonical field name",
          "aliases": [],
          "category": "date|dimension|metric|input|defect|rate|condition|other",
          "dataType": "string|number|date|percent|boolean",
          "unit": "",
          "sourceEvidence": ""
        }
      ]
    },
    "analysisContract": {
      "dimensions": [],
      "metrics": [],
      "normalizationRules": [],
      "groupingRules": [],
      "reportSections": []
    },
    "reportBuilderSpec": {
      "version": 1,
      "parseRules": [],
      "fieldMappings": [],
      "metrics": [],
      "aggregations": [],
      "reportLayout": [],
      "renderRules": [],
      "colorRules": [],
      "validationRules": [],
      "updateStrategy": ""
    },
    "analysisData": {
      "totals": {},
      "aggregates": {},
      "matrices": [],
      "topRisks": [],
      "coverage": ""
    }
  },
  "analysisMarkdown": "",
  "analysisHtml": "<h2>...</h2> or <!doctype html>...",
  "historyParameters": {},
  "historyAnalysisMarkdown": "",
  "historyAnalysisHtml": ""
}

For current_input_analysis, historyParameters/historyAnalysisHtml should match the current input scope.
For cumulative_update, parameters/analysisHtml are cumulative; historyParameters/historyAnalysisHtml may be empty when there is no separate current input.

Return ONLY valid JSON. No markdown fence. No commentary outside JSON.

You may read RequestJson and referenced text files. You may write temporary analysis scripts/files only under `tmp/daily-test-data-ai/`. Do not update SQLite and do not edit repository source files.
