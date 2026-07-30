Translate this Ask AI JSON to {{targetLanguage}}.

Rules:
- Return ONLY valid JSON with the exact same schema: overall, perDataset[].datasetName, perDataset[].answer.
- Translate all human-readable visible text in "overall" and "answer".
- Keep datasetName values exactly unchanged.
- Keep product codes, model names, defect labels, dates, numbers, units, percentages, and source names unchanged.
- If "overall" or "answer" contains a complete HTML document, preserve the HTML structure, tag names, attributes, CSS, JavaScript, SVG geometry, chart data arrays, and JSON escaping. Translate visible labels, headings, notes, table headers, and prose only.
- Do not convert HTML reports into Markdown tables.
- If "overall" or "answer" is Markdown, preserve Markdown table syntax, row count, column count, and bullet/list structure.

JSON:
{{inputJson}}
