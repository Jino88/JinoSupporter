The attached image(s) are Excel sheet screenshots containing manufacturing inspection or production data.

Parse visible tables into the same JSON array schema used by the text parser:
[
  {
    "tableName": "descriptive name",
    "columns": [{"field": "f0", "label": "Column Label"}, ...],
    "rows": [{"f0": "value", ...}, ...]
  }
]

Rules:
- Preserve row order and column order.
- Preserve visible numbers, percentages, product codes, dates, and units exactly.
- Use merged visual headers to build column labels, but do not invent missing values.
- For blank continuation cells under visible Date/Model/Type/Line/No section values, carry context down only when the table layout clearly indicates merged cells.
- Do not carry down numeric, OK, NG, total, rate, sample, or measurement columns.
- Return ONLY the JSON array. No markdown fence and no explanation.
