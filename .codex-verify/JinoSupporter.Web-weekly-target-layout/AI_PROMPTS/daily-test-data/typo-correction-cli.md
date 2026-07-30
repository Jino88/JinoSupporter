Run the Daily Test Data typo correction workflow for JinoSupporter.

RequestJson={{requestPath}}

Open and read RequestJson. It contains candidate text values extracted from newly added Daily Test data before DB save.

Return ONLY one valid JSON object:
{
  "replacements": [
    {
      "original": "exact candidate text",
      "corrected": "corrected text",
      "reason": "short reason"
    }
  ]
}

Rules:
- Correct only obvious human-language typos in labels, descriptions, section names, or Korean/English wording.
- Do not correct or normalize product codes, model names, lot IDs, VP/IR labels, file paths, sheet names, formulas, numbers, dates, units, acronyms, or abbreviations.
- Do not infer domain meaning. If a value might be a code or source label, omit it.
- original must exactly match one candidate text from RequestJson.
- corrected must be a single-line replacement for that whole candidate text.
- If no safe corrections exist, return {"replacements":[]}.
- Do not return commentary, markdown fences, or the rewritten source data.
