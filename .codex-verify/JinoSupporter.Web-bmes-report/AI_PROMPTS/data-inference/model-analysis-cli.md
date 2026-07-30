Run the Each Model Analysis workflow for JinoSupporter.
RequestJson={{requestPath}}
Database={{dbPath}}

Open the request JSON. It contains productType, analysisMode, purposeScoped, improvementPurposeScope, language, languages, reportCount, includedDatasets, sourceContextHash, createdAt, and context.
The context contains AI-authored per-report markdown. Use those AI_REPORT_MARKDOWN sections as the primary evidence. Do not rely on parameters alone.
If purposeScoped is true, analyze only the improvement purpose group named by improvementPurposeScope and do not generalize to the whole Product Type.
If purposeScoped is false, analyze all improvement purposes under the selected Product Type together.
Generate one complete model-level analysis for each language listed in request.languages. If request.languages has one item, generate and insert only that language.
The user's required report is not a generic risk summary. It must start with a practical action board: what to check/change next, how often similar evidence was reviewed, what the combined result says, and what exact next action should be taken.
After the action board, explain what reviews/tests were performed for this Product Type/model, group similar reviews together, show each review's result, then give the AI's overall model-level judgement.
Group similar reviews by purpose, defect/phenomenon, process, part/material, jig/mold/machine, changed condition, supplier/lot, or measurement type. Choose the grouping that best matches the evidence in the reports.
For each group, list the included report names and summarize each report's result separately before writing the group-level interpretation.
If analysisMode contains incremental, use the previous model analysis for the same target language when present in the context, evaluate only the new/changed reports as new evidence, but save one complete replacement model-level analysis for the same scope.
If analysisMode contains full, analyze all AI report markdown in the context for the selected scope together.

For each language, write concise Markdown only. No JSON in AnalysisMarkdown. No code fence.
The report must be table-centered. Avoid long prose; use short paragraphs only where judgement or caveats need explanation.
Required Markdown sections, translated naturally for the target language:
## Practical action board
Create the main table. Required semantic columns: Priority | Check/change item | Reviewed evidence count | Combined result | Why it matters | Next action.
The Check/change item column must be the factor to inspect or adjust, not the defect name itself.
Translate the table headers for the target language if appropriate, but preserve those six meanings.
## Review history by theme
Briefly state how many reports were reviewed, the main review themes, and the overall model-level direction.
## Similar review grouped results
Create a grouped evidence table. Required semantic columns: Review Group | Included Reports | Review Purpose/Condition | Result By Report | AI Group Judgment.
## Individual review result details
Use a compact table where possible. For each included report, show concrete result, judgement, and important caveats.
## Repeated patterns and differences
Explain what patterns repeat across groups and where results conflict or differ.
## AI overall judgment
Give the AI's overall judgement for this model/Product Type based only on the included reports.
## Follow-up review / improvement suggestions
Prioritize practical follow-up checks or improvements implied by the grouped evidence.
## Included reports
List all included report names.

Rules:
1. Do not invent facts that are not present in the request context.
2. Separate stable repeated patterns from one-off or weak-sample findings.
3. When reports conflict, describe the conflict and cite the affected report names.
4. Do not collapse everything into one summary. The grouped review history is the core deliverable.
5. Mention the included report names in the Included Reports section.
6. Preserve product codes, defect names, line names, numeric values, and units verbatim.
7. AnalysisTableMarkdown must contain only the exact main table from the Practical action board section, including the Markdown header row and separator row.

After writing the Markdown, insert exactly one row per requested language into SQLite table AiModelAnalyses in the supplied Database path.
Before inserting, ensure the table exists and includes AnalysisTableMarkdown. Create the table if needed with these columns:
Id INTEGER PRIMARY KEY AUTOINCREMENT, ProductType TEXT, AnalysisMode TEXT, Language TEXT, ReportCount INTEGER, IncludedDatasetsJson TEXT, AnalysisMarkdown TEXT, AnalysisTableMarkdown TEXT, SourceContextHash TEXT, CreatedAt TEXT.
If the table already exists but PRAGMA table_info(AiModelAnalyses) does not include AnalysisTableMarkdown, run ALTER TABLE AiModelAnalyses ADD COLUMN AnalysisTableMarkdown TEXT NOT NULL DEFAULT ''.
For every inserted row use ProductType=request.productType, AnalysisMode=request.analysisMode, Language=the exact requested language value, ReportCount=request.reportCount, IncludedDatasetsJson=json.dumps(request.includedDatasets, ensure_ascii=False), AnalysisMarkdown=the full Markdown for that language, AnalysisTableMarkdown=the exact main comparison table Markdown for that language, SourceContextHash=request.sourceContextHash, CreatedAt=current UTC ISO-8601.
Do not edit repository source files. Only read the request JSON and write the SQLite row.
Print a short completion line with productType and the inserted row ids for the requested languages.
