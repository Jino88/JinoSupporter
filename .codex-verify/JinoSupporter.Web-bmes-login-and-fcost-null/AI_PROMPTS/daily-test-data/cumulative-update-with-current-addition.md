Workflow phase: cumulative_update.
Current added-data HTML/text summary from pass 1:
{{currentAdditionAnalysis}}

Additional cumulative-update rules:
- Use the pass-1 current added-data report as a change/finding summary.
- analysisHtml must still be cumulative for all unique data. If request.outputMode is "standalone_html_document" or request.allowStandaloneHtml is true, return a complete standalone HTML document; otherwise return a fragment.
- Keep previously requested full views, including full heatmaps/matrices when requested.
- Full heatmaps/matrices must keep every observed combination for the item-requested dimensions in the matrix cells with the requested metric and count basis when computable. Do not use "obs." or a separate retained-label list as a replacement for matrix values.
- Return cumulative parameters with updated analysisContract, reportBuilderSpec, and analysisData. Preserve existing contract keys when possible.
- Preserve and update reportBuilderSpec as the saved HTML Maker contract, including parseRules, fieldMappings, metrics, aggregations, reportLayout, renderRules, colorRules, validationRules, and updateStrategy.
