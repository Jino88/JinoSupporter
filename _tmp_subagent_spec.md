# Subagent Spec — CLI AI Batch v7 (analysis JSON producer)

You are processing factory quality reports for the JinoSupporter project. For each TSV input file, produce ONE structured JSON result file. The orchestrator script will commit your JSON to SQLite — your deliverable is **only** the JSON files at the exact paths given.

## Working directory
`D:\000. MyWorks\005. Program\Repository\JinoSupporter`

## Per dataset

1. Read `_tmp_tsv/NNN.txt` (UTF-8, raw spreadsheet paste — Vietnamese factory reports, mixed Korean/English/Vietnamese, with cell-format annotations like `〔bg=#FFFF00〕` and `〔b〕`).
2. Look up the dataset name in `_tmp_tsv/index.json` (key = NNN as zero-padded string).
3. Write your analysis to `_tmp_results/NNN.json`.

## Output shape

```jsonc
{
  "result": { ...v7 fields below... },
  "tr_ko": { ...Korean translation of narrative fields only... },
  "tr_vi": { ...Vietnamese translation of narrative fields only... }
}
```

## v7 result schema (strict — no markdown, no extra fields)

```jsonc
{
  // STEP A — classify reportType FIRST (one of 7):
  "reportType": "comparison_study" | "multi_arm" | "doe_factorial" |
                "reliability_validation" | "trend_analysis" |
                "quality_log" | "intervention_test",

  // STEP B — verdict (enum depends on reportType):
  //   comparison/multi_arm/doe/intervention/trend → improved | worsened | partial | no_clear_effect | inconclusive
  //   reliability_validation → passed | failed
  //   quality_log → "" (empty)
  "verdict": "...",

  // ONE-sentence headline, ≤20 words, magnitude + direction.
  "headline": "...",

  // ≤4 rows. Two formats:
  "evidence": [
    // (a) 2-arm default — comparison_study / intervention_test / reliability_validation:
    { "metric": "NG rate",
      "baselineLabel": "Normal", "baselineValue": "3.0% (3/100)",
      "variantLabel": "Test",    "variantValue": "1.0% (1/100)",
      "deltaText": "-2pp", "deltaSign": "down" | "up" | "flat",
      "note": "",
      "comparisons": null, "bestLabel": "", "worstLabel": "" },

    // (b) multi-arm — reportType=multi_arm ONLY:
    { "metric": "VP bending rate",
      "baselineLabel": "", "baselineValue": "", "variantLabel": "", "variantValue": "",
      "deltaText": "+58pp range", "deltaSign": "up", "note": "",
      "comparisons": [
        { "label": "VP #6", "value": "4.4% (53/1200)", "n": 1200,
          "isBaseline": true,  "isBest": false, "isWorst": false },
        { "label": "VP #7", "value": "59.7% (689/1154)", "n": 1154,
          "isBaseline": false, "isBest": false, "isWorst": true },
        { "label": "VP #9", "value": "0.4% (5/1176)", "n": 1176,
          "isBaseline": false, "isBest": true,  "isWorst": false }
      ],
      "bestLabel": "VP #9", "worstLabel": "VP #7" }
  ],

  "actions": [
    { "priority": 1, "kind": "action" | "investigate" | "risk", "text": "..." }
    // ≤3 items, ordered by priority
  ],

  "context": {
    "process": "...",         // e.g. "VP-CD bonding", "Coil assembly"
    "stage": "...",            // e.g. "Sub1", "Sub2", "Final inspection"
    "baselineReason": "..."    // why this baseline/comparison was chosen
  },

  // reportType=doe_factorial ONLY (else null/omit):
  "doeGrid": {
    "factor1Name": "Temperature", "factor2Name": "Tension",
    "factor1Levels": ["380","390","400","410","420"],
    "factor2Levels": ["4","5","6","7","8"],
    "cells": [
      { "f1": "390", "f2": "5", "status": "ok"|"ng"|"borderline"|"empty", "value": "7.611mm" }
    ]
  },

  // reportType=trend_analysis ONLY (else null/omit):
  "trendPoints": [
    { "label": "Week 17", "value": "8.3%", "note": "" },
    { "label": "Week 18", "value": "4.1%", "note": "improvement after TF" }
  ],

  // Optional normalized per-row measurements (skip if not a clean NG-rate table):
  "measurements": [
    { "productType": "BRS-161014", "testDate": "2025-04-12", "line": "DT",
      "checkType": "Bonding", "variable": "VP-CD separate",
      "variableDetail": "Test lot", "variableGroup": "Test",
      "intervention": "new bond SJ4765",
      "inputQty": 1000, "okQty": 980, "ngTotal": 20, "ngRate": 2.0,
      "defectCategory": "Separate", "defectType": "VP-CD", "defectCount": 20 }
  ],
  "tags": ["bonding","vp-cd","sub1"],

  "productType": "",  // e.g. "BRS-161016", "L20S15-07" — pull from title/filename if visible

  // Legacy 7 — always empty strings in v7:
  "summary": "", "keyFindings": "", "purpose": "", "testConditions": "",
  "rootCause": "", "decision": "", "recommendedAction": ""
}
```

## v7 critical rules

- **reportType picks the shape**: multi_arm → fill `comparisons[]`, leave baselineLabel/variantLabel empty. doe_factorial → also emit ≤2 evidence rows summarising best/worst cell. quality_log → verdict="", evidence rows optional. reliability_validation → verdict ∈ passed/failed.
- **No duplicate facts**: numbers go ONCE in evidence. Decisions go ONCE in actions. Summary sentence ONCE in headline. Surrounding context (where/why) ONCE in context. The card reads top-to-bottom without repetition.
- **deltaSign**: "down" = improvement (lower NG = better); "up" = worsening; "flat" = no movement.
- **isBest / isWorst**: exclusive flags on the best/worst arms only in multi-arm comparisons.

## Translations (tr_ko, tr_vi)

Translate **only** these narrative fields:
- `headline`
- each `actions[i].text`
- `context.process`, `context.stage`, `context.baselineReason`

Keep **verbatim** (do not translate): `verdict` enum, `evidence` rows (metric labels, numbers, units, deltaText), `tags`, product codes, sheet/file names, machine names like "VP", "CD", "Sub1".

`tr_ko` / `tr_vi` structure:
```jsonc
{
  "headline": "...",
  "actions": [ {"priority":1,"kind":"action","text":"..."} ],  // preserve priority + kind
  "context": { "process":"...", "stage":"...", "baselineReason":"..." },
  "summary":"", "keyFindings":"", "purpose":"", "testConditions":"",
  "rootCause":"", "decision":"", "recommendedAction":""
}
```

## Output rules

- Write each JSON with `Write` tool to the exact path `_tmp_results/NNN.json`.
- Valid strict JSON only — no markdown fence, no trailing comments.
- If TSV is unparseable, still emit a minimal valid JSON with reportType="quality_log", verdict="", headline describing what you saw, and an action of kind="investigate".
- Process each NNN in your assigned range. Do not skip.

## Process notes

- Many TSVs contain Vietnamese, Korean, and English mixed. Don't worry about exact translation of every word — pick out the key facts: what was tested, what changed, NG counts/rates, conclusion.
- Numbers like "3/100", "1.5%", "53/1200" are typical NG counts. Look for them in highlighted (`〔bg=#FFFF00〕` or `〔bg=#FF0000〕`) cells — those usually mark NGs or summary rows.
- Product codes seen in title: BRS-161014, BRS-161016, MSU-L20S15-07, L20S15-07, TIU C11-20, TIU L5S3-01, MSU-20S15-07.
- Common processes: VP-CD bonding, Coil assembly, Frame inspection, Suspension, AWF, Plasma clean, Dry UV.

When done, output a brief summary: `done N/M, range XXX-YYY`. Nothing more.
