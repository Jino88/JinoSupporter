Analysis output rules:

- For INPUT DATA analysis, the visible product is analysisHtml. The web UI will
  display it directly in an iframe srcdoc and will not rebuild, regroup, style,
  sanitize, or fix your HTML.
- Return only JSON:
  {"parameters":{...},"analysisText":"...","analysisHtml":"..."}
- Extra diagnostic fields are allowed only if they do not replace the required
  fields. analysisHtml must be complete and self-contained.

Parameters:

- Decide one parameters object before writing analysisText or analysisHtml.
- Copy the same values into analysisText and the visible HTML summary.
- Parameters object fields must be exactly:
  reviewPurpose, tags, purpose, purposeCode, targetDefects, reviewItems, model,
  date, confidence.
- Use arrays for tags, targetDefects, and reviewItems. Use empty strings or
  arrays only when workbook text does not provide enough evidence. confidence
  is 0 to 1.
- Do not fill front parameters from filename, dataset name, source path,
  sequence number, or date suffix.

analysisText:

- Start with this block in order:
  - `검토 목적: ...`
  - `검토 태그: ...`
  - `목적: ...`
  - `대상불량: ...`
  - `검토사항: ...`
  - `Model: ...`
  - `Date: ...`
  - `Confidence: ...`
- Then summarize the judgement, visualization decision, and key evidence.

analysisHtml:

- Must include `<!doctype html>`, html/head/body, a style block, and no external
  assets.
- Use only HTML/CSS/SVG. Do not use scripts, inline events, iframes, forms,
  network resources, or app CSS classes.
- Korean by default.
- Preserve units in console-safe ASCII where needed: degC, degF, kg, mg, mm,
  um, cm, mm2, cm2, m2, m3, kPa.

Visible report shape:

1. Top dashboard summary
   - title, current question/objective, model/date, confidence
   - key judgement first
   - cards for major findings

2. 보고서별 검토 목적/결과 판정
   - table based on reportReviewMatrix
   - columns: 보고서/Sheet, 검토 항목, 검토 목적, 보이는 결과, 검토 구분,
     공정 근거, 기능 근거, 사용 판단, 한계
   - This section must make it clear whether each report points to
     공정 불량 검토, 기능 불량 검토, or 공정+기능 연계 검토.

3. 시각화 권장
   - show `판정: 확정` or `판정: 확정 불가`
   - recommended visualization
   - reason
   - exact workbook fields/groups/metrics to use
   - missing information if not confirmed

4. 현재 문제 및 전체 결론
   - answer the user's review objective directly
   - do not invent a decision absent from workbook evidence

5. 공정 불량 검토
   - process-side evidence only: process, line, jig, mold, lot, material, dry,
     UV, bonding, pressure, timing, plasma, laser, supplier, etc.

6. 기능 불량 검토
   - function/spec evidence only: final function NG, VP+CD separate NG, SPL,
     THD, F0, tension, measurement/spec, pass/fail, reliability, etc.

7. 공정-기능 연계 해석
   - explain which process conditions appear linked to function NG/spec results
   - separate strong evidence from weak or missing Normal

8. 실측 수치 집계 및 차트
   - every defect-rate/NG-rate/yield/PPM comparison must have a vertical bar
     chart and a numeric table
   - every sample-level measurement/spec comparison must use scatter/dot with
     spec/reference lines when raw samples exist
   - every two-factor matrix must use a heatmap/matrix when evidence supports it

9. 다음 진행 항목
   - practical next actions in order
   - identify whether to pursue process defect work, function defect work, or
     both

Terminology:

- Visible comparison wording must use `Normal`, `Normal 대비`, `Normal 값`, or
  `Normal 미확인`.
- Do not display `Local Control`, `Control`, `Baseline`, `Reference`, or
  `대조군` as visible labels. If the workbook uses those words, normalize the
  visible label to `Normal` and keep the original only in a short evidence note
  when needed.

Chart requirements:

- For NG rate, defect rate, yield, PPM, OK/NG ratio, and below-spec counts:
  vertical bar chart + numeric table is mandatory.
- Draw real columns rising from one baseline in a fixed-height plot area.
- Bar height must be proportional to value within that chart. Normalize the
  largest compared value to 100% height while preserving exact labels.
- Zero values must sit on the baseline with no filled column.
- Do not represent a bar chart as same-height strips, badges, underlines,
  legend marks, or fixed-height pills.
- Include exact group labels, value, delta, ratio, n/input count, and evidence
  note in the table.
- If denominators are not comparable, state that and avoid overstating delta.
- For measurement raw samples, draw one vertical-positioned dot per sample,
  show SPEC/LSL/USL/Average lines, and include avg/min/max/n in the table.
- For heatmaps, show row variable, column variable, exact cell value, and
  NG/Input count when available.

Layout and alignment:

- Use a clean dashboard layout with stable dimensions. Cards can be used for
  repeated items or summary cards, not nested inside other cards.
- For metadata rows such as `검토 항목`, `검토 목적`, `보이는 결과`, use a CSS grid
  so text aligns by label and value.
- Tables must use border-collapse, fixed or deliberate column widths,
  centered headers, right-aligned numeric columns, and top-aligned text cells.
- Long Korean text must wrap cleanly. Do not let text overlap or overflow its
  parent.
- Keep labels compact and use tabular numerals for numbers.
- Avoid one-color themes. Use restrained status colors for worse/warn/good and
  keep the background light enough for tables.

Speed and scope:

- Do not re-analyze unrelated sheets after reportReviewMatrix and visualization
  decision are complete.
- Use selected evidence blocks and source citations for final writing.
- Do not run builds/tests, edit files, call external services, or create temp
  scripts to format the final JSON.

Return only valid JSON with no markdown fence and no surrounding prose.
