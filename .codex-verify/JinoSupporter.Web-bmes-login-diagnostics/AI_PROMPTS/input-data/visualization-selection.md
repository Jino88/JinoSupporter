Visualization selection rules:

Every analyzed workbook must include a visualization decision before the final
report is written. This applies to one newly added file and to every file in
batch mode.

Decision output:

- `판정: 확정` when workbook evidence is enough to choose a chart.
- `판정: 확정 불가` when the chart cannot be chosen safely.
- recommended visualization type
- why that visualization fits
- exact workbook fields/groups/metrics to use
- missing information if not confirmed

Choose by evidence, not by filename:

- Use table headers, merged-cell labels, units, numeric patterns, row labels,
  comparison groups, and workbook notes.
- Decide whether the primary result is defect-rate/yield, continuous
  measurement/spec, time/lot trend, two-factor condition matrix, defect mix, or
  document/photo-heavy evidence.

Mandatory chart rules:

- Any comparison of NG rate, defect rate, yield rate, PPM, OK/NG ratio, or
  below-spec count must include a vertical bar chart plus a numeric table. Do
  not leave a rate comparison as table-only.
- Vertical bars must rise from a shared baseline with proportional heights. A
  lower value must visibly create a shorter column, and zero must sit on the
  baseline with no filled column.
- Use exact value labels on bars and keep the numeric table directly below or
  beside the chart.
- Use `Normal` as the visible comparison label. If the workbook has no valid
  Normal, label it `Normal 미확인` and do not claim improvement/worsening.

Recommended visualization by data type:

- NG rate / defect rate / yield / PPM:
  `vertical grouped bar + delta + numeric table`.
- Multiple defect types inside each lot/condition:
  total NG rate vertical bars plus defect-mix breakdown table or compact
  composition strip.
- Sample-level measurement values with units/spec such as kgf, mm, um, Gauss,
  Hz, dB, SPL, THD, F0, tension:
  `vertical scatter/dot distribution + SPEC/LSL/USL/Average lines + summary
  table`.
- Measurement summary only:
  summary comparison table plus clearly labeled summary chart. State that raw
  sample points are unavailable.
- Date/lot/line trend with comparable metrics:
  line/control chart plus summary bars.
- Two independent variables with one result metric and at least four
  combinations:
  heatmap/matrix as primary visualization. Rows = one variable, columns = the
  other, cell = exact metric plus NG/Input when available.
- Many NG type columns or defect names:
  Pareto or ranked bars plus cumulative/share table.
- Mostly pictures, sparse text, or ambiguous numeric meaning:
  `판정: 확정 불가`.

Scatter/dot details:

- x-axis = group/condition/sample category.
- y-axis = measured value.
- One dot per raw sample when raw samples exist.
- Draw SPEC/LSL/USL/Average as horizontal reference lines.
- Samples in the same group align on one vertical centerline. Do not use
  horizontal jitter as the main encoding.

Mark `판정: 확정 불가` when:

- numbers may be counts, rates, or measurement values but labels are unclear
- denominators/input quantities are missing for rate comparison
- Normal/test or condition mapping is implied but not explicitly mapped
- spec limits/pass criteria are missing for measurement judgement
- both defect-rate and measurement/spec tables are strong but the workbook
  objective does not identify the primary result

The recommendation must be practical and specific. Use names such as `vertical
bar + table`, `scatter/dot + SPEC line`, `heatmap + top-risk bar`, `Pareto +
defect share`, or `trend + summary bars`, not vague words like "graph".
