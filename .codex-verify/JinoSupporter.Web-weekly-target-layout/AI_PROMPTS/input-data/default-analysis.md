Default analysis when currentStep.prompt and checked autoPrompts are empty:

- Start from RequestJson.reviewPurpose when present. Treat it as the review
  question to answer, then confirm, correct, or narrow it using workbook
  evidence.
- Use the AI pre-analysis and reportReviewMatrix as the first context for every
  later section.
- Separate the conclusion into:
  - `공정 불량 검토`: process-side factors to inspect/change
  - `기능 불량 검토`: function/spec-side results to confirm
  - `공정-기능 연계 해석`: process changes that appear to affect function NG or
    measurement/spec results
- Summarize the workbook purpose and test intent.
- Identify model, material, lot/date, process, line, sample/input count,
  Normal groups, test/changed groups, and comparison blocks.
- Extract only key result tables with important NG counts, rates, OK/input,
  measurement values, spec limits, and workbook notes.
- Compare meaningful same-event groups: Normal vs test/change, before vs after,
  old vs new, lot/mold/line/condition blocks.
- Do not claim improvement/worsening when Normal or same-event comparison is
  missing. In that case, report it as ranking or absolute result review.
- Use `Normal`, `Normal 대비`, `Normal 값`, or `Normal 미확인` in visible text.
  Do not display `Local Control`, `Control`, `Baseline`, or `대조군` as the
  visible comparison label.
- State the practical decision first, then show evidence. Do not invent a
  decision not present in workbook evidence.
