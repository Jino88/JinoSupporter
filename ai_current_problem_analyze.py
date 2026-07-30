from __future__ import annotations

import argparse
import html
import json
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import create_current_problem_search_html as search_support


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SAMPLE_DIR = Path(r"C:\Users\jhbyun\Desktop\새 폴더 (4)\sample_ready")
DEFAULT_DB = Path(r"D:\000. MyWorks\002. DB\process-review.db")

sys.path.insert(0, str(SCRIPT_DIR))
import _ai_batch_helper as batch_helper  # noqa: E402
import _xlsx_render  # noqa: E402


def log(message: str) -> None:
    stamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{stamp}] {message}", flush=True)


def norm(value: Any) -> str:
    return str(value or "").casefold()


def compact(value: Any) -> str:
    return re.sub(r"[^a-z0-9가-힣]+", "", norm(value))


def tokens(text: str) -> list[str]:
    stop = {
        "report", "test", "result", "check", "date", "model", "data", "the", "and",
        "with", "from", "normal", "new", "old", "검토", "개선", "뭐해야함",
    }
    found = re.findall(r"[A-Za-z0-9가-힣.+/-]+", str(text or ""))
    out: list[str] = []
    for token in found:
        key = token.casefold()
        if len(key) < 2 or key in stop:
            continue
        if key not in out:
            out.append(key)
    return out


def row_text(row: dict[str, Any]) -> str:
    parts = [
        row.get("datasetName"),
        row.get("fileNames"),
        row.get("model"),
        row.get("aiModel"),
        row.get("date"),
        row.get("reviewPurpose"),
        row.get("purpose"),
        row.get("evidenceSummary"),
        row.get("uncertainty"),
        " ".join(row.get("targetDefects") or []),
        " ".join(row.get("reviewItems") or []),
        " ".join(row.get("tags") or []),
    ]
    return " ".join(str(x or "") for x in parts).casefold()


def field_text(row: dict[str, Any], field: str) -> str:
    value = row.get(field)
    if isinstance(value, list):
        return " ".join(str(x) for x in value).casefold()
    return str(value or "").casefold()


def detect_model(problem: str, rows: list[dict[str, Any]], explicit: str = "") -> str:
    if explicit:
        return explicit
    problem_key = compact(problem)
    models = sorted({str(row.get("model") or "") for row in rows if row.get("model")}, key=len, reverse=True)
    for model in models:
        key = compact(model)
        if key and key in problem_key:
            return model
        short = re.sub(r"s08zz|s15|l20|brs|msu|tiu|tim", "", key)
        if len(short) >= 5 and short in problem_key:
            return model
    return ""


def matched_terms(problem: str, terms: list[dict[str, str]]) -> list[dict[str, str]]:
    text = norm(problem)
    ctext = compact(problem)
    out: list[dict[str, str]] = []
    for term in terms:
        name = str(term.get("term") or "").strip()
        if not name:
            continue
        if norm(name) in text or compact(name) in ctext:
            out.append(term)
    return out[:20]


def applicable_prompt_rules(problem: str, model: str, rules: list[dict[str, str]]) -> list[dict[str, str]]:
    text = norm(problem)
    model_text = norm(model)
    out: list[dict[str, str]] = []
    for rule in rules:
        dataset = norm(rule.get("dataset"))
        rule_model = norm(rule.get("model"))
        answer = norm(rule.get("canonical"))
        if rule_model and model_text and (rule_model in model_text or model_text in rule_model):
            out.append(rule)
        elif dataset and dataset in text:
            out.append(rule)
        elif "standard" in answer and "standard" in text:
            out.append(rule)
    return out[:12]


def score_row(
    row: dict[str, Any],
    query_tokens: list[str],
    model: str,
    term_hits: list[dict[str, str]],
    excluded_datasets: set[str],
) -> tuple[float, list[str]]:
    reasons: list[str] = []
    score = 0.0
    hay = row_text(row)
    row_model = str(row.get("model") or "")
    if model:
        if row_model == model:
            score += 80
            reasons.append("model exact")
        elif row_model and (row_model in model or model in row_model):
            score += 42
            reasons.append("model related")
    for token in query_tokens:
        if token in field_text(row, "targetDefects"):
            score += 9
            reasons.append(f"defect:{token}")
        elif token in field_text(row, "reviewItems"):
            score += 7
            reasons.append(f"item:{token}")
        elif token in field_text(row, "tags"):
            score += 6
            reasons.append(f"tag:{token}")
        elif token in norm(row.get("reviewPurpose")) or token in norm(row.get("purpose")):
            score += 4
            reasons.append(f"review:{token}")
        elif token in hay:
            count = hay.count(token)
            score += min(6, count * 1.5)
            reasons.append(f"text:{token}")
    for term in term_hits:
        name = norm(term.get("term"))
        if name and name in hay:
            score += 12
            reasons.append(f"user term:{term.get('term')}")
    score += float(row.get("confidence") or 0) * 4
    if row.get("needsDetailedAnalysis"):
        score += 2
    if row.get("datasetName") in excluded_datasets:
        score -= 100
        reasons.append("user excluded")
    return score, list(dict.fromkeys(reasons))[:10]


def retrieve(sample_dir: Path, problem: str, model: str = "", top: int = 20) -> dict[str, Any]:
    state = search_support.build_state(sample_dir)
    rows: list[dict[str, Any]] = state["rows"]
    detected_model = detect_model(problem, rows, model)
    query_tokens = tokens(problem)
    term_hits = matched_terms(problem, state["terms"])
    rule_hits = applicable_prompt_rules(problem, detected_model, state["promptRules"])
    excluded = set(state["excludedDatasets"])
    scored: list[dict[str, Any]] = []
    for row in rows:
        score, reasons = score_row(row, query_tokens, detected_model, term_hits, excluded)
        if score > 2:
            scored.append({"score": round(score, 2), "reasons": reasons, "row": row})
    scored.sort(key=lambda x: x["score"], reverse=True)
    return {
        "problem": problem,
        "detectedModel": detected_model,
        "queryTokens": query_tokens,
        "userTerms": term_hits,
        "promptRules": rule_hits,
        "excludedDatasets": sorted(excluded),
        "candidates": scored[:top],
        "stateStats": state["stats"],
    }


def build_prompt(package: dict[str, Any], candidate_limit: int) -> str:
    candidates = []
    for item in package["candidates"][:candidate_limit]:
        row = item["row"]
        candidates.append({
            "score": item["score"],
            "matchReasons": item["reasons"],
            "datasetName": row.get("datasetName"),
            "model": row.get("model"),
            "date": row.get("date"),
            "purposeCode": row.get("purposeCode"),
            "reviewPurpose": row.get("reviewPurpose"),
            "purpose": row.get("purpose"),
            "targetDefects": row.get("targetDefects"),
            "reviewItems": row.get("reviewItems"),
            "tags": row.get("tags"),
            "confidence": row.get("confidence"),
            "needsDetailedAnalysis": row.get("needsDetailedAnalysis"),
            "evidenceSummary": row.get("evidenceSummary"),
            "uncertainty": row.get("uncertainty"),
        })

    guidance = {
        "matchedUserTerms": package["userTerms"],
        "matchedPromptRules": package["promptRules"],
        "excludedDatasets": package["excludedDatasets"][:12],
    }
    payload = {
        "currentProblem": package["problem"],
        "detectedModel": package["detectedModel"],
        "queryTokens": package["queryTokens"],
        "userGuidance": guidance,
        "retrievedCandidates": candidates,
    }
    return f"""
You are analyzing historical manufacturing validation/review reports for JinoSupporter.

The user has a current manufacturing problem. A retrieval step already selected related historical reports from `demo_index.json`.
You must now do the AI analysis step, not just keyword grouping.

Important rules:
- Use only the retrieved report summaries and user guidance provided below.
- Treat user guidance from `ai_term_guidance.md` and `prompt_update_requests.md` as precedence/tie-breaker rules.
- If a report says final Decision was not visible, do not invent a final pass/fail. Say that the extracted result is limited.
- First classify every retrieved candidate report one by one before making groups. Do not skip this step.
- For each report, decide:
  - what the report actually reviewed,
  - why it was reviewed,
  - what the visible result was,
  - whether the useful evidence is process defect evidence, function defect evidence, both, or unclear,
  - what should and should not be used for the current problem.
- Use report-level classification as the source of truth for later grouping. Do not create a group if the per-report purpose/result/domain does not support it.
- Group similar historical reports by review pattern/result pattern, not just by identical words.
- For NG rate / defect rate / PPM findings, do not compare absolute values across reports as if they share one baseline. Each workbook/date/condition can have its own Normal/reference rate. Describe the result as Test/changed condition versus its same-sheet/same-date Normal or 기준 조건 when that Normal is visible.
- Visible wording rule: in Korean report text, chart legends, table headers, and labels, do not write "Local Control", "local control", "Control", or "대조군". Use "Normal", "Normal 대비", "Normal 값", or "Normal 미확인" instead.
- Split process defect review from function defect review:
  - "공정 불량 검토" means manufacturing/process/assembly/appearance/vision issues such as VP/CD separate, bonding offset, over glue, material state, mold, bending, UV dry, wait dry, press JIG, line, machine, operator, AOI/Vision process NG.
  - "기능 불량 검토" means final or electrical/acoustic/functional result issues such as Function NG, NG hearing, Noise, Touch, Sigma, SPL, THD, DCR, impedance, acoustic/electrical characteristic NG.
  - If one historical report contains both process NG and function NG, keep them linked but do not merge them into one group. Use separate groups or clearly mark the group as "공정+기능 연계 검토" and explain which evidence is process-side and which is function-side.
  - Do not use a function NG rate as direct proof of process separate improvement unless the report also shows the matching process defect metric.
- For each group, explain what was checked, what the reported result/evidence suggests, and how it helps the current problem.
- Korean output. Preserve model names, process names, defect names, and dates.
- Be concise but concrete.

Return ONLY valid JSON with this exact shape:
{{
  "currentProblemSummary": "...",
  "detectedModel": "...",
  "overallConclusion": "...",
  "reportReviewMatrix": [
    {{
      "datasetName": "...",
      "reviewedItem": "what was checked in this report",
      "reviewPurpose": "why the report was run",
      "visibleResult": "what result is visible; say Decision not visible when applicable",
      "primaryReviewDomain": "공정 불량 검토|기능 불량 검토|공정+기능 연계 검토|기타/미확인",
      "processDefectEvidence": "process-side evidence only, or empty if none",
      "functionDefectEvidence": "function-side evidence only, or empty if none",
      "recommendedUseForCurrentProblem": "how to use this report for the current problem",
      "doNotUseAsEvidenceFor": "what this report must not be used to prove",
      "confidence": "high|medium|low",
      "limits": ["...", "..."]
    }}
  ],
  "reportGroups": [
    {{
      "groupTitle": "...",
      "groupType": "기능 NG 확인|조건 비교|원인 조사|공정/설비 변경 검증|재료/부품 변경 검증|기타",
      "reviewDomain": "공정 불량 검토|기능 불량 검토|공정+기능 연계 검토|기타/미확인",
      "processEvidenceSummary": "...",
      "functionEvidenceSummary": "...",
      "reportCount": 0,
      "memberReports": ["datasetName", "..."],
      "representativeReports": ["datasetName", "..."],
      "whatWasChecked": "...",
      "reviewResultSummary": "...",
      "similarityToCurrentProblem": "...",
      "usefulEvidence": ["...", "..."],
      "limits": ["...", "..."]
    }}
  ],
  "recommendedReviewPlan": [
    {{"step": 1, "item": "...", "why": "...", "relatedGroups": ["..."]}}
  ],
  "improvementIdeas": [
    {{"idea": "...", "basis": "...", "risk": "..."}}
  ],
  "missingInformationNeeded": ["...", "..."],
  "candidateReportsForDeepDive": ["datasetName", "..."]
}}

INPUT:
{json.dumps(payload, ensure_ascii=False, indent=2)}
""".strip()


MEASURE_KEYWORD_RE = re.compile(
    r"(ng|rate|defect|ppm|%|qty|pc|input|fail|ok|function|noise|touch|sigma|spl|thd|"
    r"tension|spec|normal|control|baseline|before|old|std|standard|reference|ref|can use|can not|cannot|decision|result|height|gauss|"
    r"대조|기준|비교|변경전|기존|정상|양품|"
    r"mold|ir|line|lot|vp|cd|separate|hearing)",
    re.IGNORECASE,
)


def excerpt_measurement_lines(text: str, max_chars: int = 7000) -> str:
    lines = text.splitlines()
    keep: set[int] = set()
    for idx, line in enumerate(lines):
        if MEASURE_KEYWORD_RE.search(line) or re.search(r"\d+(?:\.\d+)?\s*(?:%|ppm)\b", line, re.IGNORECASE):
            for offset in (-2, -1, 0, 1, 2):
                pos = idx + offset
                if 0 <= pos < len(lines):
                    keep.add(pos)
    selected = [lines[i] for i in sorted(keep)]
    excerpt = "\n".join(selected)
    if len(excerpt) > max_chars:
        return excerpt[:max_chars] + "\n[...MEASUREMENT CONTEXT TRUNCATED...]"
    return excerpt


def render_dataset_measurement_context(
    con: sqlite3.Connection,
    dataset_name: str,
    workbook_dir: Path,
    max_chars: int = 12000,
) -> str:
    parts = [f"### DATASET: {dataset_name}"]
    try:
        paths = batch_helper.get_excel_files(con, dataset_name, out_dir=str(workbook_dir))
    except Exception as exc:
        paths = []
        parts.append(f"[WORKBOOK_MATERIALIZE_FAILED] {exc}")

    for path in paths:
        parts.append(f"#### WORKBOOK: {Path(path).name}")
        try:
            rendered = _xlsx_render.render_workbook(path)
            parts.append(excerpt_measurement_lines(rendered, max_chars=max_chars))
        except Exception as exc:
            parts.append(f"[WORKBOOK_RENDER_FAILED] {exc}")

    if not paths:
        try:
            paste = batch_helper.get_excel_paste(con, dataset_name)
        except Exception as exc:
            paste = None
            parts.append(f"[EXCEL_PASTE_READ_FAILED] {exc}")
        if paste:
            parts.append("#### EXCEL_PASTE")
            parts.append(excerpt_measurement_lines(paste, max_chars=max_chars))
    return "\n".join(x for x in parts if x)


def group_report_names(analysis: dict[str, Any], package: dict[str, Any]) -> list[str]:
    names: list[str] = []
    for group in analysis.get("reportGroups") or []:
        if not isinstance(group, dict):
            continue
        for key in ("memberReports", "representativeReports"):
            for name in group.get(key) or []:
                if name and name not in names:
                    names.append(str(name))
    if names:
        return names
    for item in package.get("candidates") or []:
        name = ((item.get("row") or {}).get("datasetName") or "")
        if name and name not in names:
            names.append(str(name))
    return names


def build_measurement_prompt(
    db_path: Path,
    sample_dir: Path,
    package: dict[str, Any],
    analysis: dict[str, Any],
    report_limit: int = 20,
) -> tuple[str, list[str]]:
    names = group_report_names(analysis, package)[:report_limit]
    workbook_dir = sample_dir / "ai_current_problem" / "_measure_workbooks"
    workbook_dir.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path), timeout=60)
    con.execute("PRAGMA busy_timeout=60000")
    try:
        contexts = [
            render_dataset_measurement_context(con, name, workbook_dir)
            for name in names
        ]
    finally:
        con.close()

    payload = {
        "currentProblem": package.get("problem"),
        "detectedModel": package.get("detectedModel"),
        "reportReviewMatrix": analysis.get("reportReviewMatrix") or [],
        "reportGroups": analysis.get("reportGroups") or [],
        "workbookMeasurementContexts": contexts,
    }
    prompt = f"""
You are extracting numeric measurement/result data from rendered Excel workbook text.

Task:
- For each similar report group, read the workbook cell text and extract actual numeric result values.
- Use reportReviewMatrix as the source of truth for each report's purpose/result/domain. If workbook numeric text conflicts with the report-level purpose/result summary, keep the number but state the conflict/limit.
- For every extracted visualization and record, connect it back to the report's reviewedItem, reviewPurpose, visibleResult, and primaryReviewDomain when possible.
- Do not extract a chart as "current problem evidence" when reportReviewMatrix says the report should not be used for that proof. You may still include it as reference with the same limit.
- Decide the visualization type by data meaning:
  - Use "verticalBar" when data is defect rate, NG rate, PPM, NG count/ratio, or Normal-vs-Test comparison.
  - Use "scatter" when data is continuous measurements such as tension, gauss, height, impedance, DCR, SPL/THD measurement. Include specMin/specMax lines when spec values are visible.
  - Use "heatmap" when data varies by two dimensions such as mold x IR, line x condition, machine x date, or worker x defect type.
  - Use "table" only when numeric extraction is too sparse for a chart.
- Preserve actual numbers and units. Do not invent missing values.
- If a value is a percent, store value as the percent number, e.g. 13.6 for 13.6%.
- Review domain split rule:
  - Assign every visualization and record to one of: "공정 불량 검토", "기능 불량 검토", "공정+기능 연계 검토", "기타/미확인".
  - Process defect metrics include VP/CD separate, VP+CD separate, bonding separate/offset, over glue, material appearance, mold, bending, UV dry, wait dry, press JIG, AOI/Vision process NG, process yield/defect rate.
  - Function defect metrics include Function NG, NG hearing, Noise, Touch, Sigma, SPL, THD, DCR, impedance, electrical/acoustic characteristic NG, final function result.
  - Do not put process defect metrics and function defect metrics in the same chart/table just because they are from the same workbook. Create separate visualizations: one for process defect review, one for function defect review.
  - If a workbook compares process changes and final function together, add a separate "공정+기능 연계 검토" table that shows the relationship, but still keep the raw process chart and raw function chart separate.
  - Do not rank process conditions using a function NG rate unless the chart/table title and reviewDomain clearly say it is a function review.
- Scatter / continuous measurement rule:
  - For scatter charts, extract every visible individual measurement point. Do not collapse raw values into only average points.
  - If the workbook shows N samples or repeated rows/columns for tension, gauss, height, impedance, DCR, SPL/THD, include all visible sample values as records with statType "raw" and sampleIndex/sampleCount when visible.
  - Also extract avg, max, and min values when visible. If raw points are visible but avg/max/min are not, calculate avg/max/min per comparable condition from the visible raw records and include them in statSummary.
  - Compare avg-to-avg, max-to-max, and min-to-min between Test/changed/new condition and its same-sheet/same-date Normal/reference condition when visible.
  - Do not compare a Test average to a Normal max/min. Keep statistic types aligned.
  - Include count n for every condition used in scatter or summary comparisons.
  - Consistency check: if statSummary says n=10 for a condition, the records array must contain 10 raw records for that same condition unless the workbook only shows aggregate avg/max/min and no raw sample columns. Do not output partial raw records with n from the full table.
  - If raw sample columns are visible but too many to include, create a separate scatter visualization for fewer comparable conditions rather than dropping sample points. Completeness of raw points is more important than covering many conditions in one chart.
  - Do not write "all raw points" or "raw distribution" unless every visible sample point used for that condition is present as a record.
- NG rate / defect rate / PPM rule:
  - Do not interpret or rank by absolute rate alone.
  - Each dataset can have a different Normal/reference defect rate. Pair every Test/changed/new condition with its same-workbook/same-sheet/same-date Normal/reference value from the same table, model, mold, line, lot, or condition when visible.
  - Put Normal comparison fields on each non-Normal record when possible: controlLabel, controlValue, deltaVsControl, ratioVsControl, comparisonBasis, controlScope. These JSON keys are internal; visible report text must call this "Normal".
  - deltaVsControl = value - Normal value in the same unit. ratioVsControl = value / Normal value when Normal value is non-zero.
  - If the Normal value is zero, set ratioVsControl to null and explain the zero-baseline case in note/interpretation.
  - If the Normal value is not visible, set controlValue to null, write "Normal 미확인" in note/limits, and do not use that record for improvement/worsening ranking.
- If Normal value is visible, include it as a record with series "Normal". Set normalValue only when one local baseline applies to the whole chart. If Normal changes by date/condition, keep the per-record control fields instead of a single normalValue.
- Include sourceCell/sourceText when visible from rendered text.
- Korean summaries.

Return ONLY valid JSON:
{{
  "measurementSummary": "...",
  "visualizations": [
    {{
      "groupTitle": "...",
      "chartTitle": "...",
      "reviewDomain": "공정 불량 검토|기능 불량 검토|공정+기능 연계 검토|기타/미확인",
      "metricDomain": "process|function|linked|unknown",
      "reviewedItem": "what this chart/table is checking",
      "reviewPurpose": "why this evidence was reviewed",
      "visibleResult": "visible report result relevant to this chart/table",
      "chartType": "verticalBar|scatter|heatmap|table",
      "metricName": "NG rate|Defect rate|Tension|...",
      "unit": "%|PPM|g|N|...",
      "whyThisChart": "...",
      "normalValue": null,
      "specMin": null,
      "specMax": null,
      "statSummary": [
        {{
          "condition": "...",
          "series": "...",
          "n": 0,
          "avg": null,
          "max": null,
          "min": null,
          "controlCondition": "...",
          "controlN": null,
          "avgDeltaVsControl": null,
          "maxDeltaVsControl": null,
          "minDeltaVsControl": null,
          "comparisonBasis": "same sheet/date/model/mold/line/lot/condition"
        }}
      ],
      "records": [
        {{
          "datasetName": "...",
          "report": "...",
          "reviewDomain": "공정 불량 검토|기능 불량 검토|공정+기능 연계 검토|기타/미확인",
          "metricDomain": "process|function|linked|unknown",
          "reviewedItem": "...",
          "reviewPurpose": "...",
          "visibleResult": "...",
          "date": "...",
          "category": "...",
          "series": "...",
          "condition": "...",
          "line": "...",
          "mold": "...",
          "ir": "...",
          "x": "...",
          "y": "...",
          "value": 0.0,
          "unit": "%",
          "statType": "raw|avg|max|min",
          "sampleIndex": null,
          "sampleCount": null,
          "controlLabel": "Normal|Before|Old|...",
          "controlValue": null,
          "deltaVsControl": null,
          "ratioVsControl": null,
          "comparisonBasis": "same sheet/date/model/mold/line/lot/condition",
          "controlScope": "which rows/cells were used as Normal/reference",
          "sourceCell": "...",
          "sourceText": "...",
          "note": "..."
        }}
      ],
      "interpretation": "...",
      "limits": ["..."]
    }}
  ]
}}

INPUT:
{json.dumps(payload, ensure_ascii=False, indent=2)}
""".strip()
    return prompt, names


def extract_json(text: str) -> dict[str, Any]:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except Exception:
        pass
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        return json.loads(text[start:end + 1])
    raise ValueError("AI output did not contain valid JSON")


def run_codex_exec(prompt: str, out_path: Path, effort: str, timeout_sec: int, stage: str) -> str:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()
    codex_exe = shutil.which("codex") or shutil.which("codex.cmd") or shutil.which("codex.exe")
    if not codex_exe:
        raise FileNotFoundError("codex executable not found on PATH")
    cmd = [
        codex_exe,
        "exec",
        "-c",
        f'model_reasoning_effort="{effort}"',
        "--cd",
        str(SCRIPT_DIR),
        "--dangerously-bypass-approvals-and-sandbox",
        "--skip-git-repo-check",
        "--color",
        "never",
        "--output-last-message",
        str(out_path),
        "-",
    ]
    log(f"{stage}: Codex CLI started (effort={effort}, timeout={timeout_sec}s)")
    log(f"{stage}: output file -> {out_path}")
    started = time.monotonic()
    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    if proc.stdin is None or proc.stdout is None:
        raise RuntimeError("failed to open Codex CLI pipes")

    lines: queue.Queue[str | None] = queue.Queue()

    def read_stdout() -> None:
        try:
            for line in proc.stdout:
                lines.put(line)
        finally:
            lines.put(None)

    reader = threading.Thread(target=read_stdout, daemon=True)
    reader.start()
    proc.stdin.write(prompt)
    proc.stdin.close()

    output_parts: list[str] = []
    last_heartbeat = started
    stdout_closed = False
    while True:
        try:
            line = lines.get(timeout=0.5)
        except queue.Empty:
            line = ""
        if line is None:
            stdout_closed = True
        elif line:
            output_parts.append(line)
            sys.stdout.write(line)
            sys.stdout.flush()
            last_heartbeat = time.monotonic()

        now = time.monotonic()
        if proc.poll() is not None and (stdout_closed or lines.empty()):
            break
        if now - started > timeout_sec:
            proc.kill()
            log(f"{stage}: timeout after {timeout_sec}s")
            raise subprocess.TimeoutExpired(cmd, timeout_sec)
        if now - last_heartbeat >= 30:
            log(f"{stage}: still running ({int(now - started)}s elapsed)")
            last_heartbeat = now

    reader.join(timeout=2)
    output_text = "".join(output_parts)
    elapsed = int(time.monotonic() - started)
    if proc.returncode != 0:
        raise RuntimeError(f"codex exit {proc.returncode}: {output_text[-2000:]}")
    log(f"{stage}: Codex CLI finished ({elapsed}s)")
    return out_path.read_text(encoding="utf-8", errors="replace") if out_path.exists() else output_text


def call_codex(prompt: str, out_path: Path, effort: str, timeout_sec: int, stage: str) -> dict[str, Any]:
    output_text = run_codex_exec(prompt, out_path, effort, timeout_sec, stage)
    log(f"{stage}: parsing JSON")
    return extract_json(output_text)


def call_codex_text(prompt: str, out_path: Path, effort: str, timeout_sec: int, stage: str) -> str:
    output_text = run_codex_exec(prompt, out_path, effort, timeout_sec, stage)
    log(f"{stage}: received HTML/text output")
    return output_text


def clean_ai_html(text: str) -> str:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:html)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    lower = text.casefold()
    start = lower.find("<!doctype")
    if start < 0:
        start = lower.find("<html")
    end = lower.rfind("</html>")
    if start >= 0 and end >= start:
        text = text[start:end + len("</html>")]
    return text.strip()


def build_ai_html_prompt(
    problem: str,
    package: dict[str, Any],
    analysis: dict[str, Any],
    measurement: dict[str, Any] | None,
) -> str:
    payload = {
        "currentProblem": problem,
        "retrievalPackage": package,
        "groupAnalysis": analysis,
        "measurementExtraction": measurement,
    }
    return f"""
You are generating the final browser report HTML for JinoSupporter.

Do not return JSON. Do not return markdown. Return ONLY a complete standalone HTML document.

The Python runner only collected source data and called AI. You must create the actual report HTML yourself:
- Write the report structure, Korean wording, CSS, and charts directly in HTML.
- All visible Korean text must be valid, readable UTF-8 Korean. Do not copy mojibake/garbled strings such as "遺꾩꽍" from inputs; rewrite them in normal Korean.
- Use embedded CSS and inline SVG only. No external CDN, no scripts that require internet, no image links.
- Use the AI group analysis and measurement extraction below.
- Report-level judgment table:
  - Before charts, render a dense table titled "보고서별 검토 목적/결과 판정".
  - Use groupAnalysis.reportReviewMatrix.
  - Required columns: 보고서, 검토 항목, 검토 목적, 보이는 결과, 판단 영역, 현재 문제에 쓰는 방법, 쓰면 안 되는 근거, 한계.
  - This table must make it clear whether each report is mainly "공정 불량 검토", "기능 불량 검토", "공정+기능 연계 검토", or "기타/미확인".
  - If a report has no visible Decision, show "Decision 미확인" in the result/limit area.
  - If the report only supports function impact, do not present it as process root-cause proof.
- Visible wording rule:
  - Never display "Local Control", "local control", "Control", or "대조군" in the HTML report.
  - Use "Normal", "Normal 대비", "Normal 값", "Normal 기준", and "Normal 미확인" for chart legends, axis labels, table headers, notes, and interpretations.
  - JSON fields named controlLabel/controlValue/controlScope are internal source fields. When rendering them, label them as Normal.
- Split process defect review from function defect review in the report:
  - Create separate visible subsections for "공정 불량 검토" and "기능 불량 검토" before the detailed historical groups.
  - Put process charts/tables only under "공정 불량 검토": VP/CD separate, bonding separate/offset, AOI/Vision process NG, material/mold/bending/UV/wait dry/press JIG/process condition data.
  - Put function charts/tables only under "기능 불량 검토": Function NG, NG hearing, Noise, Touch, Sigma, SPL, THD, DCR, impedance, final function/acoustic/electrical characteristic data.
  - If one workbook contains both process and function results, show them as separate charts/tables and add a short "공정-기능 연계 해석" note or compact table. Do not merge their bars in one chart.
  - The chart title, table header, and interpretation must state whether it is process defect review or function defect review.
  - Never use function NG data as a substitute for process separate rate; present it as downstream function impact only.
- If measurementExtraction has chartType:
  - verticalBar: render a vertical bar chart, especially for NG rate / defect rate / PPM / Normal-vs-Test comparisons. For NG/defect rates, prefer paired Test-vs-Normal bars or delta-vs-Normal bars. Include a Normal baseline line only when normalValue is a single valid same-chart Normal baseline.
  - scatter: render a scatter plot for continuous measurements such as tension, gauss, height, impedance, DCR, SPL/THD. Include every extracted raw measurement point, not only the average. Include specMin/specMax lines when present. Also show an adjacent compact table comparing avg-to-avg, max-to-max, and min-to-min for each Test/Normal condition pair, with n counts.
  - heatmap: render a heatmap for two-dimensional comparisons such as mold x IR, line x condition, machine x date. For NG/defect rates, color by deltaVsControl or ratioVsControl when available, not by absolute rate alone.
  - table: render a compact numeric table and explain why charting is limited.
- Use all available actual extracted numbers. Do not invent missing rates or measurements.
- Rate chart rule:
  - Every defect-rate/NG-rate/PPM comparison must have a vertical bar chart and a numeric table. A rate comparison must never be shown as table-only.
  - This includes every table or section with columns or wording such as 값, NG rate, Total NG rate, Defect rate, PPM, Normal 값, 차이, 비율, or %.
  - Place the vertical bar chart immediately before the matching table.
  - The chart and the table must use the same rows and same numeric values. If the table has 8 rate rows, the chart must display those 8 rate rows.
  - Use vertical bars for both direct rate values and Normal 대비 deltas. If there are many rows, use a compact vertical bar chart with shorter wrapped labels, not a table-only fallback.
  - For Normal comparison, prefer bars colored by delta direction and keep the table with Normal 값, 차이, 비율 below the chart.
- Critical scatter comparison rule:
  - A scatter chart must display all visible sample points for each condition. Do not represent a condition with only one avg marker if raw values are available.
  - If statSummary says n=10, the chart must have 10 raw point markers for that condition. If the source has 10 Test and 10 Normal samples, draw 20 raw markers.
  - Do not show a "partial raw" scatter as if it is complete. If raw records are incomplete, use a table and state that raw point extraction is incomplete.
  - Visually distinguish raw points from avg/max/min summary markers.
  - Compare summary statistics only by the same statistic type: avg vs avg, max vs max, min vs min.
  - Always show n count per condition. If raw values were not visible and only avg/max/min were visible, state that raw distribution is unavailable.
- Critical NG/defect rate comparison rule:
  - Do not compare absolute defect rates across different reports/dates as if one global Normal baseline exists.
  - Each workbook/table/date/condition may have its own Normal/reference rate. Conclusions must be based on Normal comparison: value, controlValue, deltaVsControl, ratioVsControl, and comparisonBasis.
  - Show absolute value together with its Normal value and delta/ratio when possible.
  - If Normal is missing, label the item "Normal 미확인" and treat it as reference only, not as proof of improvement/worsening.
  - If Normal changes within the chart, do not draw one global Normal line. Use paired bars, delta bars, or a comparison table, and label it as Normal.
- Domain separation rule:
  - Before drawing a chart/table, inspect reviewDomain/metricDomain and metricName/category/report text.
  - If a visualization or record set contains both process and function metrics, split it into separate displayed blocks.
  - Use neutral colors/labels consistently but do not imply process improvement from function-only charts.
- Group similar report categories together so the user can see what the historical review results were.
- Each group must show: checked item, actual result summary, related Excel reports, chart or numeric table, interpretation, limits.
- Every chart/table section must show a short line for "검토 항목", "검토 목적", and "보이는 결과" when this is available from reportReviewMatrix or measurementExtraction.
- Make the first viewport useful: current problem, overall conclusion, and measurement summary visible near the top.
- Keep UI dense and readable like an engineering dashboard. Avoid marketing-style hero sections.
- Ensure long dataset names wrap and text does not overflow.
- Alignment and table layout rule:
  - Align all metadata rows cleanly. Do not put "검토 항목", "검토 목적", and "보이는 결과" as a loose inline sentence.
  - Use a fixed two-column or six-column grid for metadata: label cells have a fixed width and value cells fill the remaining width.
  - Table headers must be centered, text columns left-aligned, and numeric/percent columns right-aligned with tabular numbers.
  - Keep row heights consistent, vertical-align middle for compact rate tables, and avoid mixed inline labels that shift column alignment.
  - Use consistent column names: 보고서, 조건, n, 값, Normal 값, 차이, 비율, 해석.
- Add a small note when a result is limited because Decision was not visible in the extracted report.

Required sections:
1. Current problem and overall conclusion.
2. 보고서별 검토 목적/결과 판정.
3. 공정 불량 검토: actual numeric aggregation, charts, and tables.
4. 기능 불량 검토: actual numeric aggregation, charts, and tables.
5. 공정-기능 연계 해석: only when both domains appear in the same related report or review plan.
6. AI grouped historical review results.
7. Recommended review sequence.
8. Improvement ideas and risks.
9. Reports needing deep dive.

INPUT DATA:
{json.dumps(payload, ensure_ascii=False, indent=2)}
""".strip()


def esc(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def pills(values: list[str]) -> str:
    if not values:
        return '<span class="muted">-</span>'
    return "".join(f'<span class="pill">{esc(x)}</span>' for x in values)


def as_float(value: Any) -> float | None:
    try:
        if value is None or value == "":
            return None
        return float(str(value).replace(",", "").replace("%", "").strip())
    except Exception:
        return None


def svg_text(value: Any, x: float, y: float, size: int = 11, anchor: str = "middle") -> str:
    return f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size}" text-anchor="{anchor}" fill="#344054">{esc(value)}</text>'


def render_vertical_bar(viz: dict[str, Any], records: list[dict[str, Any]]) -> str:
    data = [(r, as_float(r.get("value"))) for r in records]
    data = [(r, v) for r, v in data if v is not None][:28]
    if not data:
        return ""
    width, height = 920, 300
    left, right, top, bottom = 54, 16, 24, 62
    chart_w = width - left - right
    chart_h = height - top - bottom
    max_v = max([v for _, v in data] + [as_float(viz.get("normalValue")) or 0, as_float(viz.get("specMax")) or 0, 1])
    min_v = min([0] + [v for _, v in data])
    span = max(max_v - min_v, 1)
    bar_w = max(8, chart_w / max(len(data), 1) * 0.68)
    gap = chart_w / max(len(data), 1)
    parts = [
        f'<svg class="chart-svg" viewBox="0 0 {width} {height}" role="img">',
        f'<line x1="{left}" y1="{top + chart_h}" x2="{width - right}" y2="{top + chart_h}" stroke="#98a2b3"/>',
        f'<line x1="{left}" y1="{top}" x2="{left}" y2="{top + chart_h}" stroke="#98a2b3"/>',
    ]
    normal = as_float(viz.get("normalValue"))
    if normal is not None:
        y = top + chart_h - ((normal - min_v) / span * chart_h)
        parts.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width - right}" y2="{y:.1f}" stroke="#d92d20" stroke-dasharray="5 4"/>')
        parts.append(svg_text(f'Normal {normal:g}', width - right - 2, y - 4, 11, "end"))
    for idx, (record, value) in enumerate(data):
        x = left + idx * gap + (gap - bar_w) / 2
        h = max(1, (value - min_v) / span * chart_h)
        y = top + chart_h - h
        series = str(record.get("series") or "")
        color = "#2c7da0" if "normal" not in series.casefold() else "#7f8794"
        label = record.get("category") or record.get("condition") or record.get("mold") or record.get("series") or idx + 1
        parts.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{h:.1f}" fill="{color}"><title>{esc(label)}: {value:g}</title></rect>')
        parts.append(svg_text(f'{value:g}', x + bar_w / 2, y - 4, 10))
        parts.append(svg_text(str(label)[:12], x + bar_w / 2, top + chart_h + 16, 9))
    parts.append(svg_text(str(viz.get("unit") or ""), 12, top + 6, 11, "start"))
    parts.append("</svg>")
    return "\n".join(parts)


def render_scatter(viz: dict[str, Any], records: list[dict[str, Any]]) -> str:
    data = [(idx + 1, r, as_float(r.get("value"))) for idx, r in enumerate(records) if as_float(r.get("value")) is not None][:60]
    if not data:
        return ""
    width, height = 920, 300
    left, right, top, bottom = 54, 18, 24, 56
    chart_w = width - left - right
    chart_h = height - top - bottom
    values = [v for _, _, v in data if v is not None]
    spec_min = as_float(viz.get("specMin"))
    spec_max = as_float(viz.get("specMax"))
    max_v = max(values + [spec_max or max(values), spec_min or min(values)])
    min_v = min(values + [spec_min or min(values), spec_max or max(values)])
    if max_v == min_v:
        max_v += 1
        min_v -= 1
    span = max_v - min_v
    parts = [
        f'<svg class="chart-svg" viewBox="0 0 {width} {height}" role="img">',
        f'<line x1="{left}" y1="{top + chart_h}" x2="{width - right}" y2="{top + chart_h}" stroke="#98a2b3"/>',
        f'<line x1="{left}" y1="{top}" x2="{left}" y2="{top + chart_h}" stroke="#98a2b3"/>',
    ]
    for spec, label in ((spec_min, "Spec min"), (spec_max, "Spec max")):
        if spec is None:
            continue
        y = top + chart_h - ((spec - min_v) / span * chart_h)
        parts.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width - right}" y2="{y:.1f}" stroke="#d92d20" stroke-dasharray="5 4"/>')
        parts.append(svg_text(f'{label} {spec:g}', width - right - 2, y - 4, 11, "end"))
    denom = max(len(data) - 1, 1)
    for idx, record, value in data:
        x = left + ((idx - 1) / denom * chart_w)
        y = top + chart_h - ((value - min_v) / span * chart_h)
        label = record.get("category") or record.get("condition") or record.get("series") or idx
        parts.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4.2" fill="#2c7da0"><title>{esc(label)}: {value:g}</title></circle>')
    parts.append(svg_text(str(viz.get("unit") or ""), 12, top + 6, 11, "start"))
    parts.append("</svg>")
    return "\n".join(parts)


def render_heatmap(viz: dict[str, Any], records: list[dict[str, Any]]) -> str:
    data = [(r, as_float(r.get("value"))) for r in records if as_float(r.get("value")) is not None][:120]
    if not data:
        return ""
    xs: list[str] = []
    ys: list[str] = []
    for record, _ in data:
        x = str(record.get("mold") or record.get("ir") or record.get("condition") or record.get("category") or "-")
        y = str(record.get("line") or record.get("series") or record.get("report") or "-")
        if x not in xs:
            xs.append(x)
        if y not in ys:
            ys.append(y)
    xs = xs[:16]
    ys = ys[:14]
    values = [v for _, v in data]
    min_v, max_v = min(values), max(values)
    span = max(max_v - min_v, 1)
    cell_w, cell_h = 52, 28
    left, top = 130, 44
    width = left + len(xs) * cell_w + 20
    height = top + len(ys) * cell_h + 30
    parts = [f'<svg class="chart-svg" viewBox="0 0 {width} {height}" role="img">']
    for i, x in enumerate(xs):
        parts.append(svg_text(x[:10], left + i * cell_w + cell_w / 2, 24, 10))
    for j, y_label in enumerate(ys):
        parts.append(svg_text(y_label[:18], left - 8, top + j * cell_h + 18, 10, "end"))
    for record, value in data:
        x_label = str(record.get("mold") or record.get("ir") or record.get("condition") or record.get("category") or "-")
        y_label = str(record.get("line") or record.get("series") or record.get("report") or "-")
        if x_label not in xs or y_label not in ys:
            continue
        i, j = xs.index(x_label), ys.index(y_label)
        intensity = (value - min_v) / span
        color = f'rgb({int(235 - 175 * intensity)}, {int(244 - 120 * intensity)}, {int(255 - 70 * intensity)})'
        x = left + i * cell_w
        y = top + j * cell_h
        parts.append(f'<rect x="{x}" y="{y}" width="{cell_w - 2}" height="{cell_h - 2}" fill="{color}" stroke="#fff"><title>{esc(x_label)} / {esc(y_label)}: {value:g}</title></rect>')
        parts.append(svg_text(f'{value:g}', x + cell_w / 2, y + 18, 10))
    parts.append("</svg>")
    return "\n".join(parts)


def render_record_table(records: list[dict[str, Any]]) -> str:
    if not records:
        return '<div class="muted">추출된 수치 없음</div>'
    rows = []
    for record in records[:80]:
        rows.append(
            "<tr>"
            f"<td>{esc(record.get('datasetName') or record.get('report'))}</td>"
            f"<td>{esc(record.get('category') or record.get('condition') or record.get('series'))}</td>"
            f"<td>{esc(record.get('value'))} {esc(record.get('unit'))}</td>"
            f"<td>{esc(record.get('sourceCell'))}</td>"
            f"<td>{esc(record.get('note') or record.get('sourceText'))}</td>"
            "</tr>"
        )
    return (
        '<div class="measure-table"><table><thead><tr>'
        '<th>Report</th><th>Condition</th><th>Value</th><th>Cell</th><th>Note</th>'
        '</tr></thead><tbody>'
        + "".join(rows)
        + "</tbody></table></div>"
    )


def render_measurements(measurement: dict[str, Any] | None) -> str:
    if not measurement:
        return '<section class="panel"><h2>실측 수치 취합</h2><p class="muted">수치 추출 분석이 아직 실행되지 않았습니다.</p></section>'
    visuals = measurement.get("visualizations") if isinstance(measurement.get("visualizations"), list) else []
    panels: list[str] = []
    for viz in visuals:
        if not isinstance(viz, dict):
            continue
        records = [r for r in (viz.get("records") or []) if isinstance(r, dict)]
        chart_type = str(viz.get("chartType") or "table")
        if chart_type == "verticalBar":
            chart = render_vertical_bar(viz, records)
        elif chart_type == "scatter":
            chart = render_scatter(viz, records)
        elif chart_type == "heatmap":
            chart = render_heatmap(viz, records)
        else:
            chart = ""
        if not chart:
            chart = render_record_table(records)
        panels.append(f"""
        <section class="group">
          <div class="group-head">
            <h2>{esc(viz.get("chartTitle") or viz.get("groupTitle"))}</h2>
            <span>{esc(chart_type)} / {esc(viz.get("metricName"))} {esc(viz.get("unit"))}</span>
          </div>
          <p><b>AI 차트 판단:</b> {esc(viz.get("whyThisChart"))}</p>
          <div class="chart-wrap">{chart}</div>
          <p><b>해석:</b> {esc(viz.get("interpretation"))}</p>
          <div><b>한계</b><ul>{''.join(f"<li>{esc(x)}</li>" for x in list(viz.get("limits") or []))}</ul></div>
          {render_record_table(records)}
        </section>
        """)
    return (
        '<section class="panel"><h2>실측 수치 취합</h2>'
        f'<p>{esc(measurement.get("measurementSummary"))}</p></section>'
        + ("".join(panels) if panels else '<section class="panel"><p class="muted">추출된 수치 시각화 없음</p></section>')
    )


def write_html(path: Path, problem: str, package: dict[str, Any], analysis: dict[str, Any], measurement: dict[str, Any] | None = None) -> None:
    groups = analysis.get("reportGroups") if isinstance(analysis.get("reportGroups"), list) else []
    group_html = []
    for group in groups:
        group_html.append(f"""
        <section class="group">
          <div class="group-head">
            <h2>{esc(group.get("groupTitle"))}</h2>
            <span>{esc(group.get("groupType"))} / {esc(group.get("reportCount"))}건</span>
          </div>
          <div class="grid">
            <div><b>검토 내용</b><p>{esc(group.get("whatWasChecked"))}</p></div>
            <div><b>검토 결과 요약</b><p>{esc(group.get("reviewResultSummary"))}</p></div>
            <div><b>현재 문제와의 연결</b><p>{esc(group.get("similarityToCurrentProblem"))}</p></div>
          </div>
          <div><b>대표 보고서</b><div class="chips">{pills(list(group.get("representativeReports") or []))}</div></div>
          <div><b>근거</b><ul>{''.join(f"<li>{esc(x)}</li>" for x in list(group.get("usefulEvidence") or []))}</ul></div>
          <div><b>한계</b><ul>{''.join(f"<li>{esc(x)}</li>" for x in list(group.get("limits") or []))}</ul></div>
        </section>
        """)

    plan = analysis.get("recommendedReviewPlan") if isinstance(analysis.get("recommendedReviewPlan"), list) else []
    ideas = analysis.get("improvementIdeas") if isinstance(analysis.get("improvementIdeas"), list) else []
    missing = analysis.get("missingInformationNeeded") if isinstance(analysis.get("missingInformationNeeded"), list) else []
    deep_dive = analysis.get("candidateReportsForDeepDive") if isinstance(analysis.get("candidateReportsForDeepDive"), list) else []

    path.write_text(f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI Current Problem Analysis</title>
<style>
* {{ box-sizing: border-box; }}
body {{ margin:0; background:#f5f6f8; color:#151922; font-family:"Segoe UI","Malgun Gothic",Arial,sans-serif; font-size:14px; }}
header {{ padding:14px 16px; background:#fff; border-bottom:1px solid #d4d8df; }}
h1 {{ margin:0 0 6px; font-size:20px; }}
main {{ padding:14px; display:grid; gap:12px; }}
.panel,.group {{ background:#fff; border:1px solid #d4d8df; box-shadow:0 1px 2px rgba(16,24,40,.08); padding:12px; }}
.group-head {{ display:flex; justify-content:space-between; gap:10px; align-items:start; border-bottom:1px solid #e1e5eb; padding-bottom:8px; margin-bottom:10px; }}
h2 {{ margin:0; font-size:17px; }}
.grid {{ display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px; margin:10px 0; }}
p {{ margin:6px 0 0; line-height:1.55; }}
ul {{ margin:6px 0 0 20px; padding:0; line-height:1.55; }}
.muted {{ color:#667085; }}
.chips {{ display:flex; flex-wrap:wrap; gap:5px; margin-top:6px; }}
.pill {{ display:inline-flex; border:1px solid #c9d3e2; border-radius:999px; padding:2px 7px; background:#f8fafc; color:#24364f; font-size:12px; }}
.cols {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }}
.item {{ border:1px solid #d4d8df; border-radius:6px; padding:9px; margin-top:8px; background:#fff; }}
.chart-wrap {{ width:100%; overflow:auto; border:1px solid #e1e5eb; border-radius:6px; margin:10px 0; background:#fff; }}
.chart-svg {{ display:block; min-width:720px; width:100%; height:auto; }}
.measure-table {{ max-height:320px; overflow:auto; border:1px solid #e1e5eb; margin-top:10px; }}
.measure-table table {{ width:100%; border-collapse:collapse; font-size:12px; }}
.measure-table th,.measure-table td {{ border-bottom:1px solid #edf0f4; padding:6px; text-align:left; vertical-align:top; }}
.measure-table th {{ background:#f9fafb; position:sticky; top:0; }}
@media(max-width:960px) {{ .grid,.cols {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<header>
  <h1>AI 현재 문제 분석</h1>
  <div class="muted">{esc(problem)} / generated {esc(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>
</header>
<main>
  <section class="panel">
    <b>전체 결론</b>
    <p>{esc(analysis.get("overallConclusion"))}</p>
    <div class="chips">
      <span class="pill">Detected model: {esc(analysis.get("detectedModel") or package.get("detectedModel"))}</span>
      <span class="pill">Retrieved: {len(package.get("candidates") or [])} reports</span>
      <span class="pill">AI groups: {len(groups)}</span>
    </div>
  </section>
  {render_measurements(measurement)}
  {''.join(group_html)}
  <section class="cols">
    <div class="panel">
      <h2>권장 검토 순서</h2>
      {''.join(f'<div class="item"><b>{esc(x.get("step"))}. {esc(x.get("item"))}</b><p>{esc(x.get("why"))}</p><div class="chips">{pills(list(x.get("relatedGroups") or []))}</div></div>' for x in plan)}
    </div>
    <div class="panel">
      <h2>개선 아이디어</h2>
      {''.join(f'<div class="item"><b>{esc(x.get("idea"))}</b><p>{esc(x.get("basis"))}</p><p class="muted">Risk: {esc(x.get("risk"))}</p></div>' for x in ideas)}
    </div>
  </section>
  <section class="cols">
    <div class="panel"><h2>추가로 필요한 정보</h2><ul>{''.join(f"<li>{esc(x)}</li>" for x in missing)}</ul></div>
    <div class="panel"><h2>우선 상세확인 보고서</h2><div class="chips">{pills(list(deep_dive))}</div></div>
  </section>
</main>
</body>
</html>
""", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sample-dir", type=Path, default=DEFAULT_SAMPLE_DIR)
    parser.add_argument("--problem")
    parser.add_argument("--model", default="")
    parser.add_argument("--top", type=int, default=20)
    parser.add_argument("--candidate-limit", type=int, default=16)
    parser.add_argument("--measure-report-limit", type=int, default=20)
    parser.add_argument("--effort", default="medium", choices=["minimal", "low", "medium", "high", "xhigh"])
    parser.add_argument("--measure-effort", default="medium", choices=["minimal", "low", "medium", "high", "xhigh"])
    parser.add_argument("--html-effort", default="medium", choices=["minimal", "low", "medium", "high", "xhigh"])
    parser.add_argument("--timeout-sec", type=int, default=900)
    parser.add_argument("--skip-measurements", action="store_true")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-html", type=Path)
    parser.add_argument("--from-json", type=Path, help="Reuse an existing analysis JSON and regenerate only the final HTML.")
    args = parser.parse_args()

    sample_dir = args.sample_dir
    out_dir = sample_dir / "ai_current_problem"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_prompt_path = out_dir / f"current_problem_html_prompt_{stamp}.txt"
    html_raw_path = out_dir / f"current_problem_html_raw_{stamp}.html"
    html_path = args.out_html or sample_dir / "current_problem_ai_analysis.html"

    if args.from_json:
        log("AI current problem HTML-only generation started")
        log(f"sample dir -> {sample_dir}")
        log(f"from JSON -> {args.from_json}")
        result = json.loads(args.from_json.read_text(encoding="utf-8-sig", errors="replace"))
        package = result.get("package") if isinstance(result.get("package"), dict) else {}
        analysis = result.get("analysis") if isinstance(result.get("analysis"), dict) else {}
        measurement = result.get("measurement") if isinstance(result.get("measurement"), dict) else None
        measured_reports = result.get("measuredReports") if isinstance(result.get("measuredReports"), list) else []
        problem = args.problem or str(package.get("problem") or "")
        if not problem:
            raise ValueError("--problem is required when --from-json does not contain package.problem")
        log(
            "loaded existing analysis: "
            f"candidates={len(package.get('candidates') or [])}, "
            f"groups={len(analysis.get('reportGroups') or [])}, "
            f"visualizations={len(measurement.get('visualizations') or []) if measurement else 0}"
        )
        log("building final HTML prompt")
        html_prompt = build_ai_html_prompt(problem, package, analysis, measurement)
        log(f"writing final HTML prompt -> {html_prompt_path}")
        html_prompt_path.write_text(html_prompt, encoding="utf-8")
        ai_html = clean_ai_html(call_codex_text(
            html_prompt,
            html_raw_path,
            args.html_effort,
            args.timeout_sec,
            "HTML-only final HTML generation",
        ))
        log(f"writing final HTML -> {html_path}")
        html_path.write_text(ai_html, encoding="utf-8")
        log("AI current problem HTML-only generation finished")
        print(json.dumps({
            "json": str(args.from_json),
            "html": str(html_path),
            "htmlPrompt": str(html_prompt_path),
            "htmlRaw": str(html_raw_path),
            "measuredReports": len(measured_reports),
        }, ensure_ascii=False))
        return 0

    if not args.problem:
        raise ValueError("--problem is required unless --from-json is provided")

    log("AI current problem analysis started")
    log(f"sample dir -> {sample_dir}")
    log(f"problem -> {args.problem}")
    if args.model:
        log(f"requested model -> {args.model}")
    log("retrieving similar historical reports")
    package = retrieve(sample_dir, args.problem, args.model, args.top)
    log(
        "retrieval finished: "
        f"detectedModel={package.get('detectedModel') or '-'}, "
        f"candidates={len(package.get('candidates') or [])}"
    )
    prompt = build_prompt(package, args.candidate_limit)

    prompt_path = out_dir / f"current_problem_prompt_{stamp}.txt"
    raw_path = out_dir / f"current_problem_raw_{stamp}.txt"
    measure_prompt_path = out_dir / f"current_problem_measure_prompt_{stamp}.txt"
    measure_raw_path = out_dir / f"current_problem_measure_raw_{stamp}.txt"
    json_path = args.out_json or out_dir / f"current_problem_analysis_{stamp}.json"

    log(f"writing analysis prompt -> {prompt_path}")
    prompt_path.write_text(prompt, encoding="utf-8")
    analysis = call_codex(prompt, raw_path, args.effort, args.timeout_sec, "1/3 historical analysis")
    log("1/3 historical analysis JSON parsed")

    measurement: dict[str, Any] | None = None
    measured_reports: list[str] = []
    if not args.skip_measurements:
        log(f"building measurement prompt from up to {args.measure_report_limit} reports")
        measure_prompt, measured_reports = build_measurement_prompt(
            args.db,
            sample_dir,
            package,
            analysis,
            report_limit=args.measure_report_limit,
        )
        log(f"measurement reports selected: {len(measured_reports)}")
        for name in measured_reports:
            log(f"  measure report: {name}")
        log(f"writing measurement prompt -> {measure_prompt_path}")
        measure_prompt_path.write_text(measure_prompt, encoding="utf-8")
        measurement = call_codex(
            measure_prompt,
            measure_raw_path,
            args.measure_effort,
            args.timeout_sec,
            "2/3 measurement extraction",
        )
        visual_count = len(measurement.get("visualizations") or []) if isinstance(measurement, dict) else 0
        log(f"2/3 measurement extraction JSON parsed: visualizations={visual_count}")
    else:
        log("measurement extraction skipped")

    result = {"package": package, "analysis": analysis, "measurement": measurement, "measuredReports": measured_reports}
    log(f"writing combined JSON -> {json_path}")
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    log("building final HTML prompt")
    html_prompt = build_ai_html_prompt(args.problem, package, analysis, measurement)
    log(f"writing final HTML prompt -> {html_prompt_path}")
    html_prompt_path.write_text(html_prompt, encoding="utf-8")
    ai_html = clean_ai_html(call_codex_text(
        html_prompt,
        html_raw_path,
        args.html_effort,
        args.timeout_sec,
        "3/3 final HTML generation",
    ))
    log(f"writing final HTML -> {html_path}")
    html_path.write_text(ai_html, encoding="utf-8")

    log("AI current problem analysis finished")
    print(json.dumps({
        "json": str(json_path),
        "html": str(html_path),
        "prompt": str(prompt_path),
        "raw": str(raw_path),
        "measurePrompt": str(measure_prompt_path) if measurement is not None else "",
        "measureRaw": str(measure_raw_path) if measurement is not None else "",
        "htmlPrompt": str(html_prompt_path),
        "htmlRaw": str(html_raw_path),
        "candidates": len(package.get("candidates") or []),
        "measuredReports": len(measured_reports),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
