from __future__ import annotations

import argparse
import html
import json
import re
from collections import Counter
from pathlib import Path
from typing import Any


DEFAULT_SAMPLE_DIR = Path(r"C:\Users\jhbyun\Desktop\새 폴더 (4)\sample_ready")


def read_text(path: Path, default: str = "") -> str:
    if not path.exists():
        return default
    return path.read_text(encoding="utf-8-sig", errors="replace")


def read_json(path: Path) -> Any:
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8-sig", errors="replace"))


def parse_term_guidance(path: Path) -> list[dict[str, str]]:
    text = read_text(path).strip()
    if not text:
        return []
    entries: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    current_key = ""
    for line in text.splitlines():
        raw = line.rstrip()
        if raw.startswith("### "):
            if current:
                entries.append(current)
            current = {"term": raw[4:].strip(), "meaning": "", "usage": ""}
            current_key = ""
            continue
        if current is None:
            continue
        if raw.startswith("- Meaning:"):
            current_key = "meaning"
            current["meaning"] = raw.split(":", 1)[1].strip()
        elif raw.startswith("- Usage:"):
            current_key = "usage"
            current["usage"] = raw.split(":", 1)[1].strip()
        elif raw.strip() and current_key:
            current[current_key] = (current.get(current_key, "") + "\n" + raw.strip()).strip()
    if current:
        entries.append(current)
    return [x for x in entries if x.get("term") or x.get("meaning") or x.get("usage")]


def parse_prompt_requests(path: Path) -> list[dict[str, str]]:
    text = read_text(path).strip()
    if not text or "No pending requests." in text:
        return []
    rows: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    current_key = ""
    for line in text.splitlines():
        raw = line.rstrip()
        if raw.startswith("### "):
            if current:
                rows.append(current)
            title = raw.split(".", 1)[1].strip() if "." in raw else raw[4:].strip()
            current = {"type": title}
            current_key = ""
            continue
        if current is None:
            continue
        if raw.startswith("- ") and ":" in raw:
            key, value = raw[2:].split(":", 1)
            key = key.strip()
            value = value.strip()
            current_key = key
            if key == "AI Value":
                current["original"] = value
            elif key == "User Answer":
                current["canonical"] = value
            elif key == "Dataset":
                current["dataset"] = value
            elif key == "Model":
                current["model"] = value
            elif key == "Issue":
                current["issueType"] = value
            elif key == "Question":
                current["question"] = value
            elif key == "Note":
                current["note"] = value
            else:
                current[key] = value
        elif raw.strip() and current_key == "User Answer":
            current["canonical"] = (current.get("canonical", "") + "\n" + raw.strip()).strip()
    if current:
        rows.append(current)
    return rows


def esc(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def list_value(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if isinstance(value, str):
        return [x.strip() for x in value.split("|") if x.strip()]
    return []


def compact_row(row: dict[str, Any], idx: int) -> dict[str, Any]:
    return {
        "id": idx,
        "datasetName": str(row.get("datasetName") or ""),
        "fileNames": str(row.get("fileNames") or ""),
        "model": str(row.get("model") or ""),
        "aiModel": str(row.get("aiModel") or ""),
        "modelMappingSource": str(row.get("modelMappingSource") or ""),
        "date": str(row.get("date") or ""),
        "dbReportDate": str(row.get("dbReportDate") or ""),
        "purposeCode": str(row.get("purposeCode") or ""),
        "reviewPurpose": str(row.get("reviewPurpose") or ""),
        "purpose": str(row.get("purpose") or ""),
        "targetDefects": list_value(row.get("targetDefects")),
        "reviewItems": list_value(row.get("reviewItems")),
        "tags": list_value(row.get("tags")),
        "confidence": float(row.get("confidence") or 0),
        "needsDetailedAnalysis": bool(row.get("needsDetailedAnalysis")),
        "evidenceSummary": str(row.get("evidenceSummary") or ""),
        "evidenceCells": list_value(row.get("evidenceCells"))[:16],
        "uncertainty": str(row.get("uncertainty") or ""),
    }


def grouped_counts(rows: list[dict[str, Any]], key: str, limit: int = 30) -> list[dict[str, Any]]:
    counts = Counter(str(row.get(key) or "").strip() or "(unknown)" for row in rows)
    return [{"name": name, "count": count} for name, count in counts.most_common(limit)]


def build_state(sample_dir: Path) -> dict[str, Any]:
    rows_raw = read_json(sample_dir / "demo_index.json") or []
    if not isinstance(rows_raw, list):
        rows_raw = []
    rows = [compact_row(row, idx + 1) for idx, row in enumerate(rows_raw)]
    terms = parse_term_guidance(sample_dir / "ai_term_guidance.md")
    prompts = parse_prompt_requests(sample_dir / "prompt_update_requests.md")
    excluded_datasets = [
        req.get("dataset", "")
        for req in prompts
        if "제외" in req.get("canonical", "") or "exclude" in req.get("canonical", "").casefold()
    ]
    return {
        "rows": rows,
        "terms": terms,
        "promptRules": prompts,
        "excludedDatasets": excluded_datasets,
        "modelCounts": grouped_counts(rows, "model", 80),
        "stats": {
            "rows": len(rows),
            "terms": len(terms),
            "promptRules": len(prompts),
            "excludedRules": len(excluded_datasets),
        },
    }


HTML_TEMPLATE = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Current Problem Search</title>
<style>
* { box-sizing: border-box; }
body {
  margin: 0;
  background: #f5f6f8;
  color: #151922;
  font-family: "Segoe UI", "Malgun Gothic", Arial, sans-serif;
  font-size: 14px;
}
header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 12px 14px;
  background: #fff;
  border-bottom: 1px solid #d4d8df;
}
h1 { margin: 0; font-size: 18px; }
button, input, select, textarea {
  font: inherit;
  border: 1px solid #b9c2cf;
  border-radius: 6px;
  padding: 8px 10px;
  background: #fff;
}
button { cursor: pointer; }
button.primary { background: #315fcb; color: #fff; border-color: #315fcb; }
button.subtle { color: #344054; background: #fff; }
main { padding: 12px; display: grid; gap: 12px; }
.toolbar { display: flex; gap: 8px; flex-wrap: wrap; }
.grid { display: grid; grid-template-columns: minmax(0, 1fr) 430px; gap: 12px; align-items: start; }
.stack { display: grid; gap: 12px; }
.side { display: grid; gap: 12px; position: sticky; top: 12px; }
.panel {
  background: #fff;
  border: 1px solid #d4d8df;
  box-shadow: 0 1px 2px rgba(16, 24, 40, .08);
}
.panel-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 10px 12px;
  border-bottom: 1px solid #d4d8df;
  background: #f9fafb;
  font-weight: 700;
}
.panel-body { padding: 12px; display: grid; gap: 10px; }
.stats {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 8px;
}
.stat {
  border: 1px solid #d4d8df;
  border-radius: 6px;
  padding: 10px;
  background: #fff;
}
.stat .label { color: #667085; font-size: 12px; }
.stat .value { font-size: 20px; font-weight: 800; margin-top: 4px; }
.input-grid { display: grid; grid-template-columns: minmax(0, 1fr) 180px 110px 120px 120px; gap: 8px; align-items: end; }
textarea { width: 100%; min-height: 140px; resize: vertical; line-height: 1.5; }
#packageText { min-height: 310px; font-family: Consolas, "D2Coding", monospace; font-size: 12px; }
.options { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; color: #344054; }
.options label { display: inline-flex; align-items: center; gap: 6px; }
.summary-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.summary-box { border: 1px solid #d4d8df; border-radius: 6px; padding: 9px; min-height: 64px; }
.summary-box b { display: block; margin-bottom: 5px; }
.chips { display: flex; flex-wrap: wrap; gap: 5px; }
.pill {
  display: inline-flex;
  align-items: center;
  max-width: 100%;
  border: 1px solid #c9d3e2;
  border-radius: 999px;
  padding: 2px 7px;
  background: #f8fafc;
  color: #24364f;
  font-size: 12px;
  line-height: 1.4;
}
.pill.warn { border-color: #f2c98f; background: #fff7ed; color: #9a5200; }
.pill.good { border-color: #b7d7bd; background: #f0f9f1; color: #276738; }
.muted { color: #667085; }
.small { font-size: 12px; line-height: 1.45; }
.table-wrap { border: 1px solid #d4d8df; max-height: 720px; overflow: auto; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; min-width: 1180px; }
th, td { border-bottom: 1px solid #e1e5eb; padding: 8px 9px; text-align: left; vertical-align: top; }
th { position: sticky; top: 0; background: #495565; color: #fff; z-index: 1; }
tbody tr:hover { background: #f3f7ff; }
.score { font-weight: 800; color: #143a76; }
.title { font-weight: 700; color: #101828; }
.purpose { margin-top: 4px; line-height: 1.45; }
.reason-list { display: grid; gap: 3px; }
.guidance-list { display: grid; gap: 8px; max-height: 330px; overflow: auto; }
.guidance-item { border: 1px solid #d4d8df; border-radius: 6px; padding: 8px; }
.guidance-item b { display: block; margin-bottom: 4px; }
.model-bars { display: grid; gap: 6px; max-height: 210px; overflow: auto; }
.bar-row { display: grid; grid-template-columns: 150px minmax(0, 1fr) 42px; gap: 8px; align-items: center; }
.bar-track { height: 8px; background: #edf0f4; border-radius: 999px; overflow: hidden; }
.bar-fill { height: 100%; background: #2c7da0; }
@media (max-width: 1180px) {
  .grid { grid-template-columns: 1fr; }
  .side { position: static; }
}
@media (max-width: 760px) {
  header, .input-grid, .stats, .summary-grid { display: grid; grid-template-columns: 1fr; align-items: stretch; }
}
</style>
</head>
<body>
<header>
  <h1>현재 문제 검색 데모</h1>
  <div class="toolbar">
    <button id="openControlBtn" class="subtle">분석 현황</button>
    <button id="openTermsBtn" class="subtle">용어 정리</button>
    <button id="openAiAnalysisBtn" class="subtle">최근 AI 분석</button>
    <button id="runAiCommandBtn" class="primary">AI 분석 CMD 실행</button>
    <button id="copyAiCommandBtn" class="subtle">명령 복사</button>
    <button id="copyBtn" class="primary">AI 패키지 복사</button>
  </div>
</header>
<main>
  <section class="stats">
    <div class="stat"><div class="label">Rows</div><div class="value" id="statRows">0</div></div>
    <div class="stat"><div class="label">User Terms</div><div class="value" id="statTerms">0</div></div>
    <div class="stat"><div class="label">Prompt Rules</div><div class="value" id="statRules">0</div></div>
    <div class="stat"><div class="label">Excluded Rules</div><div class="value" id="statExcluded">0</div></div>
  </section>
  <div class="grid">
    <div class="stack">
      <section class="panel">
        <div class="panel-head">
          <span>현재 문제</span>
          <span class="muted small" id="searchStatus">ready</span>
        </div>
        <div class="panel-body">
          <textarea id="problemInput">BRS-161016 VP+CD separate NG rate 상승. Bonding 조건 변경 후 과거 검증 데이터와 개선 방향 확인</textarea>
          <div class="input-grid">
            <select id="modelSelect"><option value="">모델 자동 판단</option></select>
            <select id="topSelect">
              <option value="10">Top 10</option>
              <option value="20" selected>Top 20</option>
              <option value="40">Top 40</option>
              <option value="80">Top 80</option>
            </select>
            <button id="searchBtn" class="primary">검색</button>
            <button id="copyAiCommandInlineBtn" class="primary">AI 분석</button>
            <button id="openAiAnalysisInlineBtn" class="subtle">최근 결과</button>
          </div>
          <div class="options">
            <label><input id="applyExclude" type="checkbox" checked> 사용자 제외 기준 적용</label>
            <label><input id="sameModelBoost" type="checkbox" checked> 동일 모델 우선</label>
            <label><input id="needsDetailBoost" type="checkbox" checked> 상세분석 필요 데이터 가산</label>
          </div>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head">
          <span>검색 요약</span>
          <span class="muted small" id="resultCount">0 results</span>
        </div>
        <div class="panel-body">
          <div class="summary-grid">
            <div class="summary-box"><b>감지 모델</b><div id="detectedModel" class="chips"></div></div>
            <div class="summary-box"><b>감지 용어</b><div id="detectedTerms" class="chips"></div></div>
            <div class="summary-box"><b>적용 기준</b><div id="appliedRules" class="chips"></div></div>
          </div>
          <div class="summary-box">
            <b>검토 방향 초안</b>
            <div id="suggestionBox" class="small"></div>
          </div>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head"><span>유사 과거 데이터</span><span class="muted small" id="scoreInfo"></span></div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th style="width:70px;">Score</th>
                <th style="width:245px;">Dataset</th>
                <th style="width:130px;">Model</th>
                <th style="width:250px;">Review</th>
                <th style="width:210px;">Defects</th>
                <th style="width:210px;">Items / Tags</th>
                <th>Reason</th>
              </tr>
            </thead>
            <tbody id="resultBody"></tbody>
          </table>
        </div>
      </section>
    </div>
    <aside class="side">
      <section class="panel">
        <div class="panel-head"><span>AI 상세분석 패키지</span><span class="muted small" id="packageStatus"></span></div>
        <div class="panel-body">
          <textarea id="packageText" readonly></textarea>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head"><span>사용자 지정 기준</span><span class="muted small" id="guidanceCount"></span></div>
        <div class="panel-body">
          <input id="guidanceSearch" placeholder="용어/프롬프트 기준 검색">
          <div class="guidance-list" id="guidanceList"></div>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head"><span>모델 분포</span></div>
        <div class="panel-body"><div class="model-bars" id="modelBars"></div></div>
      </section>
    </aside>
  </div>
</main>
<script id="state" type="application/json">__STATE__</script>
<script>
const state = JSON.parse(document.getElementById("state").textContent || "{}");
const sampleDir = "__SAMPLE_DIR__";
const analyzerScript = "__ANALYZER_SCRIPT__";
const aiLauncher = "jinosupporter-ai:run";
const rows = state.rows || [];
const terms = state.terms || [];
const promptRules = state.promptRules || [];
const excludedDatasets = new Set(state.excludedDatasets || []);
let lastResults = [];
let lastDetected = {};

const els = {
  problemInput: document.getElementById("problemInput"),
  modelSelect: document.getElementById("modelSelect"),
  topSelect: document.getElementById("topSelect"),
  searchBtn: document.getElementById("searchBtn"),
  applyExclude: document.getElementById("applyExclude"),
  sameModelBoost: document.getElementById("sameModelBoost"),
  needsDetailBoost: document.getElementById("needsDetailBoost"),
  resultBody: document.getElementById("resultBody"),
  resultCount: document.getElementById("resultCount"),
  searchStatus: document.getElementById("searchStatus"),
  scoreInfo: document.getElementById("scoreInfo"),
  detectedModel: document.getElementById("detectedModel"),
  detectedTerms: document.getElementById("detectedTerms"),
  appliedRules: document.getElementById("appliedRules"),
  suggestionBox: document.getElementById("suggestionBox"),
  packageText: document.getElementById("packageText"),
  packageStatus: document.getElementById("packageStatus"),
  guidanceSearch: document.getElementById("guidanceSearch"),
  guidanceList: document.getElementById("guidanceList"),
  guidanceCount: document.getElementById("guidanceCount"),
  modelBars: document.getElementById("modelBars"),
};

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[ch]));
}
function norm(value) {
  return String(value || "").toLowerCase();
}
function compact(value) {
  return norm(value).replace(/[^a-z0-9가-힣]+/g, "");
}
function tokens(text) {
  const stop = new Set(["report","test","result","check","date","model","data","the","and","with","from","normal","new","old"]);
  return [...new Set((String(text || "").match(/[A-Za-z0-9가-힣.+/-]+/g) || [])
    .map(x => x.toLowerCase())
    .filter(x => x.length >= 2 && !stop.has(x)))];
}
function rowText(row) {
  return [
    row.datasetName, row.fileNames, row.model, row.aiModel, row.date,
    row.reviewPurpose, row.purpose, row.evidenceSummary, row.uncertainty,
    ...(row.targetDefects || []), ...(row.reviewItems || []), ...(row.tags || [])
  ].join(" ").toLowerCase();
}
function fieldText(row, field) {
  const value = row[field];
  return Array.isArray(value) ? value.join(" ").toLowerCase() : String(value || "").toLowerCase();
}
function pill(value, cls="") {
  return `<span class="pill ${cls}">${esc(value)}</span>`;
}
function renderModelOptions() {
  const options = ['<option value="">모델 자동 판단</option>'];
  for (const item of state.modelCounts || []) {
    options.push(`<option value="${esc(item.name)}">${esc(item.name)} (${item.count})</option>`);
  }
  els.modelSelect.innerHTML = options.join("");
}
function renderStats() {
  document.getElementById("statRows").textContent = (state.stats?.rows || rows.length).toLocaleString();
  document.getElementById("statTerms").textContent = (state.stats?.terms || terms.length).toLocaleString();
  document.getElementById("statRules").textContent = (state.stats?.promptRules || promptRules.length).toLocaleString();
  document.getElementById("statExcluded").textContent = (state.stats?.excludedRules || excludedDatasets.size).toLocaleString();
}
function renderModelBars() {
  const max = Math.max(...(state.modelCounts || []).map(x => x.count), 1);
  els.modelBars.innerHTML = (state.modelCounts || []).slice(0, 18).map(item => `
    <div class="bar-row">
      <div class="small">${esc(item.name)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(3, item.count / max * 100)}%"></div></div>
      <div class="small muted">${item.count}</div>
    </div>`).join("");
}
function detectModel(input) {
  const selected = els.modelSelect.value;
  if (selected) return selected;
  const text = compact(input);
  const models = (state.modelCounts || []).map(x => x.name).filter(Boolean).sort((a, b) => b.length - a.length);
  for (const model of models) {
    const key = compact(model);
    if (key && text.includes(key)) return model;
    const short = key.replace(/s08zz|s15|l20|brs|msu|tiu|tim/g, "");
    if (short.length >= 5 && text.includes(short)) return model;
  }
  return "";
}
function matchedTerms(input) {
  const text = norm(input);
  const ctext = compact(input);
  return terms.filter(term => {
    const name = String(term.term || "").trim();
    if (!name) return false;
    const key = norm(name);
    return text.includes(key) || ctext.includes(compact(name));
  }).slice(0, 20);
}
function applicableRules(input, model) {
  const text = norm(input);
  return promptRules.filter(rule => {
    const dataset = norm(rule.dataset);
    const ruleModel = norm(rule.model);
    const answer = norm(rule.canonical);
    if (model && ruleModel && norm(model).includes(ruleModel)) return true;
    if (dataset && text.includes(dataset)) return true;
    if (answer.includes("standard") && text.includes("standard")) return true;
    return false;
  }).slice(0, 12);
}
function scoreRow(row, queryTokens, model, termHits) {
  const reasons = [];
  let score = 0;
  const hay = rowText(row);
  if (model) {
    const rowModel = String(row.model || "");
    if (rowModel === model) {
      score += els.sameModelBoost.checked ? 80 : 35;
      reasons.push("model exact");
    } else if (rowModel.includes(model) || model.includes(rowModel)) {
      score += els.sameModelBoost.checked ? 42 : 20;
      reasons.push("model related");
    }
  }
  for (const term of queryTokens) {
    if (fieldText(row, "targetDefects").includes(term)) { score += 9; reasons.push(`defect:${term}`); continue; }
    if (fieldText(row, "reviewItems").includes(term)) { score += 7; reasons.push(`item:${term}`); continue; }
    if (fieldText(row, "tags").includes(term)) { score += 6; reasons.push(`tag:${term}`); continue; }
    if (norm(row.reviewPurpose).includes(term) || norm(row.purpose).includes(term)) { score += 4; reasons.push(`review:${term}`); continue; }
    const count = (hay.match(new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g")) || []).length;
    if (count) { score += Math.min(6, count * 1.5); reasons.push(`text:${term}`); }
  }
  for (const term of termHits) {
    const name = norm(term.term);
    if (name && hay.includes(name)) {
      score += 12;
      reasons.push(`user term:${term.term}`);
    }
  }
  score += Number(row.confidence || 0) * 4;
  if (els.needsDetailBoost.checked && row.needsDetailedAnalysis) score += 2;
  if (els.applyExclude.checked && excludedDatasets.has(row.datasetName)) {
    score -= 100;
    reasons.push("user excluded");
  }
  return { score, reasons: [...new Set(reasons)].slice(0, 8) };
}
function search() {
  const input = els.problemInput.value.trim();
  const top = Number(els.topSelect.value || 20);
  const model = detectModel(input);
  const queryTokens = tokens(input);
  const termHits = matchedTerms(input);
  const ruleHits = applicableRules(input, model);
  const scored = [];
  for (const row of rows) {
    const result = scoreRow(row, queryTokens, model, termHits);
    if (result.score > 2) scored.push({ row, score: result.score, reasons: result.reasons });
  }
  scored.sort((a, b) => b.score - a.score);
  lastResults = scored.slice(0, top);
  lastDetected = { input, model, queryTokens, termHits, ruleHits };
  renderSearch(input, model, queryTokens, termHits, ruleHits);
}
function renderSearch(input, model, queryTokens, termHits, ruleHits) {
  els.searchStatus.textContent = `updated ${new Date().toLocaleTimeString()}`;
  els.resultCount.textContent = `${lastResults.length} results`;
  els.scoreInfo.textContent = lastResults.length ? `top score ${lastResults[0].score.toFixed(1)}` : "";
  els.detectedModel.innerHTML = model ? pill(model, "good") : '<span class="muted small">-</span>';
  els.detectedTerms.innerHTML = termHits.length ? termHits.map(x => pill(x.term, "good")).join("") : '<span class="muted small">-</span>';
  els.appliedRules.innerHTML = [
    ruleHits.length ? `${ruleHits.length} prompt rules` : "",
    els.applyExclude.checked ? `${excludedDatasets.size} exclude rules` : "",
    queryTokens.length ? `${queryTokens.length} query tokens` : "",
  ].filter(Boolean).map(x => pill(x)).join("") || '<span class="muted small">-</span>';
  els.resultBody.innerHTML = lastResults.map(item => rowHtml(item)).join("") || '<tr><td colspan="7" class="muted">검색 결과 없음</td></tr>';
  renderSuggestion();
  renderPackage();
  renderGuidance();
}
function rowHtml(item) {
  const row = item.row;
  const excluded = excludedDatasets.has(row.datasetName);
  return `<tr>
    <td><div class="score">${item.score.toFixed(1)}</div><div class="muted small">${row.confidence.toFixed(2)}</div></td>
    <td><div class="title">${esc(row.datasetName)}</div><div class="muted small">${esc(row.fileNames)}</div><div class="small">${excluded ? pill("사용자 제외", "warn") : ""}</div></td>
    <td><b>${esc(row.model)}</b><div class="muted small">AI: ${esc(row.aiModel)}</div><div class="muted small">${esc(row.date || row.dbReportDate)}</div></td>
    <td><div>${esc(row.reviewPurpose)}</div><div class="purpose muted">${esc(row.purpose)}</div></td>
    <td><div class="chips">${(row.targetDefects || []).map(x => pill(x)).join("") || '<span class="muted">-</span>'}</div></td>
    <td><div class="chips">${[...(row.reviewItems || []), ...(row.tags || [])].slice(0, 12).map(x => pill(x)).join("") || '<span class="muted">-</span>'}</div></td>
    <td><div class="reason-list small">${item.reasons.map(esc).join("<br>")}</div><div class="muted small">${esc(row.uncertainty)}</div></td>
  </tr>`;
}
function topCounts(values) {
  const map = new Map();
  for (const value of values) map.set(value, (map.get(value) || 0) + 1);
  return [...map.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "ko")).slice(0, 8);
}
function renderSuggestion() {
  const rows = lastResults.map(x => x.row);
  const defects = topCounts(rows.flatMap(x => x.targetDefects || []));
  const items = topCounts(rows.flatMap(x => x.reviewItems || []));
  const tags = topCounts(rows.flatMap(x => x.tags || []));
  const detailCount = rows.filter(x => x.needsDetailedAnalysis).length;
  const uncertaintyCount = rows.filter(x => String(x.uncertainty || "").trim()).length;
  els.suggestionBox.innerHTML = `
    <div><b>우선 검토 항목</b> ${items.map(([x,c]) => pill(`${x} ${c}`)).join("") || '<span class="muted">-</span>'}</div>
    <div style="margin-top:7px;"><b>연관 결함</b> ${defects.map(([x,c]) => pill(`${x} ${c}`)).join("") || '<span class="muted">-</span>'}</div>
    <div style="margin-top:7px;"><b>참조 태그</b> ${tags.map(([x,c]) => pill(`${x} ${c}`)).join("") || '<span class="muted">-</span>'}</div>
    <div style="margin-top:7px;" class="muted">상세분석 필요 ${detailCount}건 / 불확실성 포함 ${uncertaintyCount}건</div>`;
}
function renderPackage() {
  const payload = {
    currentProblem: lastDetected.input || "",
    detected: {
      model: lastDetected.model || "",
      queryTokens: lastDetected.queryTokens || [],
      userTerms: (lastDetected.termHits || []).map(x => ({ term: x.term, meaning: x.meaning, usage: x.usage })),
      promptRules: (lastDetected.ruleHits || []).map(x => ({
        type: x.type, dataset: x.dataset, model: x.model, userAnswer: x.canonical
      })),
    },
    retrievalPolicy: {
      source: "demo_index.json",
      userTermGuidance: "ai_term_guidance.md",
      promptRules: "prompt_update_requests.md",
      excludeUserMarkedDatasets: els.applyExclude.checked,
    },
    candidates: lastResults.slice(0, 12).map(item => ({
      score: Number(item.score.toFixed(2)),
      reasons: item.reasons,
      datasetName: item.row.datasetName,
      model: item.row.model,
      date: item.row.date,
      reviewPurpose: item.row.reviewPurpose,
      purpose: item.row.purpose,
      targetDefects: item.row.targetDefects,
      reviewItems: item.row.reviewItems,
      tags: item.row.tags,
      evidenceSummary: item.row.evidenceSummary,
      uncertainty: item.row.uncertainty,
    })),
    expectedAiOutput: [
      "현재 문제와 가장 가까운 과거 검증 사례 요약",
      "추가 확인해야 할 검토 항목",
      "개선 방향 후보",
      "모델/날짜/제외 기준 관련 주의사항",
    ],
  };
  els.packageText.value = JSON.stringify(payload, null, 2);
  els.packageStatus.textContent = `${payload.candidates.length} candidates`;
}
function renderGuidance() {
  const query = norm(els.guidanceSearch.value);
  const termItems = terms
    .filter(x => !query || [x.term, x.meaning, x.usage].join(" ").toLowerCase().includes(query))
    .slice(0, 30)
    .map(x => `<div class="guidance-item"><b>${esc(x.term)}</b><div>${esc(x.meaning).replace(/\n/g, "<br>")}</div><div class="muted small">${esc(x.usage).replace(/\n/g, "<br>")}</div></div>`);
  const ruleItems = promptRules
    .filter(x => !query || [x.type, x.dataset, x.model, x.canonical].join(" ").toLowerCase().includes(query))
    .slice(0, 20)
    .map(x => `<div class="guidance-item"><b>${esc(x.type)} / ${esc(x.model || "-")}</b><div class="muted small">${esc(x.dataset)}</div><div>${esc(x.canonical).replace(/\n/g, "<br>")}</div></div>`);
  els.guidanceCount.textContent = `${termItems.length} terms / ${ruleItems.length} rules`;
  els.guidanceList.innerHTML = [...termItems, ...ruleItems].join("") || '<div class="muted">검색 결과 없음</div>';
}
async function copyPackage() {
  try {
    await navigator.clipboard.writeText(els.packageText.value);
    els.packageStatus.textContent = "copied";
  } catch {
    els.packageText.select();
    document.execCommand("copy");
    els.packageStatus.textContent = "copied";
  }
}
function shellQuote(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
}
function buildAiCommand() {
  const problem = els.problemInput.value.trim();
  const model = detectModel(problem);
  const top = Number(els.topSelect.value || 20);
  return [
    "python",
    "-u",
    shellQuote(analyzerScript),
    "--sample-dir", shellQuote(sampleDir),
    "--problem", shellQuote(problem),
    model ? `--model ${shellQuote(model)}` : "",
    "--top", String(top),
    "--candidate-limit", "16",
    "--effort", "medium",
    "--measure-report-limit", "20",
    "--measure-effort", "high",
    "--html-effort", "medium",
    "--timeout-sec", "1200",
  ].filter(Boolean).join(" ");
}
async function copyAiCommand() {
  const cmd = buildAiCommand();
  try {
    await navigator.clipboard.writeText(cmd);
    els.packageStatus.textContent = "AI command copied";
  } catch {
    els.packageText.value = cmd;
    els.packageText.select();
    document.execCommand("copy");
    els.packageStatus.textContent = "AI command copied";
  }
}
async function runAiCommand() {
  await copyAiCommand();
  els.packageStatus.textContent = "AI command copied. Opening CMD...";
  window.location.href = aiLauncher;
}
document.getElementById("openControlBtn").addEventListener("click", () => window.open("ai_batch_control.html", "_blank", "noopener"));
document.getElementById("openTermsBtn").addEventListener("click", () => window.open("ai_term_glossary.html", "_blank", "noopener"));
document.getElementById("openAiAnalysisBtn").addEventListener("click", () => window.open("current_problem_ai_analysis.html", "_blank", "noopener"));
document.getElementById("runAiCommandBtn").addEventListener("click", runAiCommand);
document.getElementById("copyAiCommandBtn").addEventListener("click", copyAiCommand);
document.getElementById("openAiAnalysisInlineBtn").addEventListener("click", () => window.open("current_problem_ai_analysis.html", "_blank", "noopener"));
document.getElementById("copyAiCommandInlineBtn").addEventListener("click", runAiCommand);
document.getElementById("copyBtn").addEventListener("click", copyPackage);
els.searchBtn.addEventListener("click", search);
els.problemInput.addEventListener("input", () => { clearTimeout(window.__searchTimer); window.__searchTimer = setTimeout(search, 250); });
for (const el of [els.modelSelect, els.topSelect, els.applyExclude, els.sameModelBoost, els.needsDetailBoost]) el.addEventListener("change", search);
els.guidanceSearch.addEventListener("input", renderGuidance);
renderModelOptions();
renderStats();
renderModelBars();
renderGuidance();
search();
</script>
</body>
</html>
"""


def write_launcher(sample_dir: Path) -> Path:
    launcher_path = sample_dir / "current_problem_ai_analysis_launcher.cmd"
    report_path = sample_dir / "current_problem_ai_analysis.html"
    work_dir = Path(__file__).resolve().parent
    launcher_path.write_text(
        f"""@echo off
chcp 65001 >nul
title JinoSupporter AI Analysis - Codex CLI
cd /d "{work_dir}"
echo [JinoSupporter] Running AI analysis command copied by current_problem_search.html
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference = 'Stop'; $cmd = (Get-Clipboard -Raw).Trim(); if (-not $cmd) {{ throw 'Clipboard is empty. Press AI 분석 in current_problem_search.html first.' }}; if ($cmd -notmatch 'ai_current_problem_analyze\\.py') {{ throw 'Clipboard does not contain ai_current_problem_analyze.py command. Press AI 분석 in current_problem_search.html first.' }}; Write-Host $cmd; Invoke-Expression $cmd"
if errorlevel 1 (
  echo.
  echo [JinoSupporter] AI analysis failed.
) else (
  echo.
  echo [JinoSupporter] AI analysis finished. Opening latest report...
  start "" "{report_path}"
)
echo.
pause
""",
        encoding="utf-8",
    )
    return launcher_path


def write_protocol_register_script(sample_dir: Path, launcher_path: Path) -> Path:
    script_path = sample_dir / "register_current_problem_ai_protocol.cmd"
    launcher_value = str(launcher_path).replace("'", "''")
    script_path.write_text(
        f"""@echo off
chcp 65001 >nul
echo Registering jinosupporter-ai URL protocol for current Windows user...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference = 'Stop'; $root = 'HKCU:\\Software\\Classes\\jinosupporter-ai'; $launcher = '{launcher_value}'; New-Item -Path $root -Force | Out-Null; Set-Item -Path $root -Value 'URL:JinoSupporter AI Analysis'; New-ItemProperty -Path $root -Name 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null; New-Item -Path ($root + '\\DefaultIcon') -Force | Out-Null; Set-Item -Path ($root + '\\DefaultIcon') -Value '%%SystemRoot%%\\System32\\cmd.exe,0'; New-Item -Path ($root + '\\shell\\open\\command') -Force | Out-Null; Set-Item -Path ($root + '\\shell\\open\\command') -Value ('\"' + $launcher + '\" \"%%1\"')"
echo.
echo Done. Chrome may show an external protocol confirmation the first time.
pause
""",
        encoding="utf-8",
    )
    return script_path


def write_html(sample_dir: Path) -> Path:
    state = build_state(sample_dir)
    state_json = json.dumps(state, ensure_ascii=False).replace("</", "<\\/")
    html_text = (
        HTML_TEMPLATE
        .replace("__STATE__", state_json)
        .replace("__SAMPLE_DIR__", str(sample_dir).replace("\\", "\\\\"))
        .replace("__ANALYZER_SCRIPT__", str((Path(__file__).resolve().parent / "ai_current_problem_analyze.py")).replace("\\", "\\\\"))
    )
    out_path = sample_dir / "current_problem_search.html"
    out_path.write_text(html_text, encoding="utf-8")
    launcher_path = write_launcher(sample_dir)
    write_protocol_register_script(sample_dir, launcher_path)
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sample-dir", type=Path, default=DEFAULT_SAMPLE_DIR)
    args = parser.parse_args()
    out_path = write_html(args.sample_dir)
    print(json.dumps({
        "html": str(out_path),
        "launcher": str(args.sample_dir / "current_problem_ai_analysis_launcher.cmd"),
        "protocolRegister": str(args.sample_dir / "register_current_problem_ai_protocol.cmd"),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
