from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


DEFAULT_BATCH_DIR = Path(r"C:\Users\jhbyun\Desktop\새 폴더 (4)")
DEFAULT_SAMPLE_DIR = DEFAULT_BATCH_DIR / "sample_ready"


def read_json(path: Path) -> Any:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8-sig", errors="replace"))
    except Exception:
        return None


def read_jsonl_tail(path: Path, limit: int = 20) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except Exception:
                continue
            if isinstance(obj, dict):
                rows.append(obj)
    return rows[-limit:]


def read_text(path: Path, default: str = "") -> str:
    if not path.exists():
        return default
    return path.read_text(encoding="utf-8-sig", errors="replace")


def parse_prompt_requests(path: Path) -> list[dict[str, str]]:
    text = read_text(path).strip()
    if not text or "No pending requests." in text:
        return []
    rows: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    current_key = ""
    for line in text.splitlines():
        if line.startswith("### "):
            if current:
                rows.append(current)
            title = line.split(".", 1)[1].strip() if "." in line else line[4:].strip()
            current = {"type": title}
            current_key = ""
            continue
        if current is None:
            continue
        if line.startswith("- ") and ":" in line:
            key, value = line[2:].split(":", 1)
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
        elif current_key == "User Answer" and line.strip():
            current["canonical"] = (current.get("canonical", "") + "\n" + line.strip()).strip()
    if current:
        rows.append(current)
    return rows


def ensure_prompt_request_file(path: Path) -> None:
    if path.exists():
        return
    path.write_text(
        "# Prompt Update Requests\n\n"
        "HTML control page writes pending normalization or prompt update requests here.\n",
        encoding="utf-8",
    )


def ensure_term_guidance_file(path: Path) -> None:
    if path.exists():
        return
    path.write_text(
        "# AI Term Guidance\n\n"
        "Add local manufacturing terms, abbreviations, and shop-floor expressions here.\n",
        encoding="utf-8",
    )


def parse_term_guidance(path: Path) -> list[dict[str, str]]:
    text = read_text(path).strip()
    if not text:
        return []
    entries: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    for line in text.splitlines():
        if line.startswith("### "):
            if current:
                entries.append(current)
            current = {"term": line[4:].strip(), "meaning": "", "usage": ""}
            continue
        if current is None:
            continue
        if line.startswith("- Meaning:"):
            current["meaning"] = line.split(":", 1)[1].strip()
        elif line.startswith("- Usage:"):
            current["usage"] = line.split(":", 1)[1].strip()
        elif line.strip() and current.get("meaning"):
            current["meaning"] = (current["meaning"] + "\n" + line.strip()).strip()
    if current:
        entries.append(current)
    return [x for x in entries if x.get("term") or x.get("meaning")]


TERM_TOKEN_RE = re.compile(r"\b[A-Z][A-Z0-9+/#.-]{1,14}\b")
ACRONYM_HINT_RE = re.compile(r"\b[A-Z]{2,}\b|[+#/]")
WORD_HINT_RE = re.compile(
    r"\b(bako|nti|pqc|awf|gmi|mtr|xrf|iqc|sub\s*\d)\b",
    re.IGNORECASE,
)
TERM_SKIP = {
    "AFTER", "AI", "BRS", "MSU", "TIU", "REPORT", "RESULT", "TEST", "CHECK", "MODEL", "DATE",
    "NG", "OK", "PASS", "FAIL", "FAILED", "FUNCTION", "MATERIAL", "NORMAL", "SHEET",
    "THE", "AND", "WITH", "FROM", "LOT", "NEW", "OLD", "DATA", "INPUT", "OUTPUT",
    "DB", "STD", "CHECKING", "PROCESS", "PROJECT", "SUPPLIER", "FINAL", "NOT",
    "CHANGE", "IMPROVE", "LASER", "SEMI", "UP", "GUIDE", "LOOP",
}
GENERIC_TERM_SKIP = {
    "A/B", "ANALYSE", "ARRAY", "AT", "AVG", "BEFORE/AFTER", "BOND", "BONDING",
    "AMOUNT", "ANALYSE", "ANALYZE", "BAOTOU", "BENDING", "CAN", "CHEKING", "CLEAN",
    "COIL", "CUTTING", "CUTTING NG", "DECISION", "DEFORM", "DIMENSION", "DIRECTION",
    "DOME", "DRY", "DRYING", "EXCEL", "FFT", "FORMING",
    "FO DATA", "FRAME", "FUNCTION NG", "GAP", "HEARING", "HEARING NG", "LINE",
    "GASKET", "GLUE", "INK", "JPD", "JUNJIE", "KEO", "L/R", "MACHINE", "MAIN",
    "MAX", "MIC", "MIN", "MODUL", "MOLD", "NG BENDING", "NG BONDING", "NG DAMAGE",
    "NG DYNE", "NG FUNCTION", "NG HEARING", "NG OFFSET", "NG RATE", "NG SEPARATE", "NOISE",
    "OF", "OFFSET", "OK/NG", "PAD", "PASSED", "PICK-UP", "POS", "PRESS", "PROCESS", "RALON",
    "RELIABILITY", "RPM", "RUBER", "RUIJIN", "SEPARATE", "SIGMA", "SOLDER", "SPOT",
    "STANDARD", "SUSPENSION", "SYNTHETIC", "TENSION", "TOP", "TOTAL NG", "TOUCH",
    "TRADING", "TRAY", "UNIT", "USE", "USING", "UV2", "UV3", "VINA", "WELDING", "XRAY",
}
MODEL_CODE_RE = re.compile(
    r"\b(BRS|MSM|MSU|TIU|TIM|TWS|SJ|PT|GN|G[A-Z]?)[- ]?[A-Z0-9-]*\d[A-Z0-9-]*\b"
    r"|\b[A-Z]{1,4}-?\d{3,}[A-Z0-9-]*\b",
    re.IGNORECASE,
)
AI_KNOWN_TERM_KEYS = {
    "BPT", "CD", "DCR", "DECAP", "DOE", "DRYUV", "DYNEPEN", "FPCB", "GAUSS",
    "IMPEDANCE", "IQC", "IR", "JIG", "LED", "MTR", "OQC", "PQC", "RAW", "SAMPLE",
    "SPL", "SPK", "SPOTWELDING", "SPEC", "SUS", "THD", "UV", "VISIONAOI", "VISUAL",
    "XRF", "YOKE",
}


def has_candidate_hint(text: str) -> bool:
    return bool(ACRONYM_HINT_RE.search(text) or WORD_HINT_RE.search(text))


def split_list_value(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if isinstance(value, str) and value.strip():
        return [x.strip() for x in value.split("|") if x.strip()]
    return []


def clean_candidate_term(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip(" -_/.,;:()[]"))[:80]


def compact_candidate_key(text: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(text or "").upper())


def canonical_candidate(term: str) -> tuple[str, str] | None:
    term = clean_candidate_term(term)
    if not term or len(term) < 2:
        return None
    upper = term.upper()
    compact = compact_candidate_key(term)
    if MODEL_CODE_RE.search(term):
        return None
    if compact in TERM_SKIP or upper in TERM_SKIP or upper in GENERIC_TERM_SKIP or compact.isdigit():
        return None
    if compact in {compact_candidate_key(x) for x in GENERIC_TERM_SKIP}:
        return None

    if re.search(r"\bVP\s*[/+.-]?\s*CD\b", upper):
        if "SEPARATE" in upper or "SEPERATE" in upper:
            return "VP/CD separate", "VPCDSEPARATE"
        return "VP/CD", "VPCD"
    if compact.startswith("AWF"):
        return "AWF", "AWF"
    if ("SPL" in upper and "THD" in upper) or compact == "SPLTHD":
        return "SPL+THD", "SPLTHD"
    if "AOI" in upper:
        return "Vision/AOI", "VISIONAOI"
    if compact.startswith("FPCB") or compact == "FPCBA":
        return "F-PCB", "FPCB"
    if compact.startswith("BPT") or re.search(r"\bB\s*[-/]?\s*PT\b", upper):
        return "B-PT", "BPT"
    if "BAKO" in compact:
        return "Bako", "BAKO"
    if compact in {"DRYUV", "UVLED", "UVDRY"} or ("DRY" in upper and "UV" in upper):
        return "Dry UV", "DRYUV"
    if compact in {"WEAKSOLDER", "SOLDERWEAK"}:
        return "Weak solder", "WEAKSOLDER"
    if compact == "DYNEPEN" or ("DYNE" in upper and "PEN" in upper):
        return "Dyne pen", "DYNEPEN"
    if compact == "SPOTWELDING" or ("SPOT" in upper and "WELDING" in upper):
        return "Spot welding", "SPOTWELDING"
    if compact.startswith("AUDIOBUS"):
        return "AUDIOBUS", "AUDIOBUS"
    if re.search(r"\bSUB\s*\d\b", upper) or compact == "SUB":
        return "SUB process", "SUBPROCESS"

    for acronym in (
        "VP", "SPL", "THD", "JIG", "AWF", "NTI", "PQC", "GMI", "MTR", "XRF", "IQC",
        "LAI", "BKO", "CD", "SPK", "SMG", "CMG", "FRF", "RB", "IMP", "DT", "MG",
        "DOE", "FO", "MC", "PT", "UC", "TIP", "YK", "HOHD", "KR", "ASS", "RAW",
        "SUS", "CSY", "OQC", "DCR", "FS", "CPT", "RA", "TF", "VINA", "PW", "YS",
        "ME", "BP", "DECAP", "IR", "LED",
    ):
        if re.search(rf"\b{re.escape(acronym)}\b", upper):
            return acronym, acronym

    if (
        len(compact) <= 14
        and re.fullmatch(r"[A-Z]{2,}[A-Z0-9]*", compact)
        and re.fullmatch(r"[A-Za-z0-9.-]+", term)
    ):
        return term.upper(), compact
    if WORD_HINT_RE.search(term):
        return term, compact or upper
    return None


def term_ai_judgement(key: str, item: dict[str, Any]) -> tuple[bool, str]:
    if key in AI_KNOWN_TERM_KEYS:
        return False, "AI 판단 가능"
    if item["count"] < 3 and item["score"] < 40:
        return False, "빈도 낮음"
    if len(key) <= 2 and item["count"] < 20:
        return False, "AI 판단 가능"
    if len(key) == 3 and item["count"] < 8 and not item.get("aliases"):
        return False, "빈도 낮음"
    if len(key) <= 2:
        return True, "짧은 현장 약어"
    if item.get("aliases"):
        return True, "표기 변형 확인 필요"
    return True, "로컬 의미 확인 필요"


def build_term_candidates(sample_dir: Path, limit: int = 180) -> list[dict[str, Any]]:
    rows = read_json(sample_dir / "demo_index.json") or []
    if not isinstance(rows, list):
        rows = []
    existing_terms = set()
    for entry in parse_term_guidance(sample_dir / "ai_term_guidance.md"):
        existing = canonical_candidate(str(entry.get("term") or ""))
        if existing:
            existing_terms.add(existing[1])
        elif str(entry.get("term") or "").strip():
            existing_terms.add(str(entry.get("term") or "").strip().lower())
    bucket: dict[str, dict[str, Any]] = {}

    def add(term: str, row: dict[str, Any], reason: str, score: int) -> None:
        original = clean_candidate_term(term)
        canonical = canonical_candidate(original)
        if not canonical:
            return
        display, key = canonical
        if key in existing_terms or display.lower() in existing_terms:
            return
        item = bucket.setdefault(key, {
            "term": display,
            "key": key,
            "score": 0,
            "count": 0,
            "reasons": set(),
            "aliases": set(),
            "examples": [],
        })
        item["score"] += score
        item["count"] += 1
        item["reasons"].add(reason)
        if original and original.lower() != display.lower():
            item["aliases"].add(original)
        dataset = str(row.get("datasetName") or "")
        if dataset and dataset not in item["examples"] and len(item["examples"]) < 3:
            item["examples"].append(dataset)

    for row in rows:
        confidence = row.get("confidence") if isinstance(row.get("confidence"), (int, float)) else 1
        uncertain = bool(str(row.get("uncertainty") or "").strip())
        needs = bool(row.get("needsDetailedAnalysis"))
        base = 2 + (3 if uncertain else 0) + (3 if confidence < 0.8 else 0) + (1 if needs else 0)

        for field in ("targetDefects", "reviewItems", "tags"):
            for value in split_list_value(row.get(field)):
                if len(value) <= 70 and has_candidate_hint(value):
                    add(value, row, f"{field} field", base + 3)
                for token in TERM_TOKEN_RE.findall(value):
                    add(token, row, f"{field} token", base + 2)

        if uncertain:
            uncertainty = str(row.get("uncertainty") or "")
            if "Excel serial" in uncertainty or "serial" in uncertainty:
                add("Excel serial date", row, "date uncertainty", base + 5)
            if "Decision" in uncertainty:
                add("Decision", row, "decision uncertainty", base + 3)
            if "Standard" in uncertainty:
                add("Standard", row, "standard term", base + 3)

    candidates: list[dict[str, Any]] = []
    for item in bucket.values():
        needs_user_guide, ai_judgement = term_ai_judgement(str(item["key"]), item)
        candidates.append({
            "term": item["term"],
            "score": item["score"],
            "count": item["count"],
            "reasons": sorted(item["reasons"]),
            "aliases": sorted(item["aliases"])[:12],
            "examples": item["examples"],
            "needsUserGuide": needs_user_guide,
            "aiJudgement": ai_judgement,
        })
    candidates.sort(key=lambda x: (
        0 if x["needsUserGuide"] else 1,
        -int(x["score"]),
        -int(x["count"]),
        str(x["term"]),
    ))
    return candidates[:limit]


def append_plan_note(path: Path, html_name: str) -> None:
    if not path.exists():
        return
    text = read_text(path)
    marker = "## Control HTML"
    if marker in text:
        return
    note = (
        "\n\n## Control HTML\n"
        f"- `{html_name}` shows current batch progress and normalized analysis rows.\n"
        "- Use `prompt_update_requests.md` for user-selected model/term/prompt update requests.\n"
        "- Use `ai_term_guidance.md` for user-defined term explanations that should be injected into future AI analysis prompts.\n"
        "- After user confirmation, merge accepted requests into the demo prompt/update rules.\n"
    )
    path.write_text(text.rstrip() + note + "\n", encoding="utf-8")


HTML_TEMPLATE = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI Batch Control</title>
<style>
:root {
  --bg: #f5f6f8;
  --panel: #ffffff;
  --line: #d6d9de;
  --text: #161a20;
  --muted: #616b78;
  --header: #4b5563;
  --header2: #eef0f3;
  --blue: #315fcb;
  --cyan: #0e7490;
  --green: #197149;
  --orange: #a7550c;
  --red: #b42318;
  --shadow: 0 1px 2px rgba(16, 24, 40, .08);
}
* { box-sizing: border-box; }
body {
  margin: 0;
  background: var(--bg);
  color: var(--text);
  font-family: "Segoe UI", "Malgun Gothic", Arial, sans-serif;
  font-size: 14px;
  letter-spacing: 0;
}
button, input, select, textarea { font: inherit; letter-spacing: 0; }
button {
  border: 1px solid #b8bec7;
  background: #fff;
  color: #111827;
  border-radius: 6px;
  padding: 7px 10px;
  cursor: pointer;
}
button.primary { background: var(--blue); border-color: var(--blue); color: #fff; }
button.subtle { background: #f8fafc; }
button.danger { color: var(--red); border-color: #e4aaa4; }
button:disabled { opacity: .5; cursor: default; }
input, select, textarea {
  border: 1px solid #bcc3cd;
  border-radius: 6px;
  background: #fff;
  color: var(--text);
  padding: 7px 9px;
  min-width: 0;
}
textarea { resize: vertical; line-height: 1.45; }
.app {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: 12px;
  padding: 12px;
  max-width: 100vw;
  overflow-x: hidden;
}
.top {
  grid-column: 1 / -1;
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: center;
  gap: 12px;
  padding: 12px 14px;
  background: var(--panel);
  border: 1px solid var(--line);
  box-shadow: var(--shadow);
}
.title { display: flex; align-items: baseline; gap: 10px; min-width: 0; }
h1 { margin: 0; font-size: 18px; line-height: 1.2; }
.path { color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.toolbar { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
.main { min-width: 0; display: grid; gap: 12px; }
.side { min-width: 0; display: grid; gap: 12px; align-content: start; }
.panel {
  min-width: 0;
  background: var(--panel);
  border: 1px solid var(--line);
  box-shadow: var(--shadow);
  overflow: hidden;
}
.panel-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
  background: #fafbfc;
}
.panel-title { font-weight: 700; }
.panel-body { padding: 12px; }
.stats {
  display: grid;
  grid-template-columns: repeat(5, minmax(120px, 1fr));
  gap: 8px;
}
.stat {
  border: 1px solid var(--line);
  background: #fff;
  border-radius: 6px;
  padding: 10px;
}
.stat .label { color: var(--muted); font-size: 12px; }
.stat .value { font-size: 20px; font-weight: 700; margin-top: 3px; }
.progress-wrap {
  height: 16px;
  border: 1px solid #c7ccd4;
  background: #eef1f5;
  border-radius: 999px;
  overflow: hidden;
  margin-top: 10px;
}
.progress-bar { height: 100%; width: 0; background: linear-gradient(90deg, var(--blue), var(--cyan)); }
.progress-meta { display: flex; justify-content: space-between; gap: 8px; color: var(--muted); margin-top: 6px; }
.filters {
  display: grid;
  grid-template-columns: minmax(180px, 1fr) 210px 170px 120px;
  gap: 8px;
}
.split {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 290px;
  gap: 12px;
}
.bars { display: grid; gap: 7px; }
.bar-row { display: grid; grid-template-columns: minmax(0, 1fr) 40px; gap: 8px; align-items: center; }
.bar-label { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #303846; }
.bar-track { height: 8px; background: #e7e9ed; border-radius: 999px; overflow: hidden; margin-top: 3px; }
.bar-fill { height: 100%; background: #6b7280; }
.table-wrap { overflow: auto; max-height: 560px; width: 100%; max-width: 100%; }
table { width: 100%; min-width: 1120px; border-collapse: separate; border-spacing: 0; table-layout: fixed; }
th {
  position: sticky;
  top: 0;
  z-index: 1;
  background: var(--header);
  color: #fff;
  font-weight: 700;
  text-align: left;
  padding: 8px;
  border-right: 1px solid #737b87;
}
th:nth-child(1), td:nth-child(1) { width: 52px; }
th:nth-child(2), td:nth-child(2) { width: 240px; }
th:nth-child(3), td:nth-child(3) { width: 150px; }
th:nth-child(4), td:nth-child(4) { width: 280px; }
th:nth-child(5), td:nth-child(5) { width: 220px; }
th:nth-child(6), td:nth-child(6) { width: 230px; }
th:nth-child(7), td:nth-child(7) { width: 220px; }
th:nth-child(8), td:nth-child(8) { width: 72px; }
th:nth-child(9), td:nth-child(9) { width: 74px; }
td {
  padding: 8px;
  vertical-align: top;
  border-right: 1px solid #e1e4e8;
  border-bottom: 1px solid #e1e4e8;
  background: #fff;
  overflow-wrap: anywhere;
}
tr:nth-child(even) td { background: #f8f9fb; }
tr.active td { background: #eef4ff; }
.dataset { width: 240px; font-weight: 650; }
.textcell { width: 270px; }
.number { text-align: right; font-variant-numeric: tabular-nums; }
.pill-list { display: flex; flex-wrap: wrap; gap: 4px; }
.pill {
  display: inline-flex;
  align-items: center;
  max-width: 220px;
  border: 1px solid #c8ced7;
  background: #fff;
  color: #202938;
  border-radius: 999px;
  padding: 3px 7px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.pill.term { cursor: pointer; }
.pill.model { border-color: #a8b7df; background: #f4f7ff; }
.pill.warn { border-color: #e8c086; background: #fff8ed; color: #7a4100; }
.pill.bad { border-color: #efaaa4; background: #fff3f2; color: #8a1f16; }
.muted { color: var(--muted); }
.detail-grid { display: grid; gap: 9px; }
.field .name { font-size: 12px; color: var(--muted); margin-bottom: 3px; }
.field .value { white-space: pre-wrap; overflow-wrap: anywhere; line-height: 1.45; }
.issue-card {
  display: grid;
  gap: 9px;
}
.issue-title {
  font-size: 15px;
  font-weight: 750;
}
.issue-question {
  border: 1px solid #c8d4f4;
  background: #f4f7ff;
  border-radius: 6px;
  padding: 9px;
  line-height: 1.45;
}
.issue-meta {
  display: flex;
  flex-wrap: wrap;
  gap: 5px;
}
.issue-list {
  display: grid;
  gap: 6px;
  max-height: 210px;
  overflow: auto;
}
.related-list {
  display: grid;
  gap: 6px;
}
.related-item {
  border: 1px solid #bfd0f7;
  background: #f7faff;
  border-radius: 6px;
  padding: 8px;
}
.related-item button { margin-top: 6px; }
.issue-item {
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 7px;
  background: #fff;
  cursor: pointer;
}
.issue-item.active { border-color: #8aa6ea; background: #f4f7ff; }
.request-form { display: grid; gap: 8px; }
.request-form .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.hidden-field { display: none; }
.answer-context {
  border: 1px solid #d4dae3;
  background: #f8fafc;
  border-radius: 6px;
  padding: 10px;
  line-height: 1.5;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
.answer-label {
  color: #303846;
  font-weight: 700;
}
.answer-input {
  min-height: 150px;
  font-size: 15px;
  line-height: 1.5;
}
.answer-note {
  min-height: 74px;
}
.quick-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}
.quick-actions button {
  padding: 6px 8px;
  background: #f8fafc;
}
.request-list { display: grid; gap: 8px; max-height: 260px; overflow: auto; margin-top: 8px; }
.request-item {
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 8px;
  background: #fff;
}
.request-item b { display: block; margin-bottom: 3px; }
.status-line { display: flex; align-items: center; gap: 7px; color: var(--muted); }
.dot { width: 8px; height: 8px; border-radius: 50%; background: #98a2b3; }
.dot.live { background: var(--green); }
.dot.err { background: var(--red); }
.recent { display: grid; gap: 6px; max-height: 190px; overflow: auto; }
.recent-item {
  border-left: 3px solid #c1c7d0;
  padding: 4px 0 4px 8px;
  color: #374151;
}
.empty { color: var(--muted); padding: 10px 0; }
@media (max-width: 1600px) {
  .app { grid-template-columns: 1fr; }
  .side { grid-template-columns: 1fr; }
}
@media (max-width: 760px) {
  .app { padding: 8px; }
  .top { grid-template-columns: 1fr; }
  .toolbar { justify-content: flex-start; }
  .stats, .filters, .split, .side { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<div class="app">
  <header class="top">
    <div class="title">
      <h1>AI Batch Control</h1>
      <div class="path" id="pathLabel">snapshot</div>
    </div>
    <div class="toolbar">
      <button id="problemSearchBtn">현재 문제 검색</button>
      <button id="termsBtn">용어 정리</button>
      <button id="connectBtn" class="primary">폴더 연결</button>
      <button id="refreshBtn">새로고침</button>
      <button id="saveReqBtn" class="subtle hidden-field">요청 MD 저장</button>
    </div>
  </header>

  <main class="main">
    <section class="panel">
      <div class="panel-head">
        <div class="panel-title">분석 상황</div>
        <div class="status-line"><span class="dot" id="liveDot"></span><span id="liveText">snapshot</span></div>
      </div>
      <div class="panel-body">
        <div class="stats">
          <div class="stat"><div class="label">Total</div><div class="value" id="statTotal">0</div></div>
          <div class="stat"><div class="label">Done</div><div class="value" id="statDone">0</div></div>
          <div class="stat"><div class="label">Remaining</div><div class="value" id="statRemaining">0</div></div>
          <div class="stat"><div class="label">Failed</div><div class="value" id="statFailed">0</div></div>
          <div class="stat"><div class="label">Rows Loaded</div><div class="value" id="statRows">0</div></div>
        </div>
        <div class="progress-wrap"><div class="progress-bar" id="progressBar"></div></div>
        <div class="progress-meta"><span id="progressPct">0%</span><span id="updatedAt">-</span></div>
      </div>
    </section>

    <section class="panel">
      <div class="panel-head"><div class="panel-title">모델 분포</div></div>
      <div class="panel-body"><div class="bars" id="modelBars"></div></div>
    </section>

    <section class="panel">
      <div class="panel-head">
        <div class="panel-title">분석 대시보드</div>
        <div class="muted" id="filterCount">0 rows</div>
      </div>
      <div class="panel-body">
        <div class="filters">
          <input id="searchBox" placeholder="검색">
          <select id="modelFilter"><option value="">All Models</option></select>
          <select id="statusFilter">
            <option value="">All Rows</option>
            <option value="needs">Needs Detailed</option>
            <option value="low">Low Confidence</option>
            <option value="missing">Missing Model</option>
            <option value="uncertain">Uncertainty</option>
          </select>
          <select id="limitFilter">
            <option value="100">100</option>
            <option value="250">250</option>
            <option value="500">500</option>
            <option value="9999">All</option>
          </select>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>No.</th>
              <th>Dataset</th>
              <th>Model</th>
              <th>Purpose</th>
              <th>Target Defects</th>
              <th>Review Items</th>
              <th>Tags</th>
              <th>Conf.</th>
              <th></th>
            </tr>
          </thead>
          <tbody id="rowBody"></tbody>
        </table>
      </div>
    </section>
  </main>

  <aside class="side hidden-field">
    <div class="hidden-field"><span id="issueCount">0</span><div id="issuePane"></div></div>
    <div id="answerContext"></div>
    <select id="reqType"><option value="prompt">Prompt Rule</option></select>
    <input id="reqOriginal">
    <input id="reqCanonical">
    <textarea id="reqAnswer"></textarea>
    <textarea id="reqNote"></textarea>
    <button id="addReqBtn"></button>
    <button id="clearReqBtn"></button>
    <span id="reqCount">0</span>
    <span id="saveDot"></span>
    <span id="saveStatus"></span>
    <div id="requestList"></div>
  </aside>
</div>

<script id="initial-state" type="application/json">__INITIAL_STATE__</script>
<script>
const embeddedState = JSON.parse(document.getElementById("initial-state").textContent);
let summary = embeddedState.summary || {};
let progress = embeddedState.progress || [];
let rows = Array.isArray(embeddedState.rows) ? embeddedState.rows : [];
let modelPairs = Array.isArray(embeddedState.modelPairs) ? embeddedState.modelPairs : [];
let selectedRow = null;
let selectedIssue = null;
let activeAnswerIssueId = null;
let dirHandle = null;
let pollTimer = null;
let lastReadAt = null;
let requests = mergeRequests(loadRequests(), embeddedState.savedRequests || []);
const AUTO_RESOLVE_SCORE = 8;

const els = {
  pathLabel: document.getElementById("pathLabel"),
  liveDot: document.getElementById("liveDot"),
  liveText: document.getElementById("liveText"),
  statTotal: document.getElementById("statTotal"),
  statDone: document.getElementById("statDone"),
  statRemaining: document.getElementById("statRemaining"),
  statFailed: document.getElementById("statFailed"),
  statRows: document.getElementById("statRows"),
  progressBar: document.getElementById("progressBar"),
  progressPct: document.getElementById("progressPct"),
  updatedAt: document.getElementById("updatedAt"),
  searchBox: document.getElementById("searchBox"),
  modelFilter: document.getElementById("modelFilter"),
  statusFilter: document.getElementById("statusFilter"),
  limitFilter: document.getElementById("limitFilter"),
  filterCount: document.getElementById("filterCount"),
  rowBody: document.getElementById("rowBody"),
  modelBars: document.getElementById("modelBars"),
  issuePane: document.getElementById("issuePane"),
  issueCount: document.getElementById("issueCount"),
  answerContext: document.getElementById("answerContext"),
  reqType: document.getElementById("reqType"),
  reqOriginal: document.getElementById("reqOriginal"),
  reqCanonical: document.getElementById("reqCanonical"),
  reqAnswer: document.getElementById("reqAnswer"),
  reqNote: document.getElementById("reqNote"),
  reqCount: document.getElementById("reqCount"),
  saveDot: document.getElementById("saveDot"),
  saveStatus: document.getElementById("saveStatus"),
  requestList: document.getElementById("requestList"),
};

const termRules = [
  [/\bvp\s*[+/]?\s*cd\b.*\bsepar|separ.*\bvp\s*[+/]?\s*cd\b/i, "VP/CD separate"],
  [/\b(coil\s*sp|sp\s*coil)\b.*\bsepar|separ.*\b(coil\s*sp|sp\s*coil)\b/i, "Coil SP separate"],
  [/\bbond(?:ing)?\s+not\s+dry\b|\bnot\s+dry\b/i, "Bond not dry"],
  [/\bweak\s+solder\b|\bsolder\s+weak\b|\bsoldering\s+weak\b/i, "Weak solder"],
  [/\blow\s+gauss\b|\bgauss\s+low\b|\bng\s+gauss\b/i, "Low gauss"],
  [/\bng\s+function\s+high\b|\bfunction\s+ng\b|\bng\s+rate\s+function\b|\bfunction\s+check\b/i, "Function NG"],
  [/\bvp\s+bending\b|\bbending\s+vp\b/i, "VP bending"],
  [/\bcd\s+bending\b|\bbending\s+cd\b/i, "CD bending"],
  [/\bvp\s+deform\b|\bdeform\s+vp\b/i, "VP deform"],
  [/\bcoil\s+damage\b|\bdamage\s+coil\b/i, "Coil damage"],
  [/\bframe\s+damage\b|\bdamage\s+frame\b/i, "Frame damage"],
  [/\bvp\s+damage\b|\bdamage\s+vp\b/i, "VP damage"],
  [/\bdome\s+damage\b|\bdamage\s+dome\b/i, "Dome damage"],
  [/\bdome\s+offset\b|\boffset\s+dome\b/i, "Dome offset"],
  [/\bbond(?:ing)?\s+offset\b|\boffset\s+bond/i, "Bonding offset"],
  [/\bover\s+bond\b|\bover\s+glue\b|\bglue\s+over\b/i, "Over bond/glue"],
  [/\bair\s+leak\b|\bleak\s+air\b/i, "Air leak"],
  [/\bparticle\b|\bdust\b/i, "Particle"],
  [/\bburr\b/i, "Burr"],
  [/\bgap\b/i, "Gap"],
  [/\bdimension\b|\bdim\b/i, "Dimension"],
  [/\btension\b/i, "Tension"],
  [/\bplasma\b/i, "Plasma"],
  [/\bsupplier\b|\bvender\b|\bvendor\b/i, "Supplier"],
  [/\bmaterial\b|\bfilm\b|\bplate\b|\byoke\b|\bcd\b|\bcm\b|\bsm\b/i, "Material"],
  [/\bdry\s+uv\b|\buv\s+led\b|\bled\s+uv\b/i, "Dry UV"],
  [/\bmold\b/i, "Mold"],
  [/\bjig\b/i, "JIG"],
  [/\breliability\b|\bdrop\b|\bload\b|\bshock\b|\btemperature\b|\bhumidity\b/i, "Reliability"],
  [/\bng\s+rate\b|\brate\s+ng\b/i, "NG rate"],
  [/\bbefore\b.*\bafter\b|\bafter\b.*\bbefore\b/i, "Before/After"],
  [/\bdoe\b/i, "DOE"],
  [/\bvision\b|\baoi\b/i, "Vision/AOI"],
  [/\bnoise\b/i, "Noise"],
  [/\btouch\b/i, "Touch"],
  [/\bsigma\b/i, "Sigma"],
  [/\bspl\s*[+&/ ]\s*thd\b/i, "SPL+THD"],
  [/\bthd\b/i, "THD"],
  [/\bspl\b/i, "SPL"],
];

function htmlEscape(value) {
  return String(value ?? "").replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[ch]));
}

function normKey(text) {
  return String(text || "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function listValue(value) {
  if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
  if (typeof value === "string" && value.trim()) return [value.trim()];
  return [];
}

function parseJsonl(text) {
  return String(text || "").split(/\r?\n/).map(line => line.trim()).filter(Boolean).map(line => {
    try { return JSON.parse(line); } catch { return null; }
  }).filter(Boolean);
}

function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let quote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; continue; }
    if (ch === '"') { quote = !quote; continue; }
    if (ch === "," && !quote) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function parseModelCsv(text) {
  const lines = String(text || "").split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const headers = parseCsvLine(lines[0]).map(h => h.trim());
  const srcIdx = headers.indexOf("source_model");
  const dstIdx = headers.indexOf("target_model");
  if (srcIdx < 0 || dstIdx < 0) return [];
  const pairs = [];
  for (const line of lines.slice(1)) {
    const cols = parseCsvLine(line);
    const src = (cols[srcIdx] || "").trim();
    const dst = (cols[dstIdx] || "").trim();
    if (src && dst) pairs.push([src, dst]);
  }
  return pairs.sort((a, b) => normKey(b[0]).length - normKey(a[0]).length);
}

function mapModel(texts) {
  const matches = [];
  for (const rawText of texts) {
    const hay = normKey(rawText);
    if (!hay) continue;
    for (const [src, dst] of modelPairs) {
      const needle = normKey(src);
      const pos = hay.indexOf(needle);
      if (needle && pos >= 0) matches.push({ pos, len: needle.length, src, dst });
    }
  }
  matches.sort((a, b) => a.pos - b.pos || b.len - a.len);
  const targets = [];
  const sources = [];
  for (const m of matches) {
    if (!targets.includes(m.dst)) targets.push(m.dst);
    if (!sources.includes(m.src)) sources.push(m.src);
  }
  return { model: targets.slice(0, 3).join(" / "), sources: sources.slice(0, 8) };
}

function canonicalTerm(raw) {
  const text = String(raw || "").replace(/\s+/g, " ").trim();
  if (!text) return "";
  for (const [re, replacement] of termRules) {
    if (re.test(text)) return replacement;
  }
  const replacements = { ng:"NG", uv:"UV", led:"LED", vp:"VP", cd:"CD", sp:"SP", bp:"BP", sm:"SM", cm:"CM", dt:"DT", gmi:"GMI", aoi:"AOI", ir:"IR" };
  return text.replace(/^[-_./,;:\s]+|[-_./,;:\s]+$/g, "").split(/\s+/).map(part => {
    const key = part.toLowerCase().replace(/[()[\]]/g, "");
    return replacements[key] || part;
  }).join(" ");
}

function canonicalList(values) {
  const out = [];
  for (const value of values) {
    const canonical = canonicalTerm(value);
    if (canonical && !out.includes(canonical)) out.push(canonical);
  }
  return out;
}

function flattenRaw(item) {
  if (!item || typeof item !== "object") return null;
  if (!item.classification) return item;
  const cls = item.classification || {};
  const dataset = String(item.datasetName || "");
  const files = String(item.fileNames || "");
  const aiModel = String(cls.model || "").trim();
  const dbModel = String(item.dbProductType || "");
  let mapped = mapModel([dataset, files]);
  if (!mapped.model) mapped = mapModel([aiModel, dbModel]);
  return {
    datasetName: dataset,
    fileNames: files,
    dbProductType: dbModel,
    dbReportDate: String(item.dbReportDate || ""),
    aiModel,
    model: mapped.model || aiModel || dbModel,
    modelMappingSource: mapped.sources.join(" | "),
    date: String(cls.date || ""),
    purposeCode: String(cls.purposeCode || ""),
    reviewPurpose: String(cls.reviewPurpose || ""),
    purpose: String(cls.purpose || ""),
    targetDefects: canonicalList(listValue(cls.targetDefects)),
    reviewItems: canonicalList(listValue(cls.reviewItems)),
    tags: canonicalList(listValue(cls.tags)),
    confidence: typeof cls.confidence === "number" ? cls.confidence : 0,
    needsDetailedAnalysis: Boolean(cls.needsDetailedAnalysis),
    evidenceSummary: String(cls.evidenceSummary || ""),
    evidenceCells: listValue(cls.evidenceCells),
    uncertainty: String(cls.uncertainty || ""),
  };
}

async function connectFolder() {
  if (!window.showDirectoryPicker) {
    els.liveText.textContent = "folder api unavailable";
    els.liveDot.className = "dot err";
    return;
  }
  dirHandle = await window.showDirectoryPicker({ id: "ai-batch-result-root", mode: "readwrite" });
  els.pathLabel.textContent = dirHandle.name;
  await refreshFromFolder();
  if (requests.length) await autoSaveRequestsMd();
  clearInterval(pollTimer);
  pollTimer = setInterval(refreshFromFolder, 5000);
}

async function getFileText(paths) {
  if (!dirHandle) return null;
  for (const parts of paths) {
    try {
      let handle = dirHandle;
      for (const part of parts.slice(0, -1)) handle = await handle.getDirectoryHandle(part);
      const fileHandle = await handle.getFileHandle(parts[parts.length - 1]);
      const file = await fileHandle.getFile();
      return await file.text();
    } catch {}
  }
  return null;
}

async function getWritableSampleDir() {
  if (!dirHandle) return null;
  try { return await dirHandle.getDirectoryHandle("sample_ready"); } catch { return dirHandle; }
}

async function refreshFromFolder() {
  if (!dirHandle) return;
  try {
    const summaryText = await getFileText([["classification_summary.json"]]);
    if (summaryText) summary = JSON.parse(summaryText);
    const progressText = await getFileText([["classification_progress.jsonl"]]);
    if (progressText) progress = parseJsonl(progressText).slice(-20);
    const mapText = await getFileText([["sample_ready", "model_mapping_conditions.csv"], ["model_mapping_conditions.csv"]]);
    const parsedMap = parseModelCsv(mapText || "");
    if (parsedMap.length) modelPairs = parsedMap;
    const resultsText = await getFileText([["classification_results.jsonl"]]);
    if (resultsText) {
      rows = parseJsonl(resultsText).map(flattenRaw).filter(Boolean);
    } else {
      const demoText = await getFileText([["sample_ready", "demo_index.json"], ["demo_index.json"]]);
      if (demoText) rows = JSON.parse(demoText);
    }
    lastReadAt = new Date();
    els.liveDot.className = "dot live";
    els.liveText.textContent = "live";
  } catch (err) {
    els.liveDot.className = "dot err";
    els.liveText.textContent = "read error";
  }
  renderAll();
}

function fmtDate(value) {
  if (!value) return "-";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString();
}

function renderProgress() {
  const total = Number(summary.total || 0);
  const done = Number(summary.done || 0);
  const failed = Number(summary.failed || 0);
  const remaining = Number(summary.remaining || Math.max(total - done, 0));
  const pct = total ? Math.min(100, Math.round(done / total * 1000) / 10) : 0;
  els.statTotal.textContent = total.toLocaleString();
  els.statDone.textContent = done.toLocaleString();
  els.statRemaining.textContent = remaining.toLocaleString();
  els.statFailed.textContent = failed.toLocaleString();
  els.statRows.textContent = rows.length.toLocaleString();
  els.progressBar.style.width = pct + "%";
  els.progressPct.textContent = pct + "%";
  els.updatedAt.textContent = "updated " + fmtDate(summary.updatedAt || embeddedState.generatedAt);
  if (!dirHandle) {
    els.liveText.textContent = "snapshot";
    els.liveDot.className = "dot";
  } else if (lastReadAt) {
    els.liveText.textContent = "live " + lastReadAt.toLocaleTimeString();
  }
}

function countsBy(field) {
  const map = new Map();
  for (const row of rows) {
    const value = String(row[field] || "").trim() || "(blank)";
    map.set(value, (map.get(value) || 0) + 1);
  }
  return [...map.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
}

function renderFilters() {
  const current = els.modelFilter.value;
  const options = ['<option value="">All Models</option>'];
  for (const [model, count] of countsBy("model")) {
    options.push(`<option value="${htmlEscape(model)}">${htmlEscape(model)} (${count})</option>`);
  }
  els.modelFilter.innerHTML = options.join("");
  if ([...els.modelFilter.options].some(o => o.value === current)) els.modelFilter.value = current;
}

function rowSearchText(row) {
  return [
    row.datasetName, row.fileNames, row.model, row.aiModel, row.reviewPurpose, row.purpose,
    ...(row.targetDefects || []), ...(row.reviewItems || []), ...(row.tags || []), row.uncertainty
  ].join(" ").toLowerCase();
}

function filteredRows() {
  const q = els.searchBox.value.trim().toLowerCase();
  const model = els.modelFilter.value;
  const status = els.statusFilter.value;
  return rows.filter(row => {
    if (q && !rowSearchText(row).includes(q)) return false;
    if (model && String(row.model || "") !== model) return false;
    if (status === "needs" && !row.needsDetailedAnalysis) return false;
    if (status === "low" && Number(row.confidence || 0) >= 0.75) return false;
    if (status === "missing" && String(row.model || "").trim()) return false;
    if (status === "uncertain" && !String(row.uncertainty || "").trim()) return false;
    return true;
  });
}

function pills(values, field, row) {
  const arr = Array.isArray(values) ? values : [];
  if (!arr.length) return '<span class="muted">-</span>';
  return '<div class="pill-list">' + arr.map(v => `<span class="pill term" data-field="${htmlEscape(field)}" data-term="${htmlEscape(v)}">${htmlEscape(v)}</span>`).join("") + '</div>';
}

function renderTable() {
  const filtered = filteredRows();
  const limit = Number(els.limitFilter.value || 100);
  const shown = filtered.slice(0, limit);
  els.filterCount.textContent = `${filtered.length.toLocaleString()} rows`;
  els.rowBody.innerHTML = shown.map((row, idx) => {
    const globalIdx = rows.indexOf(row);
    const conf = Number(row.confidence || 0);
    const confClass = conf < 0.75 ? "bad" : row.needsDetailedAnalysis ? "warn" : "";
    return `<tr data-idx="${globalIdx}" class="${selectedRow === row ? "active" : ""}">
      <td class="number">${idx + 1}</td>
      <td class="dataset">${htmlEscape(row.datasetName || "")}</td>
      <td><span class="pill model">${htmlEscape(row.model || "-")}</span></td>
      <td class="textcell">${htmlEscape(row.reviewPurpose || row.purpose || "")}</td>
      <td>${pills(row.targetDefects, "targetDefects", row)}</td>
      <td>${pills(row.reviewItems, "reviewItems", row)}</td>
      <td>${pills(row.tags, "tags", row)}</td>
      <td><span class="pill ${confClass}">${conf.toFixed(2)}</span></td>
      <td><button class="select-row" data-idx="${globalIdx}">확인</button></td>
    </tr>`;
  }).join("");
}

function renderBars() {
  const data = countsBy("model").slice(0, 14);
  const max = Math.max(1, ...data.map(x => x[1]));
  els.modelBars.innerHTML = data.map(([model, count]) => `
    <div class="bar-row">
      <div>
        <div class="bar-label">${htmlEscape(model)}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${count / max * 100}%"></div></div>
      </div>
      <div class="number">${count}</div>
    </div>`).join("") || '<div class="empty">No data</div>';
}

function renderRecent() {
  const items = progress.slice(-10).reverse();
  els.recentList.innerHTML = items.map(item => {
    const label = item.event || "event";
    const batch = item.batch != null ? `#${item.batch}` : "";
    const done = item.done != null ? ` done ${item.done}` : "";
    const names = Array.isArray(item.datasets) ? item.datasets.slice(0, 2).join(" / ") : "";
    return `<div class="recent-item"><b>${htmlEscape(label)} ${htmlEscape(batch)}${htmlEscape(done)}</b><div class="muted">${htmlEscape(fmtDate(item.at))}</div><div>${htmlEscape(names)}</div></div>`;
  }).join("") || '<div class="empty">No progress events</div>';
}

function fieldHtml(name, value) {
  return `<div class="field"><div class="name">${htmlEscape(name)}</div><div class="value">${htmlEscape(value || "-")}</div></div>`;
}

function rowIssues(row) {
  const issues = [];
  const idx = rows.indexOf(row);
  const conf = Number(row.confidence || 0);
  const model = String(row.model || "").trim();
  const uncertainty = String(row.uncertainty || "").trim();
  const targetDefects = row.targetDefects || [];
  const reviewItems = row.reviewItems || [];
  const tags = row.tags || [];
  if (!model) {
    issues.push({
      id: `${idx}:model-missing`,
      rowIndex: idx,
      type: "model",
      severity: 1,
      title: "모델명 확인 필요",
      question: "이 파일/데이터셋은 어떤 표준 모델명으로 묶어야 합니까?",
      original: row.datasetName || row.aiModel || row.dbProductType || "",
      canonical: "",
      note: "파일명 기준 모델 매핑에 추가할 값이면 Canonical에 표준 모델명을 입력하세요.",
    });
  } else if (model.includes(" / ")) {
    issues.push({
      id: `${idx}:model-multiple`,
      rowIndex: idx,
      type: "model",
      severity: 2,
      title: "모델명 복수 매핑",
      question: "복수 모델로 잡힌 이 건은 대표 모델을 하나로 고정해야 합니까, 아니면 비교 자료로 유지해야 합니까?",
      original: model,
      canonical: model.split(" / ")[0],
      note: "대표 모델 하나만 남길 경우 Canonical을 수정하세요. 비교 자료면 '비교자료 유지'라고 적어주세요.",
    });
  }
  if (conf > 0 && conf < 0.75) {
    issues.push({
      id: `${idx}:low-confidence`,
      rowIndex: idx,
      type: "prompt",
      severity: 3,
      title: "AI 분류 신뢰도 낮음",
      question: "이 건의 목적/결함/검토항목 중 어느 쪽을 우선 보정해야 합니까?",
      original: `${row.reviewPurpose || row.purpose || ""} / ${targetDefects.join(" | ")}`,
      canonical: "",
      note: "보정 방향을 적으면 이후 프롬프트 주의사항 또는 정규화 룰에 반영합니다.",
    });
  }
  if (uncertainty) {
    issues.push({
      id: `${idx}:uncertainty`,
      rowIndex: idx,
      type: "prompt",
      severity: 4,
      title: "AI가 애매하다고 표시",
      question: "아래 uncertainty 내용은 어떤 기준으로 처리해야 합니까?",
      original: uncertainty,
      canonical: "",
      note: "예: 본문 날짜 우선, 파일명 날짜 무시, Decision 비어 있으면 판정 미확정 등.",
    });
  }
  if (!targetDefects.length && row.needsDetailedAnalysis) {
    issues.push({
      id: `${idx}:target-empty`,
      rowIndex: idx,
      type: "targetDefects",
      severity: 5,
      title: "Target Defect 비어 있음",
      question: "이 건은 결함 없음/조건 검증 자료입니까, 아니면 Target Defect를 지정해야 합니까?",
      original: row.reviewPurpose || row.purpose || row.datasetName || "",
      canonical: "",
      note: "결함 없음이면 Canonical에 'No defect / 조건 검증'처럼 입력하세요.",
    });
  }
  if (!reviewItems.length && (row.reviewPurpose || row.purpose)) {
    issues.push({
      id: `${idx}:review-empty`,
      rowIndex: idx,
      type: "reviewItems",
      severity: 6,
      title: "Review Item 비어 있음",
      question: "이 건에서 검토항목으로 묶어야 할 대표 단어는 무엇입니까?",
      original: row.reviewPurpose || row.purpose || "",
      canonical: "",
      note: "예: Vision/AOI, Tension, Reliability, NG rate, Dry UV.",
    });
  }
  if (!tags.length && (targetDefects.length || reviewItems.length)) {
    issues.push({
      id: `${idx}:tag-empty`,
      rowIndex: idx,
      type: "tags",
      severity: 7,
      title: "Tag 비어 있음",
      question: "검색용 태그로 어떤 대표 단어를 추가해야 합니까?",
      original: `${targetDefects.join(" | ")} / ${reviewItems.join(" | ")}`,
      canonical: "",
      note: "나중에 현재 문제 검색에 쓸 태그 기준입니다.",
    });
  }
  return issues;
}

function allIssues() {
  return rows.flatMap(rowOpenIssues).sort((a, b) => a.severity - b.severity || a.rowIndex - b.rowIndex);
}

function findIssue(id) {
  return allIssues().find(issue => issue.id === id) || null;
}

function issueFieldPreview(row) {
  return [
    fieldHtml("Dataset", row.datasetName),
    fieldHtml("Current Model", row.model),
    fieldHtml("AI Purpose", row.reviewPurpose || row.purpose),
    fieldHtml("Target Defects", (row.targetDefects || []).join(" | ")),
    fieldHtml("Review Items", (row.reviewItems || []).join(" | ")),
    fieldHtml("Tags", (row.tags || []).join(" | ")),
    fieldHtml("Uncertainty", row.uncertainty),
  ].join("");
}

function requestKey(req) {
  return [req.type || "", req.original || "", req.canonical || "", req.dataset || "", req.question || ""].join("\u001f");
}

function mergeRequests(primary, secondary) {
  const out = [];
  const seen = new Set();
  for (const req of [...(primary || []), ...(secondary || [])]) {
    if (!req || !String(req.canonical || "").trim()) continue;
    const key = requestKey(req);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(req);
  }
  return out;
}

function tokenSet(text) {
  return new Set(String(text || "").toLowerCase().split(/[^a-z0-9가-힣]+/).filter(x => x.length >= 3));
}

function overlapScore(a, b) {
  const aa = tokenSet(a);
  const bb = tokenSet(b);
  let score = 0;
  for (const token of aa) {
    if (bb.has(token)) score += 1;
  }
  return score;
}

function relatedRequests(issue, row) {
  if (!issue) return [];
  return requests.map(req => {
    let score = 0;
    if (req.type === issue.type) score += 3;
    if (req.issueType === issue.title) score += 4;
    if (req.question === issue.question) score += 3;
    if (req.model && row.model && req.model === row.model) score += 2;
    score += Math.min(4, overlapScore(req.original, issue.original));
    score += Math.min(3, overlapScore(req.dataset, row.datasetName));
    return { req, score };
  }).filter(x => x.score >= 4).sort((a, b) => b.score - a.score).slice(0, 4);
}

function autoResolvedMatch(issue, row) {
  const match = relatedRequests(issue, row).find(x => x.score >= AUTO_RESOLVE_SCORE && String(x.req.canonical || "").trim());
  return match || null;
}

function rowOpenIssues(row) {
  return rowIssues(row).filter(issue => !autoResolvedMatch(issue, row));
}

function rowAutoResolvedIssues(row) {
  return rowIssues(row).map(issue => ({ issue, match: autoResolvedMatch(issue, row) })).filter(x => x.match);
}

function renderIssuePane() {
  const issues = allIssues();
  const autoCount = rows.reduce((sum, row) => sum + rowAutoResolvedIssues(row).length, 0);
  els.issueCount.textContent = `${issues.length} / 자동 ${autoCount}`;
  if (!selectedIssue || !findIssue(selectedIssue.id)) {
    selectedIssue = issues[0] || null;
    selectedRow = selectedIssue ? rows[selectedIssue.rowIndex] : null;
  }
  if (!selectedIssue) {
    els.issuePane.innerHTML = `<div class="empty">확인 필요 항목 없음<br>기존 답변으로 자동정리된 항목: ${autoCount}</div>`;
    return;
  }
  if (activeAnswerIssueId !== selectedIssue.id) {
    prefillIssue(selectedIssue);
  }
  const row = rows[selectedIssue.rowIndex] || {};
  const topIssues = issues.slice(0, 20);
  const related = relatedRequests(selectedIssue, row);
  els.issuePane.innerHTML = `
    <div class="issue-title">${htmlEscape(selectedIssue.title)}</div>
    <div class="issue-question">${htmlEscape(selectedIssue.question)}</div>
    <div class="issue-meta">
      <span class="pill ${Number(row.confidence || 0) < 0.75 ? "bad" : "warn"}">conf ${Number(row.confidence || 0).toFixed(2)}</span>
      <span class="pill model">${htmlEscape(row.model || "model blank")}</span>
      <span class="pill">${htmlEscape(selectedIssue.type)}</span>
    </div>
    <button id="prefillIssueBtn" class="primary">이 질문 답변란에 넣기</button>
    ${related.length ? `<div class="field"><div class="name">기존 답변 후보</div><div class="related-list">
      ${related.map(({ req }, idx) => `<div class="related-item">
        <b>${htmlEscape(req.canonical || "")}</b>
        <div class="muted">${htmlEscape(req.issueType || req.type || "")}</div>
        <div class="muted">${htmlEscape(req.dataset || "")}</div>
        <button type="button" data-apply-related="${idx}">이 답변 적용</button>
      </div>`).join("")}
    </div></div>` : ""}
    ${issueFieldPreview(row)}
    <div class="field"><div class="name">다른 확인 필요 항목</div><div class="issue-list">
      ${topIssues.map(issue => `<div class="issue-item ${issue.id === selectedIssue.id ? "active" : ""}" data-issue-id="${htmlEscape(issue.id)}">
        <b>${htmlEscape(issue.title)}</b>
        <div class="muted">${htmlEscape((rows[issue.rowIndex] || {}).datasetName || "")}</div>
      </div>`).join("")}
    </div></div>
  `;
}

function renderRequests() {
  els.reqCount.textContent = String(requests.length);
  els.requestList.innerHTML = requests.map((req, idx) => `
    <div class="request-item">
      <b>${htmlEscape(req.issueType || req.type)}</b>
      <div>답변: ${htmlEscape(req.canonical || "-")}</div>
      <div class="muted">AI 값: ${htmlEscape(req.original || "-")}</div>
      <div class="muted">${htmlEscape(req.dataset || "")}</div>
      <div>${htmlEscape(req.note || "")}</div>
      <button data-remove-req="${idx}">삭제</button>
    </div>`).join("") || '<div class="empty">요청 없음</div>';
}

function renderAll() {
  renderProgress();
  renderFilters();
  renderTable();
  renderBars();
  renderRequests();
}

function selectRowByIndex(idx) {
  selectedRow = rows[idx] || null;
  const issues = selectedRow ? rowIssues(selectedRow) : [];
  selectedIssue = issues[0] || {
    id: `${idx}:manual`,
    rowIndex: idx,
    type: "prompt",
    title: "수동 확인",
    question: "이 건에 추가할 프롬프트/정규화 기준을 입력하세요.",
    original: selectedRow ? selectedRow.datasetName : "",
    canonical: "",
    note: "",
  };
  prefillIssue(selectedIssue);
  renderTable();
}

function prefillTerm(field, term) {
  els.reqType.value = field;
  els.reqOriginal.value = term;
  els.reqCanonical.value = term;
  els.reqAnswer.value = term;
  activeAnswerIssueId = `term:${field}:${term}`;
  els.answerContext.textContent = `항목 정리 필요\n\n필드: ${field}\nAI 값: ${term}`;
  if (selectedRow) els.reqNote.value = selectedRow.datasetName || "";
}

function prefillIssue(issue) {
  if (!issue) return;
  selectedIssue = issue;
  selectedRow = rows[issue.rowIndex] || selectedRow;
  activeAnswerIssueId = issue.id;
  els.reqType.value = issue.type || "prompt";
  els.reqOriginal.value = issue.original || "";
  els.reqCanonical.value = issue.canonical || "";
  els.reqAnswer.value = issue.canonical || "";
  els.answerContext.textContent = [
    `질문: ${issue.question || ""}`,
    "",
    `AI가 잡은 값: ${issue.original || "-"}`,
    "",
    `데이터: ${selectedRow ? selectedRow.datasetName || "" : ""}`,
    selectedRow ? `현재 모델: ${selectedRow.model || "-"}` : "",
    issue.note ? `가이드: ${issue.note}` : "",
  ].filter(Boolean).join("\n");
  els.reqNote.value = "";
}

function loadRequests() {
  try { return JSON.parse(localStorage.getItem("aiBatchControlRequests") || "[]"); } catch { return []; }
}

function setSaveStatus(kind, text) {
  els.saveStatus.textContent = text;
  els.saveDot.className = kind === "ok" ? "dot live" : kind === "err" ? "dot err" : "dot";
}

function storeRequests() {
  localStorage.setItem("aiBatchControlRequests", JSON.stringify(requests));
  setSaveStatus("pending", dirHandle ? "브라우저 임시 저장됨, md 자동 저장 대기" : "브라우저 임시 저장됨. md 저장은 폴더 연결 후 가능");
}

function addRequest() {
  const answer = els.reqAnswer.value.trim();
  const req = {
    createdAt: new Date().toISOString(),
    type: els.reqType.value,
    original: els.reqOriginal.value.trim(),
    canonical: answer || els.reqCanonical.value.trim(),
    note: els.reqNote.value.trim(),
    dataset: selectedRow ? selectedRow.datasetName : "",
    model: selectedRow ? selectedRow.model : "",
    issueType: selectedIssue ? selectedIssue.title : "",
    question: selectedIssue ? selectedIssue.question : "",
  };
  requests.unshift(req);
  els.reqAnswer.value = "";
  els.reqNote.value = "";
  storeRequests();
  renderRequests();
  autoSaveRequestsMd();
}

function buildRequestMd() {
  const lines = ["# Prompt Update Requests", "", `Generated: ${new Date().toISOString()}`, ""];
  if (!requests.length) {
    lines.push("No pending requests.");
    return lines.join("\n") + "\n";
  }
  lines.push("## Pending Requests", "");
  requests.forEach((req, idx) => {
    lines.push(`### ${idx + 1}. ${req.type}`);
    lines.push(`- AI Value: ${req.original || ""}`);
    lines.push(`- User Answer: ${req.canonical || ""}`);
    lines.push(`- Dataset: ${req.dataset || ""}`);
    lines.push(`- Model: ${req.model || ""}`);
    lines.push(`- Issue: ${req.issueType || ""}`);
    lines.push(`- Question: ${req.question || ""}`);
    lines.push(`- Note: ${req.note || ""}`);
    lines.push("");
  });
  lines.push("## Apply Target");
  lines.push("- Update model mapping rules when type is `model`.");
  lines.push("- Update canonical term rules when type is `targetDefects`, `reviewItems`, or `tags`.");
  lines.push("- Update AI batch prompt cautions when type is `prompt`.");
  return lines.join("\n") + "\n";
}

async function saveRequestsMd() {
  const md = buildRequestMd();
  if (dirHandle) {
    await writeRequestsMd(md);
    return true;
  }
  const blob = new Blob([md], { type: "text/markdown;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "prompt_update_requests.md";
  a.click();
  URL.revokeObjectURL(a.href);
  setSaveStatus("ok", "다운로드 파일로 저장됨. 폴더 연결 시 직접 md에 저장됩니다.");
  return true;
}

async function writeRequestsMd(md) {
  try {
    const sampleDir = await getWritableSampleDir();
    const fileHandle = await sampleDir.getFileHandle("prompt_update_requests.md", { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(md);
    await writable.close();
    els.liveText.textContent = "request saved";
    setSaveStatus("ok", "prompt_update_requests.md 저장 완료");
    return true;
  } catch (err) {
    setSaveStatus("err", "md 저장 실패. 폴더 연결 권한을 다시 확인하세요.");
    return false;
  }
}

async function autoSaveRequestsMd() {
  if (!dirHandle) {
    setSaveStatus("pending", "브라우저에는 저장됨. 실제 md 저장은 폴더 연결 후 `요청 MD 저장`을 누르세요.");
    return false;
  }
  return writeRequestsMd(buildRequestMd());
}

document.getElementById("connectBtn").addEventListener("click", connectFolder);
document.getElementById("refreshBtn").addEventListener("click", () => dirHandle ? refreshFromFolder() : renderAll());
document.getElementById("saveReqBtn").addEventListener("click", saveRequestsMd);
document.getElementById("problemSearchBtn").addEventListener("click", () => window.open("current_problem_search.html", "_blank", "noopener"));
document.getElementById("termsBtn").addEventListener("click", () => window.open("ai_term_glossary.html", "_blank", "noopener"));
document.getElementById("addReqBtn").addEventListener("click", addRequest);
document.getElementById("clearReqBtn").addEventListener("click", () => { requests = []; storeRequests(); renderRequests(); autoSaveRequestsMd(); });
for (const el of [els.searchBox, els.modelFilter, els.statusFilter, els.limitFilter]) {
  el.addEventListener("input", () => { renderTable(); });
}
els.rowBody.addEventListener("click", event => {
  const term = event.target.closest("[data-term]");
  if (term) {
    prefillTerm(term.dataset.field, term.dataset.term);
    return;
  }
  const btn = event.target.closest("[data-idx]");
  if (btn) selectRowByIndex(Number(btn.dataset.idx));
});
els.issuePane.addEventListener("click", event => {
  const issueItem = event.target.closest("[data-issue-id]");
  if (issueItem) {
    const issue = findIssue(issueItem.dataset.issueId);
    if (issue) {
      selectedIssue = issue;
      selectedRow = rows[issue.rowIndex] || null;
      prefillIssue(issue);
      renderTable();
      renderIssuePane();
    }
    return;
  }
  const prefillBtn = event.target.closest("#prefillIssueBtn");
  if (prefillBtn) prefillIssue(selectedIssue);
  const relatedBtn = event.target.closest("[data-apply-related]");
  if (relatedBtn && selectedIssue && selectedRow) {
    const related = relatedRequests(selectedIssue, selectedRow);
    const item = related[Number(relatedBtn.dataset.applyRelated)];
    if (item && item.req) {
      els.reqAnswer.value = item.req.canonical || "";
      els.reqNote.value = item.req.note || "";
      setSaveStatus("pending", "기존 답변을 답변란에 적용했습니다. `답변 추가`를 누르면 저장됩니다.");
    }
  }
});
const requestForm = document.querySelector(".request-form");
if (requestForm) {
  requestForm.addEventListener("click", event => {
    const quick = event.target.closest("[data-quick-answer]");
    if (!quick) return;
    els.reqAnswer.value = quick.dataset.quickAnswer || "";
  });
}
els.requestList.addEventListener("click", event => {
  const btn = event.target.closest("[data-remove-req]");
  if (!btn) return;
  requests.splice(Number(btn.dataset.removeReq), 1);
  storeRequests();
  renderRequests();
  autoSaveRequestsMd();
});

renderAll();
setSaveStatus("pending", requests.length ? "브라우저 임시 저장 답변이 있습니다. 폴더 연결 후 md 저장하세요." : "답변 없음");
</script>
</body>
</html>
"""


TERM_GLOSSARY_TEMPLATE = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI Term Guidance</title>
<style>
* { box-sizing: border-box; }
body {
  margin: 0;
  background: #f5f6f8;
  color: #161a20;
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
  border-bottom: 1px solid #d6d9de;
}
h1 { margin: 0; font-size: 18px; }
button, input, textarea {
  font: inherit;
  border: 1px solid #bcc3cd;
  border-radius: 6px;
  padding: 8px 10px;
  background: #fff;
}
button { cursor: pointer; }
button.primary { background: #315fcb; color: #fff; border-color: #315fcb; }
main { display: grid; gap: 12px; padding: 12px; }
.panel {
  background: #fff;
  border: 1px solid #d6d9de;
  box-shadow: 0 1px 2px rgba(16, 24, 40, .08);
}
.panel-head {
  padding: 10px 12px;
  border-bottom: 1px solid #d6d9de;
  font-weight: 700;
  background: #fafbfc;
}
.panel-body { padding: 12px; display: grid; gap: 8px; }
.row { display: grid; grid-template-columns: 220px minmax(0, 1fr); gap: 8px; }
textarea { min-height: 100px; resize: vertical; line-height: 1.5; }
.list { display: grid; gap: 8px; }
.candidate-list { display: grid; gap: 8px; max-height: 360px; overflow: auto; }
.candidate {
  border: 1px solid #bfd0f7;
  border-radius: 6px;
  padding: 10px;
  background: #f7faff;
}
.candidate b { display: block; margin-bottom: 5px; }
.candidate .meta { color: #616b78; font-size: 12px; }
.candidate .examples { margin-top: 5px; color: #303846; }
.item {
  border: 1px solid #d6d9de;
  border-radius: 6px;
  padding: 10px;
  background: #fff;
}
.item b { display: block; margin-bottom: 5px; }
.muted { color: #616b78; }
.toolbar { display: flex; gap: 8px; flex-wrap: wrap; }
@media (max-width: 760px) {
  header, .row { grid-template-columns: 1fr; display: grid; align-items: stretch; }
}
</style>
</head>
<body>
<header>
  <h1>AI 용어 정리</h1>
  <div class="toolbar">
    <button id="connectBtn">폴더 연결</button>
    <button id="saveBtn" class="primary">MD 저장</button>
    <button id="downloadBtn">다운로드</button>
  </div>
</header>
<main>
  <section class="panel">
    <div class="panel-head">AI가 헷갈릴 수 있는 용어 후보 <span class="muted" id="candidateCount">0</span></div>
    <div class="panel-body">
      <div class="candidate-list" id="candidateList"></div>
    </div>
  </section>
  <section class="panel">
    <div class="panel-head">용어 추가</div>
    <div class="panel-body">
      <div class="row">
        <input id="termInput" placeholder="용어 / 약어 / 현장 표현">
        <textarea id="meaningInput" placeholder="AI가 이해해야 할 뜻을 자세히 적어주세요. 예: Standard = 특정 모델 하나가 아니라 표준 기준/일반 조건으로 처리"></textarea>
      </div>
      <textarea id="usageInput" placeholder="분석할 때 적용 기준이나 예외(선택)"></textarea>
      <button id="addBtn" class="primary">용어 추가</button>
      <div class="muted" id="saveStatus">브라우저 임시 저장 상태</div>
    </div>
  </section>
  <section class="panel">
    <div class="panel-head">저장될 용어</div>
    <div class="panel-body"><div class="list" id="termList"></div></div>
  </section>
</main>
<script id="initial-terms" type="application/json">__TERMS_STATE__</script>
<script id="initial-candidates" type="application/json">__CANDIDATES_STATE__</script>
<script>
let dirHandle = null;
let entries = mergeEntries(loadEntries(), JSON.parse(document.getElementById("initial-terms").textContent || "[]"));
let candidates = JSON.parse(document.getElementById("initial-candidates").textContent || "[]");
const els = {
  termInput: document.getElementById("termInput"),
  meaningInput: document.getElementById("meaningInput"),
  usageInput: document.getElementById("usageInput"),
  termList: document.getElementById("termList"),
  candidateList: document.getElementById("candidateList"),
  candidateCount: document.getElementById("candidateCount"),
  saveStatus: document.getElementById("saveStatus"),
};
function esc(value) {
  return String(value || "").replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[ch]));
}
function loadEntries() {
  try { return JSON.parse(localStorage.getItem("aiTermGuidanceEntries") || "[]"); } catch { return []; }
}
function keyOf(entry) {
  return [entry.term || "", entry.meaning || "", entry.usage || ""].join("\u001f");
}
function mergeEntries(a, b) {
  const out = [];
  const seen = new Set();
  for (const entry of [...(a || []), ...(b || [])]) {
    if (!entry || (!String(entry.term || "").trim() && !String(entry.meaning || "").trim())) continue;
    const key = keyOf(entry);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(entry);
  }
  return out;
}
function storeEntries() {
  localStorage.setItem("aiTermGuidanceEntries", JSON.stringify(entries));
  els.saveStatus.textContent = dirHandle ? "브라우저 저장됨. MD 자동 저장 대기" : "브라우저 저장됨. 실제 AI 반영은 MD 저장 필요";
}
function buildMarkdown() {
  const lines = ["# AI Term Guidance", "", "AI analysis must use these user-defined term explanations when reading manufacturing reports.", "", "## Terms", ""];
  entries.forEach(entry => {
    lines.push(`### ${entry.term || "(no term)"}`);
    lines.push(`- Meaning: ${entry.meaning || ""}`);
    lines.push(`- Usage: ${entry.usage || ""}`);
    lines.push("");
  });
  return lines.join("\n");
}
function render() {
  els.termList.innerHTML = entries.map((entry, idx) => `
    <div class="item">
      <b>${esc(entry.term)}</b>
      <div>${esc(entry.meaning).replace(/\n/g, "<br>")}</div>
      <div class="muted">${esc(entry.usage).replace(/\n/g, "<br>")}</div>
      <button data-remove="${idx}">삭제</button>
    </div>`).join("") || '<div class="muted">등록된 용어 없음</div>';
  renderCandidates();
}
function renderCandidates() {
  const existing = new Set(entries.map(x => String(x.term || "").trim().toLowerCase()).filter(Boolean));
  const visible = candidates.filter(x => !existing.has(String(x.term || "").trim().toLowerCase())).slice(0, 80);
  els.candidateCount.textContent = `${visible.length}`;
  els.candidateList.innerHTML = visible.map((item, idx) => `
    <div class="candidate">
      <b>${esc(item.term)}</b>
      <div class="meta">count ${item.count || 0} / score ${item.score || 0} / ${(item.reasons || []).map(esc).join(", ")}</div>
      <div class="examples">${(item.examples || []).map(esc).join("<br>")}</div>
      <button type="button" data-use-candidate="${idx}">이 용어 설명하기</button>
    </div>`).join("") || '<div class="muted">현재 후보 없음</div>';
}
async function saveMd() {
  const md = buildMarkdown();
  if (!dirHandle) {
    downloadMd();
    els.saveStatus.textContent = "폴더 미연결: 다운로드로 저장했습니다.";
    return;
  }
  try {
    const sampleDir = await getSampleDir();
    const fileHandle = await sampleDir.getFileHandle("ai_term_guidance.md", { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(md);
    await writable.close();
    els.saveStatus.textContent = "ai_term_guidance.md 저장 완료";
  } catch {
    els.saveStatus.textContent = "저장 실패. 폴더 권한을 다시 연결하세요.";
  }
}
function downloadMd() {
  const blob = new Blob([buildMarkdown()], { type: "text/markdown;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "ai_term_guidance.md";
  a.click();
  URL.revokeObjectURL(a.href);
}
async function getSampleDir() {
  try { return await dirHandle.getDirectoryHandle("sample_ready"); } catch { return dirHandle; }
}
async function connectFolder() {
  if (!window.showDirectoryPicker) {
    els.saveStatus.textContent = "이 브라우저는 폴더 저장 API를 지원하지 않습니다.";
    return;
  }
  dirHandle = await window.showDirectoryPicker({ id: "ai-batch-result-root", mode: "readwrite" });
  els.saveStatus.textContent = "폴더 연결됨. 현재 용어를 ai_term_guidance.md에 저장합니다.";
  await saveMd();
}
document.getElementById("connectBtn").addEventListener("click", connectFolder);
document.getElementById("saveBtn").addEventListener("click", saveMd);
document.getElementById("downloadBtn").addEventListener("click", downloadMd);
document.getElementById("addBtn").addEventListener("click", () => {
  const entry = { term: els.termInput.value.trim(), meaning: els.meaningInput.value.trim(), usage: els.usageInput.value.trim() };
  if (!entry.term && !entry.meaning) return;
  entries.unshift(entry);
  els.termInput.value = "";
  els.meaningInput.value = "";
  els.usageInput.value = "";
  storeEntries();
  render();
  if (dirHandle) saveMd();
});
els.candidateList.addEventListener("click", event => {
  const btn = event.target.closest("[data-use-candidate]");
  if (!btn) return;
  const existing = new Set(entries.map(x => String(x.term || "").trim().toLowerCase()).filter(Boolean));
  const visible = candidates.filter(x => !existing.has(String(x.term || "").trim().toLowerCase())).slice(0, 80);
  const item = visible[Number(btn.dataset.useCandidate)];
  if (!item) return;
  els.termInput.value = item.term || "";
  els.meaningInput.value = "";
  els.usageInput.value = [
    "AI 후보 근거:",
    (item.reasons || []).join(", "),
    "",
    "예시 데이터:",
    ...(item.examples || []),
  ].join("\n");
  els.meaningInput.focus();
});
els.termList.addEventListener("click", event => {
  const btn = event.target.closest("[data-remove]");
  if (!btn) return;
  entries.splice(Number(btn.dataset.remove), 1);
  storeEntries();
  render();
  if (dirHandle) saveMd();
});
render();
</script>
</body>
</html>
"""


TERM_GLOSSARY_TEMPLATE = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI 용어 정리</title>
<style>
* { box-sizing: border-box; }
body {
  margin: 0;
  background: #f4f6f8;
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
button.ghost { background: #fff; color: #344054; }
main { padding: 12px; }
.toolbar { display: flex; gap: 8px; flex-wrap: wrap; }
.workspace {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 430px;
  gap: 12px;
  align-items: start;
}
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
  font-weight: 700;
  background: #f9fafb;
}
.panel-body { padding: 12px; display: grid; gap: 10px; }
.filters {
  display: grid;
  grid-template-columns: minmax(220px, 1fr) 170px 170px;
  gap: 8px;
}
.table-wrap {
  border: 1px solid #d4d8df;
  max-height: calc(100vh - 165px);
  overflow: auto;
  background: #fff;
}
table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  min-width: 1120px;
}
th, td {
  border-bottom: 1px solid #e1e5eb;
  padding: 8px 9px;
  vertical-align: top;
  text-align: left;
}
th {
  position: sticky;
  top: 0;
  z-index: 1;
  background: #495565;
  color: #fff;
  font-weight: 700;
}
tbody tr { cursor: pointer; }
tbody tr:hover { background: #f3f7ff; }
tbody tr.selected { background: #eaf1ff; }
.term-cell { font-weight: 700; color: #102a56; }
.num { text-align: right; font-variant-numeric: tabular-nums; }
.muted { color: #667085; }
.small { font-size: 12px; line-height: 1.45; }
.chips { display: flex; flex-wrap: wrap; gap: 4px; }
.chip {
  display: inline-flex;
  align-items: center;
  max-width: 100%;
  border: 1px solid #c9d3e2;
  border-radius: 999px;
  padding: 2px 7px;
  color: #24364f;
  background: #f8fafc;
  font-size: 12px;
  line-height: 1.4;
}
.badge {
  display: inline-flex;
  align-items: center;
  border-radius: 999px;
  padding: 3px 8px;
  font-size: 12px;
  line-height: 1.4;
  border: 1px solid #d0d5dd;
  background: #f2f4f7;
  color: #344054;
}
.badge.need {
  border-color: #f2c98f;
  background: #fff7ed;
  color: #9a5200;
}
.badge.auto {
  border-color: #b7d7bd;
  background: #f0f9f1;
  color: #276738;
}
.field { display: grid; gap: 5px; }
.field label { font-weight: 700; color: #273244; }
textarea { width: 100%; min-height: 110px; resize: vertical; line-height: 1.5; }
#meaningInput { min-height: 180px; }
.actions { display: grid; grid-template-columns: 1fr auto; gap: 8px; align-items: center; }
.saved-list { display: grid; gap: 8px; max-height: 310px; overflow: auto; }
.item {
  border: 1px solid #d4d8df;
  border-radius: 6px;
  padding: 9px;
  background: #fff;
}
.item b { display: block; margin-bottom: 5px; }
.item button { margin-top: 7px; padding: 5px 8px; }
.empty {
  padding: 18px;
  color: #667085;
  text-align: center;
}
@media (max-width: 1080px) {
  .workspace { grid-template-columns: 1fr; }
  .side { position: static; }
  .table-wrap { max-height: 460px; }
}
@media (max-width: 720px) {
  header, .filters, .actions { display: grid; grid-template-columns: 1fr; align-items: stretch; }
}
</style>
</head>
<body>
<header>
  <h1>AI 용어 정리</h1>
  <div class="toolbar">
    <button id="connectBtn">폴더 연결</button>
    <button id="saveBtn" class="primary">MD 저장</button>
    <button id="downloadBtn">다운로드</button>
  </div>
</header>
<main>
  <div class="workspace">
    <section class="panel">
      <div class="panel-head">
        <span>AI가 헷갈릴 수 있는 용어 후보</span>
        <span class="muted" id="candidateCount">0</span>
      </div>
      <div class="panel-body">
        <div class="filters">
          <input id="candidateSearch" placeholder="용어, 유사표현, 예시 데이터 검색">
          <select id="candidateMode">
            <option value="needs">확인 필요만</option>
            <option value="all">전체 보기</option>
            <option value="weak">빈도 낮음만</option>
            <option value="auto">AI 판단 제외됨</option>
          </select>
          <select id="sortSelect">
            <option value="score">점수순</option>
            <option value="count">빈도순</option>
            <option value="term">용어순</option>
          </select>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th style="width: 150px;">용어</th>
                <th style="width: 74px;">빈도</th>
                <th style="width: 74px;">점수</th>
                <th style="width: 135px;">AI 판단</th>
                <th style="width: 250px;">유사표현</th>
                <th>예시 데이터</th>
                <th style="width: 92px;">입력</th>
              </tr>
            </thead>
            <tbody id="candidateTableBody"></tbody>
          </table>
        </div>
      </div>
    </section>
    <aside class="side">
      <section class="panel">
        <div class="panel-head">
          <span>용어 설명</span>
          <span class="muted small" id="selectedInfo">선택 없음</span>
        </div>
        <div class="panel-body">
          <div class="field">
            <label for="termInput">표준 용어</label>
            <input id="termInput" placeholder="예: VP/CD, SPL+THD, Bako">
          </div>
          <div class="field">
            <label for="meaningInput">AI가 이해해야 할 의미</label>
            <textarea id="meaningInput" placeholder="이 약어가 실제 공정/검토에서 무엇을 뜻하는지 적어주세요."></textarea>
          </div>
          <div class="field">
            <label for="usageInput">분석 적용 기준 / 예외</label>
            <textarea id="usageInput" placeholder="선택한 후보의 유사표현과 예시가 자동으로 들어옵니다. 필요한 기준을 덧붙이세요."></textarea>
          </div>
          <div class="actions">
            <button id="addBtn" class="primary">용어 저장</button>
            <button id="clearBtn" class="ghost">비우기</button>
          </div>
          <div class="muted small" id="saveStatus">브라우저 임시 저장 상태</div>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head">
          <span>저장된 용어</span>
          <span class="muted" id="termCount">0</span>
        </div>
        <div class="panel-body">
          <div class="saved-list" id="termList"></div>
        </div>
      </section>
    </aside>
  </div>
</main>
<script id="initial-terms" type="application/json">__TERMS_STATE__</script>
<script id="initial-candidates" type="application/json">__CANDIDATES_STATE__</script>
<script>
let dirHandle = null;
let entries = mergeEntries(loadEntries(), JSON.parse(document.getElementById("initial-terms").textContent || "[]"));
let candidates = JSON.parse(document.getElementById("initial-candidates").textContent || "[]");
let renderedCandidates = [];
let selectedCandidateKey = "";
const els = {
  termInput: document.getElementById("termInput"),
  meaningInput: document.getElementById("meaningInput"),
  usageInput: document.getElementById("usageInput"),
  termList: document.getElementById("termList"),
  termCount: document.getElementById("termCount"),
  candidateTableBody: document.getElementById("candidateTableBody"),
  candidateCount: document.getElementById("candidateCount"),
  candidateSearch: document.getElementById("candidateSearch"),
  candidateMode: document.getElementById("candidateMode"),
  sortSelect: document.getElementById("sortSelect"),
  selectedInfo: document.getElementById("selectedInfo"),
  saveStatus: document.getElementById("saveStatus"),
};
function esc(value) {
  return String(value || "").replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[ch]));
}
function loadEntries() {
  try { return JSON.parse(localStorage.getItem("aiTermGuidanceEntries") || "[]"); } catch { return []; }
}
function keyOf(entry) {
  return [entry.term || "", entry.meaning || "", entry.usage || ""].join("\u001f");
}
function normalize(value) {
  return String(value || "").trim().toLowerCase();
}
function candidateKeys(item) {
  return [item.term, ...(item.aliases || [])].map(normalize).filter(Boolean);
}
function mergeEntries(a, b) {
  const out = [];
  const seen = new Set();
  for (const entry of [...(a || []), ...(b || [])]) {
    if (!entry || (!String(entry.term || "").trim() && !String(entry.meaning || "").trim())) continue;
    const key = keyOf(entry);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(entry);
  }
  return out;
}
function storeEntries() {
  localStorage.setItem("aiTermGuidanceEntries", JSON.stringify(entries));
  els.saveStatus.textContent = dirHandle
    ? "브라우저에 저장됨. MD 자동 저장 대기"
    : "브라우저에 저장됨. 실제 AI 반영은 MD 저장이 필요합니다.";
}
function buildMarkdown() {
  const lines = ["# AI Term Guidance", "", "AI analysis must use these user-defined term explanations when reading manufacturing reports.", "", "## Terms", ""];
  entries.forEach(entry => {
    lines.push(`### ${entry.term || "(no term)"}`);
    lines.push(`- Meaning: ${entry.meaning || ""}`);
    lines.push(`- Usage: ${entry.usage || ""}`);
    lines.push("");
  });
  return lines.join("\n");
}
function clearForm() {
  selectedCandidateKey = "";
  els.termInput.value = "";
  els.meaningInput.value = "";
  els.usageInput.value = "";
  els.selectedInfo.textContent = "선택 없음";
  renderCandidates();
}
function getVisibleCandidates() {
  const existing = new Set(entries.flatMap(entry => [entry.term, ...(entry.aliases || [])].map(normalize)).filter(Boolean));
  const query = normalize(els.candidateSearch.value);
  const mode = els.candidateMode.value;
  const filtered = candidates.filter(item => {
    if (candidateKeys(item).some(key => existing.has(key))) return false;
    if (mode === "needs" && item.needsUserGuide === false) return false;
    if (mode === "auto" && item.needsUserGuide !== false) return false;
    if (mode === "weak" && item.aiJudgement !== "빈도 낮음") return false;
    if (!query) return true;
    const haystack = [
      item.term,
      ...(item.aliases || []),
      ...(item.examples || []),
      ...(item.reasons || []),
    ].join(" ").toLowerCase();
    return haystack.includes(query);
  });
  const sortBy = els.sortSelect.value;
  filtered.sort((a, b) => {
    if (sortBy === "term") return String(a.term || "").localeCompare(String(b.term || ""), "ko");
    if (sortBy === "count") return (Number(b.count || 0) - Number(a.count || 0)) || String(a.term || "").localeCompare(String(b.term || ""), "ko");
    return (Number(b.score || 0) - Number(a.score || 0)) || (Number(b.count || 0) - Number(a.count || 0));
  });
  return filtered;
}
function renderCandidates() {
  renderedCandidates = getVisibleCandidates();
  const needsTotal = candidates.filter(x => x.needsUserGuide !== false).length;
  const autoTotal = candidates.length - needsTotal;
  const weakTotal = candidates.filter(x => x.aiJudgement === "빈도 낮음").length;
  els.candidateCount.textContent = `${renderedCandidates.length} 표시 / 확인필요 ${needsTotal} / 빈도낮음 ${weakTotal} / AI판단 ${autoTotal}`;
  if (!renderedCandidates.length) {
    els.candidateTableBody.innerHTML = '<tr><td colspan="7" class="empty">현재 후보 없음</td></tr>';
    return;
  }
  els.candidateTableBody.innerHTML = renderedCandidates.map((item, idx) => {
    const key = normalize(item.term);
    const aliases = (item.aliases || []).slice(0, 8);
    const examples = (item.examples || []).slice(0, 3);
    const reasons = (item.reasons || []).slice(0, 3);
    const selected = key && key === selectedCandidateKey ? " selected" : "";
    return `<tr class="${selected}" data-candidate-index="${idx}">
      <td class="term-cell">${esc(item.term)}</td>
      <td class="num">${Number(item.count || 0).toLocaleString()}</td>
      <td class="num">${Number(item.score || 0).toLocaleString()}</td>
      <td><span class="badge ${item.needsUserGuide === false ? "auto" : "need"}">${esc(item.aiJudgement || "")}</span></td>
      <td><div class="chips">${aliases.map(x => `<span class="chip">${esc(x)}</span>`).join("") || '<span class="muted small">-</span>'}</div></td>
      <td>
        <div>${examples.map(esc).join("<br>") || '<span class="muted">-</span>'}</div>
        <div class="muted small">${reasons.map(esc).join(", ")}</div>
      </td>
      <td><button type="button" data-select-candidate="${idx}">쓰기</button></td>
    </tr>`;
  }).join("");
}
function renderTerms() {
  els.termCount.textContent = `${entries.length}`;
  els.termList.innerHTML = entries.map((entry, idx) => `
    <div class="item">
      <b>${esc(entry.term)}</b>
      <div>${esc(entry.meaning).replace(/\n/g, "<br>")}</div>
      <div class="muted small">${esc(entry.usage).replace(/\n/g, "<br>")}</div>
      <button data-remove="${idx}">삭제</button>
    </div>`).join("") || '<div class="muted">등록된 용어 없음</div>';
}
function render() {
  renderTerms();
  renderCandidates();
}
function selectCandidate(item) {
  if (!item) return;
  selectedCandidateKey = normalize(item.term);
  els.termInput.value = item.term || "";
  els.meaningInput.value = "";
  const lines = [];
  if ((item.aliases || []).length) lines.push(`유사표현: ${(item.aliases || []).join(", ")}`);
  if ((item.reasons || []).length) lines.push(`후보 근거: ${(item.reasons || []).join(", ")}`);
  if ((item.examples || []).length) {
    lines.push("");
    lines.push("예시 데이터:");
    lines.push(...(item.examples || []));
  }
  els.usageInput.value = lines.join("\n");
  els.selectedInfo.textContent = item.term || "선택됨";
  renderCandidates();
  els.meaningInput.focus();
}
async function saveMd() {
  const md = buildMarkdown();
  if (!dirHandle) {
    downloadMd();
    els.saveStatus.textContent = "폴더 미연결 상태라 다운로드로 저장했습니다.";
    return;
  }
  try {
    const sampleDir = await getSampleDir();
    const fileHandle = await sampleDir.getFileHandle("ai_term_guidance.md", { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(md);
    await writable.close();
    els.saveStatus.textContent = "ai_term_guidance.md 저장 완료";
  } catch {
    els.saveStatus.textContent = "저장 실패. 폴더 권한을 다시 연결하세요.";
  }
}
function downloadMd() {
  const blob = new Blob([buildMarkdown()], { type: "text/markdown;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "ai_term_guidance.md";
  a.click();
  URL.revokeObjectURL(a.href);
}
async function getSampleDir() {
  try { return await dirHandle.getDirectoryHandle("sample_ready"); } catch { return dirHandle; }
}
async function connectFolder() {
  if (!window.showDirectoryPicker) {
    els.saveStatus.textContent = "이 브라우저는 폴더 저장 API를 지원하지 않습니다.";
    return;
  }
  dirHandle = await window.showDirectoryPicker({ id: "ai-batch-result-root", mode: "readwrite" });
  els.saveStatus.textContent = "폴더 연결됨. 저장하면 ai_term_guidance.md에 반영됩니다.";
}
document.getElementById("connectBtn").addEventListener("click", connectFolder);
document.getElementById("saveBtn").addEventListener("click", saveMd);
document.getElementById("downloadBtn").addEventListener("click", downloadMd);
document.getElementById("clearBtn").addEventListener("click", clearForm);
document.getElementById("addBtn").addEventListener("click", () => {
  const entry = { term: els.termInput.value.trim(), meaning: els.meaningInput.value.trim(), usage: els.usageInput.value.trim() };
  if (!entry.term && !entry.meaning) return;
  entries.unshift(entry);
  clearForm();
  storeEntries();
  render();
  if (dirHandle) saveMd();
});
els.candidateSearch.addEventListener("input", renderCandidates);
els.candidateMode.addEventListener("change", renderCandidates);
els.sortSelect.addEventListener("change", renderCandidates);
els.candidateTableBody.addEventListener("click", event => {
  const row = event.target.closest("tr[data-candidate-index]");
  if (!row) return;
  const item = renderedCandidates[Number(row.dataset.candidateIndex)];
  selectCandidate(item);
});
els.termList.addEventListener("click", event => {
  const btn = event.target.closest("[data-remove]");
  if (!btn) return;
  entries.splice(Number(btn.dataset.remove), 1);
  storeEntries();
  render();
  if (dirHandle) saveMd();
});
render();
</script>
</body>
</html>
"""


def write_term_glossary_html(sample_dir: Path) -> Path:
    terms_path = sample_dir / "ai_term_guidance.md"
    ensure_term_guidance_file(terms_path)
    terms_json = json.dumps(parse_term_guidance(terms_path), ensure_ascii=False).replace("</", "<\\/")
    candidates_json = json.dumps(build_term_candidates(sample_dir), ensure_ascii=False).replace("</", "<\\/")
    html_text = (
        TERM_GLOSSARY_TEMPLATE
        .replace("__TERMS_STATE__", terms_json)
        .replace("__CANDIDATES_STATE__", candidates_json)
    )
    out_path = sample_dir / "ai_term_glossary.html"
    out_path.write_text(html_text, encoding="utf-8")
    return out_path


def build_state(batch_dir: Path, sample_dir: Path) -> dict[str, Any]:
    rows = read_json(sample_dir / "demo_index.json") or []
    model_pairs = []
    map_text = read_text(sample_dir / "model_mapping_conditions.csv")
    for line in map_text.splitlines()[1:]:
        if not line.strip():
            continue
        parts = [part.strip() for part in line.split(",", 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            model_pairs.append(parts)
    return {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "paths": {
            "batchDir": str(batch_dir),
            "sampleDir": str(sample_dir),
        },
        "summary": read_json(batch_dir / "classification_summary.json") or {},
        "progress": read_jsonl_tail(batch_dir / "classification_progress.jsonl"),
        "rows": rows if isinstance(rows, list) else [],
        "modelPairs": model_pairs,
        "savedRequests": parse_prompt_requests(sample_dir / "prompt_update_requests.md"),
    }


def write_html(batch_dir: Path, sample_dir: Path, file_name: str) -> Path:
    sample_dir.mkdir(parents=True, exist_ok=True)
    state_json = json.dumps(build_state(batch_dir, sample_dir), ensure_ascii=False).replace("</", "<\\/")
    html = HTML_TEMPLATE.replace("__INITIAL_STATE__", state_json)
    out_path = sample_dir / file_name
    out_path.write_text(html, encoding="utf-8")
    ensure_prompt_request_file(sample_dir / "prompt_update_requests.md")
    write_term_glossary_html(sample_dir)
    append_plan_note(sample_dir / "DEMO_APPLY_PLAN.md", file_name)
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-dir", type=Path, default=DEFAULT_BATCH_DIR)
    parser.add_argument("--sample-dir", type=Path, default=DEFAULT_SAMPLE_DIR)
    parser.add_argument("--file-name", default="ai_batch_control.html")
    args = parser.parse_args()
    out_path = write_html(args.batch_dir, args.sample_dir, args.file_name)
    print(json.dumps({"html": str(out_path)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
