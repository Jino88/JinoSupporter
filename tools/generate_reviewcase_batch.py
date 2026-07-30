"""Generate source-backed ReviewCase drafts for MicroSpeaker Excel files.

This is a batch pre-analysis step. It does not modify the MicroSpeaker SQLite
database and it does not overwrite manually confirmed drafts. The output is
intended for AI verification and user review.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DB = REPO_ROOT.parent / "MicroSpeaker_ProductTech_DB" / "db" / "InputDataFinish.sqlite"
DEFAULT_AUDIT = REPO_ROOT / "REVIEWCASE_AI_AUDIT_DECISIONS.md"
DEFAULT_OUTPUT = REPO_ROOT / "REVIEWCASE_AI_DRAFTS" / "batch"
MANUAL_DRAFT_DIR = REPO_ROOT / "REVIEWCASE_AI_DRAFTS"

MAX_PAIR_CANDIDATES = 300
MAX_METRIC_CANDIDATES = 300
MAX_MEASUREMENT_CANDIDATES = 300
MAX_CONTEXT_ROWS = 80
MAX_SUBRESULTS_PER_OUTCOME = 60
MAX_MODEL_CANDIDATES = 12
MAX_MODEL_EVIDENCE = 8


DOMAIN_TERMS = {
    "supplier": ["supplier", "vendor", "doojin", "myungjin", "baotou", "boutou", "glonics", "cmg"],
    "material": ["material", "raw material", "mtr", "yk", "yoke", "film", "magnet"],
    "coating": ["coating", "plating", "tin", "polish", "non polish"],
    "equipment": ["machine", "equipment", "m/c", "mc ", "new mc", "repair", "awf"],
    "jig": ["jig", "fixture"],
    "mold": ["mold", "mould"],
    "process condition": ["condition", "method", "press", "dry", "uv", "plasma", "bonding", "assembly", "assy", "ass'y"],
    "dimension/spec": ["dimension", "spec", "thickness", "cutting", "height", "width"],
    "lot": [" lot", "lot ", "lot date"],
}

OUTCOME_TERMS = [
    ("measurement", ["tension", "strength", "pull", "spl", "thd", "dcr", "imp", "f0", "dimension", "gauss"]),
    ("function defect", ["function", "hearing", "sound", "rub", "buzz", "noise"]),
    ("reliability", ["drop", "salt", "reliability", "packing"]),
    ("process defect", ["vision", "process", "ng", "defect", "decap", "bond", "spot", "air leak", "separate"]),
]

CONTEXT_TERMS = [
    "title", "purpose", "objective", "content", "condition", "standard", "spec",
    "result", "decision", "conclusion", "remark", "note", "problem", "reason",
    "before", "after", "normal", "test", "sample", "lot", "supplier", "material",
    "machine", "m/c", "jig", "base", "laser", "coating", "gauss", "tension",
    "function", "vision", "repair",
]

MODEL_PATTERNS = [
    re.compile(r"\bBRS[-\s]?\d{4,6}(?:[-\s]?(?:DT|GMI|TF))?\b", re.IGNORECASE),
    re.compile(r"\bTIU[-\s]?C11[-\s]?20(?:[-\s]?[LR])?\b", re.IGNORECASE),
    re.compile(r"\bC11[-\s]?20(?:[-\s]?[LR])?\b", re.IGNORECASE),
    re.compile(r"\bTIU[-\s]?L5S3[-\s]?01(?:[-\s]?[LR])?\b", re.IGNORECASE),
    re.compile(r"\bL5S3[-\s]?01(?:[-\s]?[LR])?\b", re.IGNORECASE),
    re.compile(r"\b(?:MSU[-\s]?)?L?20S15[-\s]?07(?:[-\s]?(?:DT|GMI))?\b", re.IGNORECASE),
    re.compile(r"\b(?:MSU[-\s]?)?20S15[-\s]?07\b", re.IGNORECASE),
    re.compile(r"\bMSU[-\s]?201507(?:[-\s]?DT)?\b", re.IGNORECASE),
    re.compile(r"\bTIM[-\s]?\d{3}(?:[-\s]?[A-Z])?\b", re.IGNORECASE),
]

CONFIDENCE_RANK = {"low": 1, "medium": 2, "high": 3}


def now_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    uri = "file:" + db_path.as_posix().replace("?", "%3f") + "?mode=ro"
    con = sqlite3.connect(uri, uri=True, timeout=60)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA busy_timeout=60000")
    return con


def read_audit_decisions(path: Path) -> dict[int, dict[str, str]]:
    decisions: dict[int, dict[str, str]] = {}
    if not path.exists():
        return decisions

    pattern = re.compile(
        r"^\|\s*(?P<id>\d+)\s*\|\s*(?P<decision>[^|]+)\|\s*(?P<note>[^|]+)\|\s*`?(?P<file>[^|`]+)`?\s*\|"
    )
    for line in path.read_text(encoding="utf-8").splitlines():
        match = pattern.match(line)
        if not match:
            continue
        file_id = int(match.group("id"))
        decision = match.group("decision").strip()
        decision_key = decision.strip("`").lower()
        if not (decision_key.startswith("keep") or decision_key.startswith("exclude")):
            continue
        decisions[file_id] = {
            "decision": decision,
            "note": match.group("note").strip(),
            "fileName": match.group("file").strip(),
        }
    return decisions


def compact(text: Any, max_len: int = 320) -> str:
    value = re.sub(r"\s+", " ", str(text or "")).strip()
    return value if len(value) <= max_len else value[: max_len - 3].rstrip() + "..."


def normalize_model_name(value: Any) -> str:
    raw = re.sub(r"\s+", " ", str(value or "").upper()).strip(" .,_;:/\\|()[]{}")
    if not raw:
        return ""

    raw = raw.replace("_", " ")
    raw = raw.replace("BRS ", "BRS-")
    raw = raw.replace("TIU ", "TIU-")
    raw = raw.replace("MSU ", "MSU-")
    raw = re.sub(r"\s*-\s*", "-", raw)

    brs = re.search(r"\bBRS-?(\d{4,6})(?:-?(?:DT|GMI|TF))?\b", raw)
    if brs:
        return f"BRS-{brs.group(1)}"

    tiu_c = re.search(r"\b(?:TIU-?)?C11-?20(?:-?([LR]))?\b", raw)
    if tiu_c:
        suffix = f"-{tiu_c.group(1)}" if tiu_c.group(1) else ""
        return f"TIU-C11-20{suffix}"

    tiu_l = re.search(r"\b(?:TIU-?)?L5S3-?01(?:-?([LR]))?\b", raw)
    if tiu_l:
        suffix = f"-{tiu_l.group(1)}" if tiu_l.group(1) else ""
        return f"TIU-L5S3-01{suffix}"

    l20 = re.search(r"\b(?:MSU-?)?L?20S15-?07\b", raw)
    if l20:
        return "L20S15-07"

    msu201507 = re.search(r"\bMSU-?201507(?:-?DT)?\b", raw)
    if msu201507:
        return "MSU-201507"

    tim = re.search(r"\bTIM-?(\d{3})(?:-?([A-Z]))?\b", raw)
    if tim:
        suffix = f"-{tim.group(2)}" if tim.group(2) else ""
        return f"TIM-{tim.group(1)}{suffix}"

    return ""


def split_model_field(value: Any) -> list[str]:
    models = []
    for part in re.split(r"[;,\n\r|]+", str(value or "")):
        normalized = normalize_model_name(part)
        if normalized and normalized not in models:
            models.append(normalized)
    return models


def is_ambiguous_model(model: str) -> bool:
    return bool(re.fullmatch(r"BRS-\d{4}", model or ""))


def add_model_candidate(
    candidates: dict[str, dict[str, Any]],
    model: str,
    source: str,
    evidence: str,
    confidence: str,
    evidence_row: str = "",
) -> None:
    normalized = normalize_model_name(model)
    if not normalized:
        return

    row = candidates.setdefault(
        normalized,
        {
            "model": normalized,
            "confidence": confidence,
            "sources": [],
            "evidence": [],
            "evidenceRows": [],
            "ambiguous": is_ambiguous_model(normalized),
        },
    )
    if CONFIDENCE_RANK.get(confidence, 0) > CONFIDENCE_RANK.get(row.get("confidence", "low"), 0):
        row["confidence"] = confidence
    if source and source not in row["sources"]:
        row["sources"].append(source)
    if evidence and evidence not in row["evidence"] and len(row["evidence"]) < MAX_MODEL_EVIDENCE:
        row["evidence"].append(compact(evidence, 180))
    if evidence_row and evidence_row not in row["evidenceRows"] and len(row["evidenceRows"]) < MAX_MODEL_EVIDENCE:
        row["evidenceRows"].append(evidence_row)


def add_model_candidates_from_text(
    candidates: dict[str, dict[str, Any]],
    text: str,
    source: str,
    confidence: str,
    evidence_row: str = "",
) -> None:
    for pattern in MODEL_PATTERNS:
        for match in pattern.finditer(text or ""):
            add_model_candidate(candidates, match.group(0), source, text, confidence, evidence_row)


def build_model_review(file_row: sqlite3.Row, rows: list[sqlite3.Row]) -> dict[str, Any]:
    candidates: dict[str, dict[str, Any]] = {}
    source_models = str(file_row["models"] or "")
    for model in split_model_field(source_models):
        add_model_candidate(candidates, model, "files.models", source_models, "high")

    add_model_candidates_from_text(candidates, str(file_row["file_name"] or ""), "file_name", "medium")
    add_model_candidates_from_text(candidates, str(file_row["path"] or ""), "path", "low")
    add_model_candidates_from_text(candidates, str(file_row["dataset"] or ""), "dataset", "low")

    for row in rows[:80]:
        text = str(row["row_text"] or "")
        if not text.strip():
            continue
        add_model_candidates_from_text(
            candidates,
            text,
            "sheet_rows",
            "medium" if row["row_number"] <= 12 else "low",
            row_ref(row["sheet_name"], row["row_number"]),
        )

    ordered = sorted(
        candidates.values(),
        key=lambda item: (
            -CONFIDENCE_RANK.get(str(item.get("confidence") or "low"), 0),
            str(item.get("model") or ""),
        ),
    )[:MAX_MODEL_CANDIDATES]
    candidate_models = [str(item["model"]) for item in ordered]
    exact_candidates = [model for model in candidate_models if not is_ambiguous_model(model)]

    selected_models: list[str] = []
    mapping_status = "missing"
    confidence = "low"
    question = "이 엑셀의 기준 모델명을 확인해 주세요. 후보가 없으므로 원본 파일명/제목행을 보고 모델명을 입력해야 합니다."

    if len(exact_candidates) == 1 and len(candidate_models) == 1:
        selected_models = exact_candidates
        mapping_status = "confirmed"
        confidence = str(ordered[0].get("confidence") or "medium")
        question = ""
    elif len(exact_candidates) == 1 and all(is_ambiguous_model(model) or model == exact_candidates[0] for model in candidate_models):
        selected_models = exact_candidates
        mapping_status = "confirmed"
        confidence = "medium"
        question = ""
    elif candidate_models:
        mapping_status = "needs_user_mapping"
        confidence = "medium" if exact_candidates else "low"
        question = (
            "이 엑셀의 기준 모델명을 확정해 주세요. 후보: "
            + ", ".join(candidate_models[:8])
            + ". 여러 모델을 함께 다루는 파일이면 모두 알려주세요."
        )

    evidence_rows: list[str] = []
    for item in ordered:
        for evidence_row in item.get("evidenceRows", []) or []:
            if evidence_row not in evidence_rows:
                evidence_rows.append(evidence_row)

    return {
        "sourceModels": source_models,
        "selectedModels": selected_models,
        "mappingStatus": mapping_status,
        "confidence": confidence,
        "candidates": ordered,
        "evidenceRows": evidence_rows[:MAX_MODEL_EVIDENCE],
        "question": question,
        "mappingPrompt": {
            "task": "Confirm the canonical model name(s) for this Excel workbook before the ReviewCase is approved.",
            "fileId": int(file_row["file_id"]),
            "fileName": file_row["file_name"],
            "sourceModels": source_models,
            "candidateModels": candidate_models[:8],
            "question": question,
        },
    }


def normalize_key(text: Any) -> str:
    value = re.sub(r"[^a-z0-9]+", " ", str(text or "").lower())
    return re.sub(r"\s+", " ", value).strip()


def split_condition(condition: str) -> tuple[str, list[str]]:
    parts = [p.strip() for p in str(condition or "").split("|") if p.strip()]
    if not parts:
        return "", []
    return parts[0], parts[1:]


def infer_domains(*texts: str) -> list[str]:
    haystack = " ".join(t or "" for t in texts).lower()
    domains = []
    for domain, terms in DOMAIN_TERMS.items():
        if any(term in haystack for term in terms):
            domains.append(domain)
    return domains[:5] or ["unknown"]


def infer_outcome_domain(*texts: str) -> str:
    haystack = " ".join(t or "" for t in texts).lower()
    for domain, terms in OUTCOME_TERMS:
        if any(term in haystack for term in terms):
            return domain
    return "unknown"


def aggregate_judgement(effects: list[str], has_measurement: bool = False) -> str:
    values = {str(e or "").upper() for e in effects if str(e or "").strip()}
    if not values:
        return "not_judged" if has_measurement else "needs_review"
    if values == {"NO_CHANGE"}:
        return "no_change"
    if values == {"IMPROVED"}:
        return "improved"
    if values == {"WORSENED"}:
        return "worse"
    return "mixed"


def row_ref(sheet: str, row_number: int) -> str:
    return f"{sheet}!{row_number}" if sheet else str(row_number)


def parse_evidence_refs(evidence: str) -> list[str]:
    refs = []
    for match in re.finditer(r"(?P<sheet>[^!;\r\n]+)!(?P<row>\d+)", evidence or ""):
        sheet = match.group("sheet").strip()
        sheet = re.sub(r"^(?:vs|and|or|,|\(|\)|/|-)+\s*", "", sheet, flags=re.IGNORECASE).strip()
        refs.append(row_ref(sheet, int(match.group("row"))))
    return sorted(set(refs), key=lambda x: (x.rsplit("!", 1)[0], int(x.rsplit("!", 1)[1]) if "!" in x else 0))


def rate_percent(value: Any) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number * 100.0 if abs(number) <= 1.5 else number


def load_files(con: sqlite3.Connection) -> list[sqlite3.Row]:
    return list(
        con.execute(
            """
            SELECT f.file_id, f.dataset, f.path, f.file_name, f.extension, f.status, f.error,
                   f.sheet_count, f.sheet_names, f.models, f.categories, f.dates_found,
                   f.structure_family, f.structure_confidence, f.term_summary,
                   f.metric_candidate_count, f.measurement_stat_count, f.comparison_pair_count,
                   (SELECT COUNT(*) FROM sheet_rows sr WHERE sr.file_id=f.file_id) AS sheet_row_count,
                   (SELECT COUNT(*) FROM sheet_cells sc WHERE sc.file_id=f.file_id) AS sheet_cell_count
            FROM files f
            ORDER BY f.file_id
            """
        )
    )


def load_sheet_rows(con: sqlite3.Connection, file_id: int) -> list[sqlite3.Row]:
    return list(
        con.execute(
            """
            SELECT sheet_name, row_number, non_empty_count, row_text
            FROM sheet_rows
            WHERE file_id=?
            ORDER BY sheet_name, row_number
            """,
            (file_id,),
        )
    )


def load_pairs(con: sqlite3.Connection, file_id: int) -> list[sqlite3.Row]:
    return list(
        con.execute(
            """
            SELECT pair_id, table_title, compare_item, control_condition, test_condition,
                   control_input, control_ng, control_rate, test_input, test_ng, test_rate,
                   delta_rate, improvement_rate, effect_direction, evidence, pair_confidence
            FROM comparison_pairs
            WHERE file_id=?
            ORDER BY pair_id
            LIMIT ?
            """,
            (file_id, MAX_PAIR_CANDIDATES),
        )
    )


def load_metrics(con: sqlite3.Connection, file_id: int) -> list[sqlite3.Row]:
    return list(
        con.execute(
            """
            SELECT metric_id, sheet_name, row_number, table_title, condition_label,
                   input_qty, ok_qty, ng_qty, ng_rate, detail, raw_row, parse_confidence
            FROM metric_candidates
            WHERE file_id=?
            ORDER BY metric_id
            LIMIT ?
            """,
            (file_id, MAX_METRIC_CANDIDATES),
        )
    )


def load_measurements(con: sqlite3.Connection, file_id: int) -> list[sqlite3.Row]:
    return list(
        con.execute(
            """
            SELECT stat_id, sheet_name, row_number, item_label, condition_label, spec,
                   min_value, max_value, avg_value, sample_count, violation_count,
                   raw_row, parse_confidence
            FROM measurement_stats
            WHERE file_id=?
            ORDER BY stat_id
            LIMIT ?
            """,
            (file_id, MAX_MEASUREMENT_CANDIDATES),
        )
    )


def context_rows(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
    out = []
    for row in rows:
        text = str(row["row_text"] or "")
        lower = text.lower()
        if not text.strip():
            continue
        if len(out) < 12 or any(term in lower for term in CONTEXT_TERMS):
            out.append(
                {
                    "rowId": row_ref(row["sheet_name"], row["row_number"]),
                    "sheetName": row["sheet_name"],
                    "rowNumber": row["row_number"],
                    "rowText": compact(text, 420),
                }
            )
        if len(out) >= MAX_CONTEXT_ROWS:
            break
    return out


def title_from_rows(file_name: str, rows: list[sqlite3.Row]) -> str:
    for row in rows[:20]:
        text = str(row["row_text"] or "")
        if "title" in text.lower() or len(text.strip()) > 20:
            cleaned = compact(text.replace("| TITLE |", "TITLE |"), 220)
            cleaned = re.sub(r"^.*?\bTITLE\s*\|\s*", "", cleaned, flags=re.IGNORECASE)
            if cleaned:
                return cleaned
    return Path(file_name).stem.replace("_clean", "")


def purpose_from_rows(rows: list[sqlite3.Row]) -> str:
    for idx, row in enumerate(rows):
        text = str(row["row_text"] or "")
        if "purpose" not in text.lower():
            continue
        snippets = []
        for next_row in rows[idx : idx + 5]:
            candidate = compact(next_row["row_text"], 260)
            if candidate:
                snippets.append(candidate)
        return " / ".join(snippets)
    return ""


def build_changed_factor(file_row: sqlite3.Row, rows: list[sqlite3.Row], pairs: list[sqlite3.Row], metrics: list[sqlite3.Row], measurements: list[sqlite3.Row]) -> dict[str, Any]:
    controls = Counter()
    tests = Counter()
    subgroups = Counter()

    for pair in pairs:
        control, control_sub = split_condition(pair["control_condition"])
        test, test_sub = split_condition(pair["test_condition"])
        if control:
            controls[control] += 1
        if test:
            tests[test] += 1
        for item in control_sub + test_sub:
            subgroups[item] += 1

    for metric in metrics:
        condition, detail = split_condition(metric["condition_label"])
        if condition:
            if re.search(r"\bnormal\b|before", condition, re.IGNORECASE):
                controls[condition] += 1
            else:
                tests[condition] += 1
        for item in detail:
            subgroups[item] += 1

    baseline = ", ".join(label for label, _ in controls.most_common(4)) or "baseline/normal condition not explicit"
    changed = ", ".join(label for label, _ in tests.most_common(8)) or "test/changed condition not explicit"

    text_blob = " ".join(
        [
            file_row["file_name"] or "",
            file_row["term_summary"] or "",
            " ".join(str(p["table_title"] or "") for p in pairs[:20]),
            " ".join(str(m["table_title"] or "") for m in metrics[:20]),
            " ".join(str(s["item_label"] or "") for s in measurements[:20]),
        ]
    )
    domains = infer_domains(text_blob)

    evidence = []
    for row in rows[:15]:
        rid = row_ref(row["sheet_name"], row["row_number"])
        text = str(row["row_text"] or "").lower()
        if any(term in text for term in ["title", "purpose", "content", "condition", "normal", "test", "supplier", "material"]):
            evidence.append(rid)
    for pair in pairs[:20]:
        evidence.extend(parse_evidence_refs(pair["evidence"]))
    for metric in metrics[:20]:
        evidence.append(row_ref(metric["sheet_name"], metric["row_number"]))
    for stat in measurements[:20]:
        evidence.append(row_ref(stat["sheet_name"], stat["row_number"]))

    changed_factor_text = f"Condition comparison: {baseline} vs {changed}"
    if "baseline/normal condition not explicit" in baseline and "test/changed condition not explicit" in changed:
        changed_factor_text = compact(Path(file_row["file_name"]).stem.replace("_clean", ""), 180)

    return {
        "changedFactorId": "cf-1",
        "changeDomain": domains,
        "changedFactor": changed_factor_text,
        "baselineCondition": baseline,
        "changedCondition": changed,
        "subgroupKeys": [label for label, _ in subgroups.most_common(12)],
        "evidenceRows": sorted(set(evidence))[:80],
    }


def build_pair_outcomes(pairs: list[sqlite3.Row]) -> list[dict[str, Any]]:
    grouped: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for pair in pairs:
        key = normalize_key(pair["table_title"] or pair["compare_item"] or "comparison")
        grouped[key].append(pair)

    outcomes = []
    for idx, group in enumerate(grouped.values(), start=1):
        first = group[0]
        evidence = []
        subresults = []
        effects = []
        for pair in group[:MAX_SUBRESULTS_PER_OUTCOME]:
            refs = parse_evidence_refs(pair["evidence"])
            evidence.extend(refs)
            effects.append(pair["effect_direction"])
            subresults.append(
                {
                    "pairId": pair["pair_id"],
                    "controlCondition": compact(pair["control_condition"], 180),
                    "testCondition": compact(pair["test_condition"], 180),
                    "control": {
                        "input": pair["control_input"],
                        "ng": pair["control_ng"],
                        "ratePercent": rate_percent(pair["control_rate"]),
                    },
                    "test": {
                        "input": pair["test_input"],
                        "ng": pair["test_ng"],
                        "ratePercent": rate_percent(pair["test_rate"]),
                    },
                    "deltaRatePercentPoint": rate_percent(pair["delta_rate"]),
                    "effectDirection": pair["effect_direction"],
                    "evidenceRows": refs,
                }
            )

        outcomes.append(
            {
                "outcomeId": f"pair-{idx}",
                "changedFactorId": "cf-1",
                "outcomeDomain": infer_outcome_domain(first["table_title"], first["compare_item"]),
                "outcomeMetric": compact(first["table_title"] or first["compare_item"], 220),
                "comparisonRows": sorted(set(evidence)),
                "judgement": aggregate_judgement(effects),
                "subResults": subresults,
                "limitations": [] if len(group) <= MAX_SUBRESULTS_PER_OUTCOME else ["subResults truncated"],
            }
        )
    return outcomes


def build_metric_outcomes(metrics: list[sqlite3.Row], existing_domains: set[str]) -> list[dict[str, Any]]:
    grouped: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for metric in metrics:
        key = normalize_key(metric["table_title"] or "metric")
        grouped[key].append(metric)

    outcomes = []
    for idx, group in enumerate(grouped.values(), start=1):
        first = group[0]
        metric_name = compact(first["table_title"] or "extracted metric table", 220)
        if normalize_key(metric_name) in existing_domains:
            continue
        rows = [row_ref(item["sheet_name"], item["row_number"]) for item in group[:MAX_SUBRESULTS_PER_OUTCOME]]
        outcomes.append(
            {
                "outcomeId": f"metric-{idx}",
                "changedFactorId": "cf-1",
                "outcomeDomain": infer_outcome_domain(metric_name),
                "outcomeMetric": metric_name,
                "comparisonRows": sorted(set(rows)),
                "judgement": "not_judged",
                "subResults": [
                    {
                        "metricId": item["metric_id"],
                        "condition": compact(item["condition_label"], 180),
                        "input": item["input_qty"],
                        "ok": item["ok_qty"],
                        "ng": item["ng_qty"],
                        "ratePercent": rate_percent(item["ng_rate"]),
                        "evidenceRows": [row_ref(item["sheet_name"], item["row_number"])],
                    }
                    for item in group[:MAX_SUBRESULTS_PER_OUTCOME]
                ],
                "limitations": ["metric candidates need AI grouping verification"],
            }
        )
    return outcomes


def build_measurement_outcomes(measurements: list[sqlite3.Row]) -> list[dict[str, Any]]:
    grouped: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for stat in measurements:
        key = normalize_key(stat["item_label"] or "measurement")
        grouped[key].append(stat)

    outcomes = []
    for idx, group in enumerate(grouped.values(), start=1):
        first = group[0]
        rows = [row_ref(item["sheet_name"], item["row_number"]) for item in group[:MAX_SUBRESULTS_PER_OUTCOME]]
        outcomes.append(
            {
                "outcomeId": f"measurement-{idx}",
                "changedFactorId": "cf-1",
                "outcomeDomain": "measurement",
                "outcomeMetric": compact(first["item_label"] or "measurement table", 220),
                "comparisonRows": sorted(set(rows)),
                "judgement": "not_judged",
                "subResults": [
                    {
                        "statId": item["stat_id"],
                        "condition": compact(item["condition_label"], 180),
                        "spec": compact(item["spec"], 120),
                        "min": item["min_value"],
                        "max": item["max_value"],
                        "avg": item["avg_value"],
                        "sampleCount": item["sample_count"],
                        "violationCount": item["violation_count"],
                        "evidenceRows": [row_ref(item["sheet_name"], item["row_number"])],
                    }
                    for item in group[:MAX_SUBRESULTS_PER_OUTCOME]
                ],
                "limitations": ["measurement candidates need AI grouping verification"],
            }
        )
    return outcomes


def verify_refs(draft: dict[str, Any], available_refs: set[str]) -> tuple[list[str], list[str]]:
    refs = []

    def walk(value: Any) -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                if key in {"evidenceRows", "comparisonRows"} and isinstance(child, list):
                    refs.extend(str(item) for item in child)
                else:
                    walk(child)
        elif isinstance(value, list):
            for item in value:
                walk(item)

    walk(draft)
    unique = sorted(set(ref for ref in refs if "!" in ref))
    missing = [ref for ref in unique if ref not in available_refs]
    return unique, missing


def build_draft(
    con: sqlite3.Connection,
    file_row: sqlite3.Row,
    audit: dict[str, str] | None,
    manual_draft_exists: bool,
) -> dict[str, Any]:
    file_id = int(file_row["file_id"])
    rows = load_sheet_rows(con, file_id)
    available_refs = {row_ref(row["sheet_name"], row["row_number"]) for row in rows}
    model_review = build_model_review(file_row, rows)

    base = {
        "sourceFileId": file_id,
        "sourceFile": file_row["file_name"],
        "sourcePath": file_row["path"],
        "sourceModels": model_review["sourceModels"],
        "resolvedModels": model_review["selectedModels"],
        "modelMappingStatus": model_review["mappingStatus"],
        "modelReview": model_review,
        "generatedAt": now_iso(),
        "generationMode": "generic extracted-data batch preanalysis",
        "manualDraftExists": manual_draft_exists,
        "auditDecision": audit,
        "sourceStats": {
            "sheetCount": file_row["sheet_count"],
            "sheetNames": file_row["sheet_names"],
            "sheetRowCount": file_row["sheet_row_count"],
            "sheetCellCount": file_row["sheet_cell_count"],
            "comparisonPairCount": file_row["comparison_pair_count"],
            "metricCandidateCount": file_row["metric_candidate_count"],
            "measurementStatCount": file_row["measurement_stat_count"],
        },
    }

    if audit and audit.get("decision", "").lower().startswith("exclude"):
        base.update(
            {
                "reviewCaseStatus": "excluded",
                "excludeReason": audit.get("note", "user audit decision"),
                "reviewCases": [],
                "verification": {"status": "passed", "checkedEvidenceRows": [], "issues": []},
            }
        )
        return base

    if not rows or int(file_row["sheet_cell_count"] or 0) == 0:
        base.update(
            {
                "reviewCaseStatus": "excluded",
                "excludeReason": "no extracted sheet row/cell evidence",
                "reviewCases": [],
                "verification": {"status": "passed", "checkedEvidenceRows": [], "issues": []},
            }
        )
        return base

    pairs = load_pairs(con, file_id)
    metrics = load_metrics(con, file_id)
    measurements = load_measurements(con, file_id)

    if not pairs and not metrics and not measurements:
        base.update(
            {
                "reviewCaseStatus": "needs_review",
                "excludeReason": "",
                "contextRows": context_rows(rows),
                "reviewCases": [],
                "verification": {
                    "status": "needs_review",
                    "checkedEvidenceRows": [],
                    "issues": ["no comparison_pairs, metric_candidates, or measurement_stats were extracted"],
                },
            }
        )
        return base

    changed_factor = build_changed_factor(file_row, rows, pairs, metrics, measurements)
    pair_outcomes = build_pair_outcomes(pairs)
    pair_metric_keys = {normalize_key(outcome["outcomeMetric"]) for outcome in pair_outcomes}
    outcomes = pair_outcomes
    outcomes.extend(build_metric_outcomes(metrics, pair_metric_keys))
    outcomes.extend(build_measurement_outcomes(measurements))

    all_evidence = sorted(
        set(changed_factor["evidenceRows"])
        | {ref for outcome in outcomes for ref in outcome.get("comparisonRows", [])}
    )

    review_case = {
        "reviewCaseId": f"ms-{file_id}-rc-1",
        "reviewTitle": title_from_rows(file_row["file_name"], rows),
        "reviewPurpose": purpose_from_rows(rows),
        "changedFactors": [changed_factor],
        "outcomes": outcomes,
        "evidenceRows": all_evidence[:300],
        "limitations": [
            "batch draft generated from extracted DB rows; AI must verify grouping before Ask AI uses it as final evidence",
            "candidate pairs/metrics/measurements are hints and may miss merged-cell or image-only context",
        ],
    }

    base.update(
        {
            "reviewCaseStatus": "needs_ai_verification",
            "contextRows": context_rows(rows),
            "reviewCases": [review_case],
        }
    )
    checked, missing = verify_refs(base, available_refs)
    issues = []
    if missing:
        issues.append(f"{len(missing)} cited evidence rows were not found in sheet_rows")
    if len(pairs) >= MAX_PAIR_CANDIDATES:
        issues.append("comparison_pairs truncated by batch limit")
    if len(metrics) >= MAX_METRIC_CANDIDATES:
        issues.append("metric_candidates truncated by batch limit")
    if len(measurements) >= MAX_MEASUREMENT_CANDIDATES:
        issues.append("measurement_stats truncated by batch limit")

    base["verification"] = {
        "status": "needs_review" if issues else "passed_for_source_row_existence",
        "checkedEvidenceRows": checked[:500],
        "missingEvidenceRows": missing,
        "issues": issues,
    }
    return base


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_summary(path: Path, manifest: dict[str, Any]) -> None:
    counts = manifest["counts"]
    lines = [
        "# ReviewCase Batch Summary",
        "",
        f"Generated at: {manifest['generatedAt']}",
        "",
        "## Counts",
        "",
    ]
    for key in sorted(counts):
        lines.append(f"- {key}: {counts[key]}")

    lines.extend(["", "## Outputs", ""])
    lines.append(f"- Manifest: `{manifest['manifestPath']}`")
    lines.append(f"- File drafts: `{manifest['filesDir']}`")
    lines.extend(
        [
            "",
            "## Notes",
            "",
            "- This batch does not modify the MicroSpeaker SQLite database.",
            "- User-confirmed manual drafts are not overwritten.",
            "- `needs_ai_verification` drafts must be verified by AI/user before Ask AI treats them as final ReviewCases.",
            "- `model_*` counts show whether the workbook model was confirmed, missing, or queued for user mapping.",
            "- Excluded files follow `REVIEWCASE_AI_AUDIT_DECISIONS.md` or lack citeable extracted row/cell evidence.",
        ]
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--audit", type=Path, default=DEFAULT_AUDIT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--limit", type=int, default=0, help="Process only the first N files for smoke tests.")
    args = parser.parse_args()

    db_path = args.db.resolve()
    output_dir = args.output.resolve()
    files_dir = output_dir / "files"
    output_dir.mkdir(parents=True, exist_ok=True)
    files_dir.mkdir(parents=True, exist_ok=True)

    audit_decisions = read_audit_decisions(args.audit)
    con = connect_readonly(db_path)
    file_rows = load_files(con)
    if args.limit > 0:
        file_rows = file_rows[: args.limit]

    entries = []
    counts: Counter[str] = Counter()
    generated_at = now_iso()

    for file_row in file_rows:
        file_id = int(file_row["file_id"])
        manual_draft_exists = (MANUAL_DRAFT_DIR / f"{file_id}.reviewcase-draft.json").exists()
        draft = build_draft(con, file_row, audit_decisions.get(file_id), manual_draft_exists)
        draft["batchGeneratedAt"] = generated_at
        draft_path = files_dir / f"{file_id}.reviewcase-draft.json"
        write_json(draft_path, draft)

        status = draft.get("reviewCaseStatus", "unknown")
        counts[str(status)] += 1
        counts["total"] += 1
        if manual_draft_exists:
            counts["manual_draft_exists"] += 1
        if draft.get("verification", {}).get("missingEvidenceRows"):
            counts["missing_evidence_refs"] += 1
        model_review = draft.get("modelReview") or {}
        model_status = str(model_review.get("mappingStatus") or "unknown")
        counts[f"model_{model_status}"] += 1

        entries.append(
            {
                "fileId": file_id,
                "fileName": file_row["file_name"],
                "status": status,
                "sourceModels": model_review.get("sourceModels", ""),
                "resolvedModels": model_review.get("selectedModels", []),
                "modelMappingStatus": model_status,
                "modelQuestion": model_review.get("question", ""),
                "manualDraftExists": manual_draft_exists,
                "draftPath": str(draft_path.relative_to(REPO_ROOT)),
                "comparisonPairCount": file_row["comparison_pair_count"],
                "metricCandidateCount": file_row["metric_candidate_count"],
                "measurementStatCount": file_row["measurement_stat_count"],
                "sheetRowCount": file_row["sheet_row_count"],
                "verificationStatus": draft.get("verification", {}).get("status", ""),
                "issueCount": len(draft.get("verification", {}).get("issues", [])),
            }
        )

    manifest_path = output_dir / "reviewcase_batch_manifest.json"
    summary_path = output_dir / "reviewcase_batch_summary.md"
    manifest = {
        "generatedAt": generated_at,
        "databasePath": str(db_path),
        "auditDecisionPath": str(args.audit.resolve()),
        "manifestPath": str(manifest_path.relative_to(REPO_ROOT)),
        "summaryPath": str(summary_path.relative_to(REPO_ROOT)),
        "filesDir": str(files_dir.relative_to(REPO_ROOT)),
        "counts": dict(counts),
        "entries": entries,
    }
    write_json(manifest_path, manifest)
    write_summary(summary_path, manifest)
    print(json.dumps({"summaryPath": str(summary_path), "manifestPath": str(manifest_path), "counts": dict(counts)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
