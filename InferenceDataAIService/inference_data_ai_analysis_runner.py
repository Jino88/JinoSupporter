from __future__ import annotations

"""Run the deterministic packet manifest -> verified analysis DB -> HTML dashboard pipeline."""

import argparse
import hashlib
import html
import json
import math
import re
import sqlite3
import sys
import uuid
from pathlib import Path

import inference_data_ai_cli as core

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def analysis_html(data: dict) -> str:
    """Render every persisted analysis value without assuming a workbook-specific schema."""
    report = data.get("report") or {}

    def pick(item: dict, *names: str, default: object = None) -> object:
        for name in names:
            if name in item and item[name] is not None:
                return item[name]
        return default

    def present(value: object) -> bool:
        return value is not None and not (isinstance(value, str) and not value.strip())

    def raw_text(value: object, default: str = "—") -> str:
        if not present(value):
            return default
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False, separators=(",", ": "))
        return str(value)

    def esc(value: object, default: str = "—") -> str:
        return html.escape(raw_text(value, default))

    def number(value: object) -> str:
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return raw_text(value)
        rounded = round(float(value), 2)
        if rounded == 0:
            return "0"
        return f"{rounded:,.2f}".rstrip("0").rstrip(".")

    def rounded_decimal_text(value: object, default: str = "—") -> str:
        text = raw_text(value, default)
        return re.sub(
            r"(?<![\w.])-?\d+\.\d+(?![\w.])",
            lambda match: number(float(match.group(0))),
            text,
        )

    static_texts = {
        "Complete raw Sample/Average/Max/Min values were recomputed from the selected packet. No acceptance limit or release decision was present.": "선택된 패킷의 원시 표본으로 최소·최대·평균을 재계산했습니다. 허용 기준이나 출하 판정은 제공되지 않았습니다.",
        "Preserve complete raw measurement observations from the selected workbook packet.": "선택된 워크북 패킷의 완전한 원시 측정 관측값을 보존합니다.",
        "Only complete Sample/Average/Max/Min tables represented in the selected packet.": "선택된 패킷에 완전하게 포함된 표본·평균·최대·최소 표만 대상으로 합니다.",
        "Recompute source-provided raw sample summaries without inferring acceptance.": "원본 표본 요약을 재계산하되, 허용 여부는 추정하지 않습니다.",
        "Minimum, maximum, and average recomputed from the exact selected sample sequence.": "선택된 원시 표본열로부터 재계산한 최소·최대·평균입니다.",
        "Average delta calculated only against the one source-labelled Normal/control cohort.": "원본에 명시된 단일 대조군과의 평균 차이만 계산했습니다.",
        "A single explicit Normal/control cohort was present; average deltas are shown as observations.": "명시된 단일 대조군이 있어 평균 차이를 관측값으로 표시합니다.",
        "No unambiguous single Normal/control cohort was present, so no cohort comparison was inferred.": "명확한 단일 대조군이 없어 조건 비교를 추정하지 않았습니다.",
        "No acceptance limit, specification, or release decision was supplied in the selected packet.": "선택된 패킷에는 허용 기준, 규격 또는 출하 판정이 제공되지 않았습니다.",
        "Average difference is an observed calculation only; no acceptance decision was present.": "평균 차이는 관측값으로만 계산했으며, 허용 판정은 제공되지 않았습니다.",
        "The packet was recorded as an observation only; it does not contain a source-backed acceptance decision.": "패킷은 관측값으로만 기록되었으며, 근거가 있는 허용 판정은 포함하지 않습니다.",
        "Record selected packet evidence without inferring unstated workbook meaning.": "선택된 패킷 정보를 기록하되, 명시되지 않은 워크북 의미는 추정하지 않습니다.",
        "Selected source packet evidence only.": "선택된 원본 패킷 정보만 포함합니다.",
        "Preserve the selected evidence while awaiting an explicit analysis basis.": "명시적인 분석 기준이 제공될 때까지 선택된 패킷 정보를 보존합니다.",
        "No explicit Normal/control comparison basis was available.": "명시적인 대조군 비교 기준이 제공되지 않았습니다.",
        "No complete Sample/Average/Max/Min table was available in the selected packet.": "선택된 패킷에 완전한 표본·평균·최대·최소 표가 없습니다.",
        "The selected packet is incomplete; omitted rows or cells must be reviewed in the source workbook.": "선택된 패킷이 불완전하므로, 생략된 행 또는 셀은 원본 워크북에서 검토해야 합니다.",
    }
    detail_labels = {
        "sampleCount": "표본 수 (N)",
        "sampleEvidenceRange": "표본 근거",
        "recomputedSummary": "재계산 요약",
        "sampleStandardDeviation": "표본 표준편차",
        "sampleStdDev": "표본 표준편차",
        "displayedSummaryReconciliation": "표시 요약 일치",
        "average": "평균",
        "min": "최소",
        "max": "최대",
        "range": "범위",
        "sampleRange": "범위",
        "rejectCount": "불량 수",
        "noise": "노이즈",
        "touch": "터치",
        "packetSelection": "패킷 선택 상태",
    }
    status_labels = {
        "VERIFIED": "검증됨",
        "CAN_USE": "사용 가능",
        "OK": "정상",
        "REFERENCE": "기준",
        "IMPROVED": "개선",
        "REJECTED": "반려",
        "CAN_NOT_USE": "사용 불가",
        "NEEDS_REVIEW": "검토 필요",
        "OBSERVED": "관측값",
        "OBSERVED_ONLY": "관측값만",
    }
    metric_type_labels = {
        "defect_rate": "불량률",
        "measurement_summary": "Raw Measurement statistics",
        "measurement_average_comparison": "Average comparison",
        "packet_evidence": "Packet evidence",
    }

    def localized_text(value: object, default: str = "—") -> str:
        text = raw_text(value, default)
        if text.startswith("[Fresh deterministic draft: force token="):
            return "[새 결정론적 초안: 강제 생성 토큰=" + text.removeprefix("[Fresh deterministic draft: force token=")
        if text == "Complete raw measurement statistics by cohort":
            return "Complete Raw Measurement Statistics"
        technical_source_text = {
            "Complete raw measurement observations",
            "Average comparison versus explicit Normal/control",
            "Explicit Normal/control label",
            "Source table row label",
            "No explicit Normal/control comparison basis was available.",
            "Packet-backed observation",
            "Selected packet evidence",
            "No inferred condition",
        }
        if text in technical_source_text:
            return rounded_decimal_text(text)
        return rounded_decimal_text(static_texts.get(text, text))

    def localized_multiline(value: object, default: str = "—") -> str:
        return "\n".join(localized_text(line, default) for line in raw_text(value, default).splitlines())

    def localized_html(value: object, default: str = "—") -> str:
        return html.escape(localized_multiline(value, default)).replace("\n", "<br>")

    def human_label(value: object) -> str:
        key = str(value or "")
        if key in detail_labels:
            return detail_labels[key]
        if key and (key == key.upper() or key.startswith("NG ")):
            return key
        name = str(value or "").replace("_", " ")
        name = "".join((" " if char.isupper() and index and name[index - 1].islower() else "") + char for index, char in enumerate(name))
        return " ".join(name.split()).capitalize() or "Value"

    def badge(status: object) -> str:
        value = raw_text(status, "NEEDS_REVIEW").upper()
        css = "good" if value in {"VERIFIED", "IMPROVED", "CAN_USE", "OK", "REFERENCE"} else "bad" if value in {"REJECTED", "CAN_NOT_USE"} else "review"
        return f"<span class='status {css}'>{html.escape(status_labels.get(value, value))}</span>"

    def cohort_heading(review: dict, value: dict, fallback: object) -> str:
        key = pick(value, "cohort", "cohort_key", "cohortKey", default=fallback)
        cohort = next(
            (item for item in review.get("cohorts", []) if isinstance(item, dict) and pick(item, "cohort_key", "cohortKey", "key") == key),
            {},
        )
        cohort_label = pick(value, "cohort_label", "cohortLabel")
        display = cohort_label if present(cohort_label) else pick(cohort, "label", default=key)
        condition = pick(cohort, "condition_text", "conditionText", "condition")
        return f"<strong class='condition-name'>{esc(display)}</strong>" + (
            f"<br><span class='condition'>{html.escape(localized_text(condition))}</span>" if present(condition) else ""
        )

    def details_html(details: object) -> str:
        if not isinstance(details, dict) or not details:
            return ""
        hidden_audit_keys = {
            "sampleCount", "sample_count", "n", "sampleSequence", "sampleValues", "observedSamples", "rawSamples",
            "sampleEvidenceRange", "recomputedSummary", "displayedSummaryReconciliation",
        }
        rendered = "".join(
            f"<div><dt>{esc(human_label(key))}</dt><dd>{esc(number(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else value)}</dd></div>"
            for key, value in details.items()
            if key not in hidden_audit_keys
        )
        return f"<dl class='breakdown'>{rendered}</dl>" if rendered else ""

    def value_html(value: dict) -> str:
        rows: list[str] = []
        details = pick(value, "details", default={})
        measurement_details = details if isinstance(details, dict) else {}
        value_text = pick(value, "value_text", "valueText")
        value_number = pick(value, "value_number", "valueNumber")
        if present(value_text) and not re.fullmatch(r"N\s*=\s*\d+", str(value_text).strip()):
            rows.append(f"<div>{esc(rounded_decimal_text(value_text))}</div>")
        elif present(value_number):
            rows.append(f"<div><span class='value-label'>값</span> {esc(number(value_number))}</div>")

        numerator = pick(value, "numerator")
        denominator = pick(value, "denominator")
        rate_ppm = pick(value, "rate_ppm", "ratePpm")
        rate_text = ""
        if present(numerator) and present(denominator):
            rate_text = f"{number(numerator)} / {number(denominator)}"
        elif present(numerator) or present(denominator):
            rate_text = f"{number(numerator)} / {number(denominator)}"
        if present(rate_ppm):
            rate_text = f"{rate_text} ({number(rate_ppm)} ppm)" if rate_text else f"{number(rate_ppm)} ppm"
        if rate_text:
            rows.append(f"<div><span class='value-label'>비율</span> {html.escape(rate_text)}</div>")

        recomputed = measurement_details.get("recomputedSummary") if isinstance(measurement_details.get("recomputedSummary"), dict) else {}
        measurement_fields = (
            ("N", pick(measurement_details, "sampleCount", "sample_count", "n")),
            ("최소", pick(value, "min_value", "minValue", "min")),
            ("최대", pick(value, "max_value", "maxValue", "max")),
            ("평균", pick(value, "average_value", "averageValue", "average", "avg_value", "avgValue")),
            ("표준편차", pick(recomputed, "sampleStandardDeviation", "sampleStdDev")),
            ("범위", pick(recomputed, "range", "sampleRange")),
        )
        measurements = []
        for title, measurement in measurement_fields:
            if present(measurement):
                measurements.append(f"<span><b>{title}</b> {esc(number(measurement))}</span>")
        if measurements:
            rows.append("<div class='measurements'>" + "".join(measurements) + "</div>")

        detail_rows = details_html(measurement_details)
        if detail_rows:
            rows.append(detail_rows)
        status = pick(value, "result_status", "resultStatus", "status")
        if present(status):
            rows.append(f"<div class='value-status'>{badge(status)}</div>")
        return "<div class='metric-value'>" + ("".join(rows) or "—") + "</div>"

    def metric_heading(metric: dict) -> str:
        metric_type = pick(metric, "metric_type", "metricType", "type")
        unit = pick(metric, "unit")
        localized_type = metric_type_labels.get(str(metric_type or ""), raw_text(metric_type, ""))
        descriptor = " · ".join(str(value) for value in (localized_type, raw_text(unit, "")) if present(value))
        spec = pick(metric, "spec_text", "specText", "spec")
        source_table = metric_source_table_metadata(metric)
        source_table_items = []
        if present(source_table.get("caption")):
            source_table_items.append(
                "<span><b>원본 표 제목</b> " + esc(source_table["caption"]) + "</span>"
            )
        if present(source_table.get("type")):
            source_table_items.append(
                "<span><b>유형</b> " + esc(source_table["type"]) + "</span>"
            )
        return (
            f"<strong>{html.escape(localized_text(pick(metric, 'label')))}</strong>"
            + (f"<br><span class='metric-type'>{html.escape(descriptor)}</span>" if descriptor else "")
            + (f"<br><span class='metric-type'>규격: {esc(rounded_decimal_text(spec))}</span>" if present(spec) else "")
            + ("<div class='source-table-meta'>" + "".join(source_table_items) + "</div>" if source_table_items else "")
        )

    def comparison_html(comparison: dict) -> str:
        parts: list[str] = []
        for field in ("summary_text", "summary"):
            summary = pick(comparison, field)
            if present(summary):
                parts.append(f"<div>{html.escape(localized_text(summary))}</div>")
                break
        calculation = pick(comparison, "calculation_text", "calculation")
        if present(calculation):
            parts.append(f"<div class='calculation'>{esc(rounded_decimal_text(calculation))}</div>")
        details = details_html(pick(comparison, "details", default={}))
        if details:
            parts.append(details)
        return "".join(parts) or "—"

    sections = [
        f"""<header><h1>{esc(pick(report, 'title'))}</h1><p>{localized_html(pick(report, 'summary'))}</p><p class='source-file'>원본 파일: <code>{esc(pick(report, 'fileName', 'file_name', 'sourcePath', 'source_path'))}</code></p></header><main>
<section><h2>분석 요약</h2><div class='table-wrap'><table><caption>보고서 범위와 판정</caption><thead><tr><th>항목</th><th>내용</th><th>상태</th></tr></thead><tbody>
<tr><th>목적</th><td>{localized_html(pick(report, 'purpose'))}</td><td>{badge(pick(report, 'status'))}</td></tr>
<tr><th>분석 범위</th><td>{localized_html(pick(report, 'scope'))}</td><td>{badge(pick(report, 'decision'))}</td></tr>
</tbody></table></div></section>"""
    ]
    for review in data.get("reviews", []):
        if not isinstance(review, dict):
            continue
        rows: list[str] = []
        for metric in review.get("metrics", []):
            if not isinstance(metric, dict):
                continue
            values = {
                pick(value, "cohort", "cohort_key", "cohortKey"): value
                for value in metric.get("values", [])
                if isinstance(value, dict)
            }
            comparisons = [item for item in metric.get("comparisons", []) if isinstance(item, dict)]
            if comparisons:
                for comparison_index, comparison in enumerate(comparisons):
                    test_key = pick(comparison, "comparedCohort", "compared_cohort_key", "comparedCohortKey")
                    control_key = pick(comparison, "controlCohort", "control_cohort_key", "controlCohortKey")
                    test, control = values.get(test_key, {}), values.get(control_key, {})
                    heading = (
                        f"<th rowspan='{len(comparisons)}'>{metric_heading(metric)}</th>"
                        if comparison_index == 0 else ""
                    )
                    rows.append(
                        f"<tr>{heading}<td>{cohort_heading(review, test, test_key)}{value_html(test)}</td>"
                        f"<td>{cohort_heading(review, control, control_key)}{value_html(control)}</td><td>{comparison_html(comparison)}</td>"
                        f"<td>{badge(pick(comparison, 'status'))}</td></tr>"
                    )
            else:
                for value_index, (cohort_key, value) in enumerate(values.items()):
                    heading = (
                        f"<th rowspan='{len(values)}'>{metric_heading(metric)}</th>"
                        if value_index == 0 else ""
                    )
                    rows.append(
                        f"<tr>{heading}<td colspan='3'>{cohort_heading(review, value, cohort_key)}{value_html(value)}</td>"
                        f"<td>{badge(pick(value, 'result_status', 'resultStatus', 'status'))}</td></tr>"
                    )
        if not rows:
            rows.append("<tr><td colspan='5'>내보낸 지표 값이 없습니다.</td></tr>")
        conclusion_text = "<br>".join(
            localized_html(pick(item, "text", "conclusion_text", "conclusionText"))
            for item in review.get("conclusions", [])
            if isinstance(item, dict)
        ) or "생성된 결론이 없습니다."
        notes = review.get("notes")
        note_text = "<br>".join(localized_html(item) for item in notes if present(item)) if isinstance(notes, list) else ""
        notes_html = f"<div class='review-notes'><b>검토 메모</b><br>{note_text}</div>" if note_text else ""
        sections.append(
            f"<section><h2>상세 분석: {localized_html(pick(review, 'title'))}</h2>{notes_html}<div class='table-wrap'><table><caption>{localized_html(pick(review, 'summary', 'summary_text', 'summaryText'))}</caption>"
            "<thead><tr><th>지표 / 유형</th><th>시험 조건</th><th>대조 조건</th><th>차이 / 결과</th><th>상태</th></tr></thead><tbody>"
            + "".join(rows)
            + f"<tr class='highlight'><th>분석 결론</th><td colspan='4'>{conclusion_text}</td></tr>"
            + "</tbody></table></div></section>"
        )
    limitations = pick(report, "limitations", default=[])
    limits = "<br>".join(localized_html(item) for item in limitations) if isinstance(limitations, list) and limitations else "이 결과를 사용하기 전에 사람의 검토가 필요합니다."
    sections.append(
        f"<section><h2>최종 판정 및 검토 제한</h2><div class='table-wrap'><table><caption>근거 기반 종합 판단</caption><thead><tr><th>항목</th><th>판정</th><th>내용</th></tr></thead><tbody>"
        f"<tr><th>분석 결론</th><td>{badge(pick(report, 'decision'))}</td><td>{localized_html(pick(report, 'summary'))}</td></tr>"
        f"<tr><th>검토 제한</th><td>{badge(pick(report, 'status'))}</td><td>{limits}</td></tr></tbody></table></div></section></main>"
    )
    css = ":root{--bg:#f4f6fa;--panel:#fff;--line:#d8e0ea;--head:#eef2f6;--ink:#17202e;--muted:#667085;--green:#067647;--green-bg:#ecfdf3;--red:#b42318;--red-bg:#fef3f2;--amber:#b54708;--amber-bg:#fffaeb}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:'Segoe UI','Malgun Gothic',Arial,sans-serif;line-height:1.45}header{padding:18px 22px 14px;border-bottom:1px solid var(--line);background:var(--panel)}h1{margin:0;font-size:23px}header p{margin:5px 0 0;color:var(--muted);font-size:12px}main{width:min(1500px,100%);margin:0 auto;padding:16px 18px 30px}section{margin-bottom:16px}h2{margin:0 0 7px;font-size:16px}.table-wrap{overflow-x:auto;border:1px solid var(--line);border-radius:9px;background:var(--panel)}table{width:100%;border-collapse:collapse;font-size:12px}caption{padding:9px 10px;border-bottom:1px solid var(--line);color:#344054;background:#f8fafc;text-align:left;font-weight:800}th,td{padding:8px 9px;border-right:1px solid var(--line);border-bottom:1px solid var(--line);vertical-align:top;overflow-wrap:anywhere}th:last-child,td:last-child{border-right:0}thead th{color:#475467;background:var(--head);text-align:center}tbody th{min-width:150px;background:#f8fafc;text-align:left}.condition-name{display:block;color:#175cd3;font-size:11px}.condition,.metric-type,.calculation{color:var(--muted);font-size:10px;line-height:1.25}.source-table-meta{display:flex;flex-direction:column;gap:3px;margin-top:7px;padding-top:6px;border-top:1px solid var(--line);font-size:10px;font-weight:400}.source-table-meta b{color:var(--muted)}.metric-value{margin-top:5px}.value-label{color:var(--muted);font-weight:700}.measurements{display:flex;flex-wrap:wrap;gap:5px;margin-top:4px}.measurements span{padding:2px 5px;border-radius:4px;background:#f2f4f7}.breakdown{display:flex;flex-wrap:wrap;gap:4px;margin:5px 0 0}.breakdown div{display:flex;gap:4px;padding:2px 5px;border:1px solid var(--line);border-radius:4px;background:#fcfcfd}.breakdown dt{font-weight:700}.breakdown dd{margin:0}.value-status{margin-top:5px}.evidence{color:var(--muted);font-family:Consolas,monospace;font-size:10px}.source-file{font-family:Consolas,monospace;font-size:11px}.status{display:inline-block;padding:3px 7px;border-radius:999px;font-size:10px;font-weight:800}.good{color:var(--green);background:var(--green-bg)}.bad{color:var(--red);background:var(--red-bg)}.review{color:var(--amber);background:var(--amber-bg)}.highlight td,.highlight th{background:#fffdf5}@media(max-width:760px){header{padding:14px 14px 11px}h1{font-size:20px}main{padding:12px 10px 24px}section{margin-bottom:12px}.table-wrap{border-radius:7px}table{min-width:720px;font-size:12px}th,td{padding:7px 8px}tbody th{min-width:175px}.measurements{gap:4px}.source-table-meta{font-size:11px}}@media print{@page{size:A4 landscape;margin:8mm}body{background:#fff}header,main{width:100%;padding-left:0;padding-right:0}.table-wrap{overflow:visible}}"
    return "<!doctype html><html lang='ko'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>분석 대시보드</title><style>" + css + "</style></head><body>" + "".join(sections) + "</body></html>"


def normalize_manifest(data: dict) -> dict:
    """Keep supplied observation fields instead of silently losing them on import."""
    supported = {"cohort", "valueNumber", "valueText", "numerator", "denominator", "ratePpm", "min", "max", "average", "status", "details", "evidence"}
    for review in data.get("reviews", []):
        for metric in review.get("metrics", []):
            for value in metric.get("values", []):
                extras = {key: item for key, item in value.items() if key not in supported}
                if extras:
                    details = value.setdefault("details", {})
                    if isinstance(details, dict): details.update(extras)
                    if not value.get("valueText"):
                        value["valueText"] = ", ".join(f"{key}={item}" for key, item in extras.items())
                if value.get("ratePpm") is None and isinstance(value.get("ngRate"), (int, float)):
                    value["ratePpm"] = float(value["ngRate"]) * 1_000_000 if abs(float(value["ngRate"])) <= 1.5 else float(value["ngRate"])
    return data


def prepare_force_ai_draft(data: dict, draft_token: str) -> dict:
    """Give a freshly generated deterministic draft an identity that preserves curated reports."""
    report = data.get("report")
    if not isinstance(report, dict):
        raise ValueError("Fresh draft must include a report object before force-draft preparation.")
    original_key = str(report.get("key") or "").strip()
    if not original_key:
        raise ValueError("Fresh draft report.key is required before force-draft preparation.")
    report["key"] = f"{original_key}-force-ai-{draft_token}"
    scope = str(report.get("scope") or "").strip()
    report["scope"] = f"{scope}\n[Fresh deterministic draft: force token={draft_token}]".strip()
    artifacts = report.get("artifacts")
    if isinstance(artifacts, dict):
        artifacts.pop("html", None)
        artifacts.pop("markdown", None)
    return data


def is_runner_draft(report: dict) -> bool:
    """Only replace reports produced by this runner, never curated CLI reports."""
    return Path(str(report.get("manifest_path") or "")).name.startswith("workbook_")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def curated_html_path(service: Path, manifest: dict) -> Path | None:
    report = manifest.get("report", {})
    if not isinstance(report, dict):
        return None
    artifacts = report.get("artifacts", {})
    if not isinstance(artifacts, dict) or not artifacts.get("html"):
        return None
    candidate = Path(str(artifacts["html"]))
    candidate = candidate if candidate.is_absolute() else service / candidate
    try:
        candidate = candidate.resolve()
        candidate.relative_to((service / "outputs").resolve())
    except (OSError, ValueError):
        return None
    return candidate if candidate.is_file() else None


def verified_curated_report(
    service: Path,
    manifest_path: Path,
    source_path: Path,
    dataset: str,
) -> dict | None:
    """Find exactly one current VERIFIED DB report that was imported from this manifest."""
    matches: list[dict] = []
    expected_manifest = manifest_path.resolve()
    expected_source = source_path.resolve()
    for db_path in sorted((service / "outputs" / "universal-grid").glob("*.sqlite")):
        try:
            conn = sqlite3.connect(db_path)
            try:
                conn.row_factory = sqlite3.Row
                rows = conn.execute(
                    """
                    SELECT a.analysis_report_id, a.manifest_path, a.workbook_fingerprint AS report_fingerprint,
                           a.source_path, w.status AS workbook_status, w.fingerprint AS workbook_fingerprint
                    FROM analysis_reports a
                    JOIN workbooks w ON w.workbook_id=a.workbook_id
                    WHERE a.dataset=? AND a.source_path=? AND a.overall_status='VERIFIED'
                    """,
                    (dataset, str(expected_source)),
                ).fetchall()
            finally:
                conn.close()
        except sqlite3.Error:
            continue
        for row in rows:
            try:
                row_manifest = Path(str(row["manifest_path"] or "")).resolve()
            except OSError:
                continue
            if row_manifest != expected_manifest or row["workbook_status"] != "OK":
                continue
            if row["report_fingerprint"] != row["workbook_fingerprint"]:
                continue
            if str(row["workbook_fingerprint"]) != core.file_fingerprint(expected_source):
                continue
            matches.append({"database": str(db_path.resolve()), "analysisReportId": int(row["analysis_report_id"])})
    return matches[0] if len(matches) == 1 else None


def curated_reuse_for_source(
    service: Path,
    source: str,
    dataset: str,
) -> dict | None:
    """Select a single byte-identical, DB-verified curated baseline; never match by path alone."""
    requested_source = Path(source).resolve()
    if not requested_source.is_file():
        return None
    requested_hash = sha256_file(requested_source)
    matches: list[dict] = []
    for manifest_path in sorted((service / "outputs" / "analysis-manifests").glob("*_analysis.json")):
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            source_metadata = manifest.get("source", {})
            if not isinstance(source_metadata, dict):
                continue
            original_source = Path(str(source_metadata.get("sourcePath") or "")).resolve()
        except (OSError, ValueError, json.JSONDecodeError):
            continue
        if not isinstance(source_metadata, dict) or source_metadata.get("dataset") != dataset or not original_source.is_file():
            continue
        original_html = curated_html_path(service, manifest)
        verified_report = verified_curated_report(service, manifest_path, original_source, dataset)
        if not original_html or not verified_report:
            continue
        original_hash = sha256_file(original_source)
        if original_hash != requested_hash:
            continue
        matches.append(
            {
                "manifestPath": manifest_path.resolve(),
                "manifest": manifest,
                "originalSource": original_source,
                "originalHtml": original_html,
                "sha256": requested_hash,
                "verifiedReport": verified_report,
            }
        )
    return matches[0] if len(matches) == 1 else None


def curated_reuse_not_applicable(source: str) -> dict[str, str]:
    return {
        "status": "curated-reuse-not-applicable",
        "source": source,
        "reason": "No single VERIFIED curated baseline had both an existing dashboard and a SHA-256 byte-identical source.",
    }


def write_reused_curated_html(service: Path, reuse: dict, requested_source: Path, workbook_id: int) -> Path:
    """Create a new, visibly provenance-marked dashboard without changing a CLI artifact."""
    original_html = Path(reuse["originalHtml"])
    target = service / "outputs" / "analysis-rendered" / f"reused_curated_workbook_{workbook_id}_{reuse['manifestPath'].stem}.html"
    original = original_html.read_text(encoding="utf-8-sig")
    provenance = (
        "<section style='margin:12px 18px;padding:10px 12px;border:1px solid #b54708;border-radius:8px;"
        "background:#fffaeb;color:#7a2e0e;font:12px Segoe UI,Malgun Gothic,sans-serif'>"
        "<strong>큐레이션 기준 재사용</strong><br>"
        f"원본: {html.escape(str(reuse['originalSource']))}<br>"
        f"현재 파일: {html.escape(str(requested_source))}<br>"
        f"SHA-256 바이트 일치: <code>{html.escape(str(reuse['sha256']))}</code><br>"
        "이 화면은 원본 분석을 현재 파일이 원본이라고 주장하지 않으며, 동일 바이트 기준으로 재사용했음을 표시합니다."
        "</section>"
    )
    body_index = original.lower().find("<body")
    if body_index >= 0:
        body_end = original.find(">", body_index)
        rendered = original[: body_end + 1] + provenance + original[body_end + 1 :] if body_end >= 0 else provenance + original
    else:
        rendered = provenance + original
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(rendered, encoding="utf-8")
    return target


def rebind_curated_manifest(
    service: Path,
    reuse: dict,
    requested_source: Path,
    dataset: str,
    workbook: dict,
    reused_html: Path,
) -> tuple[Path, dict]:
    """Write a separate manifest bound to the current imported workbook and its provenance."""
    manifest = json.loads(json.dumps(reuse["manifest"], ensure_ascii=False))
    source = manifest.setdefault("source", {})
    source.update(
        {
            "dataset": dataset,
            "sourcePath": str(requested_source),
            "workbookId": int(workbook["workbook_id"]),
            "fingerprint": str(workbook["fingerprint"]),
        }
    )
    report = manifest.setdefault("report", {})
    provenance = {
        "mode": "sha256-byte-identical-curated-reuse",
        "originalSource": str(reuse["originalSource"]),
        "originalManifest": str(reuse["manifestPath"]),
        "originalCuratedHtml": str(reuse["originalHtml"]),
        "sha256": str(reuse["sha256"]),
        "verifiedReport": reuse["verifiedReport"],
    }
    scope = str(report.get("scope") or "").strip()
    notice = f"[큐레이션 기준 재사용] 원본={provenance['originalSource']}; SHA-256={provenance['sha256']}"
    report["scope"] = f"{scope}\n{notice}".strip()
    report["reuseProvenance"] = provenance
    artifacts = report.setdefault("artifacts", {})
    if isinstance(artifacts, dict):
        artifacts["html"] = str(reused_html)
        artifacts["originalCuratedHtml"] = provenance["originalCuratedHtml"]
    clone_path = service / "outputs" / "analysis-manifests" / "reused-curated" / f"workbook_{workbook['workbook_id']}_{reuse['manifestPath'].stem}_reused.json"
    clone_path.parent.mkdir(parents=True, exist_ok=True)
    clone_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return clone_path, provenance


def number_value(value: object) -> float | None:
    """Return a finite numeric cell/value without treating booleans as numbers."""
    if isinstance(value, bool):
        return None
    try:
        number = float(str(value).strip()) if isinstance(value, str) else float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def excel_column_label(column: int) -> str:
    label = ""
    while column > 0:
        column, remainder = divmod(column - 1, 26)
        label = chr(ord("A") + remainder) + label
    return label


SOURCE_TABLE_METADATA_NOTE_KIND = "source-table-metadata-v1"


def source_cell_text(cell: object) -> str:
    """Return one stored cell value as displayable source text without interpretation."""
    if not isinstance(cell, dict):
        return ""
    value = cell.get("value", cell.get("value_text", ""))
    return str(value).strip() if value is not None else ""


def cells_by_column(cells: object) -> dict[int, dict]:
    """Index packet/grid cell dictionaries by their original source column."""
    result: dict[int, dict] = {}
    if not isinstance(cells, list):
        return result
    for cell in cells:
        if not isinstance(cell, dict):
            continue
        column = number_value(cell.get("column", cell.get("col_number")))
        if column is not None and column > 0:
            result[int(column)] = cell
    return result


def complete_measurement_header_columns(cells: object) -> dict[str, int] | None:
    """Recognize the explicit Sample/Average/Max/Min header used by the safe raw-table parser."""
    headers: dict[str, int] = {}
    type_column: int | None = None
    for column, cell in cells_by_column(cells).items():
        header = source_cell_text(cell).casefold()
        if not header:
            continue
        headers[header] = column
        if header == "type" or header.startswith("type(") or header.startswith("type "):
            type_column = column
    average_column = headers.get("average") or headers.get("avg")
    max_column, min_column = headers.get("max"), headers.get("min")
    sample_column = next((column for name, column in headers.items() if name.startswith("sample")), None)
    if not all((average_column, max_column, min_column, sample_column)):
        return None
    result = {
        "average": int(average_column),
        "max": int(max_column),
        "min": int(min_column),
        "sample": int(sample_column),
    }
    if type_column is not None:
        result["type"] = type_column
    return result


def normalized_header_text(value: object) -> str:
    """Normalize a source header only enough to compare its visible words."""
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).split())


def defect_rate_header_columns(cells: object) -> dict[str, int] | None:
    """Recognize one explicit Input/OK/Total NG/NG rate source table header.

    This intentionally does not guess from loose mentions of NG or percentages.  A
    row must name all four accounting fields so each displayed rate can be checked
    against the exact source row before a deterministic draft is generated.
    """
    headers: dict[str, int] = {}
    type_column: int | None = None
    for column, cell in cells_by_column(cells).items():
        header = normalized_header_text(source_cell_text(cell))
        if not header:
            continue
        if header == "type":
            type_column = column
        elif header == "input":
            headers["input"] = column
        elif header == "ok":
            headers["ok"] = column
        elif header in {"total ng", "total defect"}:
            headers["total_ng"] = column
        elif header in {"ng rate", "defect rate"}:
            headers["rate"] = column
    if not all(name in headers for name in ("input", "ok", "total_ng", "rate")):
        return None
    result = {name: int(column) for name, column in headers.items()}
    if type_column is not None:
        result["type"] = type_column
    return result


def source_rate_ratio(value: float) -> float | None:
    """Interpret a stored Excel rate as a ratio or a visible percentage value."""
    if value < 0:
        return None
    return value if value <= 1 else value / 100 if value <= 100 else None


def source_table_metadata_from_rows(
    rows: dict[int, list[dict]],
    header_row: int,
    header_columns: dict[str, int],
    first_data_row: int,
) -> dict[str, str]:
    """Keep only adjacent, explicit table title/type facts; never infer their meaning."""
    metadata: dict[str, str] = {}
    caption_values = [source_cell_text(cell) for cell in rows.get(header_row - 1, [])]
    captions = [value for value in caption_values if value]
    if len(captions) == 1:
        metadata["caption"] = captions[0]
    type_column = header_columns.get("type")
    if type_column is not None:
        type_text = source_cell_text(cells_by_column(rows.get(first_data_row, [])).get(type_column))
        if type_text:
            metadata["type"] = type_text
    return metadata


def common_source_table_metadata(measurements: list[dict[str, object]]) -> dict[str, str]:
    """Return metadata only when every measurement in one rendered metric agrees."""
    if not measurements:
        return {}
    common: dict[str, str] = {}
    for field in ("caption", "type"):
        values = []
        for measurement in measurements:
            source_table = measurement.get("sourceTable")
            if not isinstance(source_table, dict):
                values = []
                break
            text = str(source_table.get(field) or "").strip()
            if not text:
                values = []
                break
            values.append(text)
        if values and len(set(values)) == 1:
            common[field] = values[0]
    return common


def source_table_metadata(value: object) -> dict[str, str]:
    """Normalize the two supported source-table display fields."""
    if not isinstance(value, dict):
        return {}
    metadata: dict[str, str] = {}
    for field in ("caption", "type"):
        text = str(value.get(field) or "").strip()
        if text:
            metadata[field] = text
    return metadata


def metric_source_table_metadata(metric: object) -> dict[str, str]:
    """Read source-table metadata from a direct manifest or the DB-persisted metric notes."""
    if not isinstance(metric, dict):
        return {}
    direct = source_table_metadata(metric.get("sourceTable", metric.get("source_table")))
    if direct:
        return direct
    notes = metric.get("notes")
    note_items = notes if isinstance(notes, list) else [notes]
    for note in note_items:
        if isinstance(note, dict) and note.get("kind") == SOURCE_TABLE_METADATA_NOTE_KIND:
            metadata = source_table_metadata(note)
            if metadata:
                return metadata
    return {}


def source_table_note(metadata: dict[str, str]) -> list[dict[str, str]]:
    """Use existing metric notes_json to persist source-table fields through the universal DB."""
    return [{"kind": SOURCE_TABLE_METADATA_NOTE_KIND, **metadata}] if metadata else []


def source_table_metadata_from_grid(
    conn: sqlite3.Connection,
    workbook_id: int,
    evidence_items: object,
) -> dict[str, str]:
    """Recover explicit title/type values for older reports from their source-grid table.

    This is intentionally a read-only compatibility path.  It requires the same
    unambiguous summary header as packet detection, its immediately preceding
    single-cell caption, and the source Type column on the evidence's first row.
    """
    if not isinstance(evidence_items, list):
        return {}
    for item in evidence_items:
        if not isinstance(item, dict):
            continue
        sheet_name = str(item.get("sheet_name", item.get("sheet")) or "").strip()
        start_row = number_value(item.get("start_row", item.get("startRow")))
        if start_row is None:
            range_address = str(item.get("range_address", item.get("range")) or "")
            match = re.search(r"[A-Za-z]+(\d+)", range_address)
            start_row = float(match.group(1)) if match else None
        if not sheet_name or start_row is None or start_row < 2:
            continue
        first_data_row = int(start_row)
        source_rows = core.dict_rows(
            conn,
            """
            SELECT row_number, cells_json
            FROM grid_sheet_rows
            WHERE workbook_id=? AND sheet_name=? AND row_number BETWEEN ? AND ?
            ORDER BY row_number
            """,
            (workbook_id, sheet_name, max(1, first_data_row - 12), first_data_row),
        )
        rows: dict[int, list[dict]] = {}
        for source_row in source_rows:
            try:
                cells = json.loads(source_row.get("cells_json") or "[]")
            except (TypeError, json.JSONDecodeError):
                continue
            if isinstance(cells, list):
                rows[int(source_row["row_number"])] = cells
        for header_row in range(first_data_row - 1, max(0, first_data_row - 13), -1):
            header_columns = complete_measurement_header_columns(rows.get(header_row, []))
            if header_columns is None:
                continue
            metadata = source_table_metadata_from_rows(rows, header_row, header_columns, first_data_row)
            if metadata:
                return metadata
    return {}


def enrich_export_with_source_table_metadata(conn: sqlite3.Connection, export: dict) -> dict:
    """Add read-only source-table display metadata to legacy DB exports when absent."""
    report = export.get("report") if isinstance(export.get("report"), dict) else {}
    workbook_id = number_value(report.get("workbookId", report.get("workbook_id")))
    if workbook_id is None:
        return export
    for review in export.get("reviews", []) if isinstance(export.get("reviews"), list) else []:
        if not isinstance(review, dict):
            continue
        for metric in review.get("metrics", []) if isinstance(review.get("metrics"), list) else []:
            if not isinstance(metric, dict) or metric_source_table_metadata(metric):
                continue
            metadata = source_table_metadata_from_grid(conn, int(workbook_id), metric.get("evidence"))
            if metadata:
                metric["sourceTable"] = metadata
    return export


def packet_complete_raw_measurements(packet: dict) -> list[dict[str, object]]:
    """Find complete, source-self-consistent Sample/Average/Max/Min tables in one packet.

    This deliberately recognizes only an unambiguous rectangular measurement layout.  Other
    Excel layouts stay governed by the universal evidence checks instead of being guessed.
    """
    selection = packet.get("packetSelection")
    if not isinstance(selection, dict) or any(selection.get(key) for key in ("rowTruncated", "cellTruncated", "dataTruncated")):
        return []
    sheet_rows = packet.get("sheetRows")
    if not isinstance(sheet_rows, list):
        return []
    grouped: dict[tuple[object, object], list[dict]] = {}
    for row in sheet_rows:
        if isinstance(row, dict):
            grouped.setdefault((row.get("sheet_index"), row.get("sheet_name")), []).append(row)

    measurements: list[dict[str, object]] = []
    for (sheet_index, sheet_name), rows in grouped.items():
        rows.sort(key=lambda row: int(row.get("row_number") or 0))
        for header_index, header_row in enumerate(rows):
            cells = header_row.get("cells")
            if not isinstance(cells, list):
                continue
            header_columns = complete_measurement_header_columns(cells)
            if header_columns is None:
                continue
            average_column = header_columns["average"]
            max_column, min_column = header_columns["max"], header_columns["min"]
            sample_column = header_columns["sample"]

            current: dict[str, object] | None = None
            for row in rows[header_index + 1 :]:
                row_cells = cells_by_column(row.get("cells"))
                average = number_value(row_cells.get(average_column, {}).get("value"))
                maximum = number_value(row_cells.get(max_column, {}).get("value"))
                minimum = number_value(row_cells.get(min_column, {}).get("value"))
                label = source_cell_text(row_cells.get(average_column - 1))
                if label and all(value is not None for value in (average, maximum, minimum)):
                    if current:
                        measurements.append(current)
                    source_rows = {
                        int(candidate.get("row_number") or 0): candidate.get("cells", [])
                        for candidate in rows
                        if isinstance(candidate, dict) and int(candidate.get("row_number") or 0) > 0
                    }
                    current = {
                        "sheetIndex": sheet_index,
                        "sheetName": sheet_name,
                        "label": label,
                        "average": average,
                        "min": minimum,
                        "max": maximum,
                        "samples": [],
                        "sampleColumn": sample_column,
                        "firstSampleRow": int(row.get("row_number") or 0),
                        "lastSampleRow": int(row.get("row_number") or 0),
                        "maxSampleColumn": sample_column,
                        "sourceTable": source_table_metadata_from_rows(
                            source_rows,
                            int(header_row.get("row_number") or 0),
                            header_columns,
                            int(row.get("row_number") or 0),
                        ),
                    }
                if current:
                    samples = current["samples"]
                    if isinstance(samples, list):
                        for column, cell in row_cells.items():
                            if column < sample_column:
                                continue
                            observed = number_value(cell.get("value"))
                            if observed is not None:
                                samples.append(observed)
                                current["lastSampleRow"] = int(row.get("row_number") or 0)
                                current["maxSampleColumn"] = max(int(current["maxSampleColumn"]), column)
            if current:
                measurements.append(current)

    complete: list[dict[str, object]] = []
    for measurement in measurements:
        samples = measurement["samples"]
        if not isinstance(samples, list) or len(samples) < 3:
            continue
        average = sum(samples) / len(samples)
        if (
            not math.isclose(float(measurement["average"]), average, rel_tol=0, abs_tol=0.01)
            or not math.isclose(float(measurement["min"]), min(samples), rel_tol=0, abs_tol=1e-9)
            or not math.isclose(float(measurement["max"]), max(samples), rel_tol=0, abs_tol=1e-9)
        ):
            continue
        measurement["sampleEvidenceRange"] = (
            f"{excel_column_label(int(measurement['sampleColumn']))}{int(measurement['firstSampleRow'])}:"
            f"{excel_column_label(int(measurement['maxSampleColumn']))}{int(measurement['lastSampleRow'])}"
        )
        complete.append(measurement)
    return complete


def defect_breakdown_details(
    header_cells: object,
    subheader_cells: object,
    row_cells: dict[int, dict],
    header_columns: dict[str, int],
) -> dict[str, float]:
    """Keep only source-labelled NG breakdown cells from one validated row."""
    headers = cells_by_column(header_cells)
    subheaders = cells_by_column(subheader_cells)
    result: dict[str, float] = {}
    excluded = set(header_columns.values())
    source_columns = sorted(headers)
    detail_columns = sorted(set(headers) | set(subheaders))
    for column in detail_columns:
        if column in excluded:
            continue
        parent_column = max((candidate for candidate in source_columns if candidate <= column), default=None)
        if parent_column is None:
            continue
        parent = source_cell_text(headers[parent_column])
        parent_key = normalized_header_text(parent)
        if not parent_key.startswith("ng") or parent_key in {"ng rate", "ngrate"}:
            continue
        observed = number_value(row_cells.get(column, {}).get("value"))
        if observed is None:
            continue
        child = source_cell_text(subheaders.get(column))
        label = f"{parent} · {child}" if child else parent
        result[label] = observed
    return result


def packet_complete_defect_rates(packet: dict) -> list[dict[str, object]]:
    """Find complete, reconciled Input/OK/Total NG/NG rate tables in one packet.

    The detector accepts only fully retained packet data and requires a direct
    source rate for every cohort.  It recomputes Total NG / Input only to verify
    that source value; it never derives an acceptance threshold or business
    decision from the rate.
    """
    selection = packet.get("packetSelection")
    if not isinstance(selection, dict) or any(selection.get(key) for key in ("rowTruncated", "cellTruncated", "dataTruncated")):
        return []
    sheet_rows = packet.get("sheetRows")
    if not isinstance(sheet_rows, list):
        return []
    grouped: dict[tuple[object, object], list[dict]] = {}
    for row in sheet_rows:
        if isinstance(row, dict):
            grouped.setdefault((row.get("sheet_index"), row.get("sheet_name")), []).append(row)

    defects: list[dict[str, object]] = []
    for (sheet_index, sheet_name), rows in grouped.items():
        rows.sort(key=lambda row: int(row.get("row_number") or 0))
        source_rows = {
            int(row.get("row_number") or 0): row.get("cells", [])
            for row in rows
            if int(row.get("row_number") or 0) > 0
        }
        for header_index, header_row in enumerate(rows):
            header_cells = header_row.get("cells")
            header_columns = defect_rate_header_columns(header_cells)
            if header_columns is None:
                continue
            header_row_number = int(header_row.get("row_number") or 0)
            if header_row_number <= 0:
                continue
            subheader_cells: object = []
            if header_index + 1 < len(rows) and int(rows[header_index + 1].get("row_number") or 0) == header_row_number + 1:
                subheader_cells = rows[header_index + 1].get("cells", [])
            source_table = source_table_metadata_from_rows(
                source_rows,
                header_row_number,
                header_columns,
                header_row_number + 1,
            )
            table_defects: list[dict[str, object]] = []
            invalid_table = False
            for row in rows[header_index + 1 :]:
                row_number = int(row.get("row_number") or 0)
                row_cells = cells_by_column(row.get("cells"))
                input_count = number_value(row_cells.get(header_columns["input"], {}).get("value"))
                ok_count = number_value(row_cells.get(header_columns["ok"], {}).get("value"))
                total_ng = number_value(row_cells.get(header_columns["total_ng"], {}).get("value"))
                reported_rate = number_value(row_cells.get(header_columns["rate"], {}).get("value"))
                label = source_cell_text(row_cells.get(header_columns.get("type", -1)))
                accounting_values = (input_count, ok_count, total_ng, reported_rate)
                if not label or not any(value is not None for value in accounting_values):
                    continue
                if any(value is None for value in accounting_values):
                    invalid_table = True
                    break
                assert input_count is not None and ok_count is not None and total_ng is not None and reported_rate is not None
                ratio = source_rate_ratio(reported_rate)
                computed_ratio = total_ng / input_count if input_count > 0 else None
                if (
                    ratio is None
                    or computed_ratio is None
                    or input_count < 0
                    or ok_count < 0
                    or total_ng < 0
                    or total_ng > input_count
                    or not math.isclose(ok_count + total_ng, input_count, rel_tol=0, abs_tol=1e-9)
                    or not math.isclose(ratio, computed_ratio, rel_tol=0, abs_tol=0.0005)
                ):
                    invalid_table = True
                    break
                first_column = min(cells_by_column(header_cells), default=header_columns["input"])
                last_column = max(cells_by_column(header_cells), default=header_columns["rate"])
                metadata = dict(source_table)
                type_text = source_cell_text(row_cells.get(header_columns.get("type", -1)))
                if type_text:
                    metadata["type"] = type_text
                table_defects.append(
                    {
                        "sheetIndex": sheet_index,
                        "sheetName": sheet_name,
                        "label": label,
                        "input": input_count,
                        "ok": ok_count,
                        "totalNg": total_ng,
                        "rateRatio": computed_ratio,
                        "reportedRate": reported_rate,
                        "details": defect_breakdown_details(header_cells, subheader_cells, row_cells, header_columns),
                        "rowNumber": row_number,
                        "tableTopRow": max(1, header_row_number - 1),
                        "firstColumn": first_column,
                        "lastColumn": last_column,
                        "evidenceRange": (
                            f"{excel_column_label(first_column)}{max(1, header_row_number - 1)}:"
                            f"{excel_column_label(last_column)}{row_number}"
                        ),
                        "sourceTable": metadata,
                    }
                )
            if not invalid_table:
                defects.extend(table_defects)
    return defects


def metric_value_number(value: dict, *names: str) -> float | None:
    for name in names:
        result = number_value(value.get(name))
        if result is not None:
            return result
    return None


def measurement_value_labels(review: dict, value: dict) -> str:
    cohort_key = next(
        (
            value.get(name)
            for name in ("cohort", "cohort_key", "cohortKey")
            if value.get(name) is not None
        ),
        "",
    )
    labels = [str(cohort_key or "")]
    for name in ("cohort_label", "cohortLabel", "label"):
        if value.get(name):
            labels.append(str(value[name]))
    for cohort in review.get("cohorts", []) if isinstance(review.get("cohorts"), list) else []:
        if not isinstance(cohort, dict):
            continue
        if cohort_key in (cohort.get("key"), cohort.get("cohort_key"), cohort.get("cohortKey")):
            labels.extend(str(cohort.get(name) or "") for name in ("label", "condition", "condition_text", "conditionText"))
    return " ".join(labels).casefold()


def labels_describe_same_measurement(expected: str, candidate: str) -> bool:
    expected_numbers = set(re.findall(r"\d+(?:\.\d+)?", expected))
    candidate_numbers = set(re.findall(r"\d+(?:\.\d+)?", candidate))
    if expected_numbers:
        return expected_numbers <= candidate_numbers
    return expected.casefold() in candidate or candidate in expected.casefold()


def sample_sequence_from_details(details: dict) -> list[float] | None:
    for name in ("sampleSequence", "sampleValues", "observedSamples", "rawSamples"):
        sequence = details.get(name)
        if not isinstance(sequence, list):
            continue
        converted = [number_value(item) for item in sequence]
        if all(item is not None for item in converted):
            return [float(item) for item in converted if item is not None]
    return None


def validate_complete_raw_measurement_details(packet: dict, manifest: dict) -> None:
    """Reject a manifest that compresses a complete packet measurement table.

    The host derives the expected values solely from the selected packet.  This checks the
    output contract before database import, so a plausible average-only draft cannot replace a
    detailed raw-measurement analysis.
    """
    expected_measurements = packet_complete_raw_measurements(packet)
    if not expected_measurements:
        return
    candidates: list[tuple[dict, dict]] = []
    for review in manifest.get("reviews", []) if isinstance(manifest.get("reviews"), list) else []:
        if not isinstance(review, dict):
            continue
        for metric in review.get("metrics", []) if isinstance(review.get("metrics"), list) else []:
            if not isinstance(metric, dict):
                continue
            for value in metric.get("values", []) if isinstance(metric.get("values"), list) else []:
                if isinstance(value, dict):
                    candidates.append((review, value))

    errors: list[str] = []
    for expected in expected_measurements:
        label = str(expected["label"])
        matching = [
            (review, value)
            for review, value in candidates
            if labels_describe_same_measurement(label, measurement_value_labels(review, value))
        ]
        if not matching:
            errors.append(f"{label}: missing a standalone cohort measurement value")
            continue
        review, value = next(
            (
                (review, value)
                for review, value in matching
                if isinstance(value.get("details"), dict) and sample_sequence_from_details(value["details"])
            ),
            matching[0],
        )
        details = value.get("details")
        if not isinstance(details, dict):
            errors.append(f"{label}: details must contain the complete sample sequence")
            continue
        sequence = sample_sequence_from_details(details)
        if sequence is None:
            errors.append(f"{label}: details.sampleSequence is required for the complete raw table")
            continue
        source_samples = expected["samples"]
        if not isinstance(source_samples, list) or len(sequence) != len(source_samples):
            errors.append(f"{label}: sampleSequence count does not match the selected packet")
            continue
        if any(not math.isclose(actual, float(source), rel_tol=0, abs_tol=1e-9) for actual, source in zip(sequence, source_samples)):
            errors.append(f"{label}: sampleSequence differs from the selected packet")
            continue
        sample_count = metric_value_number(details, "sampleCount", "sample_count", "n")
        if sample_count is None or not math.isclose(sample_count, len(sequence), rel_tol=0, abs_tol=1e-9):
            errors.append(f"{label}: details.sampleCount must equal the selected sample count")
        expected_range = str(expected["sampleEvidenceRange"])
        evidence_range = str(details.get("sampleEvidenceRange") or "")
        if evidence_range not in (expected_range, f"{expected['sheetName']}!{expected_range}"):
            errors.append(f"{label}: details.sampleEvidenceRange must equal {expected_range}")
        expected_stats = {
            "average": sum(sequence) / len(sequence),
            "min": min(sequence),
            "max": max(sequence),
        }
        manifest_stats = {
            "average": metric_value_number(value, "average", "averageValue", "avg", "avgValue"),
            "min": metric_value_number(value, "min", "minValue"),
            "max": metric_value_number(value, "max", "maxValue"),
        }
        for name, calculated in expected_stats.items():
            observed = manifest_stats[name]
            if observed is None or not math.isclose(observed, calculated, rel_tol=0, abs_tol=0.01):
                errors.append(f"{label}: manifest {name} does not match its sampleSequence")
        recomputed = details.get("recomputedSummary")
        if not isinstance(recomputed, dict):
            errors.append(f"{label}: details.recomputedSummary is required")
            continue
        for name, calculated in expected_stats.items():
            observed = metric_value_number(recomputed, name, f"{name}Value")
            if observed is None or not math.isclose(observed, calculated, rel_tol=0, abs_tol=0.01):
                errors.append(f"{label}: recomputedSummary.{name} does not match sampleSequence")
        expected_sd = math.sqrt(sum((sample - expected_stats["average"]) ** 2 for sample in sequence) / (len(sequence) - 1))
        standard_deviation = metric_value_number(recomputed, "sampleStandardDeviation", "sampleStdDev")
        if standard_deviation is None or not math.isclose(standard_deviation, expected_sd, rel_tol=0, abs_tol=0.01):
            errors.append(f"{label}: recomputedSummary.sampleStandardDeviation is missing or incorrect")
        observed_range = metric_value_number(recomputed, "range", "sampleRange")
        if observed_range is None or not math.isclose(observed_range, expected_stats["max"] - expected_stats["min"], rel_tol=0, abs_tol=1e-9):
            errors.append(f"{label}: recomputedSummary.range is missing or incorrect")
        reconciliation = details.get("displayedSummaryReconciliation")
        if reconciliation not in (True, "MATCH", "MATCHES"):
            errors.append(f"{label}: details.displayedSummaryReconciliation must explicitly confirm the match")
    if errors:
        raise ValueError("Complete raw measurement contract failed: " + "; ".join(errors))


def canonical_key(text: object, fallback: str, used: set[str]) -> str:
    """Create a stable, duplicate-safe manifest key without translating source labels."""
    key = re.sub(r"[^a-z0-9]+", "-", str(text or "").casefold()).strip("-") or fallback
    candidate = key
    suffix = 2
    while candidate in used:
        candidate = f"{key}-{suffix}"
        suffix += 1
    used.add(candidate)
    return candidate


def is_explicit_control_label(label: object) -> bool:
    """Recognize only labels that explicitly identify a Normal/control cohort."""
    text = str(label or "").casefold()
    return bool(re.search(r"\b(?:normal|control)\b|정상|대조", text))


def unique_evidence(items: list[dict]) -> list[dict]:
    """Keep evidence ordering deterministic while avoiding duplicate ranges."""
    result: list[dict] = []
    seen: set[tuple[str, str]] = set()
    for item in items:
        sheet = str(item.get("sheet") or "").strip()
        range_address = str(item.get("range") or "").strip()
        if not sheet or not range_address or (sheet, range_address) in seen:
            continue
        seen.add((sheet, range_address))
        result.append({"sheet": sheet, "range": range_address, "role": str(item.get("role") or "SOURCE")})
    return result


def packet_evidence(packet: dict) -> list[dict]:
    """Return one safe packet-backed grid range for an observation-only fallback."""
    sheet_rows = packet.get("sheetRows")
    if isinstance(sheet_rows, list):
        for row in sheet_rows:
            if not isinstance(row, dict):
                continue
            sheet = str(row.get("sheet_name") or row.get("sheetName") or "").strip()
            cells = row.get("cells")
            if not sheet or not isinstance(cells, list):
                continue
            for cell in cells:
                if not isinstance(cell, dict):
                    continue
                address = str(cell.get("address") or "").strip()
                if address:
                    return [{"sheet": sheet, "range": address, "role": "PACKET"}]
                column = number_value(cell.get("column", cell.get("col_number")))
                row_number = number_value(row.get("row_number", row.get("rowNumber")))
                if column and row_number and column > 0 and row_number > 0:
                    return [{"sheet": sheet, "range": f"{excel_column_label(int(column))}{int(row_number)}", "role": "PACKET"}]

    sheets = packet.get("sheets")
    if isinstance(sheets, list):
        for sheet in sheets:
            if not isinstance(sheet, dict):
                continue
            name = str(sheet.get("sheet_name") or sheet.get("sheetName") or "").strip()
            top = number_value(sheet.get("used_top", sheet.get("usedTop")))
            left = number_value(sheet.get("used_left", sheet.get("usedLeft")))
            if name and top and left and top > 0 and left > 0:
                return [{"sheet": name, "range": f"{excel_column_label(int(left))}{int(top)}", "role": "PACKET"}]
    return []


def canonical_defect_rate_manifest(title: str, defects: list[dict[str, object]]) -> dict:
    """Build an observation-only manifest from validated source rate rows."""
    used_keys: set[str] = set()
    cohorts: list[dict] = []
    values: list[dict] = []
    evidence_bounds: dict[tuple[str, int, int, int], int] = {}
    for index, defect in enumerate(defects, start=1):
        label = str(defect.get("label") or f"Cohort {index}").strip()
        cohort_key = canonical_key(label, f"cohort-{index}", used_keys)
        sheet_name = str(defect["sheetName"])
        top_row = int(defect["tableTopRow"])
        first_column = int(defect["firstColumn"])
        last_column = int(defect["lastColumn"])
        evidence_key = (sheet_name, top_row, first_column, last_column)
        evidence_bounds[evidence_key] = max(evidence_bounds.get(evidence_key, 0), int(defect["rowNumber"]))
        cohorts.append(
            {
                "key": cohort_key,
                "role": "OBSERVED",
                "label": label,
                "condition": "Source table Type value",
                "sortOrder": index,
            }
        )
        details = defect.get("details") if isinstance(defect.get("details"), dict) else {}
        values.append(
            {
                "cohort": cohort_key,
                "numerator": float(defect["totalNg"]),
                "denominator": float(defect["input"]),
                "ratePpm": float(defect["rateRatio"]) * 1_000_000,
                "status": "OBSERVED_ONLY",
                "details": details,
            }
        )

    evidence = unique_evidence([
        {
            "sheet": sheet_name,
            "range": f"{excel_column_label(first_column)}{top_row}:{excel_column_label(last_column)}{bottom_row}",
            "role": "DEFECT_RATE_SOURCE",
        }
        for (sheet_name, top_row, first_column, last_column), bottom_row in evidence_bounds.items()
    ])
    source_table = common_source_table_metadata(defects)
    conclusion = "Total NG / Input rates were recomputed from the same source rows. No acceptance threshold or release decision was inferred."
    metric = {
        "key": "source-total-ng-rate",
        "label": "Total NG / NG rate",
        "type": "defect_rate",
        "unit": "ppm",
        "spec": "",
        "definition": "Total NG divided by Input, reconciled with the NG rate stored in each source row.",
        "status": "OBSERVED_ONLY",
        "evidence": evidence,
        "sourceTable": source_table,
        "notes": source_table_note(source_table),
        "values": values,
        "comparisons": [],
    }
    return {
        "schemaVersion": "universal-analysis-v1",
        "source": {"dataset": "", "sourcePath": "", "workbookId": 0, "fingerprint": ""},
        "report": {
            "key": "packet-canonical-defect-rate-observation",
            "title": title,
            "type": "packet_defect_rate_observation",
            "purpose": "Preserve source-backed Total NG / Input observations from the selected workbook packet.",
            "scope": "Only complete Input/OK/Total NG/NG rate rows represented in the selected packet.",
            "status": "NEEDS_REVIEW",
            "decision": "OBSERVED_ONLY",
            "summary": conclusion,
            "limitations": ["No acceptance limit, specification, or release decision was supplied in the selected packet."],
            "artifacts": {},
            "evidence": evidence,
            "conclusions": [{"key": "packet-defect-rate-observation", "verdict": "NEEDS_REVIEW", "text": conclusion, "evidence": evidence}],
        },
        "reviews": [{
            "key": "source-defect-rates",
            "sortOrder": 1,
            "title": "Source-backed Total NG rate observations",
            "type": "defect_rate_observation",
            "objective": "Recompute visible source rates without inferring acceptance.",
            "comparisonBasis": "No comparison or acceptance basis was inferred.",
            "status": "NEEDS_REVIEW",
            "decision": "OBSERVED_ONLY",
            "summary": conclusion,
            "notes": [],
            "evidence": evidence,
            "cohorts": cohorts,
            "metrics": [metric],
            "conclusions": [{"key": "source-defect-rate-observation", "verdict": "NEEDS_REVIEW", "text": conclusion, "evidence": evidence}],
        }],
    }


def canonical_manifest_from_packet(packet: dict) -> dict:
    """Build a deterministic, observation-only universal-analysis-v1 manifest.

    This path intentionally derives only source-visible facts.  A complete raw
    Sample/Average/Max/Min table receives one value per cohort; every other
    packet shape becomes a clearly limited NEEDS_REVIEW record rather than an
    invented analysis.
    """
    workbook = packet.get("workbook") if isinstance(packet.get("workbook"), dict) else {}
    title = str(workbook.get("file_name") or workbook.get("fileName") or "Packet-backed workbook analysis").strip()
    defects = packet_complete_defect_rates(packet)
    if defects:
        return canonical_defect_rate_manifest(title, defects)
    complete = packet_complete_raw_measurements(packet)

    if complete:
        used_keys: set[str] = set()
        cohorts: list[dict] = []
        values: list[dict] = []
        evidence_items: list[dict] = []
        for index, measurement in enumerate(complete, start=1):
            label = str(measurement.get("label") or f"Cohort {index}").strip()
            cohort_key = canonical_key(label, f"cohort-{index}", used_keys)
            role = "CONTROL" if is_explicit_control_label(label) else "TEST"
            samples = [float(item) for item in measurement["samples"] if isinstance(item, (int, float)) and not isinstance(item, bool)]
            average = sum(samples) / len(samples)
            minimum, maximum = min(samples), max(samples)
            sample_sd = math.sqrt(sum((sample - average) ** 2 for sample in samples) / (len(samples) - 1))
            evidence = {
                "sheet": str(measurement["sheetName"]),
                "range": str(measurement["sampleEvidenceRange"]),
                "role": "RAW_SAMPLES",
            }
            evidence_items.append(evidence)
            cohorts.append(
                {
                    "key": cohort_key,
                    "role": role,
                    "label": label,
                    "condition": "Explicit Normal/control label" if role == "CONTROL" else "Source table row label",
                    "sortOrder": index,
                }
            )
            values.append(
                {
                    "cohort": cohort_key,
                    "min": minimum,
                    "max": maximum,
                    "average": average,
                    "valueText": f"N={len(samples)}",
                    "status": "OBSERVED",
                    "details": {
                        "sampleCount": len(samples),
                        "sampleSequence": samples,
                        "sampleEvidenceRange": str(measurement["sampleEvidenceRange"]),
                        "recomputedSummary": {
                            "average": average,
                            "min": minimum,
                            "max": maximum,
                            "sampleStandardDeviation": sample_sd,
                            "range": maximum - minimum,
                        },
                        "displayedSummaryReconciliation": "MATCH",
                    },
                }
            )

        evidence = unique_evidence(evidence_items)
        control_values = [value for cohort, value in zip(cohorts, values) if cohort["role"] == "CONTROL"]
        comparisons: list[dict] = []
        if len(control_values) == 1:
            control = control_values[0]
            control_key = str(control["cohort"])
            control_average = float(control["average"])
            for index, value in enumerate(values, start=1):
                compared_key = str(value["cohort"])
                if compared_key == control_key:
                    continue
                delta = round(float(value["average"]) - control_average, 10)
                direction = "HIGHER" if delta > 0 else "LOWER" if delta < 0 else "NO_CHANGE"
                comparison: dict[str, object] = {
                    "key": f"average-vs-{control_key}-{index}",
                    "comparedCohort": compared_key,
                    "controlCohort": control_key,
                    "deltaValue": delta,
                    "deltaUnit": "",
                    "direction": direction,
                    "status": "OBSERVED_ONLY",
                    "summary": "Average difference is an observed calculation only; no acceptance decision was present.",
                    "calculation": f"{value['average']} - {control_average} = {delta}",
                    "evidence": evidence,
                }
                if control_average != 0:
                    comparison["relativeDeltaPercent"] = round(delta * 100 / control_average, 10)
                comparisons.append(comparison)

        control_note = (
            "A single explicit Normal/control cohort was present; average deltas are shown as observations."
            if len(control_values) == 1
            else "No unambiguous single Normal/control cohort was present, so no cohort comparison was inferred."
        )
        conclusion = "Complete raw Sample/Average/Max/Min values were recomputed from the selected packet. No acceptance limit or release decision was present."
        source_table = common_source_table_metadata(complete)
        metrics: list[dict] = [{
            "key": "complete-raw-measurement-statistics",
            "label": "Complete Raw Measurement Statistics",
            "type": "measurement_summary",
            "unit": "",
            "spec": "",
            "definition": "Minimum, maximum, and average recomputed from the exact selected sample sequence.",
            "status": "OBSERVED",
            "evidence": evidence,
            "sourceTable": source_table,
            "notes": source_table_note(source_table),
            "values": values,
            "comparisons": [],
        }]
        if comparisons:
            comparison_values = [
                {
                    "cohort": value["cohort"],
                    "min": value["min"],
                    "max": value["max"],
                    "average": value["average"],
                    "valueText": value["valueText"],
                    "status": "OBSERVED",
                    "details": {"sampleEvidenceRange": value["details"]["sampleEvidenceRange"]},
                }
                for value in values
            ]
            metrics.append({
                "key": "average-comparison-vs-explicit-control",
                "label": "Average comparison versus explicit Normal/control",
                "type": "measurement_average_comparison",
                "unit": "",
                "spec": "",
                "definition": "Average delta calculated only against the one source-labelled Normal/control cohort.",
                "status": "OBSERVED_ONLY",
                "evidence": evidence,
                "sourceTable": source_table,
                "notes": source_table_note(source_table),
                "values": comparison_values,
                "comparisons": comparisons,
            })
        return {
            "schemaVersion": "universal-analysis-v1",
            "source": {"dataset": "", "sourcePath": "", "workbookId": 0, "fingerprint": ""},
            "report": {
                "key": "packet-canonical-raw-measurements",
                "title": title,
                "type": "packet_raw_measurement_observation",
                "purpose": "Preserve complete raw measurement observations from the selected workbook packet.",
                "scope": "Only complete Sample/Average/Max/Min tables represented in the selected packet.",
                "status": "NEEDS_REVIEW",
                "decision": "OBSERVED_ONLY",
                "summary": conclusion,
                "limitations": ["No acceptance limit, specification, or release decision was supplied in the selected packet.", control_note],
                "artifacts": {},
                "evidence": evidence,
                "conclusions": [{"key": "packet-observation", "verdict": "NEEDS_REVIEW", "text": conclusion, "evidence": evidence}],
            },
            "reviews": [{
                "key": "complete-raw-measurements",
                "sortOrder": 1,
                "title": "Complete raw measurement observations",
                "type": "complete_raw_measurement_observation",
                "objective": "Recompute source-provided raw sample summaries without inferring acceptance.",
                "comparisonBasis": control_note,
                "status": "NEEDS_REVIEW",
                "decision": "OBSERVED_ONLY",
                "summary": conclusion,
                "notes": [],
                "evidence": evidence,
                "cohorts": cohorts,
                "metrics": metrics,
                "conclusions": [{"key": "raw-measurement-observation", "verdict": "NEEDS_REVIEW", "text": conclusion, "evidence": evidence}],
            }],
        }

    evidence = packet_evidence(packet)
    if not evidence:
        raise ValueError("The selected packet has no source grid range for a canonical observation manifest.")
    selection = packet.get("packetSelection") if isinstance(packet.get("packetSelection"), dict) else {}
    truncation_note = "The selected packet is incomplete; omitted rows or cells must be reviewed in the source workbook." if selection.get("dataTruncated") else "No complete Sample/Average/Max/Min table was available in the selected packet."
    conclusion = "The packet was recorded as an observation only; it does not contain a source-backed acceptance decision."
    return {
        "schemaVersion": "universal-analysis-v1",
        "source": {"dataset": "", "sourcePath": "", "workbookId": 0, "fingerprint": ""},
        "report": {
            "key": "packet-canonical-needs-review",
            "title": title,
            "type": "packet_observation",
            "purpose": "Record selected packet evidence without inferring unstated workbook meaning.",
            "scope": "Selected source packet evidence only.",
            "status": "NEEDS_REVIEW",
            "decision": "OBSERVED_ONLY",
            "summary": conclusion,
            "limitations": [truncation_note, "No acceptance limit, specification, or release decision was supplied in the selected packet."],
            "artifacts": {},
            "evidence": evidence,
            "conclusions": [{"key": "packet-needs-review", "verdict": "NEEDS_REVIEW", "text": conclusion, "evidence": evidence}],
        },
        "reviews": [{
            "key": "packet-observation",
            "sortOrder": 1,
            "title": "Packet-backed observation",
            "type": "packet_observation",
            "objective": "Preserve the selected evidence while awaiting an explicit analysis basis.",
            "comparisonBasis": "No explicit Normal/control comparison basis was available.",
            "status": "NEEDS_REVIEW",
            "decision": "OBSERVED_ONLY",
            "summary": conclusion,
            "notes": [truncation_note],
            "evidence": evidence,
            "cohorts": [{"key": "packet-observation", "role": "OBSERVED", "label": "Selected packet evidence", "condition": "No inferred condition"}],
            "metrics": [{
                "key": "packet-evidence",
                "label": "Selected packet evidence",
                "type": "packet_evidence",
                "unit": "",
                "spec": "",
                "definition": "A pointer to source cells retained in the selected packet.",
                "status": "OBSERVED_ONLY",
                "evidence": evidence,
                "values": [{"cohort": "packet-observation", "valueText": "Source evidence retained in the selected packet.", "status": "OBSERVED_ONLY", "details": {"packetSelection": selection}}],
                "comparisons": [],
            }],
            "conclusions": [{"key": "packet-observation", "verdict": "NEEDS_REVIEW", "text": conclusion, "evidence": evidence}],
        }],
    }


def bind_manifest_to_workbook(manifest: dict, workbook: dict, dataset: str) -> dict:
    """Bind the deterministic manifest to the selected host workbook before import."""
    source = manifest.setdefault("source", {})
    if not isinstance(source, dict):
        raise ValueError("Generated manifest source must be an object.")
    source.update(
        {
            "dataset": dataset,
            "sourcePath": str(workbook["source_path"]),
            "workbookId": int(workbook["workbook_id"]),
            "fingerprint": str(workbook["fingerprint"]),
        }
    )
    return manifest


def render_existing(args: argparse.Namespace) -> int:
    """Render the analysis reports already verified in the universal DB as HTML."""
    service = Path(args.service_dir).resolve()
    db_path = Path(args.db).resolve()
    rendered: list[dict[str, object]] = []
    with core.connect_rw(db_path) as conn:
        query = "SELECT analysis_report_id FROM analysis_reports WHERE dataset=?"
        parameters: list[object] = [args.dataset]
        if args.report_id is not None:
            query += " AND analysis_report_id=?"
            parameters.append(args.report_id)
        query += " ORDER BY analysis_report_id"
        reports = core.dict_rows(
            conn,
            query,
            tuple(parameters),
        )
        for report in reports:
            report_id = int(report["analysis_report_id"])
            export = enrich_export_with_source_table_metadata(conn, core.build_analysis_export(conn, report_id))
            html_path = service / "outputs" / "analysis-rendered" / f"analysis_report_{report_id}.html"
            html_path.parent.mkdir(parents=True, exist_ok=True)
            html_path.write_text(analysis_html(export), encoding="utf-8")
            conn.execute("UPDATE analysis_reports SET dashboard_html_path=? WHERE analysis_report_id=?", (str(html_path), report_id))
            rendered.append({"analysisReportId": report_id, "html": str(html_path)})
        conn.commit()
    print(json.dumps({"status": "ok", "rendered": rendered}, ensure_ascii=False))
    return 0


def run(args: argparse.Namespace) -> int:
    if args.force_ai_draft and args.reuse_curated:
        raise SystemExit("--force-ai-draft cannot be combined with --reuse-curated.")
    service = Path(args.service_dir).resolve(); db_path = Path(args.db).resolve(); source = str(Path(args.source).resolve())
    with core.connect_rw(db_path) as conn:
        workbook = core.first_dict(conn, "SELECT * FROM workbooks WHERE source_path=? ORDER BY workbook_id DESC LIMIT 1", (source,))
        if not workbook:
            raise SystemExit(f"Source workbook is not indexed: {source}")
        existing = core.first_dict(conn, "SELECT analysis_report_id, dashboard_html_path, manifest_path FROM analysis_reports WHERE workbook_id=? AND overall_status <> 'STALE' ORDER BY analysis_report_id DESC LIMIT 1", (int(workbook["workbook_id"]),))
        if existing and not args.force_ai_draft:
            if args.replace_auto_draft and is_runner_draft(existing):
                conn.execute("DELETE FROM analysis_reports WHERE analysis_report_id=?", (int(existing["analysis_report_id"]),))
                conn.commit()
            else:
                print(json.dumps({"status": "skipped", "analysisReportId": existing["analysis_report_id"], "html": existing["dashboard_html_path"], "reason": "A curated CLI analysis exists; it is preserved."}, ensure_ascii=False))
                return 0
        packet = core.build_universal_packet(conn, int(workbook["workbook_id"]), args.row_limit, args.cell_limit)
    draft_token = uuid.uuid4().hex[:12] if args.force_ai_draft else ""
    draft_suffix = f"_force_ai_draft_{draft_token}" if draft_token else "_ai_draft"
    manifest_path = service / "outputs" / "analysis-manifests" / f"workbook_{workbook['workbook_id']}{draft_suffix}.json"
    reuse = curated_reuse_for_source(service, source, args.dataset) if args.reuse_curated else None
    reused_html: Path | None = None
    reuse_provenance: dict | None = None
    if args.reuse_curated:
        if not reuse:
            print(json.dumps(curated_reuse_not_applicable(source), ensure_ascii=False))
        else:
            reused_html = write_reused_curated_html(service, reuse, Path(source), int(workbook["workbook_id"]))
            manifest_path, reuse_provenance = rebind_curated_manifest(
                service, reuse, Path(source), args.dataset, workbook, reused_html
            )
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if reused_html else None
    if manifest is None:
        manifest = bind_manifest_to_workbook(canonical_manifest_from_packet(packet), workbook, args.dataset)
        validate_complete_raw_measurement_details(packet, manifest)
        if args.force_ai_draft:
            manifest = prepare_force_ai_draft(manifest, draft_token)
        core.write_json(manifest_path, manifest)
    with core.connect_rw(db_path) as conn:
        imported = core.import_analysis_manifest(conn, manifest_path, manifest, args.dataset)
        report_id = int(imported["analysisReportId"]); verification = core.verify_analysis_report(conn, report_id)
        if not verification["ok"]:
            raise SystemExit("Analysis verification failed: " + "; ".join(verification["errors"]))
        if reused_html:
            html_path = reused_html
        else:
            export = enrich_export_with_source_table_metadata(conn, core.build_analysis_export(conn, report_id))
            html_path = service / "outputs" / "analysis-rendered" / f"analysis_report_{report_id}.html"; html_path.parent.mkdir(parents=True, exist_ok=True); html_path.write_text(analysis_html(export), encoding="utf-8")
        conn.execute("UPDATE analysis_reports SET dashboard_html_path=? WHERE analysis_report_id=?", (str(html_path), report_id)); conn.commit()
    result = {"status": "ok", "analysisReportId": report_id, "manifest": str(manifest_path), "html": str(html_path)}
    if not reused_html:
        result["generator"] = "deterministic-packet-canonical-v1"
    if reuse_provenance:
        result["curatedReuse"] = reuse_provenance
    if args.force_ai_draft:
        result["forceAiDraft"] = {"token": draft_token, "originalReportsPreserved": True}
    print(json.dumps(result, ensure_ascii=False))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate, verify, and render a packet-backed analysis draft.")
    parser.add_argument("--service-dir", required=True); parser.add_argument("--db", required=True); parser.add_argument("--source"); parser.add_argument("--dataset", required=True)
    parser.add_argument("--row-limit", type=int, default=500); parser.add_argument("--cell-limit", type=int, default=12000)
    parser.add_argument("--replace-auto-draft", action="store_true", help="Replace only a previous draft generated by this runner.")
    draft_mode = parser.add_mutually_exclusive_group()
    draft_mode.add_argument("--reuse-curated", action="store_true", help="Reuse a byte-identical, VERIFIED curated baseline when available; otherwise log why a normal draft is used.")
    draft_mode.add_argument("--force-ai-draft", action="store_true", help="Compatibility flag: create a fresh, separately keyed deterministic draft without changing curated reports or artifacts.")
    parser.add_argument("--render-existing", action="store_true", help="Render existing verified/curated DB analyses as HTML without generating a new draft.")
    parser.add_argument("--report-id", type=int, help="With --render-existing, render only this analysis report ID.")
    args = parser.parse_args()
    if args.report_id is not None and not args.render_existing:
        parser.error("--report-id requires --render-existing")
    if args.render_existing:
        return render_existing(args)
    if not args.source:
        parser.error("--source is required unless --render-existing is used")
    return run(args)


if __name__ == "__main__": raise SystemExit(main())
