from __future__ import annotations

"""Run the AI draft -> verified analysis DB -> HTML dashboard pipeline."""

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


def legacy_analysis_html(data: dict) -> str:
    report = data["report"]
    def esc(value: object) -> str: return html.escape(str(value or ""))
    def evidence(items: list[dict]) -> str:
        return "; ".join(f"{x.get('sheet', x.get('sheet_name', ''))}!{x.get('range', x.get('range_address', ''))}" for x in items)
    def badge(status: object) -> str:
        value = str(status or "NEEDS_REVIEW").upper()
        css = "good" if value in {"VERIFIED", "IMPROVED", "CAN_USE"} else "bad" if value in {"REJECTED", "CAN_NOT_USE"} else "review"
        return f"<span class='status {css}'>{esc(value)}</span>"
    def cohort_heading(review: dict, value: dict, fallback: object) -> str:
        key = value.get("cohort", value.get("cohort_key", fallback))
        cohort = next((item for item in review.get("cohorts", []) if item.get("cohort_key", item.get("key")) == key), {})
        label = value.get("cohort_label") or cohort.get("label") or key or fallback
        condition = cohort.get("condition_text") or ""
        return f"<strong class='cohort'>{esc(label)}</strong>" + (f"<br><span class='condition'>{esc(condition)}</span>" if condition else "")
    sections = [f"""<header><h1>{esc(report.get('title'))}</h1><p>{esc(report.get('summary'))}</p><p class='source-file'>원본 Excel: <code>{esc(report.get('fileName', report.get('sourcePath')))}</code></p></header><main>
<section><h2>분석 요약</h2><div class='table-wrap'><table><caption>보고서 및 분석 판정</caption><thead><tr><th>항목</th><th>내용</th><th>상태</th><th>근거</th></tr></thead><tbody>
<tr><th>시험 목적</th><td>{esc(report.get('purpose'))}</td><td>{badge(report.get('status'))}</td><td class='evidence'>{esc(evidence(report.get('evidence', [])))}</td></tr>
<tr><th>분석 범위</th><td>{esc(report.get('scope'))}</td><td>{badge(report.get('decision'))}</td><td class='evidence'>{esc(evidence(report.get('evidence', [])))}</td></tr>
</tbody></table></div></section>"""]
    key_rows = []
    for review in data.get("reviews", []):
        for metric in review.get("metrics", []):
            values = {v.get("cohort", v.get("cohort_key")): v for v in metric.get("values", [])}
            for comparison in metric.get("comparisons", []):
                test_key = comparison.get("comparedCohort", comparison.get("compared_cohort_key"))
                control_key = comparison.get("controlCohort", comparison.get("control_cohort_key"))
                test, control = values.get(test_key, {}), values.get(control_key, {})
                def value_text(value: dict) -> str:
                    rate = value.get("rate_ppm", value.get("ratePpm"))
                    if isinstance(rate, (int, float)): return f"{value.get('numerator', '')} / {value.get('denominator', '')} ({rate:,.0f} ppm)"
                    return str(value.get("value_text") or value.get("valueText") or "")
                key_rows.append(f"<tr><th scope='row'>{esc(metric.get('label'))}</th><td class='num'>{cohort_heading(review, test, test_key)}<br>{esc(value_text(test))}</td><td class='num'>{cohort_heading(review, control, control_key)}<br>{esc(value_text(control))}</td><td>{esc(comparison.get('summary', comparison.get('summary_text')))}</td><td>{badge(comparison.get('status'))}</td><td class='evidence'>{esc(evidence(comparison.get('evidence', [])))}</td></tr>")
    if key_rows:
        sections.append("<section><h2>핵심 결과 비교</h2><div class='table-wrap'><table><caption>변경 조건과 대조 조건의 주요 결과</caption><thead><tr><th>지표</th><th>Test</th><th>Control</th><th>차이</th><th>판정</th><th>근거</th></tr></thead><tbody>" + "".join(key_rows) + "</tbody></table></div></section>")
    for review in data.get("reviews", []):
        rows = []
        cohort_labels = {c.get("cohort_key", c.get("key")): c.get("label", c.get("cohort_key", c.get("key"))) for c in review.get("cohorts", [])}
        for metric in review.get("metrics", []):
            values = {v.get("cohort", v.get("cohort_key")): v for v in metric.get("values", [])}
            for comparison in metric.get("comparisons", []):
                test_key = comparison.get("comparedCohort", comparison.get("compared_cohort_key"))
                control_key = comparison.get("controlCohort", comparison.get("control_cohort_key"))
                test, control = values.get(test_key, {}), values.get(control_key, {})
                test_value = test.get("value_text") or test.get("valueText") or (f"{test.get('numerator', '')} / {test.get('denominator', '')} = {test.get('rate_ppm', test.get('ratePpm', '')):,.0f} ppm" if isinstance(test.get('rate_ppm', test.get('ratePpm')), (int, float)) else "")
                control_value = control.get("value_text") or control.get("valueText") or (f"{control.get('numerator', '')} / {control.get('denominator', '')} = {control.get('rate_ppm', control.get('ratePpm', '')):,.0f} ppm" if isinstance(control.get('rate_ppm', control.get('ratePpm')), (int, float)) else "")
                rows.append(f"<tr><th>{esc(metric.get('label'))}</th><td>{cohort_heading(review, test, test_key)}<br>{esc(test_value)}</td><td>{cohort_heading(review, control, control_key)}<br>{esc(control_value)}</td><td>{esc(comparison.get('summary', comparison.get('summary_text')))}</td><td>{badge(comparison.get('status'))}</td><td class='evidence'>{esc(evidence(comparison.get('evidence', [])))}</td></tr>")
            if not metric.get("comparisons"):
                for value in metric.get("values", []):
                    displayed = value.get("value_text") or value.get("valueText") or value.get("value_number") or value.get("valueNumber") or json.dumps(value.get("details", {}), ensure_ascii=False)
                    key = value.get("cohort_key", value.get("cohort")); rows.append(f"<tr><th>{esc(metric.get('label'))}: {esc(cohort_labels.get(key, key))}</th><td colspan='3'>{esc(displayed)}</td><td>{badge(value.get('result_status', value.get('status')))}</td><td class='evidence'>{esc(evidence(metric.get('evidence', [])))}</td></tr>")
        conclusions = "<br>".join(esc(x.get("text", x.get("conclusion_text"))) for x in review.get("conclusions", [])) or "No conclusion generated."
        sections.append(f"<section><h2>상세 분석: {esc(review.get('title'))}</h2><div class='table-wrap'><table><caption>{esc(review.get('summary'))}</caption><thead><tr><th>지표 / 조건</th><th>Test / 관측값</th><th>Control</th><th>차이 / 결과</th><th>판정</th><th>근거</th></tr></thead><tbody>{''.join(rows)}<tr class='highlight'><th>분석 결론</th><td colspan='4'>{conclusions}</td><td class='evidence'>{esc(evidence(review.get('evidence', [])))}</td></tr></tbody></table></div></section>")
    limits = "<br>".join(esc(x) for x in report.get("limitations", [])) or "Human review is required before using this result."
    sections.append(f"<section><h2>최종 판정 및 검토 제한</h2><div class='table-wrap'><table><caption>데이터 기반 종합 판단</caption><thead><tr><th>항목</th><th>판정</th><th>내용</th></tr></thead><tbody><tr><th>분석 결론</th><td>{badge(report.get('decision'))}</td><td>{esc(report.get('summary'))}</td></tr><tr><th>검토 제한</th><td>{badge(report.get('status'))}</td><td>{limits}</td></tr></tbody></table></div></section></main>")
    css = ":root{--bg:#f4f6fa;--panel:#fff;--line:#d8e0ea;--head:#eef2f6;--ink:#17202e;--muted:#667085;--blue:#175cd3;--green:#067647;--green-bg:#ecfdf3;--red:#b42318;--red-bg:#fef3f2;--amber:#b54708;--amber-bg:#fffaeb}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:'Segoe UI','Malgun Gothic',Arial,sans-serif;line-height:1.45}header{padding:18px 22px 14px;border-bottom:1px solid var(--line);background:var(--panel)}h1{margin:0;font-size:23px}header p{margin:5px 0 0;color:var(--muted);font-size:12px}main{width:min(1500px,100%);margin:0 auto;padding:16px 18px 30px}section{margin-bottom:16px}h2{margin:0 0 7px;font-size:16px}.table-wrap{overflow-x:auto;border:1px solid var(--line);border-radius:9px;background:var(--panel)}table{width:100%;border-collapse:collapse;font-size:12px}caption{padding:9px 10px;border-bottom:1px solid var(--line);color:#344054;background:#f8fafc;text-align:left;font-weight:800}th,td{padding:8px 9px;border-right:1px solid var(--line);border-bottom:1px solid var(--line);vertical-align:top;overflow-wrap:anywhere}th:last-child,td:last-child{border-right:0}thead th{color:#475467;background:var(--head);text-align:center}tbody th{min-width:130px;background:#f8fafc;text-align:left}.cohort{color:#175cd3;font-size:11px}.condition{color:#667085;font-size:10px;line-height:1.25}.evidence{color:var(--muted);font-family:Consolas,monospace;font-size:10px}.source-file{font-family:Consolas,monospace;font-size:11px}.status{display:inline-block;padding:3px 7px;border-radius:999px;font-size:10px;font-weight:800}.good{color:var(--green);background:var(--green-bg)}.bad{color:var(--red);background:var(--red-bg)}.review{color:var(--amber);background:var(--amber-bg)}.highlight td,.highlight th{background:#fffdf5}@media print{@page{size:A4 landscape;margin:8mm}body{background:#fff}header,main{width:100%;padding-left:0;padding-right:0}.table-wrap{overflow:visible}}"
    return "<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>Analysis dashboard</title><style>" + css + "</style></head><body>" + "".join(sections) + "</body></html>"


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
        return f"{value:,.0f}" if float(value).is_integer() else f"{value:,}"

    def human_label(value: object) -> str:
        name = str(value or "").replace("_", " ")
        name = "".join((" " if char.isupper() and index and name[index - 1].islower() else "") + char for index, char in enumerate(name))
        return " ".join(name.split()).capitalize() or "Value"

    def evidence(items: object) -> str:
        references: list[str] = []
        for item in items if isinstance(items, list) else []:
            if not isinstance(item, dict):
                continue
            sheet = raw_text(pick(item, "sheet", "sheet_name", "sheetName"), "")
            range_address = raw_text(pick(item, "range", "range_address", "rangeAddress"), "")
            reference = "!".join(part for part in (sheet, range_address) if part)
            if reference:
                references.append(reference)
        return "; ".join(references) or "—"

    def badge(status: object) -> str:
        value = raw_text(status, "NEEDS_REVIEW").upper()
        css = "good" if value in {"VERIFIED", "IMPROVED", "CAN_USE", "OK", "REFERENCE"} else "bad" if value in {"REJECTED", "CAN_NOT_USE"} else "review"
        return f"<span class='status {css}'>{html.escape(value)}</span>"

    def cohort_heading(review: dict, value: dict, fallback: object) -> str:
        key = pick(value, "cohort", "cohort_key", "cohortKey", default=fallback)
        cohort = next(
            (item for item in review.get("cohorts", []) if isinstance(item, dict) and pick(item, "cohort_key", "cohortKey", "key") == key),
            {},
        )
        cohort_label = pick(value, "cohort_label", "cohortLabel")
        display = cohort_label if present(cohort_label) else pick(cohort, "label", default=key)
        condition = pick(cohort, "condition_text", "conditionText", "condition")
        return f"<strong class='cohort'>{esc(display)}</strong>" + (
            f"<br><span class='condition'>{esc(condition)}</span>" if present(condition) else ""
        )

    def details_html(details: object) -> str:
        if not isinstance(details, dict) or not details:
            return ""
        rendered = "".join(
            f"<div><dt>{esc(human_label(key))}</dt><dd>{esc(number(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else value)}</dd></div>"
            for key, value in details.items()
        )
        return f"<dl class='breakdown'>{rendered}</dl>"

    def value_html(value: dict) -> str:
        rows: list[str] = []
        value_text = pick(value, "value_text", "valueText")
        value_number = pick(value, "value_number", "valueNumber")
        if present(value_text):
            rows.append(f"<div>{esc(value_text)}</div>")
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

        measurement_fields = (
            ("최소", ("min_value", "minValue", "min")),
            ("최대", ("max_value", "maxValue", "max")),
            ("평균", ("average_value", "averageValue", "average", "avg_value", "avgValue")),
        )
        measurements = []
        for title, names in measurement_fields:
            measurement = pick(value, *names)
            if present(measurement):
                measurements.append(f"<span><b>{title}</b> {esc(number(measurement))}</span>")
        if measurements:
            rows.append("<div class='measurements'>" + "".join(measurements) + "</div>")

        detail_rows = details_html(pick(value, "details", default={}))
        if detail_rows:
            rows.append(detail_rows)
        status = pick(value, "result_status", "resultStatus", "status")
        if present(status):
            rows.append(f"<div class='value-status'>{badge(status)}</div>")
        return "<div class='metric-value'>" + ("".join(rows) or "—") + "</div>"

    def metric_heading(metric: dict) -> str:
        metric_type = pick(metric, "metric_type", "metricType", "type")
        unit = pick(metric, "unit")
        descriptor = " · ".join(raw_text(value, "") for value in (metric_type, unit) if present(value))
        spec = pick(metric, "spec_text", "specText", "spec")
        return (
            f"<strong>{esc(pick(metric, 'label'))}</strong>"
            + (f"<br><span class='metric-type'>{html.escape(descriptor)}</span>" if descriptor else "")
            + (f"<br><span class='metric-type'>규격: {esc(spec)}</span>" if present(spec) else "")
        )

    def comparison_html(comparison: dict) -> str:
        parts: list[str] = []
        for field in ("summary_text", "summary"):
            summary = pick(comparison, field)
            if present(summary):
                parts.append(f"<div>{esc(summary)}</div>")
                break
        calculation = pick(comparison, "calculation_text", "calculation")
        if present(calculation):
            parts.append(f"<div class='calculation'>{esc(calculation)}</div>")
        details = details_html(pick(comparison, "details", default={}))
        if details:
            parts.append(details)
        return "".join(parts) or "—"

    report_evidence = evidence(pick(report, "evidence", default=[]))
    sections = [
        f"""<header><h1>{esc(pick(report, 'title'))}</h1><p>{esc(pick(report, 'summary'))}</p><p class='source-file'>원본 Excel: <code>{esc(pick(report, 'fileName', 'file_name', 'sourcePath', 'source_path'))}</code></p></header><main>
<section><h2>분석 요약</h2><div class='table-wrap'><table><caption>보고서 범위와 판정</caption><thead><tr><th>항목</th><th>내용</th><th>상태</th><th>근거</th></tr></thead><tbody>
<tr><th>목적</th><td>{esc(pick(report, 'purpose'))}</td><td>{badge(pick(report, 'status'))}</td><td class='evidence'>{html.escape(report_evidence)}</td></tr>
<tr><th>분석 범위</th><td>{esc(pick(report, 'scope'))}</td><td>{badge(pick(report, 'decision'))}</td><td class='evidence'>{html.escape(report_evidence)}</td></tr>
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
                for comparison in comparisons:
                    test_key = pick(comparison, "comparedCohort", "compared_cohort_key", "comparedCohortKey")
                    control_key = pick(comparison, "controlCohort", "control_cohort_key", "controlCohortKey")
                    test, control = values.get(test_key, {}), values.get(control_key, {})
                    rows.append(
                        f"<tr><th>{metric_heading(metric)}</th><td>{cohort_heading(review, test, test_key)}{value_html(test)}</td>"
                        f"<td>{cohort_heading(review, control, control_key)}{value_html(control)}</td><td>{comparison_html(comparison)}</td>"
                        f"<td>{badge(pick(comparison, 'status'))}</td><td class='evidence'>{html.escape(evidence(pick(comparison, 'evidence', default=[])))}</td></tr>"
                    )
            else:
                for cohort_key, value in values.items():
                    rows.append(
                        f"<tr><th>{metric_heading(metric)}</th><td colspan='3'>{cohort_heading(review, value, cohort_key)}{value_html(value)}</td>"
                        f"<td>{badge(pick(value, 'result_status', 'resultStatus', 'status'))}</td>"
                        f"<td class='evidence'>{html.escape(evidence(pick(metric, 'evidence', default=[])))}</td></tr>"
                    )
        if not rows:
            rows.append("<tr><td colspan='6'>내보낸 지표 값이 없습니다.</td></tr>")
        conclusion_text = "<br>".join(
            esc(pick(item, "text", "conclusion_text", "conclusionText"))
            for item in review.get("conclusions", [])
            if isinstance(item, dict)
        ) or "생성된 결론이 없습니다."
        sections.append(
            f"<section><h2>상세 분석: {esc(pick(review, 'title'))}</h2><div class='table-wrap'><table><caption>{esc(pick(review, 'summary', 'summary_text', 'summaryText'))}</caption>"
            "<thead><tr><th>지표 / 유형</th><th>Test / 관측값</th><th>Control</th><th>차이 / 결과</th><th>상태</th><th>근거</th></tr></thead><tbody>"
            + "".join(rows)
            + f"<tr class='highlight'><th>분석 결론</th><td colspan='4'>{conclusion_text}</td><td class='evidence'>{html.escape(evidence(pick(review, 'evidence', default=[])))}</td></tr>"
            + "</tbody></table></div></section>"
        )
    limitations = pick(report, "limitations", default=[])
    limits = "<br>".join(esc(item) for item in limitations) if isinstance(limitations, list) and limitations else "이 결과를 사용하기 전에 사람의 검토가 필요합니다."
    sections.append(
        f"<section><h2>최종 판정 및 검토 제한</h2><div class='table-wrap'><table><caption>근거 기반 종합 판단</caption><thead><tr><th>항목</th><th>판정</th><th>내용</th></tr></thead><tbody>"
        f"<tr><th>분석 결론</th><td>{badge(pick(report, 'decision'))}</td><td>{esc(pick(report, 'summary'))}</td></tr>"
        f"<tr><th>검토 제한</th><td>{badge(pick(report, 'status'))}</td><td>{limits}</td></tr></tbody></table></div></section></main>"
    )
    css = ":root{--bg:#f4f6fa;--panel:#fff;--line:#d8e0ea;--head:#eef2f6;--ink:#17202e;--muted:#667085;--green:#067647;--green-bg:#ecfdf3;--red:#b42318;--red-bg:#fef3f2;--amber:#b54708;--amber-bg:#fffaeb}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:'Segoe UI','Malgun Gothic',Arial,sans-serif;line-height:1.45}header{padding:18px 22px 14px;border-bottom:1px solid var(--line);background:var(--panel)}h1{margin:0;font-size:23px}header p{margin:5px 0 0;color:var(--muted);font-size:12px}main{width:min(1500px,100%);margin:0 auto;padding:16px 18px 30px}section{margin-bottom:16px}h2{margin:0 0 7px;font-size:16px}.table-wrap{overflow-x:auto;border:1px solid var(--line);border-radius:9px;background:var(--panel)}table{width:100%;border-collapse:collapse;font-size:12px}caption{padding:9px 10px;border-bottom:1px solid var(--line);color:#344054;background:#f8fafc;text-align:left;font-weight:800}th,td{padding:8px 9px;border-right:1px solid var(--line);border-bottom:1px solid var(--line);vertical-align:top;overflow-wrap:anywhere}th:last-child,td:last-child{border-right:0}thead th{color:#475467;background:var(--head);text-align:center}tbody th{min-width:150px;background:#f8fafc;text-align:left}.cohort{display:block;color:#175cd3;font-size:11px}.condition,.metric-type,.calculation{color:var(--muted);font-size:10px;line-height:1.25}.metric-value{margin-top:5px}.value-label{color:var(--muted);font-weight:700}.measurements{display:flex;flex-wrap:wrap;gap:5px;margin-top:4px}.measurements span{padding:2px 5px;border-radius:4px;background:#f2f4f7}.breakdown{display:flex;flex-wrap:wrap;gap:4px;margin:5px 0 0}.breakdown div{display:flex;gap:4px;padding:2px 5px;border:1px solid var(--line);border-radius:4px;background:#fcfcfd}.breakdown dt{font-weight:700}.breakdown dd{margin:0}.value-status{margin-top:5px}.evidence{color:var(--muted);font-family:Consolas,monospace;font-size:10px}.source-file{font-family:Consolas,monospace;font-size:11px}.status{display:inline-block;padding:3px 7px;border-radius:999px;font-size:10px;font-weight:800}.good{color:var(--green);background:var(--green-bg)}.bad{color:var(--red);background:var(--red-bg)}.review{color:var(--amber);background:var(--amber-bg)}.highlight td,.highlight th{background:#fffdf5}@media print{@page{size:A4 landscape;margin:8mm}body{background:#fff}header,main{width:100%;padding-left:0;padding-right:0}.table-wrap{overflow:visible}}"
    return "<!doctype html><html lang='ko'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>분석 대시보드</title><style>" + css + "</style></head><body>" + "".join(sections) + "</body></html>"


def normalize_manifest(data: dict) -> dict:
    """Keep AI-provided observation fields instead of silently losing them on import."""
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
    """Give a newly generated draft an identity that cannot replace a curated report."""
    report = data.get("report")
    if not isinstance(report, dict):
        raise ValueError("AI draft must include a report object before force-draft preparation.")
    original_key = str(report.get("key") or "").strip()
    if not original_key:
        raise ValueError("AI draft report.key is required before force-draft preparation.")
    report["key"] = f"{original_key}-force-ai-{draft_token}"
    scope = str(report.get("scope") or "").strip()
    report["scope"] = f"{scope}\n[Fresh AI draft: force-ai token={draft_token}]".strip()
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
            columns = {
                str(cell.get("value") or "").strip().casefold(): int(cell.get("column") or 0)
                for cell in cells
                if isinstance(cell, dict) and int(cell.get("column") or 0) > 0
            }
            average_column = columns.get("average") or columns.get("avg")
            max_column, min_column = columns.get("max"), columns.get("min")
            sample_column = next((column for name, column in columns.items() if name.startswith("sample")), None)
            if not all((average_column, max_column, min_column, sample_column)):
                continue

            current: dict[str, object] | None = None
            for row in rows[header_index + 1 :]:
                row_cells = {
                    int(cell.get("column") or 0): cell
                    for cell in row.get("cells", [])
                    if isinstance(cell, dict) and int(cell.get("column") or 0) > 0
                }
                average = number_value(row_cells.get(average_column, {}).get("value"))
                maximum = number_value(row_cells.get(max_column, {}).get("value"))
                minimum = number_value(row_cells.get(min_column, {}).get("value"))
                label = str(row_cells.get(average_column - 1, {}).get("value") or "").strip()
                if label and all(value is not None for value in (average, maximum, minimum)):
                    if current:
                        measurements.append(current)
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
    """Reject an AI draft that compresses a complete packet measurement table.

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


def canonical_manifest_from_packet(packet: dict) -> dict:
    """Build a deterministic, observation-only universal-analysis-v1 manifest.

    This path intentionally derives only source-visible facts.  A complete raw
    Sample/Average/Max/Min table receives one value per cohort; every other
    packet shape becomes a clearly limited NEEDS_REVIEW record rather than an
    invented analysis.
    """
    workbook = packet.get("workbook") if isinstance(packet.get("workbook"), dict) else {}
    title = str(workbook.get("file_name") or workbook.get("fileName") or "Packet-backed workbook analysis").strip()
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
                delta = float(value["average"]) - control_average
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
                    comparison["relativeDeltaPercent"] = delta * 100 / control_average
                comparisons.append(comparison)

        control_note = (
            "A single explicit Normal/control cohort was present; average deltas are shown as observations."
            if len(control_values) == 1
            else "No unambiguous single Normal/control cohort was present, so no cohort comparison was inferred."
        )
        conclusion = "Complete raw Sample/Average/Max/Min values were recomputed from the selected packet. No acceptance limit or release decision was present."
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
                "metrics": [{
                    "key": "complete-raw-measurement",
                    "label": "Complete raw measurement summary",
                    "type": "measurement_summary",
                    "unit": "",
                    "spec": "",
                    "definition": "Minimum, maximum, and average recomputed from the exact selected sample sequence.",
                    "status": "OBSERVED",
                    "evidence": evidence,
                    "values": values,
                    "comparisons": comparisons,
                }],
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
        raise ValueError("AI draft source must be an object.")
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
        reports = core.dict_rows(
            conn,
            "SELECT analysis_report_id FROM analysis_reports WHERE dataset=? ORDER BY analysis_report_id",
            (args.dataset,),
        )
        for report in reports:
            report_id = int(report["analysis_report_id"])
            export = core.build_analysis_export(conn, report_id)
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
            export = core.build_analysis_export(conn, report_id)
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
    draft_mode.add_argument("--force-ai-draft", action="store_true", help="Create a fresh, separately keyed AI draft without changing any curated report or artifact.")
    parser.add_argument("--render-existing", action="store_true", help="Render existing verified/curated DB analyses as HTML without calling AI.")
    args = parser.parse_args()
    if args.render_existing:
        return render_existing(args)
    if not args.source:
        parser.error("--source is required unless --render-existing is used")
    return run(args)


if __name__ == "__main__": raise SystemExit(main())
