from __future__ import annotations

"""Run the AI draft -> verified analysis DB -> HTML dashboard pipeline."""

import argparse
import html
import json
import shutil
import sqlite3
import subprocess
from pathlib import Path

import inference_data_ai_cli as core


def json_object(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end < start:
        raise ValueError("AI did not return an analysis manifest JSON object.")
    value = json.loads(text[start : end + 1])
    if not isinstance(value, dict):
        raise ValueError("AI result must be a JSON object.")
    return value


def analysis_html(data: dict) -> str:
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


def is_runner_draft(report: dict) -> bool:
    """Only replace reports produced by this runner, never curated CLI reports."""
    return Path(str(report.get("manifest_path") or "")).name.startswith("workbook_")


def curated_manifest_for_source(service: Path, source: str) -> Path | None:
    """Return the hand-reviewed CLI manifest for this exact source, if present."""
    expected = str(Path(source).resolve())
    for manifest_path in sorted((service / "outputs" / "analysis-manifests").glob("*_analysis.json")):
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest_source = str(Path(str(manifest.get("source", {}).get("sourcePath") or "")).resolve())
        except (OSError, ValueError, json.JSONDecodeError):
            continue
        if manifest_source == expected:
            return manifest_path
    return None


def cli_calibration_paths(service: Path) -> list[Path]:
    """The curated CLI analyses define the required report quality and structure."""
    return sorted((service / "outputs" / "analysis-manifests").glob("*_analysis.json"))


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
    service = Path(args.service_dir).resolve(); db_path = Path(args.db).resolve(); source = str(Path(args.source).resolve())
    with core.connect_rw(db_path) as conn:
        workbook = core.first_dict(conn, "SELECT * FROM workbooks WHERE source_path=? ORDER BY workbook_id DESC LIMIT 1", (source,))
        if not workbook:
            raise SystemExit(f"Source workbook is not indexed: {source}")
        existing = core.first_dict(conn, "SELECT analysis_report_id, dashboard_html_path, manifest_path FROM analysis_reports WHERE workbook_id=? AND overall_status <> 'STALE' ORDER BY analysis_report_id DESC LIMIT 1", (int(workbook["workbook_id"]),))
        if existing:
            if args.replace_auto_draft and is_runner_draft(existing):
                conn.execute("DELETE FROM analysis_reports WHERE analysis_report_id=?", (int(existing["analysis_report_id"]),))
                conn.commit()
            else:
                print(json.dumps({"status": "skipped", "analysisReportId": existing["analysis_report_id"], "html": existing["dashboard_html_path"], "reason": "A curated CLI analysis exists; it is preserved."}, ensure_ascii=False))
                return 0
        packet = core.build_universal_packet(conn, int(workbook["workbook_id"]), args.row_limit, args.cell_limit)
    packet_path = service / "outputs" / "analysis-inputs" / f"workbook_{workbook['workbook_id']}_analysis_input.json"; core.write_json(packet_path, packet)
    manifest_path = service / "outputs" / "analysis-manifests" / f"workbook_{workbook['workbook_id']}_ai_draft.json"
    raw_path = manifest_path.with_suffix(".raw.txt")
    curated = curated_manifest_for_source(service, source)
    if curated:
        manifest = json.loads(curated.read_text(encoding="utf-8"))
        manifest_path = curated
        print(json.dumps({"status": "using-curated-cli-manifest", "manifest": str(curated)}, ensure_ascii=False))
    else:
        manifest = None
    codex = shutil.which("codex")
    if manifest is None and not codex:
        raise SystemExit("Codex CLI was not found on PATH; install/login to Codex CLI first.")
    if manifest is None:
        examples = "\n".join(f"- {path}" for path in cli_calibration_paths(service)) or "- No curated example is available."
        prompt = f"""Read {packet_path}. Create one source-backed universal-analysis-v1 JSON manifest for this Excel report.

The report must match the quality and structure of the hand-reviewed CLI analyses below. Read them as calibration examples before writing the JSON:
{examples}

Required analysis contract:
1. Establish report purpose, scope, comparison basis, cohorts, metrics, comparison deltas, decision, limitations, and exact Excel evidence ranges.
2. Keep separate sheets, tables, metrics, and conditions separate. Never combine incompatible Normal/Test groups or denominators.
3. For every NG-rate/ppm conclusion, cite numerator and denominator evidence and calculate ppm/delta exactly. For measurements, preserve units, spec, sample grouping, and min/max/average evidence.
4. A report can be VERIFIED only when its source has an explicit decision or complete, matched comparison evidence. Otherwise use NEEDS_REVIEW with concrete missing conditions. Do not infer causality from a filename.
5. Include source.dataset/sourcePath/fingerprint exactly from packet workbook. Every numeric claim and conclusion needs existing packet evidence.

Return JSON only, with no markdown."""
        result = subprocess.run([codex, "exec", "--sandbox", "read-only", "-C", str(service), "-o", str(raw_path), prompt], cwd=service, text=True, encoding="utf-8", errors="replace")
        if result.returncode:
            raise SystemExit(f"Codex CLI analysis failed with exit code {result.returncode}.")
        manifest = normalize_manifest(json_object(raw_path.read_text(encoding="utf-8"))); core.write_json(manifest_path, manifest)
    with core.connect_rw(db_path) as conn:
        imported = core.import_analysis_manifest(conn, manifest_path, manifest, args.dataset)
        report_id = int(imported["analysisReportId"]); verification = core.verify_analysis_report(conn, report_id)
        if not verification["ok"]:
            raise SystemExit("Analysis verification failed: " + "; ".join(verification["errors"]))
        export = core.build_analysis_export(conn, report_id)
        html_path = service / "outputs" / "analysis-rendered" / f"analysis_report_{report_id}.html"; html_path.parent.mkdir(parents=True, exist_ok=True); html_path.write_text(analysis_html(export), encoding="utf-8")
        conn.execute("UPDATE analysis_reports SET dashboard_html_path=? WHERE analysis_report_id=?", (str(html_path), report_id)); conn.commit()
    print(json.dumps({"status": "ok", "analysisReportId": report_id, "manifest": str(manifest_path), "html": str(html_path)}, ensure_ascii=False))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate, verify, and render an AI analysis draft.")
    parser.add_argument("--service-dir", required=True); parser.add_argument("--db", required=True); parser.add_argument("--source"); parser.add_argument("--dataset", required=True)
    parser.add_argument("--row-limit", type=int, default=500); parser.add_argument("--cell-limit", type=int, default=12000)
    parser.add_argument("--replace-auto-draft", action="store_true", help="Replace only a previous draft generated by this runner.")
    parser.add_argument("--render-existing", action="store_true", help="Render existing verified/curated DB analyses as HTML without calling AI.")
    args = parser.parse_args()
    if args.render_existing:
        return render_existing(args)
    if not args.source:
        parser.error("--source is required unless --render-existing is used")
    return run(args)


if __name__ == "__main__": raise SystemExit(main())
