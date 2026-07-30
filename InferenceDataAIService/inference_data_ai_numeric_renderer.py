#!/usr/bin/env python3
"""Render one numeric-only HTML review page per workbook from numeric-capture.sqlite.

The renderer never reads an Excel file and never invents a review result.  It
only formats source-backed numeric facts and same-date Test–Normal comparisons
already present in the batch-scoped database.
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import sqlite3
import sys
from datetime import UTC, datetime
from pathlib import Path

import inference_data_ai_structure_scan as structure


RENDERER_VERSION = "numeric-renderer-v1"


def utc_now() -> str:
    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


def atomic_write_text(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(value, encoding="utf-8")
    os.replace(temporary, path)


def atomic_write_json(path: Path, value: object) -> None:
    atomic_write_text(path, json.dumps(value, ensure_ascii=False, indent=2) + "\n")


def resolve_batch_directory(service_dir: Path, batch_id: str) -> Path:
    return structure.resolve_batch_directory(service_dir, structure.safe_batch_id(batch_id))


def number(value: object, digits: int = 2) -> str:
    if value is None:
        return "—"
    result = f"{float(value):,.{digits}f}"
    return result.rstrip("0").rstrip(".") if "." in result else result


def ppm(value: object) -> str:
    return "—" if value is None else number(float(value) * 1_000_000, 2)


def escaped(value: object) -> str:
    return html.escape(str(value or "—"))


def status_text(value: object) -> str:
    labels = {
        "OBSERVED": "관측",
        "NEEDS_REVIEW": "검토 필요",
        "VALID": "유효",
        "NO_SAME_DAY_NORMAL": "동일 날짜 Normal 없음",
        "NORMAL_AMBIGUOUS": "Normal 중복",
        "TEST_AMBIGUOUS": "Test 중복",
        "NO_COMPARISON_NEEDS_REVIEW": "비교 검토 필요",
        "PENDING": "원본 적재 대기",
        "TRUNCATED": "원본 적재 보류",
        "CHANGED": "원본 변경됨",
        "QUARANTINED": "원본 확인 필요",
        "FAILED_RETRYABLE": "원본 적재 실패",
    }
    return labels.get(str(value), str(value))


def table(headers: list[str], rows: list[list[str]], *, empty: str = "표시할 숫자 데이터가 없습니다.") -> str:
    head = "".join(f"<th>{escaped(item)}</th>" for item in headers)
    if not rows:
        body = f"<tr><td class='empty' colspan='{len(headers)}'>{escaped(empty)}</td></tr>"
    else:
        body = "".join("<tr>" + "".join(f"<td>{item}</td>" for item in row) + "</tr>" for row in rows)
    return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def open_database(path: Path) -> sqlite3.Connection:
    if not path.is_file():
        raise ValueError(f"Numeric capture database is missing: {path}")
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys=ON")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS numeric_html_reports (
            workbook_id INTEGER PRIMARY KEY REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
            source_fingerprint TEXT NOT NULL,
            capture_status TEXT NOT NULL,
            renderer_version TEXT NOT NULL,
            report_path TEXT NOT NULL,
            rendered_at TEXT NOT NULL
        )
        """
    )
    connection.commit()
    return connection


def workbook_report(connection: sqlite3.Connection, workbook: sqlite3.Row, relative_path: str) -> str:
    workbook_id = int(workbook["workbook_id"])
    capture_status = str(workbook["capture_status"])
    defect_rows = list(
        connection.execute(
            """
            SELECT measurement_date, condition_role, condition_label, input_value, total_ng_value,
                   computed_ng_rate, fact_status
            FROM numeric_review_facts
            WHERE workbook_id=?
            ORDER BY measurement_date, row_index, fact_id
            """,
            (workbook_id,),
        )
    )
    measurement_rows = list(
        connection.execute(
            """
            SELECT measurement_date, condition_label, sample_value, average_value,
                   minimum_value, maximum_value, fact_status
            FROM measurement_summary_facts
            WHERE workbook_id=?
            ORDER BY measurement_date, row_index, fact_id
            """,
            (workbook_id,),
        )
    )
    comparison_rows = list(
        connection.execute(
            """
            SELECT measurement_date, test_ng_rate, normal_ng_rate, absolute_delta,
                   comparison_status
            FROM test_normal_comparisons
            WHERE workbook_id=?
            ORDER BY measurement_date, comparison_id
            """,
            (workbook_id,),
        )
    )
    unclassified_count = int(
        connection.execute(
            """
            SELECT COUNT(*)
            FROM numeric_table_reviews AS review
            JOIN numeric_table_candidates AS candidate ON candidate.table_id=review.table_id
            JOIN captured_sheets AS sheet ON sheet.sheet_id=candidate.sheet_id
            WHERE sheet.workbook_id=? AND review.extraction_status IN ('NOT_IMPLEMENTED', 'NEEDS_REVIEW')
            """,
            (workbook_id,),
        ).fetchone()[0]
    )
    defect_table = table(
        ["측정일", "조건", "Input", "Total NG", "불량률 (ppm)", "상태"],
        [
            [
                escaped(row["measurement_date"]),
                escaped(row["condition_label"] or row["condition_role"]),
                number(row["input_value"]),
                number(row["total_ng_value"]),
                ppm(row["computed_ng_rate"]),
                escaped(status_text(row["fact_status"])),
            ]
            for row in defect_rows
        ],
        empty="추출된 불량률 숫자 데이터가 없습니다.",
    )
    measurement_table = table(
        ["측정일", "조건", "N", "Average", "Min", "Max", "상태"],
        [
            [
                escaped(row["measurement_date"]),
                escaped(row["condition_label"]),
                number(row["sample_value"], 0),
                number(row["average_value"]),
                number(row["minimum_value"]),
                number(row["maximum_value"]),
                escaped(status_text(row["fact_status"])),
            ]
            for row in measurement_rows
        ],
        empty="추출된 측정 통계 숫자 데이터가 없습니다.",
    )
    comparison_table = table(
        ["측정일", "Test 불량률 (ppm)", "Normal 불량률 (ppm)", "차이 (ppm)", "상태"],
        [
            [
                escaped(row["measurement_date"]),
                ppm(row["test_ng_rate"]),
                ppm(row["normal_ng_rate"]),
                ppm(row["absolute_delta"]),
                escaped(status_text(row["comparison_status"])),
            ]
            for row in comparison_rows
        ],
        empty="동일 날짜의 Test–Normal 비교 결과가 없습니다.",
    )
    pending = ""
    if capture_status != "CAPTURED":
        pending = f"<p class='notice'>원본 숫자 적재 상태: <strong>{escaped(status_text(capture_status))}</strong></p>"
    if unclassified_count:
        pending += f"<p class='notice'>추출 규칙 검토 필요 숫자 표: <strong>{unclassified_count}</strong>개</p>"
    title = Path(relative_path).name
    return f"""<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>{escaped(title)} — 숫자 검토</title>
<style>
body{{font-family:'Segoe UI',sans-serif;margin:22px;color:#162d50;background:#f7f9fc}}main{{max-width:1600px;margin:auto}}h1{{font-size:22px;margin:0 0 18px}}h2{{font-size:17px;margin:28px 0 8px}}table{{border-collapse:collapse;width:100%;background:white;font-size:13px}}th,td{{border:1px solid #d5dfeb;padding:8px;text-align:left;vertical-align:middle}}th{{background:#eaf0f7;color:#183b67}}td.empty{{color:#5d6d7e;text-align:center;padding:16px}}.notice{{padding:10px 12px;background:#fff8e8;border:1px solid #f1dfad;border-radius:6px}}.sub{{color:#53657a;font-size:13px;margin-top:-12px}}</style></head>
<body><main><h1>검토 결과 — {escaped(title)}</h1><p class='sub'>숫자 표 기반 관측값</p>{pending}
<h2>불량률</h2>{defect_table}
<h2>Test–Normal 비교</h2>{comparison_table}
<h2>측정 통계</h2>{measurement_table}
</main></body></html>"""


def run(args: argparse.Namespace) -> int:
    service_dir = Path(args.service_dir).resolve()
    if not service_dir.is_dir():
        raise ValueError(f"Service directory does not exist: {service_dir}")
    batch_dir = resolve_batch_directory(service_dir, args.structure_batch)
    database_path = batch_dir / "numeric-capture.sqlite"
    connection = open_database(database_path)
    report_dir = batch_dir / "numeric-reports"
    try:
        workbooks = list(connection.execute("SELECT * FROM capture_workbooks ORDER BY relative_path"))
        if not workbooks:
            raise ValueError("Numeric capture database has no workbook snapshot rows.")
        index_rows: list[tuple[str, str, str]] = []
        for workbook in workbooks:
            relative_path = str(workbook["relative_path"])
            stable_id = hashlib.sha256(relative_path.casefold().encode("utf-8", errors="replace")).hexdigest()[:20]
            report_path = report_dir / f"{stable_id}.html"
            atomic_write_text(report_path, workbook_report(connection, workbook, relative_path))
            relative_report = report_path.relative_to(batch_dir).as_posix()
            connection.execute(
                """
                INSERT INTO numeric_html_reports(workbook_id, source_fingerprint, capture_status, renderer_version, report_path, rendered_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(workbook_id) DO UPDATE SET
                    source_fingerprint=excluded.source_fingerprint,
                    capture_status=excluded.capture_status,
                    renderer_version=excluded.renderer_version,
                    report_path=excluded.report_path,
                    rendered_at=excluded.rendered_at
                """,
                (
                    workbook["workbook_id"],
                    workbook["snapshot_fingerprint"],
                    workbook["capture_status"],
                    RENDERER_VERSION,
                    relative_report,
                    utc_now(),
                ),
            )
            index_rows.append((relative_path, relative_report, str(workbook["capture_status"])))
        connection.commit()
        statuses = {
            str(row[0]): int(row[1])
            for row in connection.execute("SELECT capture_status, COUNT(*) FROM numeric_html_reports GROUP BY capture_status")
        }
    finally:
        connection.close()
    index_body = "".join(
        f"<tr><td>{escaped(path)}</td><td>{escaped(status_text(status))}</td><td><a href='{html.escape(report)}'>열기</a></td></tr>"
        for path, report, status in index_rows
    )
    index_html = f"""<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>숫자 검토 보고서 목록</title>
<style>body{{font-family:'Segoe UI',sans-serif;margin:24px;color:#162d50}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #d5dfeb;padding:8px;text-align:left}}th{{background:#eaf0f7}}</style></head>
<body><h1>Excel별 숫자 검토 보고서</h1><table><thead><tr><th>Excel</th><th>적재 상태</th><th>보고서</th></tr></thead><tbody>{index_body}</tbody></table></body></html>"""
    atomic_write_text(batch_dir / "numeric-report-index.html", index_html)
    summary = {
        "schemaVersion": "numeric-render-summary-v1",
        "rendererVersion": RENDERER_VERSION,
        "generatedAt": utc_now(),
        "reportCount": len(index_rows),
        "captureStatusCounts": statuses,
        "index": "numeric-report-index.html",
        "reportDirectory": "numeric-reports",
        "limitations": [
            "Reports render only numeric facts and same-date Test–Normal comparisons stored in the batch database.",
            "Source coordinates, formula text, raw cell samples, and DB paths are intentionally omitted from report HTML.",
        ],
    }
    atomic_write_json(batch_dir / "numeric-render-summary.json", summary)
    print(json.dumps({"status": "ok", "batchDirectory": str(batch_dir), "summary": summary}, ensure_ascii=False))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Render one numeric-only HTML review page per workbook in a capture batch.")
    parser.add_argument("--service-dir", required=True)
    parser.add_argument("--structure-batch", required=True)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run(args)
    except (ValueError, RuntimeError) as exc:
        print(f"numeric-renderer error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
