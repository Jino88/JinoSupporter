from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_DB = Path(r"D:\000. MyWorks\002. DB\process-review.db")
DEFAULT_OUT = Path(r"C:\Users\jhbyun\Desktop\새 폴더 (4)")

sys.path.insert(0, str(SCRIPT_DIR))
import _ai_batch_helper as batch_helper  # noqa: E402
import _xlsx_render  # noqa: E402


@dataclass(frozen=True)
class DatasetItem:
    id: int
    dataset_name: str
    product_type: str
    report_date: str
    created_at: str
    batch_excluded: int
    file_count: int
    file_names: str
    file_size: int


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def ensure_utf8() -> None:
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")
    if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
        sys.stderr.reconfigure(encoding="utf-8")


def read_done(results_path: Path) -> set[str]:
    done: set[str] = set()
    if not results_path.exists():
        return done
    with results_path.open("r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except Exception:
                continue
            name = str(obj.get("datasetName") or "").strip()
            if name:
                done.add(name)
    return done


def append_jsonl(path: Path, obj: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(obj, ensure_ascii=False, separators=(",", ":")) + "\n")


def load_items(db_path: Path) -> list[DatasetItem]:
    con = sqlite3.connect(str(db_path), timeout=60)
    con.execute("PRAGMA busy_timeout=60000")
    try:
        rows = con.execute(
            """
            SELECT
                r.Id,
                r.DatasetName,
                COALESCE(r.ProductType, ''),
                COALESCE(r.ReportDate, ''),
                COALESCE(r.CreatedAt, ''),
                COALESCE(r.BatchExcluded, 0),
                COUNT(f.Id) AS FileCount,
                COALESCE(GROUP_CONCAT(f.FileName, ' | '), '') AS FileNames,
                COALESCE(SUM(f.FileSize), 0) AS FileSize
            FROM RawReports r
            LEFT JOIN RawReportFiles f ON f.DatasetName = r.DatasetName
            GROUP BY r.Id, r.DatasetName, r.ProductType, r.ReportDate, r.CreatedAt, r.BatchExcluded
            ORDER BY r.Id
            """
        ).fetchall()
    finally:
        con.close()

    return [
        DatasetItem(
            id=int(r[0]),
            dataset_name=str(r[1] or ""),
            product_type=str(r[2] or ""),
            report_date=str(r[3] or ""),
            created_at=str(r[4] or ""),
            batch_excluded=int(r[5] or 0),
            file_count=int(r[6] or 0),
            file_names=str(r[7] or ""),
            file_size=int(r[8] or 0),
        )
        for r in rows
    ]


def render_dataset_text(con: sqlite3.Connection, item: DatasetItem, workbook_dir: Path) -> str:
    parts: list[str] = []
    try:
        paths = batch_helper.get_excel_files(con, item.dataset_name, out_dir=str(workbook_dir))
    except Exception as exc:
        paths = []
        parts.append(f"[WORKBOOK_MATERIALIZE_FAILED] {exc}")

    for path in paths:
        parts.append(f"### WORKBOOK: {Path(path).name}")
        try:
            parts.append(_xlsx_render.render_workbook(path))
        except Exception as exc:
            parts.append(f"[WORKBOOK_RENDER_FAILED] {exc}")

    if not paths:
        try:
            paste = batch_helper.get_excel_paste(con, item.dataset_name)
        except Exception as exc:
            paste = None
            parts.append(f"[EXCEL_PASTE_READ_FAILED] {exc}")
        if paste:
            parts.append("### EXCEL_PASTE")
            parts.append(paste)

    return "\n\n".join(p for p in parts if p)


KEYWORD_RE = re.compile(
    r"(purpose|result|content|condition|model|date|line|lot|process|ng|defect|fail|"
    r"hearing|noise|touch|sigma|spl|thd|reliability|improve|before|after|issue|"
    r"sample|qty|rate|judg|decision|check|test|normal|control)",
    re.IGNORECASE,
)


def make_excerpt(text: str, max_chars: int) -> tuple[str, bool]:
    text = (text or "").replace("\x00", "").strip()
    if len(text) <= max_chars:
        return text, False

    head_len = max(2000, int(max_chars * 0.55))
    tail_len = max(1000, int(max_chars * 0.15))
    mid_budget = max_chars - head_len - tail_len - 500

    head = text[:head_len]
    tail = text[-tail_len:]

    selected: list[str] = []
    used = 0
    for line in text[head_len:-tail_len].splitlines():
        if not KEYWORD_RE.search(line):
            continue
        line = line[:700]
        add = len(line) + 1
        if used + add > mid_budget:
            break
        selected.append(line)
        used += add

    excerpt = (
        head
        + "\n\n[...TRUNCATED: keyword-bearing middle lines...]\n"
        + "\n".join(selected)
        + "\n\n[...TRUNCATED: tail...]\n"
        + tail
    )
    return excerpt[:max_chars], True


def read_user_prompt_guidance(out_dir: Path, max_chars: int = 12000) -> str:
    candidates = [
        ("USER_PROMPT_UPDATE_REQUESTS", out_dir / "sample_ready" / "prompt_update_requests.md"),
        ("USER_PROMPT_UPDATE_REQUESTS", out_dir / "prompt_update_requests.md"),
        ("USER_TERM_GUIDANCE", out_dir / "sample_ready" / "ai_term_guidance.md"),
        ("USER_TERM_GUIDANCE", out_dir / "ai_term_guidance.md"),
    ]
    sections: list[str] = []
    for label, path in candidates:
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8-sig", errors="replace").strip()
        if not text or "No pending requests." in text:
            continue
        sections.append(f"## {label}\n{text}")
    merged = "\n\n".join(sections).strip()
    if len(merged) > max_chars:
        return merged[:max_chars] + "\n\n[...USER GUIDANCE TRUNCATED...]"
    return merged


def build_prompt(chunk: list[dict[str, Any]], user_guidance: str = "") -> str:
    header = """
You are doing first-pass AI classification for manufacturing Excel review reports.

This runner intentionally follows the JinoSupporter INPUT DATA (BATCH) prompt cautions:
- The review parameters must be AI-analyzed from Excel/workbook text.
- Do not use review_index, file name, dataset name, source path, sequence number, DB productType, or DB reportDate as a filename-only substitute for workbook-text analysis.
- Use workbook text, sheet names, cell coordinates, merged ranges, inherited merged-cell context, headers, result rows, notes, and evidence cells as the source of truth.
- Do not treat blank follower cells of merged Date/No/Note ranges as missing context.
- Do not invent a decision, defect, model, date, or review item when workbook text does not provide enough evidence.
- If workbook text conflicts with identifier metadata, prioritize workbook text and describe the conflict in uncertainty.

Task:
- Classify each workbook from WORKBOOK_TEXT evidence only.
- File name, dataset name, DB productType, and DB reportDate are identifiers/context only.
- This is a fast first-pass classification, not a full report.
- Korean output values are preferred, but preserve model names, defect names, process names, and dates verbatim when visible.

Return ONLY valid JSON with this exact shape:
{
  "items": [
    {
      "datasetName": "...",
      "reviewPurpose": "...",
      "tags": ["..."],
      "purpose": "...",
      "purposeCode": "1|2|3|4|",
      "targetDefects": ["..."],
      "reviewItems": ["..."],
      "model": "...",
      "date": "...",
      "confidence": 0.0,
      "evidenceSummary": "...",
      "evidenceCells": ["Sheet!A1", "..."],
      "needsDetailedAnalysis": true,
      "uncertainty": "..."
    }
  ]
}

purposeCode guide:
1 = validation/reliability/test result confirmation
2 = defect or NG phenomenon investigation
3 = improvement/change/before-after comparison
4 = monitoring/status/reporting/summary
Use "" only when the workbook text is insufficient.

Rules:
- datasetName must exactly match the input datasetName.
- confidence must be 0..1.
- tags, targetDefects, and reviewItems must be arrays.
- If evidence is weak, keep fields empty and explain uncertainty.
- Do not include markdown fences or prose outside JSON.
""".strip()

    body: list[str] = [header]
    if user_guidance.strip():
        body.extend([
            "\nUSER_CONFIRMED_RULES_FROM_CONTROL_HTML:",
            "The following markdown contains user answers for prior ambiguous classifications and user-defined term explanations.",
            "Use these answers and term explanations as precedent for same or similar model/date/exclusion/uncertainty/terminology cases.",
            "Apply them as normalization and tie-breaker rules, but do not fabricate workbook evidence.",
            user_guidance.strip(),
        ])
    body.append("\nINPUT_WORKBOOKS:")
    for idx, entry in enumerate(chunk, 1):
        meta = entry["meta"]
        body.append(f"\n--- ITEM {idx} ---")
        body.append(json.dumps(meta, ensure_ascii=False, separators=(",", ":")))
        body.append("WORKBOOK_TEXT:")
        body.append(entry["text"])
        body.append("--- END ITEM ---")
    return "\n".join(body)


def extract_json(text: str) -> dict[str, Any]:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw)

    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            return parsed
        if isinstance(parsed, list):
            return {"items": parsed}
    except Exception:
        pass

    start = raw.find("{")
    end = raw.rfind("}")
    if start >= 0 and end > start:
        candidate = raw[start : end + 1]
        parsed = json.loads(candidate)
        if isinstance(parsed, dict):
            return parsed

    start = raw.find("[")
    end = raw.rfind("]")
    if start >= 0 and end > start:
        candidate = raw[start : end + 1]
        parsed = json.loads(candidate)
        if isinstance(parsed, list):
            return {"items": parsed}

    raise ValueError("No JSON object/array found in AI output")


def call_codex(prompt: str, out_path: Path, effort: str, timeout_sec: int) -> dict[str, Any]:
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
    proc = subprocess.run(
        cmd,
        input=prompt,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=timeout_sec,
    )
    output_text = out_path.read_text(encoding="utf-8", errors="replace") if out_path.exists() else proc.stdout
    if proc.returncode != 0:
        raise RuntimeError(f"codex exit {proc.returncode}: {proc.stdout[-2000:]}")
    return extract_json(output_text)


def normalize_ai_item(raw: dict[str, Any], expected_name: str) -> dict[str, Any]:
    def str_list(value: Any) -> list[str]:
        if isinstance(value, list):
            return [str(x).strip() for x in value if str(x).strip()]
        if isinstance(value, str) and value.strip():
            return [value.strip()]
        return []

    def text(key: str) -> str:
        return str(raw.get(key) or "").strip()

    try:
        confidence = float(raw.get("confidence") or 0)
    except Exception:
        confidence = 0.0
    confidence = max(0.0, min(1.0, confidence))

    return {
        "datasetName": text("datasetName") or expected_name,
        "reviewPurpose": text("reviewPurpose"),
        "tags": str_list(raw.get("tags")),
        "purpose": text("purpose"),
        "purposeCode": text("purposeCode"),
        "targetDefects": str_list(raw.get("targetDefects")),
        "reviewItems": str_list(raw.get("reviewItems")),
        "model": text("model"),
        "date": text("date"),
        "confidence": confidence,
        "evidenceSummary": text("evidenceSummary"),
        "evidenceCells": str_list(raw.get("evidenceCells")),
        "needsDetailedAnalysis": bool(raw.get("needsDetailedAnalysis", True)),
        "uncertainty": text("uncertainty"),
    }


def write_summary(out_dir: Path, total: int, done: int, failed: int, started_at: str) -> None:
    summary = {
        "startedAt": started_at,
        "updatedAt": utc_now(),
        "total": total,
        "done": done,
        "failed": failed,
        "remaining": max(0, total - done),
    }
    (out_dir / "classification_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="AI first-pass classifier for process-review DB workbooks.")
    p.add_argument("--db", default=str(DEFAULT_DB))
    p.add_argument("--out-dir", default=str(DEFAULT_OUT))
    p.add_argument("--limit", type=int, default=0, help="Optional max datasets for smoke runs.")
    p.add_argument("--chunk-size", type=int, default=4)
    p.add_argument("--per-item-chars", type=int, default=14000)
    p.add_argument("--max-prompt-chars", type=int, default=90000)
    p.add_argument("--effort", default="low", choices=["minimal", "low", "medium", "high", "xhigh"])
    p.add_argument("--timeout-sec", type=int, default=900)
    p.add_argument("--keep-workbooks", action="store_true")
    return p.parse_args()


def main() -> int:
    ensure_utf8()
    args = parse_args()
    db_path = Path(args.db)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    results_path = out_dir / "classification_results.jsonl"
    failed_path = out_dir / "classification_failed.jsonl"
    progress_path = out_dir / "classification_progress.jsonl"
    prompt_dir = out_dir / "prompts"
    raw_dir = out_dir / "raw_ai_outputs"
    workbook_dir = out_dir / "_workbook_cache"
    prompt_dir.mkdir(parents=True, exist_ok=True)
    raw_dir.mkdir(parents=True, exist_ok=True)
    workbook_dir.mkdir(parents=True, exist_ok=True)

    started_at = utc_now()
    items = load_items(db_path)
    if args.limit > 0:
        items = items[: args.limit]
    total = len(items)
    done_names = read_done(results_path)
    failed_count = 0

    append_jsonl(progress_path, {
        "at": started_at,
        "event": "started",
        "db": str(db_path),
        "outDir": str(out_dir),
        "total": total,
        "alreadyDone": len(done_names),
    })

    con = sqlite3.connect(str(db_path), timeout=60)
    con.execute("PRAGMA busy_timeout=60000")
    current: list[dict[str, Any]] = []
    current_chars = 0
    batch_no = 0

    def flush() -> None:
        nonlocal current, current_chars, batch_no, failed_count, done_names
        if not current:
            return
        batch_no += 1
        user_guidance = read_user_prompt_guidance(out_dir)
        prompt = build_prompt(current, user_guidance)
        prompt_path = prompt_dir / f"classify_batch_{batch_no:04d}.txt"
        raw_path = raw_dir / f"classify_batch_{batch_no:04d}.txt"
        prompt_path.write_text(prompt, encoding="utf-8")
        append_jsonl(progress_path, {
            "at": utc_now(),
            "event": "batch_started",
            "batch": batch_no,
            "count": len(current),
            "datasets": [x["meta"]["datasetName"] for x in current],
        })
        try:
            parsed = call_codex(prompt, raw_path, args.effort, args.timeout_sec)
            ai_items = parsed.get("items") if isinstance(parsed, dict) else None
            if not isinstance(ai_items, list):
                raise ValueError("AI JSON missing items array")
            by_name: dict[str, dict[str, Any]] = {}
            for raw_item in ai_items:
                if isinstance(raw_item, dict):
                    name = str(raw_item.get("datasetName") or "").strip()
                    if name:
                        by_name[name] = raw_item

            for entry in current:
                meta = entry["meta"]
                dataset = meta["datasetName"]
                raw_item = by_name.get(dataset)
                if raw_item is None:
                    raise ValueError(f"AI output missing dataset: {dataset}")
                classification = normalize_ai_item(raw_item, dataset)
                record = {
                    "runStartedAt": started_at,
                    "classifiedAt": utc_now(),
                    "dbPath": str(db_path),
                    **meta,
                    "textChars": entry["textChars"],
                    "textExcerptChars": len(entry["text"]),
                    "textTruncated": entry["textTruncated"],
                    "classification": classification,
                }
                append_jsonl(results_path, record)
                done_names.add(dataset)

            append_jsonl(progress_path, {
                "at": utc_now(),
                "event": "batch_done",
                "batch": batch_no,
                "done": len(done_names),
            })
        except Exception as exc:
            failed_count += len(current)
            msg = str(exc)
            append_jsonl(progress_path, {
                "at": utc_now(),
                "event": "batch_failed",
                "batch": batch_no,
                "reason": msg,
            })
            for entry in current:
                append_jsonl(failed_path, {
                    "at": utc_now(),
                    "batch": batch_no,
                    "datasetName": entry["meta"]["datasetName"],
                    "reason": msg,
                })
        finally:
            write_summary(out_dir, total, len(done_names), failed_count, started_at)
            current = []
            current_chars = 0

    try:
        for item in items:
            if item.dataset_name in done_names:
                continue
            text = render_dataset_text(con, item, workbook_dir)
            if not text.strip():
                failed_count += 1
                append_jsonl(failed_path, {
                    "at": utc_now(),
                    "datasetName": item.dataset_name,
                    "reason": "No extractable workbook text or excel_paste text",
                })
                continue
            excerpt, truncated = make_excerpt(text, int(args.per_item_chars))
            meta = {
                "id": item.id,
                "datasetName": item.dataset_name,
                "fileNames": item.file_names,
                "fileCount": item.file_count,
                "fileSize": item.file_size,
                "dbProductType": item.product_type,
                "dbReportDate": item.report_date,
                "dbCreatedAt": item.created_at,
                "batchExcluded": item.batch_excluded,
            }
            entry = {
                "meta": meta,
                "text": excerpt,
                "textChars": len(text),
                "textTruncated": truncated,
            }
            projected = current_chars + len(excerpt)
            if current and (
                len(current) >= int(args.chunk_size)
                or projected >= int(args.max_prompt_chars)
            ):
                flush()
            current.append(entry)
            current_chars += len(excerpt)
        flush()
    finally:
        con.close()
        if not args.keep_workbooks:
            shutil.rmtree(workbook_dir, ignore_errors=True)

    write_summary(out_dir, total, len(done_names), failed_count, started_at)
    append_jsonl(progress_path, {
        "at": utc_now(),
        "event": "finished",
        "done": len(done_names),
        "failed": failed_count,
        "total": total,
    })
    return 0 if len(done_names) == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
