"""Run AI verification over generated MicroSpeaker ReviewCase drafts.

Inputs:
- REVIEWCASE_AI_DRAFTS/batch/reviewcase_batch_manifest.json
- REVIEWCASE_AI_DRAFTS/batch/files/*.reviewcase-draft.json

Outputs:
- REVIEWCASE_AI_DRAFTS/verified/files/*.reviewcase-ai-verification.json
- REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_manifest.json
- REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_summary.md

The script never modifies the MicroSpeaker SQLite database or source Excel
files. It reads an OpenAI API key from OPENAI_API_KEY or the app settings DB,
but never writes the key to logs or output files.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.request
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SETTINGS_DB = Path(r"D:\000. MyWorks\002. DB\process-review.db")
DEFAULT_BATCH_MANIFEST = REPO_ROOT / "REVIEWCASE_AI_DRAFTS" / "batch" / "reviewcase_batch_manifest.json"
DEFAULT_OUTPUT = REPO_ROOT / "REVIEWCASE_AI_DRAFTS" / "verified"
MANUAL_DRAFT_DIR = REPO_ROOT / "REVIEWCASE_AI_DRAFTS"
DEFAULT_MODEL = "gpt-5.5"
RESPONSES_URL = "https://api.openai.com/v1/responses"


JSON_OBJECT_RE = re.compile(r"\{.*\}", re.DOTALL)


def now_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def app_setting(settings_db: Path, keys: list[str]) -> str:
    if not settings_db.exists():
        return ""
    try:
        con = sqlite3.connect(f"file:{settings_db.as_posix()}?mode=ro", uri=True)
        con.row_factory = sqlite3.Row
        placeholders = ",".join("?" for _ in keys)
        row = con.execute(
            f"SELECT Value FROM AppSettings WHERE Key IN ({placeholders}) AND length(Value)>0 ORDER BY CASE Key "
            + " ".join(f"WHEN ? THEN {idx}" for idx, _ in enumerate(keys))
            + " END LIMIT 1",
            [*keys, *keys],
        ).fetchone()
        return str(row["Value"]).strip() if row else ""
    except Exception:
        return ""


def resolve_api_key(settings_db: Path) -> str:
    return (
        os.environ.get("OPENAI_API_KEY", "").strip()
        or app_setting(settings_db, ["OpenAI:ApiKey", "Codex:ApiKey"])
    )


def resolve_model(settings_db: Path) -> str:
    return (
        os.environ.get("OPENAI_REVIEWCASE_MODEL", "").strip()
        or app_setting(settings_db, ["OpenAI:ReviewCaseModel", "Codex:ReviewCaseModel", "OpenAI:TranslateModel", "Codex:TranslateModel"])
        or DEFAULT_MODEL
    )


def clean_json_text(text: str) -> str:
    value = (text or "").strip()
    if value.startswith("```"):
        lines = value.splitlines()
        if lines and lines[0].lstrip().startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        value = "\n".join(lines).strip()
    if not value.startswith("{"):
        match = JSON_OBJECT_RE.search(value)
        if match:
            value = match.group(0)
    return value


def extract_output_text(raw: str) -> tuple[str, dict[str, int | None]]:
    doc = json.loads(raw)
    usage_obj = doc.get("usage") if isinstance(doc, dict) else None
    usage = {
        "input_tokens": usage_obj.get("input_tokens") if isinstance(usage_obj, dict) else None,
        "output_tokens": usage_obj.get("output_tokens") if isinstance(usage_obj, dict) else None,
        "total_tokens": usage_obj.get("total_tokens") if isinstance(usage_obj, dict) else None,
    }
    if isinstance(doc, dict) and isinstance(doc.get("output_text"), str):
        return doc["output_text"], usage

    chunks: list[str] = []
    for item in doc.get("output", []) if isinstance(doc, dict) else []:
        for part in item.get("content", []) if isinstance(item, dict) else []:
            if isinstance(part, dict) and part.get("type") == "output_text":
                chunks.append(str(part.get("text") or ""))
    return "".join(chunks), usage


def compact_draft_for_ai(draft: dict[str, Any], max_chars: int) -> dict[str, Any]:
    """Keep all top-level judgement data, but cap oversized repeated subResults."""
    result = json.loads(json.dumps(draft, ensure_ascii=False))
    for review_case in result.get("reviewCases", []) or []:
        for outcome in review_case.get("outcomes", []) or []:
            subresults = outcome.get("subResults")
            if isinstance(subresults, list) and len(subresults) > 30:
                outcome["subResults"] = subresults[:20]
                outcome["subResultsTruncatedForAiReview"] = {
                    "originalCount": len(subresults),
                    "includedHeadCount": 20,
                    "note": "Full draft file keeps all subResults; AI review receives a capped sample plus comparisonRows/evidenceRows.",
                }

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if max_chars > 0 and len(text) > max_chars:
        result["aiInputTruncation"] = {
            "originalChars": len(text),
            "maxChars": max_chars,
            "note": "Tail of compacted draft was omitted from the AI prompt. Use needs_review if this prevents verification.",
        }
        text = json.dumps(result, ensure_ascii=False, indent=2)
        if len(text) > max_chars:
            text = text[:max_chars] + "\n...[TRUNCATED_AI_INPUT]"
            return {"truncatedDraftText": text}
    return result


def build_prompt(draft: dict[str, Any], compacted: dict[str, Any]) -> str:
    calibration = {
        "userConfirmedExample": {
            "fileId": 721,
            "primaryChangeFactor": "Normal vs Test YK MJ vs Test YK Doojin",
            "subgroupPolicy": "Baotou/Boutou remains subgroup context only unless current evidence proves it is the primary changed factor.",
        },
        "generalRules": [
            "Use calibration as reasoning guidance, never as filename or keyword hardcoding.",
            "Candidate pairs, metrics, and measurements are hints. Evidence rows and source context are authority.",
            "Preserve row-level grouping keys such as date, test round, machine, side, cavity, lot, supplier, magnet type, and secondary conditions.",
            "Keep multiple outcome domains under one changed-factor context when the workbook is one validation report.",
            "Return needs_review if changed factor grouping, baseline/test condition, or outcome grouping is ambiguous.",
            "Return excluded only for non-comparable analysis, image-only/empty/reference-only, or no citeable changed-condition evidence.",
        ],
    }
    payload = {
        "task": "Verify this generated ReviewCase draft before Ask AI is allowed to use it as final evidence.",
        "requiredOutput": {
            "sourceFileId": draft.get("sourceFileId"),
            "aiReviewCaseStatus": "verified | needs_review | excluded",
            "verificationStatus": "passed | needs_review | failed",
            "approvedForAskAi": "boolean",
            "confidence": "high | medium | low",
            "summary": "one concise sentence",
            "issues": ["maximum 5 short blocking or non-blocking issues"],
            "requiredUserQuestions": ["maximum 3 short questions only if user judgement is truly needed"],
            "correctionPlan": ["maximum 5 specific fixes needed before verification"],
            "evidencePolicy": "short note on whether evidence row citations are sufficient",
        },
        "calibration": calibration,
        "draft": compacted,
    }
    return (
        "You are an expert manufacturing review-history verifier. "
        "Return only a single compact JSON object. Do not use markdown. "
        "Keep every string concise; do not quote source rows verbatim. "
        "Use at most 5 issues, 3 questions, and 5 correction steps.\n\n"
        + json.dumps(payload, ensure_ascii=False, indent=2)
    )


def verification_json_schema() -> dict[str, Any]:
    return {
        "type": "json_schema",
        "name": "reviewcase_ai_verification",
        "strict": True,
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "sourceFileId": {"type": "integer"},
                "aiReviewCaseStatus": {"type": "string", "enum": ["verified", "needs_review", "excluded"]},
                "verificationStatus": {"type": "string", "enum": ["passed", "needs_review", "failed"]},
                "approvedForAskAi": {"type": "boolean"},
                "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                "summary": {"type": "string"},
                "issues": {"type": "array", "items": {"type": "string"}},
                "requiredUserQuestions": {"type": "array", "items": {"type": "string"}},
                "correctionPlan": {"type": "array", "items": {"type": "string"}},
                "evidencePolicy": {"type": "string"},
            },
            "required": [
                "sourceFileId",
                "aiReviewCaseStatus",
                "verificationStatus",
                "approvedForAskAi",
                "confidence",
                "summary",
                "issues",
                "requiredUserQuestions",
                "correctionPlan",
                "evidencePolicy",
            ],
        },
    }


def effective_draft_path(entry: dict[str, Any]) -> Path:
    file_id = int(entry["fileId"])
    manual_path = MANUAL_DRAFT_DIR / f"{file_id}.reviewcase-draft.json"
    if manual_path.exists():
        return manual_path
    return REPO_ROOT / entry["draftPath"]


def call_openai(api_key: str, model: str, prompt: str, timeout: int, max_output_tokens: int) -> tuple[dict[str, Any], dict[str, int | None], str]:
    body = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": [{"type": "input_text", "text": prompt}],
            }
        ],
        "text": {"format": verification_json_schema()},
        "max_output_tokens": max_output_tokens,
        "store": False,
    }
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        RESPONSES_URL,
        data=data,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        raw = response.read().decode("utf-8", errors="replace")

    output_text, usage = extract_output_text(raw)
    parsed = json.loads(clean_json_text(output_text))
    return parsed, usage, output_text


def verify_with_retry(api_key: str, model: str, prompt: str, timeout: int, max_output_tokens: int, retries: int) -> tuple[dict[str, Any], dict[str, int | None], str]:
    delay = 2.0
    last_error = ""
    for attempt in range(retries + 1):
        try:
            return call_openai(api_key, model, prompt, timeout, max_output_tokens)
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            last_error = f"HTTP {exc.code}: {body[:1000]}"
            if exc.code not in {408, 409, 429, 500, 502, 503, 504}:
                break
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
        if attempt < retries:
            time.sleep(delay)
            delay = min(delay * 2, 30)
    raise RuntimeError(last_error)


def select_entries(manifest: dict[str, Any], file_ids: set[int], statuses: set[str], limit: int, output_dir: Path, force: bool) -> list[dict[str, Any]]:
    entries = []
    for entry in manifest.get("entries", []):
        file_id = int(entry["fileId"])
        if file_ids and file_id not in file_ids:
            continue
        if statuses and entry.get("status") not in statuses:
            continue
        out_path = output_dir / "files" / f"{file_id}.reviewcase-ai-verification.json"
        if out_path.exists() and not force:
            continue
        entries.append(entry)
        if limit > 0 and len(entries) >= limit:
            break
    return entries


def update_manifest(output_dir: Path, model: str) -> dict[str, Any]:
    files_dir = output_dir / "files"
    rows = []
    counts: Counter[str] = Counter()
    total_usage = Counter()
    for path in sorted(files_dir.glob("*.reviewcase-ai-verification.json"), key=lambda p: int(p.name.split(".", 1)[0])):
        data = read_json(path)
        ai = data.get("aiVerification", {})
        status = ai.get("aiReviewCaseStatus") or data.get("status") or "unknown"
        counts[str(status)] += 1
        counts["total"] += 1
        usage = data.get("usage") or {}
        for key in ("input_tokens", "output_tokens", "total_tokens"):
            if isinstance(usage.get(key), int):
                total_usage[key] += usage[key]
        rows.append(
            {
                "fileId": data.get("sourceFileId"),
                "sourceFile": data.get("sourceFile"),
                "status": status,
                "approvedForAskAi": bool(ai.get("approvedForAskAi")),
                "confidence": ai.get("confidence", ""),
                "issueCount": len(ai.get("issues") or []),
                "path": str(path.relative_to(REPO_ROOT)),
            }
        )

    manifest_path = output_dir / "reviewcase_ai_verification_manifest.json"
    summary_path = output_dir / "reviewcase_ai_verification_summary.md"
    manifest = {
        "updatedAt": now_iso(),
        "model": model,
        "counts": dict(counts),
        "usage": dict(total_usage),
        "filesDir": str(files_dir.relative_to(REPO_ROOT)),
        "entries": rows,
    }
    write_json(manifest_path, manifest)

    lines = [
        "# ReviewCase AI Verification Summary",
        "",
        f"Updated at: {manifest['updatedAt']}",
        f"Model: `{model}`",
        "",
        "## Counts",
        "",
    ]
    for key in sorted(counts):
        lines.append(f"- {key}: {counts[key]}")
    if total_usage:
        lines.extend(["", "## Token Usage", ""])
        for key in sorted(total_usage):
            lines.append(f"- {key}: {total_usage[key]}")
    lines.extend(["", "## Outputs", "", f"- Manifest: `{manifest_path.relative_to(REPO_ROOT)}`", f"- Files: `{files_dir.relative_to(REPO_ROOT)}`"])
    summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return manifest


def process_entry(entry: dict[str, Any], args: argparse.Namespace, api_key: str, model: str) -> dict[str, Any]:
    file_id = int(entry["fileId"])
    draft_path = effective_draft_path(entry)
    draft = read_json(draft_path)
    compacted = compact_draft_for_ai(draft, args.max_input_chars)
    prompt = build_prompt(draft, compacted)
    out_path = args.output.resolve() / "files" / f"{file_id}.reviewcase-ai-verification.json"

    ai_result, usage, _ = verify_with_retry(
        api_key,
        model,
        prompt,
        args.timeout,
        args.max_output_tokens,
        args.retries,
    )
    record = {
        "sourceFileId": file_id,
        "sourceFile": draft.get("sourceFile"),
        "sourceDraftPath": str(draft_path.relative_to(REPO_ROOT)),
        "manualDraftUsed": draft_path.parent == MANUAL_DRAFT_DIR,
        "verifiedAt": now_iso(),
        "model": model,
        "aiVerification": ai_result,
        "usage": usage,
    }
    write_json(out_path, record)
    return {
        "fileId": file_id,
        "ok": True,
        "status": ai_result.get("aiReviewCaseStatus"),
        "path": str(out_path),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, default=DEFAULT_BATCH_MANIFEST)
    parser.add_argument("--settings-db", type=Path, default=DEFAULT_SETTINGS_DB)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--file-id", type=int, action="append", default=[])
    parser.add_argument("--status", action="append", default=["needs_ai_verification"])
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--timeout", type=int, default=180)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--sleep", type=float, default=0.0)
    parser.add_argument("--workers", type=int, default=1)
    parser.add_argument("--max-input-chars", type=int, default=120000)
    parser.add_argument("--max-output-tokens", type=int, default=2048)
    args = parser.parse_args()

    api_key = resolve_api_key(args.settings_db)
    if not api_key:
        raise SystemExit("OpenAI API key was not found in OPENAI_API_KEY or app settings DB.")
    model = resolve_model(args.settings_db)

    manifest = read_json(args.manifest)
    output_dir = args.output.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    entries = select_entries(
        manifest,
        set(args.file_id),
        set(args.status or []),
        args.limit,
        output_dir,
        args.force,
    )
    if not entries:
        update_manifest(output_dir, model)
        print(json.dumps({"processed": 0, "message": "no matching entries"}, ensure_ascii=False))
        return 0

    processed = 0
    failures = []
    workers = max(1, min(args.workers, 8))

    if workers == 1:
        for entry in entries:
            file_id = int(entry["fileId"])
            try:
                result = process_entry(entry, args, api_key, model)
                processed += 1
                print(json.dumps({k: result[k] for k in ("fileId", "status", "path")}, ensure_ascii=False), flush=True)
            except Exception as exc:
                failures.append({"fileId": file_id, "error": str(exc)})
                fail_path = output_dir / "failures.jsonl"
                with fail_path.open("a", encoding="utf-8") as fh:
                    fh.write(json.dumps({"fileId": file_id, "error": str(exc), "time": now_iso()}, ensure_ascii=False) + "\n")
                print(json.dumps({"fileId": file_id, "error": str(exc)}, ensure_ascii=False), flush=True)
            if args.sleep > 0:
                time.sleep(args.sleep)
    else:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_map = {}
            for entry in entries:
                future = executor.submit(process_entry, entry, args, api_key, model)
                future_map[future] = int(entry["fileId"])
                if args.sleep > 0:
                    time.sleep(args.sleep)
            for future in as_completed(future_map):
                file_id = future_map[future]
                try:
                    result = future.result()
                    processed += 1
                    print(json.dumps({k: result[k] for k in ("fileId", "status", "path")}, ensure_ascii=False), flush=True)
                except Exception as exc:
                    failures.append({"fileId": file_id, "error": str(exc)})
                    fail_path = output_dir / "failures.jsonl"
                    with fail_path.open("a", encoding="utf-8") as fh:
                        fh.write(json.dumps({"fileId": file_id, "error": str(exc), "time": now_iso()}, ensure_ascii=False) + "\n")
                    print(json.dumps({"fileId": file_id, "error": str(exc)}, ensure_ascii=False), flush=True)

    updated = update_manifest(output_dir, model)
    print(json.dumps({"processed": processed, "failed": len(failures), "counts": updated["counts"]}, ensure_ascii=False))
    return 1 if failures else 0


if __name__ == "__main__":
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        import io

        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    raise SystemExit(main())
