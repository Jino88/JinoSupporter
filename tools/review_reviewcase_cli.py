"""Review MicroSpeaker ReviewCase drafts one workbook at a time.

This helper is intentionally human-in-the-loop. It shows one Excel-derived
ReviewCase draft, model candidates, AI verification issues, and the exact
command to record the user's decision. It does not modify source Excel files or
the MicroSpeaker SQLite database.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
from typing import Any

from generate_reviewcase_batch import (
    DEFAULT_DB,
    DEFAULT_OUTPUT,
    REPO_ROOT,
    build_model_review,
    connect_readonly,
    load_sheet_rows,
)


DEFAULT_VERIFIED_DIR = REPO_ROOT / "REVIEWCASE_AI_DRAFTS" / "verified"
DEFAULT_DECISIONS = REPO_ROOT / "REVIEWCASE_AI_DRAFTS" / "manual_review" / "reviewcase_cli_decisions.jsonl"


def now_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_jsonl(path: Path, row: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")


def compact(value: Any, max_len: int = 220) -> str:
    text = " ".join(str(value or "").split())
    return text if len(text) <= max_len else text[: max_len - 3].rstrip() + "..."


def repo_path(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else REPO_ROOT / path


def load_manifest(manifest_path: Path) -> dict[str, Any]:
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest not found: {manifest_path}")
    return read_json(manifest_path)


def load_decisions(path: Path) -> dict[int, dict[str, Any]]:
    latest: dict[int, dict[str, Any]] = {}
    if not path.exists():
        return latest
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            row = json.loads(line)
        except json.JSONDecodeError:
            continue
        file_id = row.get("fileId")
        if isinstance(file_id, int):
            latest[file_id] = row
    return latest


def verified_path(file_id: int, verified_dir: Path) -> Path:
    return verified_dir / "files" / f"{file_id}.reviewcase-ai-verification.json"


def load_verified(file_id: int, verified_dir: Path) -> dict[str, Any] | None:
    path = verified_path(file_id, verified_dir)
    if not path.exists():
        return None
    try:
        return read_json(path)
    except (OSError, json.JSONDecodeError):
        return None


def load_draft(entry: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    draft_path = repo_path(str(entry["draftPath"]))
    return draft_path, read_json(draft_path)


def compute_model_review(db_path: Path, draft: dict[str, Any]) -> dict[str, Any]:
    existing = draft.get("modelReview")
    if isinstance(existing, dict):
        return existing

    file_id = int(draft["sourceFileId"])
    with connect_readonly(db_path) as con:
        row = con.execute(
            """
            SELECT file_id, dataset, path, file_name, models, categories
            FROM files
            WHERE file_id=?
            """,
            (file_id,),
        ).fetchone()
        if row is None:
            return {
                "sourceModels": draft.get("sourceModels", ""),
                "selectedModels": draft.get("resolvedModels", []),
                "mappingStatus": draft.get("modelMappingStatus", "missing"),
                "confidence": "low",
                "candidates": [],
                "question": "Model mapping could not be computed because the source file row was not found.",
            }
        rows = load_sheet_rows(con, file_id)
        return build_model_review(row, rows)


def review_priority(draft: dict[str, Any], model_review: dict[str, Any], verified: dict[str, Any] | None) -> int:
    model_status = str(model_review.get("mappingStatus") or "")
    if model_status in {"missing", "needs_user_mapping"}:
        return 0

    ai = (verified or {}).get("aiVerification") or {}
    if ai.get("requiredUserQuestions"):
        return 1
    if ai and not ai.get("approvedForAskAi", False):
        return 2
    if str(draft.get("reviewCaseStatus") or "") in {"needs_review", "needs_ai_verification"}:
        return 3
    return 9


def select_entry(args: argparse.Namespace, manifest: dict[str, Any], decisions: dict[int, dict[str, Any]]) -> dict[str, Any]:
    entries = manifest.get("entries") or []
    if args.file_id:
        for entry in entries:
            if int(entry.get("fileId", -1)) == args.file_id:
                return entry
        raise ValueError(f"File id not found in manifest: {args.file_id}")

    search = (args.search or "").lower().strip()
    best: tuple[int, int, dict[str, Any]] | None = None
    for idx, entry in enumerate(entries):
        file_id = int(entry.get("fileId", -1))
        if not args.include_resolved and file_id in decisions:
            continue
        if search and search not in str(entry.get("fileName", "")).lower():
            continue
        try:
            _, draft = load_draft(entry)
            model_review = compute_model_review(args.db, draft)
        except Exception:
            model_review = {}
            draft = {"reviewCaseStatus": entry.get("status", "")}
        verified = load_verified(file_id, args.verified_dir)
        priority = review_priority(draft, model_review, verified)
        candidate = (priority, idx, entry)
        if best is None or candidate[:2] < best[:2]:
            best = candidate

    if best is None:
        raise ValueError("No matching unresolved ReviewCase draft was found.")
    return best[2]


def print_model_review(model_review: dict[str, Any]) -> None:
    print("Model mapping")
    print(f"  status: {model_review.get('mappingStatus', '')}")
    print(f"  sourceModels: {model_review.get('sourceModels', '')}")
    selected = model_review.get("selectedModels") or []
    print(f"  selectedModels: {', '.join(selected) if selected else '(none)'}")
    if model_review.get("question"):
        print(f"  question: {model_review.get('question')}")

    candidates = model_review.get("candidates") or []
    if candidates:
        print("  candidates:")
        for item in candidates[:8]:
            model = item.get("model", "")
            confidence = item.get("confidence", "")
            sources = ", ".join(item.get("sources") or [])
            ambiguous = " ambiguous" if item.get("ambiguous") else ""
            print(f"    - {model} ({confidence}; {sources}{ambiguous})")
            for evidence in (item.get("evidence") or [])[:2]:
                print(f"      evidence: {compact(evidence, 140)}")


def print_review_case(draft: dict[str, Any]) -> None:
    print("ReviewCase draft")
    print(f"  status: {draft.get('reviewCaseStatus', '')}")
    print(f"  title: {compact(draft.get('reviewTitle') or first_case_value(draft, 'reviewTitle'), 180)}")
    print(f"  purpose: {compact(draft.get('reviewPurpose') or first_case_value(draft, 'reviewPurpose'), 180)}")

    cases = draft.get("reviewCases") or []
    if not cases and any(key in draft for key in ("changedFactors", "outcomes")):
        cases = [draft]
    for case_idx, case in enumerate(cases[:2], start=1):
        print(f"  case {case_idx}: {compact(case.get('reviewTitle', ''), 160)}")
        factors = case.get("changedFactors") or []
        for factor in factors[:4]:
            text = factor.get("changedFactor") or factor.get("changeKey") or ""
            before = factor.get("baselineCondition") or factor.get("beforeCondition") or ""
            after = factor.get("changedCondition") or ", ".join(factor.get("afterConditions") or [])
            print(f"    factor: {compact(text, 150)}")
            if before or after:
                print(f"      compare: {compact(before, 80)} -> {compact(after, 100)}")
        outcomes = case.get("outcomes") or []
        for outcome in outcomes[:6]:
            metric = outcome.get("outcomeMetric") or outcome.get("outcomeKey") or ""
            judgement = outcome.get("judgement") or outcome.get("sourceJudgement") or ""
            summary = outcome.get("resultSummary") or ""
            print(f"    outcome: {compact(metric, 150)} [{compact(judgement, 40)}]")
            if summary:
                print(f"      summary: {compact(summary, 170)}")


def first_case_value(draft: dict[str, Any], key: str) -> str:
    cases = draft.get("reviewCases") or []
    if cases and isinstance(cases[0], dict):
        return str(cases[0].get(key) or "")
    return ""


def print_verified(verified: dict[str, Any] | None) -> None:
    print("AI verification")
    if not verified:
        print("  not generated")
        return
    ai = verified.get("aiVerification") or {}
    print(f"  status: {ai.get('aiReviewCaseStatus', '')} / {ai.get('verificationStatus', '')}")
    print(f"  approvedForAskAi: {ai.get('approvedForAskAi', False)}")
    print(f"  summary: {compact(ai.get('summary', ''), 180)}")
    for issue in (ai.get("issues") or [])[:5]:
        print(f"  issue: {compact(issue, 180)}")
    for question in (ai.get("requiredUserQuestions") or [])[:3]:
        print(f"  user question: {compact(question, 180)}")


def show_entry(args: argparse.Namespace, entry: dict[str, Any]) -> None:
    file_id = int(entry["fileId"])
    draft_path, draft = load_draft(entry)
    model_review = compute_model_review(args.db, draft)
    verified = load_verified(file_id, args.verified_dir)

    print(f"File {file_id}")
    print(f"  name: {entry.get('fileName', '')}")
    print(f"  draft: {draft_path.relative_to(REPO_ROOT)}")
    if verified_path(file_id, args.verified_dir).exists():
        print(f"  verification: {verified_path(file_id, args.verified_dir).relative_to(REPO_ROOT)}")
    print()
    print_model_review(model_review)
    print()
    print_review_case(draft)
    print()
    print_verified(verified)
    print()
    print("Record decision command")
    print(
        "  python tools/review_reviewcase_cli.py --record "
        f"--file-id {file_id} --model \"<canonical model>\" "
        "--review-status verified --note \"<short note>\""
    )


def record_decision(args: argparse.Namespace) -> None:
    if not args.file_id:
        raise ValueError("--record requires --file-id")
    row = {
        "recordedAt": now_iso(),
        "fileId": args.file_id,
        "model": args.model or "",
        "reviewStatus": args.review_status or "",
        "note": args.note or "",
    }
    write_jsonl(args.decisions, row)
    print(json.dumps(row, ensure_ascii=False, indent=2))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, default=DEFAULT_OUTPUT / "reviewcase_batch_manifest.json")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--verified-dir", type=Path, default=DEFAULT_VERIFIED_DIR)
    parser.add_argument("--decisions", type=Path, default=DEFAULT_DECISIONS)
    parser.add_argument("--file-id", type=int, default=0)
    parser.add_argument("--search", default="")
    parser.add_argument("--include-resolved", action="store_true")
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--model", default="")
    parser.add_argument("--review-status", choices=["verified", "needs_review", "excluded", "skip"], default="")
    parser.add_argument("--note", default="")
    args = parser.parse_args()

    if args.record:
        record_decision(args)
        return 0

    manifest = load_manifest(args.manifest)
    decisions = load_decisions(args.decisions)
    entry = select_entry(args, manifest, decisions)
    show_entry(args, entry)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
