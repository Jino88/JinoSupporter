"""Audit new Capture v2 workbooks against the historical structure catalog.

The audit is deterministic and AI-free.  Its thresholds indicate candidate
handling only; no result is approved until a recipe has passed replay.
"""

from __future__ import annotations

import argparse
import json
import os
import sqlite3
from collections import Counter
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_recipe_matcher import fingerprint_similarity
from inference_data_ai_structure_fingerprint import (
    build_structure_fingerprint_from_database,
    validate_structure_fingerprint,
)


INCREMENTAL_MATCH_AUDIT_SCHEMA_VERSION = "excel-incremental-match-audit-v1"
INCREMENTAL_MATCH_AUDIT_ENGINE_VERSION = "historical-top-k-audit-v1.0"


class IncrementalMatchAuditError(RuntimeError):
    """Raised when audit inputs are incomplete or incompatible."""


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise IncrementalMatchAuditError(f"JSON root must be an object: {path}")
    return value


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_text(
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _open_read_only_database(path: Path) -> sqlite3.Connection:
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise FileNotFoundError(resolved)
    connection = sqlite3.connect(
        resolved.as_uri() + "?mode=ro",
        uri=True,
        timeout=30,
    )
    connection.execute("PRAGMA query_only=ON")
    return connection


def rank_catalog_structures(
    fingerprint: dict[str, Any],
    structures: Iterable[dict[str, Any]],
    *,
    top_k: int = 3,
) -> list[dict[str, Any]]:
    """Rank historical structures without treating a score as approval."""

    validate_structure_fingerprint(fingerprint)
    ranked: list[dict[str, Any]] = []
    for structure in structures:
        representative = structure.get("fingerprint")
        if not isinstance(representative, dict):
            raise IncrementalMatchAuditError(
                "Historical structure requires fingerprint."
            )
        comparison = fingerprint_similarity(fingerprint, representative)
        sample = (structure.get("members") or [{}])[0]
        ranked.append(
            {
                "structureId": str(structure.get("structureId") or ""),
                "fingerprintSha256": str(
                    structure.get("fingerprintSha256") or ""
                ),
                "historicalFileCount": int(structure.get("fileCount") or 0),
                "score": float(comparison["score"]),
                "fingerprintExact": bool(comparison["fingerprintExact"]),
                "components": comparison["components"],
                "sampleFileName": str(sample.get("fileName") or ""),
                "sampleRevisionId": int(sample.get("revisionId") or 0),
            }
        )
    ranked.sort(
        key=lambda item: (
            -float(item["score"]),
            -int(item["historicalFileCount"]),
            str(item["structureId"]),
        )
    )
    return ranked[: max(int(top_k), 0)]


def _candidate_action(best: dict[str, Any] | None, margin: float) -> str:
    if best is None:
        return "NEW_TEMPLATE_REQUIRED"
    if best["fingerprintExact"]:
        return "EXACT_STRUCTURE_REPLAY_CANDIDATE"
    score = float(best["score"])
    if score >= 0.97 and margin >= 0.03:
        return "HIGH_CONFIDENCE_REPLAY_CANDIDATE"
    if score >= 0.90:
        return "TOP_K_AI_REVIEW"
    if score >= 0.75:
        return "VARIANT_PATCH_REVIEW"
    return "NEW_TEMPLATE_REQUIRED"


def audit_incremental_matches(
    *,
    database_path: str | Path,
    historical_catalog_path: str | Path,
    preflight_manifest_path: str | Path,
    top_k: int = 3,
    limit: int | None = None,
    generated_at: str | None = None,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    catalog_path = Path(historical_catalog_path).expanduser().resolve()
    manifest_path = Path(preflight_manifest_path).expanduser().resolve()
    catalog = _read_json(catalog_path)
    manifest = _read_json(manifest_path)
    structures = catalog.get("structures")
    if not isinstance(structures, list) or not structures:
        raise IncrementalMatchAuditError("Historical catalog has no structures.")
    items = [
        item
        for item in manifest.get("items") or []
        if int(item.get("captureRevisionId") or 0) > 0
        and str(item.get("status") or "")
        not in {"CAPTURE_FAILED", "EXCLUDED_FORM"}
    ]
    items.sort(
        key=lambda item: (
            str(item.get("relativePath") or ""),
            str(item.get("contentSha256") or ""),
        )
    )
    if limit is not None:
        items = items[: max(int(limit), 0)]

    results: list[dict[str, Any]] = []
    action_counts: Counter[str] = Counter()
    score_buckets: Counter[str] = Counter()
    with closing(_open_read_only_database(database)) as connection:
        for item in items:
            fingerprint = build_structure_fingerprint_from_database(
                connection,
                int(item["captureRevisionId"]),
            )
            ranked = rank_catalog_structures(
                fingerprint,
                structures,
                top_k=top_k,
            )
            best = ranked[0] if ranked else None
            second_score = float(ranked[1]["score"]) if len(ranked) > 1 else 0.0
            margin = round(
                (float(best["score"]) if best is not None else 0.0)
                - second_score,
                6,
            )
            action = _candidate_action(best, margin)
            action_counts[action] += 1
            score = float(best["score"]) if best is not None else 0.0
            if best is not None and best["fingerprintExact"]:
                bucket = "EXACT"
            elif score >= 0.97:
                bucket = "0.97-1.00"
            elif score >= 0.90:
                bucket = "0.90-0.97"
            elif score >= 0.75:
                bucket = "0.75-0.90"
            elif score >= 0.60:
                bucket = "0.60-0.75"
            else:
                bucket = "BELOW-0.60"
            score_buckets[bucket] += 1
            results.append(
                {
                    "relativePath": str(item.get("relativePath") or ""),
                    "fileName": str(item.get("fileName") or ""),
                    "contentSha256": str(item.get("contentSha256") or ""),
                    "captureRevisionId": int(item["captureRevisionId"]),
                    "preflightFamilyId": str(item.get("formFamilyId") or ""),
                    "preflightRegistryDecision": str(
                        item.get("registryDecision") or ""
                    ),
                    "fingerprintSha256": fingerprint["fingerprintSha256"],
                    "candidateAction": action,
                    "topScore": score,
                    "topTwoMargin": margin,
                    "topCandidates": ranked,
                }
            )
    scores = [float(item["topScore"]) for item in results]
    margins = [float(item["topTwoMargin"]) for item in results]
    return {
        "schemaVersion": INCREMENTAL_MATCH_AUDIT_SCHEMA_VERSION,
        "engineVersion": INCREMENTAL_MATCH_AUDIT_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            "databasePath": str(database),
            "historicalCatalogPath": str(catalog_path),
            "preflightManifestPath": str(manifest_path),
            "historicalStructureCount": len(structures),
            "eligibleWorkbookCount": len(items),
            "topK": int(top_k),
            "limit": limit,
            "databaseMode": "READ_ONLY",
            "aiCalls": 0,
        },
        "summary": {
            "auditedWorkbookCount": len(results),
            "candidateActionCounts": dict(sorted(action_counts.items())),
            "topScoreBuckets": dict(sorted(score_buckets.items())),
            "minimumTopScore": min(scores, default=0.0),
            "averageTopScore": (
                round(sum(scores) / len(scores), 6) if scores else 0.0
            ),
            "maximumTopScore": max(scores, default=0.0),
            "averageTopTwoMargin": (
                round(sum(margins) / len(margins), 6) if margins else 0.0
            ),
            "positiveTopTwoMarginCount": sum(margin > 0 for margin in margins),
        },
        "items": results,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Audit incremental workbooks against historical structures without AI."
    )
    parser.add_argument("--db", required=True)
    parser.add_argument("--historical-catalog", required=True)
    parser.add_argument("--preflight-manifest", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--top-k", type=int, default=3)
    parser.add_argument("--limit", type=int)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    result = audit_incremental_matches(
        database_path=arguments.db,
        historical_catalog_path=arguments.historical_catalog,
        preflight_manifest_path=arguments.preflight_manifest,
        top_k=arguments.top_k,
        limit=arguments.limit,
    )
    output = Path(arguments.out).expanduser().resolve()
    _write_json(output, result)
    print(
        json.dumps(
            {
                "output": str(output),
                "summary": result["summary"],
                "aiCalls": 0,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = [
    "INCREMENTAL_MATCH_AUDIT_ENGINE_VERSION",
    "INCREMENTAL_MATCH_AUDIT_SCHEMA_VERSION",
    "IncrementalMatchAuditError",
    "audit_incremental_matches",
    "rank_catalog_structures",
]
