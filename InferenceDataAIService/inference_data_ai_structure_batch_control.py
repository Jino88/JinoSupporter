"""Fail-closed batch control for exact table-structure recipe reuse."""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_table_structure_catalog import (
    table_structure_fingerprint,
)
from inference_data_ai_table_recipe_proposal import semantic_header_signature


RECIPE_REGISTRY_SCHEMA_VERSION = "excel-structure-recipe-registry-v1"
BATCH_CONTROL_SCHEMA_VERSION = "excel-structure-reuse-batch-control-v1"
BATCH_CONTROL_ENGINE_VERSION = "structure-budget-controller-v1.0"

_VERIFIED_REPLAY_STATUS = (
    "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_NEEDS_CANONICAL_REVIEW"
)


class StructureBatchControlError(RuntimeError):
    """Raised when registry identity or AI budget state is unsafe."""


def _read_json(path: str | Path) -> dict[str, Any]:
    source = Path(path).expanduser().resolve()
    with source.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise StructureBatchControlError(
            f"JSON root must be an object: {source}"
        )
    return value


def _write_json(path: str | Path, value: Any) -> Path:
    target = Path(path).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.{os.getpid()}.tmp")
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
    os.replace(temporary, target)
    return target


def _json_files(path: str | Path | None) -> list[Path]:
    if path is None:
        return []
    root = Path(path).expanduser().resolve()
    return sorted(root.glob("*.json")) if root.is_dir() else []


def build_recipe_registry(
    *,
    priority_report: dict[str, Any],
    recipe_root: str | Path,
    replay_root: str | Path,
    telemetry_root: str | Path,
    generated_at: str | None = None,
) -> dict[str, Any]:
    """Register only recipes whose exact-structure replay fully passed."""

    queue_by_structure = {
        str(item.get("tableStructureId") or ""): item
        for item in priority_report.get("queue") or []
    }
    replay_by_recipe = {
        str(value.get("recipeId") or ""): (path, value)
        for path in _json_files(replay_root)
        for value in [_read_json(path)]
        if value.get("recipeId")
    }
    telemetry_by_structure = {
        str(value.get("tableStructureId") or ""): (path, value)
        for path in _json_files(telemetry_root)
        for value in [_read_json(path)]
        if value.get("tableStructureId")
    }
    entries: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []
    for recipe_path in _json_files(recipe_root):
        recipe = _read_json(recipe_path)
        structure_id = str(
            (recipe.get("match") or {}).get("tableStructureId") or ""
        )
        recipe_id = str(recipe.get("recipeId") or "")
        reasons: list[str] = []
        queue_item = queue_by_structure.get(structure_id)
        replay_pair = replay_by_recipe.get(recipe_id)
        telemetry_pair = telemetry_by_structure.get(structure_id)
        decision = recipe.get("decision") or {}
        decision_ai_calls = int(decision.get("aiCallCount") or 0)
        decision_source = str(decision.get("source") or "")
        if queue_item is None:
            reasons.append("STRUCTURE_NOT_IN_PRIORITY_REPORT")
        elif str((recipe.get("match") or {}).get("fingerprintSha256") or "") != str(
            queue_item.get("fingerprintSha256") or ""
        ):
            reasons.append("FINGERPRINT_MISMATCH")
        elif list(
            (recipe.get("match") or {}).get(
                "semanticHeaderSignature"
            )
            or []
        ) != list(
            (queue_item or {}).get("semanticHeaderSignature") or []
        ):
            reasons.append("SEMANTIC_HEADER_SIGNATURE_MISMATCH")
        if recipe.get("status") != _VERIFIED_REPLAY_STATUS:
            reasons.append("RECIPE_REPLAY_NOT_VERIFIED")
        if replay_pair is None:
            reasons.append("REPLAY_ARTIFACT_MISSING")
        else:
            replay = replay_pair[1]
            if (
                replay.get("status") != _VERIFIED_REPLAY_STATUS
                or int((replay.get("summary") or {}).get("failed") or 0) != 0
                or int((replay.get("summary") or {}).get("passed") or 0)
                != int((replay.get("summary") or {}).get("memberCount") or 0)
            ):
                reasons.append("REPLAY_ARTIFACT_FAILED")
        if decision_ai_calls == 0:
            if decision_source not in {
                "HISTORICAL_989_CONSENSUS",
                "SOURCE_OWNED_SINGLE_MEASURE_HEADER",
                "VERIFIED_SIGNATURE_PROPAGATION",
            }:
                reasons.append("ZERO_AI_DECISION_PROVENANCE_UNSAFE")
            if (
                decision_source == "VERIFIED_SIGNATURE_PROPAGATION"
                and not str(decision.get("sourceRecipeId") or "")
            ):
                reasons.append("PROPAGATION_SOURCE_RECIPE_MISSING")
        elif telemetry_pair is None:
            reasons.append("DECISION_TELEMETRY_MISSING")
        else:
            telemetry = telemetry_pair[1]
            if (
                telemetry.get("status") != "SUCCEEDED"
                or int(telemetry.get("aiCallsAttempted") or 0) != 1
                or int(telemetry.get("aiCallsSucceeded") or 0) != 1
                or int(telemetry.get("retryCount") or 0) != 0
            ):
                reasons.append("DECISION_TELEMETRY_UNSAFE")
        if reasons:
            rejected.append(
                {
                    "recipeFile": str(recipe_path),
                    "recipeId": recipe_id,
                    "tableStructureId": structure_id,
                    "reasons": reasons,
                }
            )
            continue
        replay_path, replay = replay_pair
        telemetry_path = telemetry_pair[0] if telemetry_pair else None
        entries.append(
            {
                "tableStructureId": structure_id,
                "fingerprintSha256": str(
                    (recipe.get("match") or {}).get(
                        "fingerprintSha256"
                    )
                    or ""
                ),
                "semanticHeaderSha256": str(
                    (recipe.get("match") or {}).get(
                        "semanticHeaderSha256"
                    )
                    or ""
                ),
                "semanticHeaderSignature": list(
                    (recipe.get("match") or {}).get(
                        "semanticHeaderSignature"
                    )
                    or []
                ),
                "recipeId": recipe_id,
                "recipeVersion": int(recipe.get("recipeVersion") or 0),
                "recipeFile": str(recipe_path),
                "replayFile": str(replay_path),
                "telemetryFile": (
                    str(telemetry_path) if telemetry_path else None
                ),
                "status": "EXACT_REPLAY_READY_NEEDS_CANONICAL_REVIEW",
                "tableCount": int(
                    (replay.get("summary") or {}).get("memberCount") or 0
                ),
                "workbookCount": int(
                    (queue_item or {}).get("workbookCount") or 0
                ),
                "deterministicFactCount": int(
                    (replay.get("summary") or {}).get(
                        "deterministicFactCount"
                    )
                    or 0
                ),
                "deterministicCellFactCount": int(
                    (replay.get("summary") or {}).get(
                        "deterministicCellFactCount"
                    )
                    or 0
                ),
                "decisionSource": decision_source,
                "decisionAiCalls": decision_ai_calls,
            }
        )
    entries.sort(key=lambda item: str(item["tableStructureId"]))
    return {
        "schemaVersion": RECIPE_REGISTRY_SCHEMA_VERSION,
        "engineVersion": BATCH_CONTROL_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "policy": {
            "matchMode": "EXACT_TABLE_STRUCTURE_ONLY",
            "canonicalReviewRequired": True,
            "fileLevelAiEnabled": False,
            "aiMayWriteValues": False,
        },
        "summary": {
            "registeredRecipeCount": len(entries),
            "registeredTableCount": sum(
                int(item["tableCount"]) for item in entries
            ),
            "registeredWorkbookReferences": sum(
                int(item["workbookCount"]) for item in entries
            ),
            "rejectedRecipeCount": len(rejected),
        },
        "recipes": entries,
        "rejected": rejected,
    }


def build_batch_control(
    *,
    priority_report: dict[str, Any],
    registry: dict[str, Any],
    telemetry_root: str | Path,
    max_ai_calls: int,
    decision_root: str | Path | None = None,
    generated_at: str | None = None,
) -> dict[str, Any]:
    """Allocate remaining AI calls by structure without starting any call."""

    budget = max(int(max_ai_calls), 0)
    telemetry_values = [_read_json(path) for path in _json_files(telemetry_root)]
    attempted_by_structure: dict[str, int] = {}
    succeeded_by_structure: dict[str, int] = {}
    failed_structures: set[str] = set()
    for telemetry in telemetry_values:
        structure_id = str(telemetry.get("tableStructureId") or "")
        if not structure_id:
            continue
        attempted = int(telemetry.get("aiCallsAttempted") or 0)
        succeeded = int(telemetry.get("aiCallsSucceeded") or 0)
        attempted_by_structure[structure_id] = (
            attempted_by_structure.get(structure_id, 0) + attempted
        )
        succeeded_by_structure[structure_id] = (
            succeeded_by_structure.get(structure_id, 0) + succeeded
        )
        if telemetry.get("status") == "FAILED":
            failed_structures.add(structure_id)
    duplicated = sorted(
        structure_id
        for structure_id, calls in attempted_by_structure.items()
        if calls > 1
    )
    if duplicated:
        raise StructureBatchControlError(
            "Structure AI call limit already violated: "
            + ", ".join(duplicated)
        )
    consumed = sum(attempted_by_structure.values())
    remaining = max(budget - consumed, 0)
    registered = {
        str(item.get("tableStructureId") or ""): item
        for item in registry.get("recipes") or []
    }
    explicit_quarantines: set[str] = set()
    if decision_root is not None:
        decisions_by_structure: dict[str, set[str]] = {}
        for path in _json_files(decision_root):
            decision = _read_json(path)
            structure_id = str(
                decision.get("targetTableStructureId") or ""
            )
            if not structure_id:
                continue
            decisions_by_structure.setdefault(structure_id, set()).add(
                str(decision.get("decision") or "")
            )
        conflicting = sorted(
            structure_id
            for structure_id, modes in decisions_by_structure.items()
            if len(modes) > 1
        )
        if conflicting:
            raise StructureBatchControlError(
                "Conflicting structure decisions: " + ", ".join(conflicting)
            )
        explicit_quarantines = {
            structure_id
            for structure_id, modes in decisions_by_structure.items()
            if modes == {"QUARANTINE"}
        }
    actions: list[dict[str, Any]] = []
    newly_authorized = 0
    for item in priority_report.get("queue") or []:
        structure_id = str(item.get("tableStructureId") or "")
        if structure_id in registered:
            action = "EXACT_RECIPE_REPLAY_READY"
            reason = "VERIFIED_RECIPE_REGISTERED"
        elif structure_id in failed_structures:
            action = "QUARANTINED_NO_RETRY"
            reason = "PRIOR_AI_CALL_FAILED"
        elif attempted_by_structure.get(structure_id, 0) >= 1:
            action = "DECISION_COMPLETE_REPLAY_NOT_REGISTERED"
            reason = "STRUCTURE_AI_CALL_ALREADY_CONSUMED"
        elif structure_id in explicit_quarantines:
            action = "SOURCE_PRECHECK_QUARANTINED"
            reason = "EXPLICIT_SOURCE_OWNED_QUARANTINE_DECISION"
        elif item.get("status") != "PROPOSAL_READY":
            action = "MANUAL_REVIEW_REQUIRED"
            reason = ",".join(item.get("safetyReasons") or [])
        elif newly_authorized < remaining:
            action = "AI_DECISION_AUTHORIZED"
            reason = "WITHIN_STRUCTURE_AND_RUN_BUDGET"
            newly_authorized += 1
        else:
            action = "AI_BUDGET_WAIT"
            reason = "RUN_AI_BUDGET_EXHAUSTED"
        actions.append(
            {
                "rank": int(item.get("rank") or 0),
                "tableStructureId": structure_id,
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "action": action,
                "reason": reason,
                "structureAiCallsAttempted": attempted_by_structure.get(
                    structure_id,
                    0,
                ),
                "structureAiCallsSucceeded": succeeded_by_structure.get(
                    structure_id,
                    0,
                ),
            }
        )
    counts: dict[str, int] = {}
    for item in actions:
        counts[item["action"]] = counts.get(item["action"], 0) + 1
    return {
        "schemaVersion": BATCH_CONTROL_SCHEMA_VERSION,
        "engineVersion": BATCH_CONTROL_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "policy": {
            "oldFileLevelPipelineEnabled": False,
            "fileLevelAiEnabled": False,
            "maxAiCallsPerStructure": 1,
            "retryCount": 0,
            "stopWhenRunBudgetExhausted": True,
            "exactFingerprintRequiredForReplay": True,
            "canonicalReviewRequiredBeforeApproval": True,
        },
        "budget": {
            "maxAiCalls": budget,
            "consumedAiCalls": consumed,
            "remainingAiCalls": remaining,
            "newlyAuthorizedAiCalls": newly_authorized,
        },
        "summary": {
            "structureCount": len(actions),
            "actionCounts": dict(sorted(counts.items())),
            "registeredRecipeCount": len(registered),
            "fileLevelAiCalls": 0,
        },
        "actions": actions,
    }


def match_registered_recipe(
    table: dict[str, Any],
    registry: dict[str, Any],
) -> dict[str, Any]:
    """Resolve an exact registered recipe without using AI."""

    fingerprint = table_structure_fingerprint(table)
    matches = [
        item
        for item in registry.get("recipes") or []
        if str(item.get("fingerprintSha256") or "")
        == str(fingerprint.get("fingerprintSha256") or "")
        and list(item.get("semanticHeaderSignature") or [])
        == semantic_header_signature(table)
    ]
    if len(matches) == 1:
        return {
            "status": "EXACT_RECIPE_MATCH",
            "aiCalls": 0,
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "recipe": matches[0],
        }
    if len(matches) > 1:
        raise StructureBatchControlError(
            "Multiple registered recipes share one exact fingerprint."
        )
    return {
        "status": "NO_REGISTERED_RECIPE",
        "aiCalls": 0,
        "fingerprintSha256": fingerprint["fingerprintSha256"],
        "recipe": None,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build fail-closed structure recipe batch state."
    )
    parser.add_argument("--priority-report", required=True)
    parser.add_argument("--recipe-root", required=True)
    parser.add_argument("--replay-root", required=True)
    parser.add_argument("--telemetry-root", required=True)
    parser.add_argument("--decision-root")
    parser.add_argument("--max-ai-calls", type=int, default=1)
    parser.add_argument("--registry-output", required=True)
    parser.add_argument("--control-output", required=True)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    report = _read_json(arguments.priority_report)
    registry = build_recipe_registry(
        priority_report=report,
        recipe_root=arguments.recipe_root,
        replay_root=arguments.replay_root,
        telemetry_root=arguments.telemetry_root,
    )
    control = build_batch_control(
        priority_report=report,
        registry=registry,
        telemetry_root=arguments.telemetry_root,
        max_ai_calls=arguments.max_ai_calls,
        decision_root=arguments.decision_root,
    )
    _write_json(arguments.registry_output, registry)
    _write_json(arguments.control_output, control)
    print(
        json.dumps(
            {
                "registry": registry["summary"],
                "budget": control["budget"],
                "actions": control["summary"]["actionCounts"],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
