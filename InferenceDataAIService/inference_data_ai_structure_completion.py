"""Resume and finish the bounded structure-recipe queue."""

from __future__ import annotations

import argparse
import json
import os
import time
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_structure_batch_control import build_recipe_registry
from inference_data_ai_table_recipe_proposal import (
    _read_json,
    _write_json,
    adapt_decision_to_priority_item,
    compile_structure_recipe,
    load_representative_captured_cells,
    replay_structure_recipe,
    run_codex_table_recipe_decision,
)


COMPLETION_STATE_SCHEMA_VERSION = "excel-structure-completion-state-v1"
COMPLETION_ENGINE_VERSION = "bounded-structure-completion-v1.0"
VERIFIED_REPLAY_STATUS = (
    "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_NEEDS_CANONICAL_REVIEW"
)


class StructureCompletionError(RuntimeError):
    """Raised when the completion run cannot preserve budget invariants."""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _decision_candidates(
    decision_root: Path,
    structure_id: str,
) -> list[Path]:
    return sorted(decision_root.glob(structure_id + "*.json"))


def _decision_for_structure(
    decision_root: Path,
    structure_id: str,
) -> dict[str, Any] | None:
    candidates = _decision_candidates(decision_root, structure_id)
    valid: list[dict[str, Any]] = []
    for path in candidates:
        value = _read_json(path)
        if value.get("targetTableStructureId") == structure_id:
            valid.append(value)
    if len(valid) > 1:
        modes = {str(value.get("decision") or "") for value in valid}
        if len(modes) > 1:
            raise StructureCompletionError(
                f"Conflicting decisions for {structure_id}."
            )
    return valid[0] if valid else None


def _telemetry_by_structure(
    telemetry_root: Path,
) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for path in sorted(telemetry_root.glob("*.json")):
        value = _read_json(path)
        structure_id = str(value.get("tableStructureId") or "")
        if not structure_id:
            continue
        if structure_id in result:
            raise StructureCompletionError(
                "More than one AI telemetry artifact exists for "
                + structure_id
            )
        result[structure_id] = value
    return result


def _state_template(
    *,
    priority_report_path: Path,
    max_total_ai_calls: int,
) -> dict[str, Any]:
    return {
        "schemaVersion": COMPLETION_STATE_SCHEMA_VERSION,
        "engineVersion": COMPLETION_ENGINE_VERSION,
        "startedAt": _now(),
        "updatedAt": _now(),
        "status": "RUNNING",
        "inputs": {
            "priorityReport": str(priority_report_path),
            "maxTotalAiCalls": int(max_total_ai_calls),
            "maxAiCallsPerStructure": 1,
            "retryCount": 0,
        },
        "outcomes": {},
        "events": [],
        "summary": {},
    }


def _event(
    state: dict[str, Any],
    *,
    structure_id: str,
    status: str,
    detail: str,
) -> None:
    state["events"].append(
        {
            "at": _now(),
            "tableStructureId": structure_id,
            "status": status,
            "detail": detail,
        }
    )
    state["events"] = state["events"][-500:]


def _summarize(
    state: dict[str, Any],
    *,
    priority_report: dict[str, Any],
    telemetry: dict[str, dict[str, Any]],
    registry: dict[str, Any],
) -> None:
    queue = list(priority_report.get("queue") or [])
    queue_ids = {
        str(value.get("tableStructureId") or "")
        for value in queue
        if str(value.get("tableStructureId") or "")
    }
    completed_ids = set(state["outcomes"]) & queue_ids
    outcome_counts = Counter(
        str(state["outcomes"][structure_id].get("status") or "")
        for structure_id in completed_ids
    )
    registered_ids = {
        str(value.get("tableStructureId") or "")
        for value in registry.get("recipes") or []
    }
    state["summary"] = {
        "queueStructureCount": len(queue),
        "completedStructureCount": len(completed_ids),
        "unresolvedStructureCount": len(queue) - len(completed_ids),
        "outcomeCounts": dict(sorted(outcome_counts.items())),
        "registeredRecipeCount": len(registered_ids),
        "registeredTableCount": int(
            (registry.get("summary") or {}).get("registeredTableCount") or 0
        ),
        "registeredWorkbookReferences": int(
            (registry.get("summary") or {}).get(
                "registeredWorkbookReferences"
            )
            or 0
        ),
        "aiCallsAttempted": sum(
            int(value.get("aiCallsAttempted") or 0)
            for value in telemetry.values()
        ),
        "aiCallsSucceeded": sum(
            int(value.get("aiCallsSucceeded") or 0)
            for value in telemetry.values()
        ),
        "aiCallsFailed": sum(
            value.get("status") == "FAILED"
            for value in telemetry.values()
        ),
        "retryCount": sum(
            int(value.get("retryCount") or 0)
            for value in telemetry.values()
        ),
        "fileLevelAiCalls": 0,
    }
    state["updatedAt"] = _now()


def _write_state(
    state_path: Path,
    state: dict[str, Any],
    *,
    priority_report: dict[str, Any],
    telemetry_root: Path,
    registry: dict[str, Any],
) -> None:
    _summarize(
        state,
        priority_report=priority_report,
        telemetry=_telemetry_by_structure(telemetry_root),
        registry=registry,
    )
    _write_json(state_path, state)


def _recipe_paths(
    structure_id: str,
    *,
    recipe_root: Path,
    replay_root: Path,
) -> tuple[Path, Path]:
    suffix = structure_id.removeprefix("table-structure-")
    return (
        recipe_root / f"structure-recipe-{suffix}.json",
        replay_root / f"structure-recipe-{suffix}.replay.json",
    )


def _refresh_registry(
    *,
    priority_report: dict[str, Any],
    recipe_root: Path,
    replay_root: Path,
    telemetry_root: Path,
    registry_path: Path,
) -> dict[str, Any]:
    registry = build_recipe_registry(
        priority_report=priority_report,
        recipe_root=recipe_root,
        replay_root=replay_root,
        telemetry_root=telemetry_root,
    )
    _write_json(registry_path, registry)
    return registry


def _registered_source_decisions(
    *,
    registry: dict[str, Any],
    queue_by_id: dict[str, dict[str, Any]],
    decision_root: Path,
) -> dict[str, list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]]:
    result: dict[
        str,
        list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]],
    ] = defaultdict(list)
    for entry in registry.get("recipes") or []:
        structure_id = str(entry.get("tableStructureId") or "")
        item = queue_by_id.get(structure_id)
        decision = _decision_for_structure(decision_root, structure_id)
        if item is None or decision is None:
            continue
        result[str(item.get("semanticHeaderSha256") or "")].append(
            (entry, item, decision)
        )
    return result


def _record_registered_outcomes(
    state: dict[str, Any],
    registry: dict[str, Any],
) -> None:
    for entry in registry.get("recipes") or []:
        structure_id = str(entry["tableStructureId"])
        current_status = str(
            (state["outcomes"].get(structure_id) or {}).get("status") or ""
        )
        if current_status.startswith("REGISTERED"):
            continue
        decision_source = str(entry.get("decisionSource") or "")
        if decision_source == "VERIFIED_SIGNATURE_PROPAGATION":
            outcome_status = "REGISTERED_PROPAGATED"
        elif decision_source == "BOUNDED_AI_STRUCTURE_DECISION":
            outcome_status = "REGISTERED_AI_REPLAY"
        else:
            outcome_status = "REGISTERED"
        state["outcomes"][structure_id] = {
            "status": outcome_status,
            "decisionSource": decision_source,
            "decisionAiCalls": int(entry.get("decisionAiCalls") or 0),
            "tableCount": int(entry.get("tableCount") or 0),
            "workbookCount": int(entry.get("workbookCount") or 0),
            "recipeId": str(entry.get("recipeId") or ""),
        }


def complete_structure_queue(
    *,
    priority_report_path: str | Path,
    recipe_root: str | Path,
    replay_root: str | Path,
    decision_root: str | Path,
    telemetry_root: str | Path,
    quarantine_root: str | Path,
    registry_path: str | Path,
    state_path: str | Path,
    max_total_ai_calls: int,
    reasoning_effort: str = "low",
    timeout_seconds: int = 600,
) -> dict[str, Any]:
    priority_path = Path(priority_report_path).expanduser().resolve()
    recipes = Path(recipe_root).expanduser().resolve()
    replays = Path(replay_root).expanduser().resolve()
    decisions = Path(decision_root).expanduser().resolve()
    telemetry = Path(telemetry_root).expanduser().resolve()
    quarantine = Path(quarantine_root).expanduser().resolve()
    registry_file = Path(registry_path).expanduser().resolve()
    state_file = Path(state_path).expanduser().resolve()
    for root in (
        recipes,
        replays,
        decisions,
        telemetry,
        quarantine,
    ):
        root.mkdir(parents=True, exist_ok=True)
    priority_report = _read_json(priority_path)
    queue = list(priority_report.get("queue") or [])
    queue_by_id = {
        str(item["tableStructureId"]): item for item in queue
    }
    if state_file.is_file():
        state = _read_json(state_file)
        if state.get("schemaVersion") != COMPLETION_STATE_SCHEMA_VERSION:
            raise StructureCompletionError("Invalid completion state schema.")
        state["status"] = "RUNNING"
        state.pop("completedAt", None)
    else:
        state = _state_template(
            priority_report_path=priority_path,
            max_total_ai_calls=max_total_ai_calls,
        )
    state.setdefault("inputs", {})
    state["inputs"]["priorityReport"] = str(priority_path)
    state["inputs"]["maxTotalAiCalls"] = int(max_total_ai_calls)
    state["inputs"]["maxAiCallsPerStructure"] = 1
    state["inputs"]["retryCount"] = 0
    registry = _refresh_registry(
        priority_report=priority_report,
        recipe_root=recipes,
        replay_root=replays,
        telemetry_root=telemetry,
        registry_path=registry_file,
    )
    _record_registered_outcomes(state, registry)
    telemetry_values = _telemetry_by_structure(telemetry)
    for structure_id, value in telemetry_values.items():
        if (
            value.get("status") == "FAILED"
            and structure_id not in state["outcomes"]
        ):
            item = queue_by_id.get(structure_id) or {}
            state["outcomes"][structure_id] = {
                "status": "QUARANTINED_AI_FAILURE_NO_RETRY",
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "error": str(value.get("error") or ""),
            }
    for item in queue:
        structure_id = str(item["tableStructureId"])
        if (
            item.get("status") != "PROPOSAL_READY"
            and structure_id not in state["outcomes"]
        ):
            state["outcomes"][structure_id] = {
                "status": "QUARANTINED_PRECHECK",
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "reasons": list(item.get("safetyReasons") or []),
            }
    _write_state(
        state_file,
        state,
        priority_report=priority_report,
        telemetry_root=telemetry,
        registry=registry,
    )

    for item in queue:
        structure_id = str(item["tableStructureId"])
        if structure_id in state["outcomes"]:
            continue
        existing_decision = _decision_for_structure(decisions, structure_id)
        if (
            existing_decision is not None
            and existing_decision.get("decision") == "QUARANTINE"
        ):
            telemetry_value = _telemetry_by_structure(telemetry).get(
                structure_id
            ) or {}
            decision_ai_calls = int(
                telemetry_value.get("aiCallsAttempted") or 0
            )
            state["outcomes"][structure_id] = {
                "status": (
                    "QUARANTINED_AI_DECISION"
                    if decision_ai_calls
                    else "QUARANTINED_SOURCE_PRECHECK"
                ),
                "decisionAiCalls": decision_ai_calls,
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "rationale": str(
                    existing_decision.get("rationale") or ""
                ),
            }
            _event(
                state,
                structure_id=structure_id,
                status=state["outcomes"][structure_id]["status"],
                detail=(
                    "existing explicit quarantine decision; "
                    f"AI calls={decision_ai_calls}; retry=0"
                ),
            )
            _write_state(
                state_file,
                state,
                priority_report=priority_report,
                telemetry_root=telemetry,
                registry=registry,
            )
            continue

        registry = _refresh_registry(
            priority_report=priority_report,
            recipe_root=recipes,
            replay_root=replays,
            telemetry_root=telemetry,
            registry_path=registry_file,
        )
        sources = _registered_source_decisions(
            registry=registry,
            queue_by_id=queue_by_id,
            decision_root=decisions,
        ).get(str(item.get("semanticHeaderSha256") or ""), [])
        propagated = False
        propagation_errors: list[str] = []
        for source_entry, source_item, source_decision in sources:
            if source_item["tableStructureId"] == structure_id:
                continue
            try:
                decision = adapt_decision_to_priority_item(
                    source_decision,
                    source_item=source_item,
                    target_item=item,
                )
                recipe = compile_structure_recipe(
                    decision,
                    priority_item=item,
                    representative_captured_cells=(
                        load_representative_captured_cells(
                            priority_report,
                            item,
                        )
                    ),
                    decision_ai_calls=0,
                    decision_source="VERIFIED_SIGNATURE_PROPAGATION",
                )
                recipe["decision"]["sourceRecipeId"] = source_entry[
                    "recipeId"
                ]
                recipe["decision"]["sourceTableStructureId"] = source_item[
                    "tableStructureId"
                ]
                replay = replay_structure_recipe(
                    recipe=recipe,
                    priority_report=priority_report,
                )
                if replay.get("status") != VERIFIED_REPLAY_STATUS:
                    raise StructureCompletionError(
                        "Propagated recipe replay did not fully pass."
                    )
                recipe_file, replay_file = _recipe_paths(
                    structure_id,
                    recipe_root=recipes,
                    replay_root=replays,
                )
                decision_file = decisions / (
                    structure_id + ".propagated.json"
                )
                _write_json(decision_file, decision)
                _write_json(recipe_file, recipe)
                _write_json(replay_file, replay)
                state["outcomes"][structure_id] = {
                    "status": "REGISTERED_PROPAGATED",
                    "decisionSource": "VERIFIED_SIGNATURE_PROPAGATION",
                    "decisionAiCalls": 0,
                    "sourceRecipeId": source_entry["recipeId"],
                    "tableCount": int(item.get("tableCount") or 0),
                    "workbookCount": int(item.get("workbookCount") or 0),
                    "recipeId": recipe["recipeId"],
                }
                _event(
                    state,
                    structure_id=structure_id,
                    status="REGISTERED_PROPAGATED",
                    detail=(
                        "AI 0; source recipe "
                        + str(source_entry["recipeId"])
                    ),
                )
                propagated = True
                break
            except Exception as exc:
                propagation_errors.append(
                    f"{source_item['tableStructureId']}:"
                    f"{type(exc).__name__}:{exc}"
                )
        if propagated:
            registry = _refresh_registry(
                priority_report=priority_report,
                recipe_root=recipes,
                replay_root=replays,
                telemetry_root=telemetry,
                registry_path=registry_file,
            )
            _write_state(
                state_file,
                state,
                priority_report=priority_report,
                telemetry_root=telemetry,
                registry=registry,
            )
            continue

        telemetry_values = _telemetry_by_structure(telemetry)
        attempted = sum(
            int(value.get("aiCallsAttempted") or 0)
            for value in telemetry_values.values()
        )
        if attempted >= int(max_total_ai_calls):
            state["status"] = "AI_BUDGET_EXHAUSTED"
            _event(
                state,
                structure_id=structure_id,
                status="AI_BUDGET_EXHAUSTED",
                detail=f"{attempted}/{max_total_ai_calls}",
            )
            continue
        telemetry_file = telemetry / f"{structure_id}.ai.json"
        decision_file = decisions / f"{structure_id}.decision.json"
        if telemetry_file.is_file():
            state["outcomes"][structure_id] = {
                "status": "QUARANTINED_EXISTING_TELEMETRY_NO_RETRY",
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
            }
            continue
        try:
            decision = run_codex_table_recipe_decision(
                priority_report=priority_report,
                table_structure_id=structure_id,
                output_path=decision_file,
                telemetry_path=telemetry_file,
                reasoning_effort=reasoning_effort,
                timeout_seconds=timeout_seconds,
            )
        except Exception as exc:
            state["outcomes"][structure_id] = {
                "status": "QUARANTINED_AI_FAILURE_NO_RETRY",
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "errorType": type(exc).__name__,
                "error": str(exc),
                "propagationErrors": propagation_errors[-3:],
            }
            _event(
                state,
                structure_id=structure_id,
                status="QUARANTINED_AI_FAILURE_NO_RETRY",
                detail=f"{type(exc).__name__}: {exc}",
            )
            _write_state(
                state_file,
                state,
                priority_report=priority_report,
                telemetry_root=telemetry,
                registry=registry,
            )
            continue

        if decision["decision"] == "QUARANTINE":
            _write_json(
                quarantine / f"{structure_id}.decision.json",
                decision,
            )
            state["outcomes"][structure_id] = {
                "status": "QUARANTINED_AI_DECISION",
                "decisionAiCalls": 1,
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "rationale": str(decision.get("rationale") or ""),
            }
            replay = None
            recipe = None
        else:
            try:
                recipe = compile_structure_recipe(
                    decision,
                    priority_item=item,
                    representative_captured_cells=(
                        load_representative_captured_cells(
                            priority_report,
                            item,
                        )
                    ),
                    decision_ai_calls=1,
                    decision_source="BOUNDED_AI_STRUCTURE_DECISION",
                )
                replay = replay_structure_recipe(
                    recipe=recipe,
                    priority_report=priority_report,
                )
            except Exception as exc:
                recipe = None
                replay = None
                state["outcomes"][structure_id] = {
                    "status": "QUARANTINED_RECIPE_CONTRACT_FAILURE",
                    "decisionAiCalls": 1,
                    "tableCount": int(item.get("tableCount") or 0),
                    "workbookCount": int(item.get("workbookCount") or 0),
                    "errorType": type(exc).__name__,
                    "error": str(exc),
                }
        if (
            decision["decision"] != "QUARANTINE"
            and replay is not None
            and replay.get("status") == VERIFIED_REPLAY_STATUS
        ):
            recipe_file, replay_file = _recipe_paths(
                structure_id,
                recipe_root=recipes,
                replay_root=replays,
            )
            _write_json(recipe_file, recipe)
            _write_json(replay_file, replay)
            state["outcomes"][structure_id] = {
                "status": "REGISTERED_AI_REPLAY",
                "decisionSource": "BOUNDED_AI_STRUCTURE_DECISION",
                "decisionAiCalls": 1,
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "recipeId": recipe["recipeId"],
            }
        elif (
            decision["decision"] != "QUARANTINE"
            and recipe is not None
            and replay is not None
        ):
            quarantine_recipe = quarantine / (
                recipe["recipeId"] + ".failed.json"
            )
            quarantine_replay = quarantine / (
                recipe["recipeId"] + ".failed.replay.json"
            )
            _write_json(quarantine_recipe, recipe)
            _write_json(quarantine_replay, replay)
            state["outcomes"][structure_id] = {
                "status": "QUARANTINED_REPLAY_FAILED",
                "decisionAiCalls": 1,
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "replaySummary": replay.get("summary") or {},
            }
        _event(
            state,
            structure_id=structure_id,
            status=state["outcomes"][structure_id]["status"],
            detail=(
                f"tables={int(item.get('tableCount') or 0)}; "
                "AI calls=1; retry=0"
            ),
        )
        registry = _refresh_registry(
            priority_report=priority_report,
            recipe_root=recipes,
            replay_root=replays,
            telemetry_root=telemetry,
            registry_path=registry_file,
        )
        _write_state(
            state_file,
            state,
            priority_report=priority_report,
            telemetry_root=telemetry,
            registry=registry,
        )

    registry = _refresh_registry(
        priority_report=priority_report,
        recipe_root=recipes,
        replay_root=replays,
        telemetry_root=telemetry,
        registry_path=registry_file,
    )
    _record_registered_outcomes(state, registry)
    _summarize(
        state,
        priority_report=priority_report,
        telemetry=_telemetry_by_structure(telemetry),
        registry=registry,
    )
    if state["summary"]["unresolvedStructureCount"] == 0:
        state["status"] = "COMPLETED"
        state["completedAt"] = _now()
    elif state.get("status") != "AI_BUDGET_EXHAUSTED":
        state["status"] = "INCOMPLETE"
    _write_json(state_file, state)
    return state


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Finish a structure-recipe queue with bounded AI."
    )
    parser.add_argument("--priority-report", required=True)
    parser.add_argument("--recipe-root", required=True)
    parser.add_argument("--replay-root", required=True)
    parser.add_argument("--decision-root", required=True)
    parser.add_argument("--telemetry-root", required=True)
    parser.add_argument("--quarantine-root", required=True)
    parser.add_argument("--registry-output", required=True)
    parser.add_argument("--state-output", required=True)
    parser.add_argument("--max-total-ai-calls", type=int, required=True)
    parser.add_argument("--reasoning-effort", default="low")
    parser.add_argument("--timeout-seconds", type=int, default=600)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    state = complete_structure_queue(
        priority_report_path=arguments.priority_report,
        recipe_root=arguments.recipe_root,
        replay_root=arguments.replay_root,
        decision_root=arguments.decision_root,
        telemetry_root=arguments.telemetry_root,
        quarantine_root=arguments.quarantine_root,
        registry_path=arguments.registry_output,
        state_path=arguments.state_output,
        max_total_ai_calls=arguments.max_total_ai_calls,
        reasoning_effort=arguments.reasoning_effort,
        timeout_seconds=arguments.timeout_seconds,
    )
    print(json.dumps(state["summary"], ensure_ascii=False))
    return 0 if state["status"] == "COMPLETED" else 2


if __name__ == "__main__":
    raise SystemExit(main())
