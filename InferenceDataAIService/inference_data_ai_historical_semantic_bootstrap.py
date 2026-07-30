"""Bootstrap new structure recipes from consistent historical semantics."""

from __future__ import annotations

import argparse
import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_table_recipe_proposal import (
    PROPOSAL_DECISION_SCHEMA_VERSION,
    _column_number,
    _read_json,
    _write_json,
    compile_structure_recipe,
    replay_structure_recipe,
    semantic_header_signature,
    _semantic_header_sha256,
)


HISTORICAL_SEMANTIC_CATALOG_SCHEMA_VERSION = (
    "excel-historical-semantic-contract-catalog-v1"
)
HISTORICAL_BOOTSTRAP_REPORT_SCHEMA_VERSION = (
    "excel-historical-semantic-bootstrap-report-v1"
)
HISTORICAL_BOOTSTRAP_ENGINE_VERSION = "historical-989-consensus-v1.0"


class HistoricalSemanticBootstrapError(RuntimeError):
    """Raised when historical semantics cannot safely produce a recipe."""


def _normalized_metric_name(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _axis_relative_columns(
    table: dict[str, Any],
) -> dict[str, int]:
    minimum = int((table.get("bounds") or {}).get("minColumn") or 0)
    return {
        str(column.get("columnId") or ""): (
            _column_number(str(column.get("column") or "")) - minimum
        )
        for column in table.get("numericColumns") or []
        if column.get("columnId")
        and not str(column.get("columnRole") or "").startswith("AGGREGATE_")
    }


def _analysis_by_table_id(
    analysis: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    return {
        str(table.get("tableId") or ""): table
        for table in analysis.get("tables") or []
        if table.get("tableId")
    }


def _target_columns(
    signature: list[dict[str, Any]],
) -> dict[int, dict[str, Any]]:
    return {
        int(column["relativeColumn"]): column
        for column in signature
        if str(column.get("columnRole") or "") == "MEASURE_VALUE"
    }


def build_historical_semantic_contract_catalog(
    *,
    priority_report: dict[str, Any],
    historical_batch_root: str | Path,
    minimum_support: int = 2,
    minimum_type_consistency: float = 0.9,
    generated_at: str | None = None,
) -> dict[str, Any]:
    old_root = Path(historical_batch_root).expanduser().resolve()
    target_items: dict[str, dict[str, Any]] = {}
    for item in priority_report.get("queue") or []:
        target_items.setdefault(
            str(item.get("semanticHeaderSha256") or ""),
            item,
        )
    observations: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for request_path in sorted((old_root / "requests").glob("*.json")):
        analysis_path = old_root / "analyses" / request_path.name
        if not analysis_path.is_file():
            continue
        request = _read_json(request_path)
        analysis = _read_json(analysis_path)
        analysis_tables = _analysis_by_table_id(analysis)
        for table in request.get("tables") or []:
            signature = semantic_header_signature(table)
            signature_sha256 = _semantic_header_sha256(signature)
            target_item = target_items.get(signature_sha256)
            if target_item is None:
                continue
            semantic = analysis_tables.get(
                str(table.get("tableId") or "")
            )
            if semantic is None:
                continue
            axes = _axis_relative_columns(table)
            metrics_by_relative: dict[int, list[dict[str, str]]] = defaultdict(
                list
            )
            for metric in semantic.get("metrics") or []:
                axis_refs = list(metric.get("axisRefs") or [])
                if len(axis_refs) != 1 or axis_refs[0] not in axes:
                    continue
                metrics_by_relative[axes[axis_refs[0]]].append(
                    {
                        "canonicalName": _normalized_metric_name(
                            metric.get("name")
                        ),
                        "unit": str(metric.get("unit") or ""),
                    }
                )
            targets = set(_target_columns(signature))
            full_mapping = (
                set(metrics_by_relative) == targets
                and all(
                    len(metrics_by_relative[relative]) == 1
                    for relative in targets
                )
            )
            observations[signature_sha256].append(
                {
                    "requestFile": request_path.name,
                    "tableId": str(table.get("tableId") or ""),
                    "tableType": str(semantic.get("type") or "UNASSESSED"),
                    "confidence": str(
                        semantic.get("confidence") or "UNASSESSED"
                    ),
                    "fullMapping": full_mapping,
                    "metrics": (
                        {
                            str(relative): metrics_by_relative[relative][0]
                            for relative in sorted(targets)
                        }
                        if full_mapping
                        else {}
                    ),
                }
            )

    contracts: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []
    for signature_sha256, values in sorted(observations.items()):
        target_item = target_items[signature_sha256]
        type_counts = Counter(value["tableType"] for value in values)
        dominant_type, dominant_count = type_counts.most_common(1)[0]
        type_consistency = dominant_count / len(values)
        full = [
            value
            for value in values
            if value["fullMapping"]
            and value["tableType"] == dominant_type
        ]
        reasons: list[str] = []
        if dominant_type != "DESCRIPTIVE":
            reasons.append("DOMINANT_TYPE_NOT_DESCRIPTIVE")
        if type_consistency < float(minimum_type_consistency):
            reasons.append("TYPE_CONSISTENCY_BELOW_THRESHOLD")
        if len(full) < int(minimum_support):
            reasons.append("FULL_MAPPING_SUPPORT_BELOW_MINIMUM")
        if reasons:
            rejected.append(
                {
                    "semanticHeaderSha256": signature_sha256,
                    "historicalMemberCount": len(values),
                    "dominantTableType": dominant_type,
                    "typeConsistency": round(type_consistency, 6),
                    "fullMappingSupport": len(full),
                    "reasons": reasons,
                }
            )
            continue

        target_columns = _target_columns(
            target_item["semanticHeaderSignature"]
        )
        metric_columns: list[dict[str, Any]] = []
        for relative, target_column in sorted(target_columns.items()):
            names = Counter(
                str(value["metrics"][str(relative)]["canonicalName"])
                for value in full
            )
            units = Counter(
                str(value["metrics"][str(relative)]["unit"])
                for value in full
            )
            fallback_headers = list(target_column.get("headerTexts") or [])
            metric_columns.append(
                {
                    "relativeColumn": relative,
                    "canonicalName": (
                        names.most_common(1)[0][0]
                        or (
                            fallback_headers[-1]
                            if fallback_headers
                            else f"metric-{relative}"
                        )
                    ),
                    "unit": units.most_common(1)[0][0],
                    "nameSupport": names.most_common(1)[0][1],
                    "unitSupport": units.most_common(1)[0][1],
                }
            )
        contracts.append(
            {
                "semanticHeaderSha256": signature_sha256,
                "semanticHeaderSignature": target_item[
                    "semanticHeaderSignature"
                ],
                "status": "HISTORICAL_CONSENSUS_READY",
                "tableType": dominant_type,
                "historicalMemberCount": len(values),
                "fullMappingSupport": len(full),
                "typeConsistency": round(type_consistency, 6),
                "metricColumns": metric_columns,
                "sourceExamples": [
                    {
                        "requestFile": value["requestFile"],
                        "tableId": value["tableId"],
                    }
                    for value in full[:5]
                ],
            }
        )
    return {
        "schemaVersion": HISTORICAL_SEMANTIC_CATALOG_SCHEMA_VERSION,
        "engineVersion": HISTORICAL_BOOTSTRAP_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            "historicalBatchRoot": str(old_root),
            "targetSemanticSignatureCount": len(target_items),
            "minimumSupport": int(minimum_support),
            "minimumTypeConsistency": float(minimum_type_consistency),
            "aiCalls": 0,
        },
        "summary": {
            "observedTargetSignatureCount": len(observations),
            "readyContractCount": len(contracts),
            "rejectedContractCount": len(rejected),
            "historicalMemberCount": sum(
                len(values) for values in observations.values()
            ),
            "aiCalls": 0,
        },
        "contracts": contracts,
        "rejected": rejected,
    }


def _decision_from_contract(
    *,
    item: dict[str, Any],
    contract: dict[str, Any],
) -> dict[str, Any]:
    representative = item.get("representativeTable") or {}
    titles = list(representative.get("titleCandidates") or [])
    metrics = [
        {
            "relativeColumn": int(metric["relativeColumn"]),
            "canonicalName": str(metric["canonicalName"]),
            "unit": str(metric["unit"]),
        }
        for metric in contract.get("metricColumns") or []
    ]
    return {
        "schemaVersion": PROPOSAL_DECISION_SCHEMA_VERSION,
        "targetTableStructureId": item["tableStructureId"],
        "targetFingerprintSha256": item["fingerprintSha256"],
        "decision": "NEW_RECIPE",
        "historicalSourceTableStructureId": "",
        "confidence": (
            "HIGH"
            if float(contract.get("typeConsistency") or 0) == 1.0
            and int(contract.get("fullMappingSupport") or 0) >= 3
            else "MEDIUM"
        ),
        "rationale": (
            "AI-free historical consensus from "
            f"{int(contract.get('fullMappingSupport') or 0)} prior tables "
            "with the same semantic metric-header signature."
        ),
        "semanticContract": {
            "title": (
                str(titles[0])
                if titles
                else "Historical descriptive result"
            ),
            "tableType": "DESCRIPTIVE",
            "studyGroup": (
                "historical-signature-"
                + str(contract["semanticHeaderSha256"])[:12]
            ),
            "groups": [],
            "metricColumns": metrics,
            "comparisonRelations": [],
            "limitations": [
                "Semantic labels came from historical consensus; all values, "
                "source text, statistics, and evidence remain code-owned."
            ],
        },
    }


def bootstrap_historical_structure_recipes(
    *,
    priority_report: dict[str, Any],
    semantic_catalog: dict[str, Any],
    recipe_root: str | Path,
    replay_root: str | Path,
    decision_root: str | Path,
    telemetry_root: str | Path,
    generated_at: str | None = None,
) -> dict[str, Any]:
    recipes = Path(recipe_root).expanduser().resolve()
    replays = Path(replay_root).expanduser().resolve()
    decisions = Path(decision_root).expanduser().resolve()
    telemetry = Path(telemetry_root).expanduser().resolve()
    contracts = {
        str(value["semanticHeaderSha256"]): value
        for value in semantic_catalog.get("contracts") or []
        if value.get("status") == "HISTORICAL_CONSENSUS_READY"
    }
    telemetry_structures = {
        str(value.get("tableStructureId") or "")
        for path in sorted(telemetry.glob("*.json"))
        for value in [_read_json(path)]
        if value.get("tableStructureId")
    }
    items: list[dict[str, Any]] = []
    for item in priority_report.get("queue") or []:
        structure_id = str(item.get("tableStructureId") or "")
        contract = contracts.get(str(item.get("semanticHeaderSha256") or ""))
        if contract is None:
            continue
        if structure_id in telemetry_structures:
            items.append(
                {
                    "tableStructureId": structure_id,
                    "status": "SKIPPED_EXISTING_AI_TELEMETRY",
                }
            )
            continue
        decision = _decision_from_contract(item=item, contract=contract)
        recipe_file = recipes / (
            "structure-recipe-"
            + structure_id.removeprefix("table-structure-")
            + ".json"
        )
        replay_file = replays / (
            "structure-recipe-"
            + structure_id.removeprefix("table-structure-")
            + ".replay.json"
        )
        decision_file = decisions / (
            structure_id + ".historical.json"
        )
        try:
            recipe = compile_structure_recipe(
                decision,
                priority_item=item,
                generated_at=generated_at,
                decision_ai_calls=0,
                decision_source="HISTORICAL_989_CONSENSUS",
            )
            replay = replay_structure_recipe(
                recipe=recipe,
                priority_report=priority_report,
                generated_at=generated_at,
            )
            _write_json(decision_file, decision)
            _write_json(recipe_file, recipe)
            _write_json(replay_file, replay)
            items.append(
                {
                    "tableStructureId": structure_id,
                    "semanticHeaderSha256": item[
                        "semanticHeaderSha256"
                    ],
                    "status": (
                        "REGISTERABLE"
                        if replay.get("status")
                        == (
                            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_"
                            "NEEDS_CANONICAL_REVIEW"
                        )
                        else "REPLAY_FAILED"
                    ),
                    "tableCount": int(item.get("tableCount") or 0),
                    "workbookCount": int(item.get("workbookCount") or 0),
                    "recipeFile": str(recipe_file),
                    "replayFile": str(replay_file),
                    "replaySummary": replay.get("summary") or {},
                }
            )
        except Exception as exc:
            items.append(
                {
                    "tableStructureId": structure_id,
                    "semanticHeaderSha256": item.get(
                        "semanticHeaderSha256"
                    ),
                    "status": "BOOTSTRAP_FAILED",
                    "errorType": type(exc).__name__,
                    "error": str(exc),
                }
            )
    counts = Counter(str(item["status"]) for item in items)
    return {
        "schemaVersion": HISTORICAL_BOOTSTRAP_REPORT_SCHEMA_VERSION,
        "engineVersion": HISTORICAL_BOOTSTRAP_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "summary": {
            "candidateStructureCount": len(items),
            "statusCounts": dict(sorted(counts.items())),
            "registerableTableCount": sum(
                int(item.get("tableCount") or 0)
                for item in items
                if item["status"] == "REGISTERABLE"
            ),
            "registerableWorkbookReferences": sum(
                int(item.get("workbookCount") or 0)
                for item in items
                if item["status"] == "REGISTERABLE"
            ),
            "aiCalls": 0,
        },
        "items": items,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Bootstrap new recipes from historical 989 semantics."
    )
    parser.add_argument("--priority-report", required=True)
    parser.add_argument("--historical-batch-root", required=True)
    parser.add_argument("--semantic-catalog-output", required=True)
    parser.add_argument("--recipe-root", required=True)
    parser.add_argument("--replay-root", required=True)
    parser.add_argument("--decision-root", required=True)
    parser.add_argument("--telemetry-root", required=True)
    parser.add_argument("--report-output", required=True)
    parser.add_argument("--minimum-support", type=int, default=2)
    parser.add_argument("--minimum-type-consistency", type=float, default=0.9)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    priority_report = _read_json(arguments.priority_report)
    catalog = build_historical_semantic_contract_catalog(
        priority_report=priority_report,
        historical_batch_root=arguments.historical_batch_root,
        minimum_support=arguments.minimum_support,
        minimum_type_consistency=arguments.minimum_type_consistency,
    )
    report = bootstrap_historical_structure_recipes(
        priority_report=priority_report,
        semantic_catalog=catalog,
        recipe_root=arguments.recipe_root,
        replay_root=arguments.replay_root,
        decision_root=arguments.decision_root,
        telemetry_root=arguments.telemetry_root,
    )
    _write_json(arguments.semantic_catalog_output, catalog)
    _write_json(arguments.report_output, report)
    print(
        json.dumps(
            {
                "catalog": catalog["summary"],
                "bootstrap": report["summary"],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
