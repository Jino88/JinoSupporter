"""Bootstrap fail-closed descriptive recipes for one-measure table structures."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_table_recipe_proposal import (
    PROPOSAL_DECISION_SCHEMA_VERSION,
    _read_json,
    _write_json,
    compile_structure_recipe,
    load_representative_captured_cells,
    replay_structure_recipe,
    validate_table_recipe_decision,
)


SINGLE_MEASURE_BOOTSTRAP_SCHEMA_VERSION = (
    "excel-single-measure-bootstrap-report-v1"
)
SINGLE_MEASURE_BOOTSTRAP_ENGINE_VERSION = (
    "source-owned-single-measure-bootstrap-v1.0"
)
SOURCE_OWNED_DECISION_SOURCE = "SOURCE_OWNED_SINGLE_MEASURE_HEADER"

_GENERIC_HEADER_KEYS = {
    "",
    "data",
    "me",
    "measure",
    "measurement",
    "result",
    "value",
}


class SingleMeasureBootstrapError(RuntimeError):
    """Raised when the explicit single-measure contract is violated."""


def _header_key(value: str) -> str:
    return re.sub(r"[\W_]+", "", value.casefold(), flags=re.UNICODE)


def _source_unit(header: str) -> str:
    normalized = header.casefold()
    if "%" in header or "percent" in normalized or "rate" in normalized:
        return "percent"
    if re.search(r"\bpcs?\b", normalized) or "q'ty" in normalized:
        return "pcs"
    return "source-formatted number"


def build_source_owned_single_measure_decision(
    priority_item: dict[str, Any],
) -> dict[str, Any]:
    """Create a descriptive decision or an explicit generic-header quarantine."""

    measures = [
        column
        for column in (
            priority_item.get("representativeTable") or {}
        ).get("numericColumns")
        or []
        if str(column.get("columnRole") or "") == "MEASURE_VALUE"
    ]
    if len(measures) != 1:
        raise SingleMeasureBootstrapError(
            "Expected exactly one source measure column."
        )
    headers = [
        " ".join(str(value).split())
        for value in measures[0].get("headerTexts") or []
        if " ".join(str(value).split())
    ]
    canonical_name = " | ".join(dict.fromkeys(headers))
    generic = not canonical_name or _header_key(canonical_name) in (
        _GENERIC_HEADER_KEYS
    )
    representative = priority_item.get("representativeTable") or {}
    titles = [
        str(value)
        for value in representative.get("titleCandidates") or []
        if str(value).strip()
    ]
    decision = {
        "schemaVersion": PROPOSAL_DECISION_SCHEMA_VERSION,
        "targetTableStructureId": priority_item["tableStructureId"],
        "targetFingerprintSha256": priority_item["fingerprintSha256"],
        "decision": "QUARANTINE" if generic else "NEW_RECIPE",
        "historicalSourceTableStructureId": "",
        "confidence": "MEDIUM" if not generic else "LOW",
        "rationale": (
            "The only source measure header is generic and cannot name a "
            "queryable metric without invention."
            if generic
            else (
                "AI-free descriptive mapping from the sole source-authored "
                "measure header. Values, statistics, and evidence remain "
                "program-owned."
            )
        ),
        "semanticContract": {
            "title": (
                titles[0]
                if titles
                else (
                    "Single source measure"
                    if generic
                    else f"{canonical_name} descriptive result"
                )
            ),
            "tableType": "DESCRIPTIVE",
            "studyGroup": (
                "source-owned-single-measure-"
                + str(priority_item.get("semanticHeaderSha256") or "")[:12]
            ),
            "groups": [],
            "metricColumns": (
                []
                if generic
                else [
                    {
                        "relativeColumn": int(
                            measures[0]["relativeColumn"]
                        ),
                        "canonicalName": canonical_name,
                        "unit": _source_unit(canonical_name),
                    }
                ]
            ),
            "comparisonRelations": [],
            "limitations": [
                (
                    "The source header is insufficient for a semantic metric; "
                    "the structure remains quarantined."
                    if generic
                    else (
                        "The source supplies one descriptive measure only; "
                        "comparison and effect interpretation are not eligible."
                    )
                )
            ],
        },
    }
    return validate_table_recipe_decision(
        decision,
        priority_item=priority_item,
    )


def bootstrap_source_owned_single_measure_recipes(
    *,
    priority_report: dict[str, Any],
    recipe_root: str | Path,
    replay_root: str | Path,
    decision_root: str | Path,
    generated_at: str | None = None,
) -> dict[str, Any]:
    recipes = Path(recipe_root).expanduser().resolve()
    replays = Path(replay_root).expanduser().resolve()
    decisions = Path(decision_root).expanduser().resolve()
    for root in (recipes, replays, decisions):
        root.mkdir(parents=True, exist_ok=True)
    items: list[dict[str, Any]] = []
    for item in priority_report.get("queue") or []:
        structure_id = str(item.get("tableStructureId") or "")
        decision = build_source_owned_single_measure_decision(item)
        decision_file = decisions / f"{structure_id}.source-owned.json"
        _write_json(decision_file, decision)
        if decision["decision"] == "QUARANTINE":
            items.append(
                {
                    "tableStructureId": structure_id,
                    "status": "QUARANTINED_GENERIC_MEASURE_HEADER",
                    "tableCount": int(item.get("tableCount") or 0),
                    "workbookCount": int(item.get("workbookCount") or 0),
                    "decisionFile": str(decision_file),
                    "aiCalls": 0,
                }
            )
            continue
        suffix = structure_id.removeprefix("table-structure-")
        recipe_file = recipes / f"structure-recipe-{suffix}.json"
        replay_file = replays / f"structure-recipe-{suffix}.replay.json"
        recipe = compile_structure_recipe(
            decision,
            priority_item=item,
            representative_captured_cells=load_representative_captured_cells(
                priority_report,
                item,
            ),
            generated_at=generated_at,
            decision_ai_calls=0,
            decision_source=SOURCE_OWNED_DECISION_SOURCE,
        )
        replay = replay_structure_recipe(
            recipe=recipe,
            priority_report=priority_report,
            generated_at=generated_at,
        )
        _write_json(recipe_file, recipe)
        _write_json(replay_file, replay)
        items.append(
            {
                "tableStructureId": structure_id,
                "status": (
                    "REGISTERABLE_SOURCE_OWNED"
                    if replay.get("status")
                    == (
                        "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_"
                        "NEEDS_CANONICAL_REVIEW"
                    )
                    else "REPLAY_FAILED"
                ),
                "tableCount": int(item.get("tableCount") or 0),
                "workbookCount": int(item.get("workbookCount") or 0),
                "decisionFile": str(decision_file),
                "recipeFile": str(recipe_file),
                "replayFile": str(replay_file),
                "replaySummary": replay.get("summary") or {},
                "aiCalls": 0,
            }
        )
    counts = Counter(str(item["status"]) for item in items)
    return {
        "schemaVersion": SINGLE_MEASURE_BOOTSTRAP_SCHEMA_VERSION,
        "engineVersion": SINGLE_MEASURE_BOOTSTRAP_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "summary": {
            "structureCount": len(items),
            "statusCounts": dict(sorted(counts.items())),
            "registerableStructureCount": sum(
                item["status"] == "REGISTERABLE_SOURCE_OWNED"
                for item in items
            ),
            "registerableTableCount": sum(
                int(item.get("tableCount") or 0)
                for item in items
                if item["status"] == "REGISTERABLE_SOURCE_OWNED"
            ),
            "quarantinedStructureCount": sum(
                item["status"] == "QUARANTINED_GENERIC_MEASURE_HEADER"
                for item in items
            ),
            "aiCalls": 0,
        },
        "items": items,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Build source-owned one-measure recipes and quarantine generic "
            "headers without AI."
        )
    )
    parser.add_argument("--priority-report", required=True)
    parser.add_argument("--recipe-root", required=True)
    parser.add_argument("--replay-root", required=True)
    parser.add_argument("--decision-root", required=True)
    parser.add_argument("--output", required=True)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    report = bootstrap_source_owned_single_measure_recipes(
        priority_report=_read_json(arguments.priority_report),
        recipe_root=arguments.recipe_root,
        replay_root=arguments.replay_root,
        decision_root=arguments.decision_root,
    )
    _write_json(arguments.output, report)
    print(json.dumps(report["summary"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
