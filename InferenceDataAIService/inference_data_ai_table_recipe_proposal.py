"""Prioritize repeated table structures and replay one bounded semantic recipe.

The expensive operation in this module is deliberately limited to one optional
AI decision per table structure. Numeric values, statistics, and evidence never
enter that decision: deterministic code reads those fields from the table-first
request when the approved structure recipe is replayed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence

from inference_data_ai_table_first import (
    CONFIDENCE_LEVELS,
    GROUP_ROLES,
    TABLE_TYPES,
)
from inference_data_ai_table_structure_catalog import (
    table_structure_fingerprint,
)


PRIORITY_REPORT_SCHEMA_VERSION = "excel-table-recipe-priority-report-v1"
PROPOSAL_DECISION_SCHEMA_VERSION = "excel-table-recipe-proposal-decision-v1"
STRUCTURE_RECIPE_SCHEMA_VERSION = "excel-table-structure-recipe-v1"
STRUCTURE_REPLAY_SCHEMA_VERSION = "excel-table-structure-replay-v1"
PROPOSAL_ENGINE_VERSION = "bounded-structure-recipe-proposal-v1.0"

DECISIONS = (
    "REUSE_HISTORICAL_RECIPE",
    "NEW_RECIPE",
    "QUARANTINE",
)

_COORDINATE_PATTERN = re.compile(r"^([A-Z]{1,3})([0-9]+)$", re.IGNORECASE)


class TableRecipeProposalError(RuntimeError):
    """Raised when a structure proposal or replay is unsafe."""


def _read_json(path: str | Path) -> dict[str, Any]:
    source = Path(path).expanduser().resolve()
    with source.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise TableRecipeProposalError(f"JSON root must be an object: {source}")
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


def _json_value(value: Any) -> Any:
    if value is None:
        return None
    return json.loads(str(value))


def _canonical(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def _jaccard(left: set[str], right: set[str]) -> float:
    if not left and not right:
        return 1.0
    union = left | right
    return len(left & right) / len(union) if union else 0.0


def _set_of(values: Iterable[Any]) -> set[str]:
    return {_canonical(value) for value in values}


def _preview_features(fingerprint: dict[str, Any]) -> set[str]:
    return _set_of(
        {
            "relativeRow": int(row.get("relativeRow") or 0),
            "relativeColumn": int(cell.get("relativeColumn") or 0),
            "kind": str(cell.get("kind") or ""),
            "merged": bool(cell.get("merged")),
        }
        for row in fingerprint.get("previewPatterns") or []
        for cell in row.get("cells") or []
    )


def _numeric_features(fingerprint: dict[str, Any]) -> set[str]:
    return _set_of(
        {
            "relativeColumn": int(column.get("relativeColumn") or 0),
            "columnRole": str(column.get("columnRole") or ""),
            "numberFormatRoles": list(column.get("numberFormatRoles") or []),
        }
        for column in fingerprint.get("numericColumns") or []
    )


def _series_features(fingerprint: dict[str, Any]) -> set[str]:
    return _set_of(fingerprint.get("numericSeries") or [])


def _ratio(left: int, right: int) -> float:
    maximum = max(abs(int(left)), abs(int(right)), 1)
    return min(abs(int(left)), abs(int(right))) / maximum


def _prepared_similarity(fingerprint: dict[str, Any]) -> dict[str, Any]:
    return {
        "fingerprintSha256": str(
            fingerprint.get("fingerprintSha256") or ""
        ),
        "numericColumnCount": int(
            fingerprint.get("numericColumnCount") or 0
        ),
        "geometry": tuple(
            str(fingerprint.get(field) or "")
            for field in (
                "rowCountBucket",
                "columnCountBucket",
                "sourceCellCountBucket",
                "numericCellCountBucket",
            )
        ),
        "numericFeatures": _numeric_features(fingerprint),
        "previewFeatures": _preview_features(fingerprint),
        "mergeFeatures": _set_of(fingerprint.get("mergeShapes") or []),
        "headerFeatures": {
            str(value) for value in fingerprint.get("headerTokens") or []
        },
        "seriesFeatures": _series_features(fingerprint),
    }


def _score_prepared_similarity(
    target: dict[str, Any],
    candidate: dict[str, Any],
) -> dict[str, Any]:
    target_numeric = int(target["numericColumnCount"])
    candidate_numeric = int(candidate["numericColumnCount"])
    hard_gate_failures: list[str] = []
    if target_numeric < 2:
        hard_gate_failures.append("TARGET_HAS_FEWER_THAN_TWO_NUMERIC_COLUMNS")
    if candidate_numeric < 2:
        hard_gate_failures.append("CANDIDATE_HAS_FEWER_THAN_TWO_NUMERIC_COLUMNS")

    geometry = sum(
        left == right
        for left, right in zip(
            target["geometry"],
            candidate["geometry"],
            strict=True,
        )
    ) / len(target["geometry"])
    numeric_layout = (
        0.75
        * _jaccard(
            target["numericFeatures"],
            candidate["numericFeatures"],
        )
        + 0.25 * _ratio(target_numeric, candidate_numeric)
    )
    preview = _jaccard(
        target["previewFeatures"],
        candidate["previewFeatures"],
    )
    merges = _jaccard(
        target["mergeFeatures"],
        candidate["mergeFeatures"],
    )
    headers = _jaccard(
        target["headerFeatures"],
        candidate["headerFeatures"],
    )
    series = _jaccard(
        target["seriesFeatures"],
        candidate["seriesFeatures"],
    )
    components = {
        "geometry": round(geometry, 6),
        "numericLayout": round(numeric_layout, 6),
        "previewLayout": round(preview, 6),
        "mergeGeometry": round(merges, 6),
        "headerTokens": round(headers, 6),
        "numericSeries": round(series, 6),
    }
    score = (
        0.20 * geometry
        + 0.35 * numeric_layout
        + 0.15 * preview
        + 0.10 * merges
        + 0.15 * headers
        + 0.05 * series
    )
    return {
        "score": round(score, 6) if not hard_gate_failures else 0.0,
        "hardGatePassed": not hard_gate_failures,
        "hardGateFailures": hard_gate_failures,
        "components": components,
        "fingerprintExact": (
            target["fingerprintSha256"] == candidate["fingerprintSha256"]
        ),
    }


def table_structure_similarity(
    target: dict[str, Any],
    candidate: dict[str, Any],
) -> dict[str, Any]:
    """Score two value-free fingerprints for bounded historical routing."""

    return _score_prepared_similarity(
        _prepared_similarity(target),
        _prepared_similarity(candidate),
    )


def _column_number(value: str) -> int:
    result = 0
    if not value or not value.isalpha():
        raise TableRecipeProposalError(f"Invalid column label: {value}")
    for character in value.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _relative_coordinate(
    coordinate: Any,
    *,
    min_row: int,
    min_column: int,
) -> dict[str, int]:
    match = _COORDINATE_PATTERN.fullmatch(str(coordinate or ""))
    if match is None:
        raise TableRecipeProposalError(
            f"Invalid table coordinate: {coordinate}"
        )
    return {
        "relativeRow": int(match.group(2)) - min_row,
        "relativeColumn": _column_number(match.group(1)) - min_column,
    }


def redact_representative_table(table: dict[str, Any]) -> dict[str, Any]:
    """Keep semantic labels and geometry while removing all source numbers."""

    bounds = table.get("bounds") or {}
    min_row = int(bounds.get("minRow") or 0)
    min_column = int(bounds.get("minColumn") or 0)
    if min_row < 1 or min_column < 1:
        raise TableRecipeProposalError("Representative bounds are invalid.")
    numeric_columns: list[dict[str, Any]] = []
    for column in table.get("numericColumns") or []:
        numeric_columns.append(
            {
                "relativeColumn": (
                    _column_number(str(column.get("column") or ""))
                    - min_column
                ),
                "columnRole": str(column.get("columnRole") or ""),
                "headerTexts": [
                    str(value) for value in column.get("headerTexts") or []
                ],
                "numberFormats": [
                    str(value) for value in column.get("numberFormats") or []
                ],
            }
        )
    preview_rows: list[dict[str, Any]] = []
    for row in table.get("previewRows") or []:
        cells: list[dict[str, Any]] = []
        for cell in row.get("cells") or []:
            kind = str(cell.get("kind") or "")
            redacted = {
                **_relative_coordinate(
                    cell.get("coordinate"),
                    min_row=min_row,
                    min_column=min_column,
                ),
                "kind": kind,
                "merged": bool(cell.get("mergeRange")),
            }
            if kind == "TEXT" and cell.get("value") not in (None, ""):
                redacted["text"] = str(cell["value"])
            cells.append(redacted)
        preview_rows.append(
            {
                "relativeRow": int(row.get("row") or 0) - min_row,
                "omittedCellCount": int(row.get("omittedCellCount") or 0),
                "cells": cells,
            }
        )
    return {
        "tableId": str(table.get("tableId") or ""),
        "sheet": str(table.get("sheet") or ""),
        "range": str(table.get("range") or ""),
        "rowCount": int(bounds.get("maxRow") or 0) - min_row + 1,
        "columnCount": int(bounds.get("maxColumn") or 0) - min_column + 1,
        "titleCandidates": [
            str(value) for value in table.get("titleCandidates") or []
        ],
        "rowLabels": [
            {
                "relativeRow": int(row.get("row") or 0) - min_row,
                "labels": [str(value) for value in row.get("labels") or []],
            }
            for row in table.get("rowLabels") or []
        ],
        "numericColumns": sorted(
            numeric_columns,
            key=lambda item: int(item["relativeColumn"]),
        ),
        "numericSeries": [
            {
                "columnRole": str(series.get("columnRole") or ""),
                "memberColumnCount": len(
                    series.get("columnIds")
                    or series.get("memberColumnIds")
                    or series.get("columns")
                    or []
                ),
                "numberFormats": [
                    str(value)
                    for value in series.get("numberFormats") or []
                ],
            }
            for series in table.get("numericSeries") or []
        ],
        "previewRows": preview_rows,
        "valuePolicy": {
            "rawValuesIncluded": False,
            "statisticsIncluded": False,
            "aiMayWriteValues": False,
        },
    }


def _table_by_id(request: dict[str, Any], table_id: str) -> dict[str, Any]:
    matches = [
        table
        for table in request.get("tables") or []
        if str(table.get("tableId") or "") == table_id
    ]
    if len(matches) != 1:
        raise TableRecipeProposalError(
            f"Expected one table {table_id}, found {len(matches)}."
        )
    return matches[0]


def semantic_header_signature(table: dict[str, Any]) -> list[dict[str, Any]]:
    """Return the ordered, value-free semantic identity of numeric axes."""

    bounds = table.get("bounds") or {}
    min_column = int(bounds.get("minColumn") or 0)
    if min_column < 1:
        raise TableRecipeProposalError("Table bounds are invalid.")
    result = [
        {
            "relativeColumn": (
                _column_number(str(column.get("column") or ""))
                - min_column
            ),
            "columnRole": str(column.get("columnRole") or ""),
            "headerTexts": [
                " ".join(str(value).casefold().split())
                for value in column.get("headerTexts") or []
            ],
        }
        for column in table.get("numericColumns") or []
        if (
            str(column.get("columnRole") or "") == "MEASURE_VALUE"
            or str(column.get("columnRole") or "").startswith("AGGREGATE_")
        )
    ]
    return sorted(result, key=lambda item: int(item["relativeColumn"]))


def _semantic_header_sha256(signature: list[dict[str, Any]]) -> str:
    return hashlib.sha256(_canonical(signature).encode("utf-8")).hexdigest()


def _measure_column_count(fingerprint: dict[str, Any]) -> int:
    return sum(
        str(column.get("columnRole") or "") == "MEASURE_VALUE"
        for column in fingerprint.get("numericColumns") or []
    )


def _safety_status(
    structure: dict[str, Any],
    *,
    minimum_measure_columns: int = 2,
) -> tuple[str, list[str]]:
    fingerprint = structure.get("fingerprint") or {}
    numeric_count = int(fingerprint.get("numericColumnCount") or 0)
    measure_count = _measure_column_count(fingerprint)
    header_tokens = {
        str(value).upper() for value in fingerprint.get("headerTokens") or []
    }
    reasons: list[str] = []
    minimum = max(int(minimum_measure_columns), 1)
    if measure_count < minimum:
        return "NOT_PARAMETER_TABLE", [
            f"FEWER_THAN_{minimum}_MEASURE_COLUMNS"
        ]
    if numeric_count > 32:
        reasons.append("HIGH_DIMENSIONAL_NUMERIC_MATRIX")
    if numeric_count > 20 and (
        "FREQUENCY" in header_tokens or "HZ" in header_tokens
    ):
        reasons.append("POSSIBLE_RAW_FREQUENCY_MATRIX")
    return (
        ("REVIEW_BEFORE_PROPOSAL" if reasons else "PROPOSAL_READY"),
        reasons,
    )


def _priority_score(structure: dict[str, Any]) -> float:
    fingerprint = structure.get("fingerprint") or {}
    workbooks = int(structure.get("workbookCount") or 0)
    tables = int(structure.get("tableCount") or 0)
    numeric = int(fingerprint.get("numericColumnCount") or 0)
    measures = _measure_column_count(fingerprint)
    repetition = min(workbooks / 10.0, 1.0)
    coverage = min(tables / 15.0, 1.0)
    measure_density = min(measures / max(numeric, 1), 1.0)
    compactness = 1.0 if numeric <= 20 else max(0.0, 1 - (numeric - 20) / 80)
    header_signal = min(
        len(fingerprint.get("headerTokens") or []) / 12.0,
        1.0,
    )
    return round(
        0.45 * repetition
        + 0.15 * coverage
        + 0.15 * measure_density
        + 0.15 * compactness
        + 0.10 * header_signal,
        6,
    )


def _historical_candidate(
    structure: dict[str, Any],
    similarity: dict[str, Any],
    *,
    verified_recipes: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    structure_id = str(structure.get("tableStructureId") or "")
    recipe = verified_recipes.get(structure_id)
    return {
        "tableStructureId": structure_id,
        "fingerprintSha256": str(structure.get("fingerprintSha256") or ""),
        "score": similarity["score"],
        "hardGatePassed": similarity["hardGatePassed"],
        "hardGateFailures": similarity["hardGateFailures"],
        "components": similarity["components"],
        "tableCount": int(structure.get("tableCount") or 0),
        "workbookCount": int(structure.get("workbookCount") or 0),
        "dominantSemanticType": str(
            structure.get("dominantSemanticType") or "UNASSESSED"
        ),
        "semanticConsistency": float(
            structure.get("semanticConsistency") or 0.0
        ),
        "headerTokens": list(
            (structure.get("fingerprint") or {}).get("headerTokens") or []
        )[:40],
        "numericColumns": list(
            (structure.get("fingerprint") or {}).get("numericColumns") or []
        ),
        "verifiedRecipe": (
            {
                "recipeId": str(recipe.get("recipeId") or ""),
                "recipeVersion": int(recipe.get("recipeVersion") or 0),
                "status": str(recipe.get("status") or ""),
                "semanticContract": recipe.get("semanticContract") or {},
            }
            if recipe
            else None
        ),
    }


def _verified_recipes(recipe_root: str | Path | None) -> dict[str, dict[str, Any]]:
    if recipe_root is None:
        return {}
    root = Path(recipe_root).expanduser().resolve()
    if not root.is_dir():
        return {}
    result: dict[str, dict[str, Any]] = {}
    for path in sorted(root.glob("*.json")):
        value = _read_json(path)
        if value.get("status") != "VERIFIED_HISTORICAL_REPLAY":
            continue
        structure_id = str(
            (value.get("match") or {}).get("tableStructureId") or ""
        )
        if structure_id:
            result[structure_id] = value
    return result


def build_table_recipe_priority_report(
    *,
    new_catalog_path: str | Path,
    historical_catalog_path: str | Path,
    new_batch_root: str | Path,
    historical_recipe_root: str | Path | None = None,
    top_k: int = 3,
    minimum_measure_columns: int = 2,
    generated_at: str | None = None,
) -> dict[str, Any]:
    """Build an AI-free queue with one redacted representative per structure."""

    new_catalog_file = Path(new_catalog_path).expanduser().resolve()
    historical_catalog_file = (
        Path(historical_catalog_path).expanduser().resolve()
    )
    new_root = Path(new_batch_root).expanduser().resolve()
    new_catalog = _read_json(new_catalog_file)
    historical_catalog = _read_json(historical_catalog_file)
    verified = _verified_recipes(historical_recipe_root)
    historical_structures = [
        (structure, _prepared_similarity(structure.get("fingerprint") or {}))
        for structure in historical_catalog.get("structures") or []
        if int((structure.get("fingerprint") or {}).get("numericColumnCount") or 0)
        >= 2
    ]
    request_cache: dict[str, dict[str, Any]] = {}

    queue: list[dict[str, Any]] = []
    for structure in new_catalog.get("structures") or []:
        if int(structure.get("workbookCount") or 0) < 2:
            continue
        fingerprint = structure.get("fingerprint") or {}
        safety_status, safety_reasons = _safety_status(
            structure,
            minimum_measure_columns=minimum_measure_columns,
        )
        if safety_status == "NOT_PARAMETER_TABLE":
            continue
        base_structure_id = str(structure.get("tableStructureId") or "")
        semantic_variants: dict[str, dict[str, Any]] = {}
        for member in structure.get("members") or []:
            request_file = str(member["requestFile"])
            request = request_cache.get(request_file)
            if request is None:
                request = _read_json(
                    new_root / "requests" / request_file
                )
                request_cache[request_file] = request
            table = _table_by_id(request, str(member["tableId"]))
            signature = semantic_header_signature(table)
            signature_sha256 = _semantic_header_sha256(signature)
            variant = semantic_variants.setdefault(
                signature_sha256,
                {
                    "signature": signature,
                    "members": [],
                    "tables": [],
                },
            )
            variant["members"].append(member)
            variant["tables"].append(table)
        if not semantic_variants:
            continue
        primary_member = list(structure.get("members") or [])[0]
        primary_request = request_cache[str(primary_member["requestFile"])]
        primary_signature_sha256 = _semantic_header_sha256(
            semantic_header_signature(
                _table_by_id(
                    primary_request,
                    str(primary_member["tableId"]),
                )
            )
        )
        ranked: list[dict[str, Any]] = []
        prepared_target = _prepared_similarity(fingerprint)
        for historical, prepared_historical in historical_structures:
            similarity = _score_prepared_similarity(
                prepared_target,
                prepared_historical,
            )
            if similarity["hardGatePassed"]:
                ranked.append(
                    _historical_candidate(
                        historical,
                        similarity,
                        verified_recipes=verified,
                    )
                )
        ranked.sort(
            key=lambda item: (
                -float(item["score"]),
                -int(item["workbookCount"]),
                str(item["tableStructureId"]),
            )
        )
        for signature_sha256, variant in semantic_variants.items():
            members = list(variant["members"])
            workbook_count = len(
                {
                    str(member.get("contentSha256") or "")
                    or str(member.get("revisionUid") or "")
                    or str(member.get("requestFile") or "")
                    for member in members
                }
            )
            if workbook_count < 2:
                continue
            variant_structure_id = (
                base_structure_id
                if signature_sha256 == primary_signature_sha256
                else (
                    base_structure_id
                    + "-semantic-"
                    + signature_sha256[:8]
                )
            )
            variant_for_priority = {
                **structure,
                "tableCount": len(members),
                "workbookCount": workbook_count,
            }
            representative_member = members[0]
            representative_table = variant["tables"][0]
            queue.append(
                {
                    "rank": 0,
                    "priorityScore": _priority_score(
                        variant_for_priority
                    ),
                    "status": safety_status,
                    "safetyReasons": safety_reasons,
                    "tableStructureId": variant_structure_id,
                    "baseTableStructureId": base_structure_id,
                    "fingerprintSha256": str(
                        structure.get("fingerprintSha256") or ""
                    ),
                    "semanticHeaderSha256": signature_sha256,
                    "semanticHeaderSignature": variant["signature"],
                    "tableCount": len(members),
                    "workbookCount": workbook_count,
                    "measureColumnCount": _measure_column_count(fingerprint),
                    "numericColumnCount": int(
                        fingerprint.get("numericColumnCount") or 0
                    ),
                    "representativeMember": {
                        key: representative_member.get(key)
                        for key in (
                            "fileName",
                            "requestFile",
                            "tableId",
                            "sheet",
                            "range",
                        )
                    },
                    "representativeTable": redact_representative_table(
                        representative_table
                    ),
                    "historicalTopK": ranked[: max(1, int(top_k))],
                }
            )
    queue.sort(
        key=lambda item: (
            item["status"] != "PROPOSAL_READY",
            -float(item["priorityScore"]),
            -int(item["workbookCount"]),
            str(item["tableStructureId"]),
        )
    )
    for rank, item in enumerate(queue, start=1):
        item["rank"] = rank
    return {
        "schemaVersion": PRIORITY_REPORT_SCHEMA_VERSION,
        "engineVersion": PROPOSAL_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            "newCatalog": str(new_catalog_file),
            "historicalCatalog": str(historical_catalog_file),
            "newBatchRoot": str(new_root),
            "sourceDatabase": (
                str(
                    (
                        _read_json(new_root / "report.json").get(
                            "inputs"
                        )
                        or {}
                    ).get("databasePath")
                    or ""
                )
                if (new_root / "report.json").is_file()
                else ""
            ),
            "historicalRecipeRoot": (
                str(Path(historical_recipe_root).expanduser().resolve())
                if historical_recipe_root is not None
                else None
            ),
            "historicalCandidateCount": len(historical_structures),
            "topK": max(1, int(top_k)),
            "minimumMeasureColumns": max(
                int(minimum_measure_columns),
                1,
            ),
            "aiCalls": 0,
        },
        "summary": {
            "queuedStructureCount": len(queue),
            "proposalReadyStructureCount": sum(
                item["status"] == "PROPOSAL_READY" for item in queue
            ),
            "reviewBeforeProposalCount": sum(
                item["status"] == "REVIEW_BEFORE_PROPOSAL"
                for item in queue
            ),
            "coveredTableCount": sum(
                int(item["tableCount"]) for item in queue
            ),
            "coveredWorkbookReferences": sum(
                int(item["workbookCount"]) for item in queue
            ),
            "aiCalls": 0,
        },
        "queue": queue,
    }


def build_priority_extension_report(
    *,
    expanded_priority_report: dict[str, Any],
    baseline_priority_report: dict[str, Any],
    generated_at: str | None = None,
) -> dict[str, Any]:
    """Select only structures newly authorized by an expanded safety policy."""

    baseline_ids = {
        str(item.get("tableStructureId") or "")
        for item in baseline_priority_report.get("queue") or []
    }
    queue = [
        json.loads(json.dumps(item))
        for item in expanded_priority_report.get("queue") or []
        if str(item.get("tableStructureId") or "") not in baseline_ids
    ]
    return {
        "schemaVersion": PRIORITY_REPORT_SCHEMA_VERSION,
        "engineVersion": PROPOSAL_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            **json.loads(
                json.dumps(expanded_priority_report.get("inputs") or {})
            ),
            "baselineQueueStructureCount": len(baseline_ids),
            "extensionSelection": "EXPANDED_MINUS_BASELINE_STRUCTURE_ID",
        },
        "summary": {
            "queuedStructureCount": len(queue),
            "proposalReadyStructureCount": sum(
                item["status"] == "PROPOSAL_READY" for item in queue
            ),
            "reviewBeforeProposalCount": sum(
                item["status"] == "REVIEW_BEFORE_PROPOSAL"
                for item in queue
            ),
            "coveredTableCount": sum(
                int(item["tableCount"]) for item in queue
            ),
            "coveredWorkbookReferences": sum(
                int(item["workbookCount"]) for item in queue
            ),
            "aiCalls": 0,
        },
        "queue": queue,
    }


def _priority_item(
    report: dict[str, Any],
    table_structure_id: str,
) -> dict[str, Any]:
    matches = [
        item
        for item in report.get("queue") or []
        if str(item.get("tableStructureId") or "") == table_structure_id
    ]
    if len(matches) != 1:
        raise TableRecipeProposalError(
            f"Expected one queued structure {table_structure_id}, "
            f"found {len(matches)}."
        )
    return matches[0]


def table_recipe_decision_schema() -> dict[str, Any]:
    text_array = {"type": "array", "items": {"type": "string"}}
    group = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "label": {"type": "string"},
            "role": {"type": "string", "enum": list(GROUP_ROLES)},
            "basis": {"type": "string"},
        },
        "required": ["label", "role", "basis"],
    }
    metric = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "relativeColumn": {"type": "integer", "minimum": 0},
            "canonicalName": {"type": "string"},
            "unit": {"type": "string"},
        },
        "required": ["relativeColumn", "canonicalName", "unit"],
    }
    relation = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "leftGroup": {"type": "string"},
            "rightGroup": {"type": "string"},
            "basis": {"type": "string"},
        },
        "required": ["leftGroup", "rightGroup", "basis"],
    }
    semantic_contract = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "title": {"type": "string"},
            "tableType": {"type": "string", "enum": list(TABLE_TYPES)},
            "studyGroup": {"type": "string"},
            "groups": {"type": "array", "items": group},
            "metricColumns": {"type": "array", "items": metric},
            "comparisonRelations": {"type": "array", "items": relation},
            "limitations": text_array,
        },
        "required": [
            "title",
            "tableType",
            "studyGroup",
            "groups",
            "metricColumns",
            "comparisonRelations",
            "limitations",
        ],
    }
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": PROPOSAL_DECISION_SCHEMA_VERSION,
            },
            "targetTableStructureId": {"type": "string"},
            "targetFingerprintSha256": {"type": "string"},
            "decision": {"type": "string", "enum": list(DECISIONS)},
            "historicalSourceTableStructureId": {"type": "string"},
            "confidence": {
                "type": "string",
                "enum": list(CONFIDENCE_LEVELS),
            },
            "rationale": {"type": "string"},
            "semanticContract": semantic_contract,
        },
        "required": [
            "schemaVersion",
            "targetTableStructureId",
            "targetFingerprintSha256",
            "decision",
            "historicalSourceTableStructureId",
            "confidence",
            "rationale",
            "semanticContract",
        ],
    }


def build_table_recipe_decision_prompt(
    report: dict[str, Any],
    *,
    table_structure_id: str,
) -> str:
    item = _priority_item(report, table_structure_id)
    request = {
        "policy": {
            "oneDecisionForWholeExactStructure": True,
            "rawValuesIncluded": False,
            "statisticsIncluded": False,
            "aiMayWriteValues": False,
            "codeExtractsValuesAndEvidence": True,
            "reuseHistoricalOnlyWhenSemanticsAndRelativeAxesFit": True,
            "quarantineWhenAmbiguous": True,
        },
        "target": {
            key: item[key]
            for key in (
                "tableStructureId",
                "fingerprintSha256",
                "tableCount",
                "workbookCount",
                "representativeTable",
            )
        },
        "historicalTopK": item.get("historicalTopK") or [],
    }
    return (
        "Decide one semantic extraction recipe for one exact repeated Excel "
        "table structure. The representative contains labels and geometry only; "
        "all numeric/formula values and statistics were removed. Do not invent, "
        "calculate, reproduce, or request values. Deterministic code will map "
        "relative metric columns to each matching table and copy the captured "
        "numeric facts and exact source ranges.\n"
        "Choose REUSE_HISTORICAL_RECIPE only if one supplied historical candidate "
        "has the same semantic purpose and compatible relative metric axes. "
        "Otherwise choose NEW_RECIPE when labels are sufficient, or QUARANTINE "
        "when they are not. A descriptive result table is DESCRIPTIVE, not "
        "SUPPORTING. Use COMPARISON only for an explicit contrast between two "
        "source-authored conditions. Metric relativeColumn values must reference "
        "numericColumns supplied in target.representativeTable. Do not treat dates, "
        "sample IDs, input counts, or OK counts as metrics unless the labels clearly "
        "make them measured outcomes. Every comparisonRelations leftGroup and "
        "rightGroup must exactly copy two distinct label values declared in groups. "
        "Return no fields outside the schema.\n\n"
        "STRUCTURE_REQUEST_JSON:\n"
        + json.dumps(request, ensure_ascii=False, separators=(",", ":"))
    )


def validate_table_recipe_decision(
    decision: dict[str, Any],
    *,
    priority_item: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(decision, dict):
        raise TableRecipeProposalError("Recipe decision must be an object.")
    expected_keys = set(table_recipe_decision_schema()["required"])
    if set(decision) != expected_keys:
        raise TableRecipeProposalError(
            "Recipe decision contains missing or additional fields."
        )
    if decision.get("schemaVersion") != PROPOSAL_DECISION_SCHEMA_VERSION:
        raise TableRecipeProposalError("Invalid decision schemaVersion.")
    if decision.get("targetTableStructureId") != priority_item.get(
        "tableStructureId"
    ):
        raise TableRecipeProposalError("Decision structure identity mismatch.")
    if decision.get("targetFingerprintSha256") != priority_item.get(
        "fingerprintSha256"
    ):
        raise TableRecipeProposalError("Decision fingerprint mismatch.")
    if decision.get("decision") not in DECISIONS:
        raise TableRecipeProposalError("Invalid recipe decision.")
    if decision.get("confidence") not in CONFIDENCE_LEVELS:
        raise TableRecipeProposalError("Invalid decision confidence.")
    for field in ("rationale", "historicalSourceTableStructureId"):
        if not isinstance(decision.get(field), str):
            raise TableRecipeProposalError(f"{field} must be a string.")

    historical_ids = {
        str(item.get("tableStructureId") or "")
        for item in priority_item.get("historicalTopK") or []
    }
    historical_source = decision["historicalSourceTableStructureId"]
    if decision["decision"] == "REUSE_HISTORICAL_RECIPE":
        if historical_source not in historical_ids:
            raise TableRecipeProposalError(
                "Historical reuse must reference a supplied Top-K candidate."
            )
    elif historical_source:
        raise TableRecipeProposalError(
            "Only historical reuse may declare a historical source."
        )

    contract = decision.get("semanticContract")
    required_contract_keys = {
        "title",
        "tableType",
        "studyGroup",
        "groups",
        "metricColumns",
        "comparisonRelations",
        "limitations",
    }
    if not isinstance(contract, dict) or set(contract) != required_contract_keys:
        raise TableRecipeProposalError("Invalid semanticContract fields.")
    if contract.get("tableType") not in TABLE_TYPES:
        raise TableRecipeProposalError("Invalid semantic table type.")
    for field in ("title", "studyGroup"):
        if not isinstance(contract.get(field), str):
            raise TableRecipeProposalError(f"semanticContract.{field} invalid.")
    for field in (
        "groups",
        "metricColumns",
        "comparisonRelations",
        "limitations",
    ):
        if not isinstance(contract.get(field), list):
            raise TableRecipeProposalError(
                f"semanticContract.{field} must be a list."
            )
    if any(not isinstance(value, str) for value in contract["limitations"]):
        raise TableRecipeProposalError("limitations must contain strings.")

    group_labels: set[str] = set()
    normalized_group_labels: dict[str, str] = {}
    for group in contract["groups"]:
        if (
            not isinstance(group, dict)
            or set(group) != {"label", "role", "basis"}
            or not isinstance(group.get("label"), str)
            or not group["label"].strip()
            or group.get("role") not in GROUP_ROLES
            or not isinstance(group.get("basis"), str)
        ):
            raise TableRecipeProposalError("Invalid semantic group.")
        group_label = group["label"]
        normalized_label = " ".join(group_label.split()).casefold()
        if (
            group_label in group_labels
            or normalized_label in normalized_group_labels
        ):
            raise TableRecipeProposalError("Duplicate semantic group label.")
        group_labels.add(group_label)
        normalized_group_labels[normalized_label] = group_label

    allowed_columns = {
        int(column["relativeColumn"]): str(column.get("columnRole") or "")
        for column in (
            priority_item.get("representativeTable") or {}
        ).get("numericColumns") or []
    }
    seen_columns: set[int] = set()
    for metric in contract["metricColumns"]:
        if (
            not isinstance(metric, dict)
            or set(metric) != {"relativeColumn", "canonicalName", "unit"}
            or not isinstance(metric.get("relativeColumn"), int)
            or metric["relativeColumn"] not in allowed_columns
            or metric["relativeColumn"] in seen_columns
            or not isinstance(metric.get("canonicalName"), str)
            or not metric["canonicalName"].strip()
            or not isinstance(metric.get("unit"), str)
        ):
            raise TableRecipeProposalError("Invalid metric column mapping.")
        seen_columns.add(metric["relativeColumn"])
    if decision["decision"] != "QUARANTINE" and not seen_columns:
        raise TableRecipeProposalError(
            "An executable recipe requires at least one metric column."
        )

    relations = contract["comparisonRelations"]
    if contract["tableType"] == "COMPARISON" and not relations:
        raise TableRecipeProposalError(
            "COMPARISON requires at least one relation."
        )
    if contract["tableType"] != "COMPARISON" and relations:
        raise TableRecipeProposalError(
            "Only COMPARISON may declare relations."
        )
    for relation in relations:
        if (
            not isinstance(relation, dict)
            or set(relation) != {"leftGroup", "rightGroup", "basis"}
            or not isinstance(relation.get("leftGroup"), str)
            or not isinstance(relation.get("rightGroup"), str)
            or not isinstance(relation.get("basis"), str)
        ):
            raise TableRecipeProposalError("Invalid comparison relation.")
        resolved_groups: list[str] = []
        for field in ("leftGroup", "rightGroup"):
            group_reference = relation[field]
            if group_reference in group_labels:
                resolved_groups.append(group_reference)
                continue
            normalized_reference = " ".join(
                group_reference.split()
            ).casefold()
            resolved = normalized_group_labels.get(normalized_reference)
            if resolved is None:
                raise TableRecipeProposalError(
                    "Invalid comparison relation."
                )
            resolved_groups.append(resolved)
        if resolved_groups[0] == resolved_groups[1]:
            raise TableRecipeProposalError("Invalid comparison relation.")
        relation["leftGroup"], relation["rightGroup"] = resolved_groups
    return decision


def _codex_command(command: Sequence[str] | None) -> list[str]:
    if command:
        return list(command)
    executable = shutil.which("codex.cmd" if os.name == "nt" else "codex")
    if not executable:
        raise TableRecipeProposalError("Codex CLI executable was not found.")
    return [executable]


def run_codex_table_recipe_decision(
    *,
    priority_report: dict[str, Any],
    table_structure_id: str,
    output_path: str | Path,
    telemetry_path: str | Path,
    reasoning_effort: str = "low",
    model: str | None = None,
    codex_command: Sequence[str] | None = None,
    timeout_seconds: int = 600,
) -> dict[str, Any]:
    """Run exactly one bounded AI call; this function never retries."""

    item = _priority_item(priority_report, table_structure_id)
    if item.get("status") != "PROPOSAL_READY":
        raise TableRecipeProposalError(
            f"Structure status does not authorize AI: {item.get('status')}"
        )
    prompt = build_table_recipe_decision_prompt(
        priority_report,
        table_structure_id=table_structure_id,
    )
    started = time.monotonic()
    started_at = datetime.now(timezone.utc)
    telemetry: dict[str, Any] = {
        "engineVersion": PROPOSAL_ENGINE_VERSION,
        "tableStructureId": table_structure_id,
        "startedAt": started_at.isoformat(timespec="seconds"),
        "aiCallBudget": 1,
        "aiCallsAttempted": 1,
        "aiCallsSucceeded": 0,
        "retryCount": 0,
        "reasoningEffort": reasoning_effort,
        "model": model or "CODEX_DEFAULT",
        "promptBytes": len(prompt.encode("utf-8")),
        "status": "RUNNING",
    }
    decision: Any = None
    try:
        with tempfile.TemporaryDirectory(
            prefix="table-recipe-proposal-"
        ) as temp_dir:
            schema_path = Path(temp_dir) / "decision.schema.json"
            last_message_path = Path(temp_dir) / "last-message.json"
            schema_path.write_text(
                json.dumps(
                    table_recipe_decision_schema(),
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            command = [
                *_codex_command(codex_command),
                "exec",
                "--ephemeral",
                "--sandbox",
                "read-only",
                "--output-schema",
                str(schema_path),
                "--output-last-message",
                str(last_message_path),
                "-c",
                f'model_reasoning_effort="{reasoning_effort}"',
            ]
            if model:
                command.extend(["--model", model])
            command.append("-")
            completed = subprocess.run(
                command,
                input=prompt,
                text=True,
                encoding="utf-8",
                errors="replace",
                capture_output=True,
                timeout=timeout_seconds,
                check=False,
            )
            telemetry["exitCode"] = completed.returncode
            if completed.returncode != 0:
                detail = (completed.stderr or completed.stdout or "").strip()
                raise TableRecipeProposalError(
                    "Codex recipe decision failed with exit code "
                    f"{completed.returncode}: {detail[-2000:]}"
                )
            if not last_message_path.is_file():
                raise TableRecipeProposalError(
                    "Codex did not produce a recipe decision."
                )
            output_bytes = last_message_path.read_bytes()
            telemetry["outputBytes"] = len(output_bytes)
            try:
                decision = json.loads(output_bytes)
            except json.JSONDecodeError as exc:
                raise TableRecipeProposalError(
                    "Codex recipe decision is not valid JSON."
                ) from exc
        validated = validate_table_recipe_decision(
            decision,
            priority_item=item,
        )
        _write_json(output_path, validated)
        telemetry["aiCallsSucceeded"] = 1
        telemetry["status"] = "SUCCEEDED"
        return validated
    except Exception as exc:
        telemetry["status"] = "FAILED"
        telemetry["errorType"] = type(exc).__name__
        telemetry["error"] = str(exc)[-2000:]
        if isinstance(decision, dict):
            telemetry["rejectedDecision"] = decision
        raise
    finally:
        telemetry["finishedAt"] = datetime.now(timezone.utc).isoformat(
            timespec="seconds"
        )
        telemetry["durationMs"] = round(
            (time.monotonic() - started) * 1000
        )
        _write_json(telemetry_path, telemetry)


def compile_structure_recipe(
    decision: dict[str, Any],
    *,
    priority_item: dict[str, Any],
    representative_captured_cells: list[dict[str, Any]] | None = None,
    generated_at: str | None = None,
    decision_ai_calls: int = 1,
    decision_source: str = "BOUNDED_AI_STRUCTURE_DECISION",
) -> dict[str, Any]:
    validated = validate_table_recipe_decision(
        decision,
        priority_item=priority_item,
    )
    recipe_id = "structure-recipe-" + str(
        priority_item["tableStructureId"]
    ).removeprefix("table-structure-")
    role_by_relative = {
        int(column["relativeColumn"]): str(column.get("columnRole") or "")
        for column in (
            priority_item.get("representativeTable") or {}
        ).get("numericColumns") or []
    }
    source_contract = validated["semanticContract"]
    representative = priority_item.get("representativeTable") or {}
    text_locations: dict[str, list[dict[str, Any]]] = {}

    def add_text_location(
        text: Any,
        *,
        relative_row: int,
        relative_column: int,
        authority: str,
    ) -> None:
        normalized = " ".join(str(text or "").casefold().split())
        if not normalized:
            return
        location = {
            "relativeRow": int(relative_row),
            "relativeColumn": int(relative_column),
            "sourceExampleText": str(text),
            "calculationAuthority": authority,
        }
        locations = text_locations.setdefault(normalized, [])
        coordinate = (
            int(location["relativeRow"]),
            int(location["relativeColumn"]),
        )
        if all(
            (
                int(existing["relativeRow"]),
                int(existing["relativeColumn"]),
            )
            != coordinate
            for existing in locations
        ):
            locations.append(location)

    for row in representative.get("previewRows") or []:
        for cell in row.get("cells") or []:
            add_text_location(
                cell.get("text"),
                relative_row=int(cell.get("relativeRow") or 0),
                relative_column=int(cell.get("relativeColumn") or 0),
                authority="CODE_FROM_CAPTURED_TABLE_PREVIEW",
            )
    if representative_captured_cells is not None:
        min_row = 1
        min_column = 1
        representative_member = priority_item.get("representativeMember") or {}
        bounds = representative.get("bounds") or {}
        if not bounds:
            range_text = str(representative_member.get("range") or "")
            range_match = re.fullmatch(
                r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",
                range_text.upper(),
            )
            if range_match:
                min_row = int(range_match.group(2))
                min_column = _column_number(range_match.group(1))
        else:
            min_row = int(bounds.get("minRow") or 1)
            min_column = int(bounds.get("minColumn") or 1)
        for cell in representative_captured_cells:
            text = _captured_text_value(cell)
            if text is None:
                continue
            add_text_location(
                text,
                relative_row=int(cell.get("row") or 0) - min_row,
                relative_column=int(cell.get("column") or 0) - min_column,
                authority="CODE_FROM_CAPTURE_V2_DATABASE",
            )

    def resolve_group_location(label: str) -> dict[str, Any]:
        normalized = " ".join(str(label).casefold().split())
        exact = sorted(
            text_locations.get(normalized) or [],
            key=lambda value: (
                int(value["relativeRow"]),
                int(value["relativeColumn"]),
            ),
        )
        if exact:
            return {**exact[0], "matchMode": "EXACT_SOURCE_TEXT"}

        label_tokens = set(re.findall(r"[^\W_]+%?", normalized, re.UNICODE))
        if len(label_tokens) < 2:
            raise TableRecipeProposalError(
                "Group label must resolve to one representative source cell: "
                + str(label)
            )
        candidates: list[tuple[int, str, dict[str, Any]]] = []
        for source_text, locations in text_locations.items():
            source_tokens = set(
                re.findall(r"[^\W_]+%?", source_text, re.UNICODE)
            )
            if not label_tokens.issubset(source_tokens):
                continue
            for location in locations:
                candidates.append(
                    (
                        len(source_tokens - label_tokens),
                        source_text,
                        location,
                    )
                )
        if not candidates:
            raise TableRecipeProposalError(
                "Group label must resolve to one representative source cell: "
                + str(label)
            )
        candidates.sort(
            key=lambda value: (
                value[0],
                value[1],
                int(value[2]["relativeRow"]),
                int(value[2]["relativeColumn"]),
            )
        )
        best_extra_tokens = candidates[0][0]
        best_source_texts = {
            source_text
            for extra_tokens, source_text, _ in candidates
            if extra_tokens == best_extra_tokens
        }
        if len(best_source_texts) != 1:
            raise TableRecipeProposalError(
                "Group label source text is ambiguous: " + str(label)
            )
        return {
            **candidates[0][2],
            "matchMode": "SOURCE_TOKEN_SUBSET",
        }

    compiled_groups: list[dict[str, Any]] = []
    selector_by_label: dict[str, str] = {}
    for index, group in enumerate(source_contract["groups"], start=1):
        location = resolve_group_location(str(group["label"]))
        selector_id = f"group-{index}"
        selector_by_label[str(group["label"])] = selector_id
        compiled_groups.append(
            {
                "selectorId": selector_id,
                "canonicalExampleLabel": str(group["label"]),
                "sourceSelector": location,
                "role": str(group["role"]),
                "basis": str(group["basis"]),
            }
        )
    compiled_relations = [
        {
            "leftSelectorId": selector_by_label[relation["leftGroup"]],
            "rightSelectorId": selector_by_label[relation["rightGroup"]],
            "basis": relation["basis"],
        }
        for relation in source_contract["comparisonRelations"]
    ]
    contract = {
        "titleMode": "SOURCE_FIRST_TITLE_CANDIDATE",
        "titleExample": source_contract["title"],
        "tableType": source_contract["tableType"],
        "studyGroup": source_contract["studyGroup"],
        "groups": compiled_groups,
        "metricColumns": json.loads(
            json.dumps(source_contract["metricColumns"])
        ),
        "comparisonRelations": compiled_relations,
        "limitations": list(source_contract["limitations"]),
    }
    for metric in contract["metricColumns"]:
        metric["expectedColumnRole"] = role_by_relative[
            int(metric["relativeColumn"])
        ]
    return {
        "schemaVersion": STRUCTURE_RECIPE_SCHEMA_VERSION,
        "engineVersion": PROPOSAL_ENGINE_VERSION,
        "recipeId": recipe_id,
        "recipeVersion": 1,
        "status": (
            "QUARANTINED_BY_STRUCTURE_DECISION"
            if validated["decision"] == "QUARANTINE"
            else "DETERMINISTIC_REPLAY_PENDING"
        ),
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "match": {
            "tableStructureId": priority_item["tableStructureId"],
            "baseTableStructureId": priority_item.get(
                "baseTableStructureId",
                priority_item["tableStructureId"],
            ),
            "fingerprintSha256": priority_item["fingerprintSha256"],
            "semanticHeaderSha256": priority_item.get(
                "semanticHeaderSha256"
            ),
            "semanticHeaderSignature": list(
                priority_item.get("semanticHeaderSignature") or []
            ),
            "matchMode": "EXACT_TABLE_STRUCTURE_ONLY",
        },
        "decision": {
            "mode": validated["decision"],
            "historicalSourceTableStructureId": validated[
                "historicalSourceTableStructureId"
            ],
            "confidence": validated["confidence"],
            "rationale": validated["rationale"],
            "source": decision_source,
            "aiCallCount": max(int(decision_ai_calls), 0),
            "retryCount": 0,
        },
        "semanticContract": contract,
        "valueOwnership": {
            "values": "CODE_FROM_TABLE_FIRST_REQUEST",
            "statistics": "CODE_FROM_TABLE_FIRST_REQUEST",
            "evidence": "EXACT_SOURCE_COLUMN_RANGE",
            "aiMayWriteValues": False,
        },
    }


def adapt_decision_to_priority_item(
    source_decision: dict[str, Any],
    *,
    source_item: dict[str, Any],
    target_item: dict[str, Any],
) -> dict[str, Any]:
    """Adapt one verified semantic contract to a compatible source layout."""

    validated = validate_table_recipe_decision(
        source_decision,
        priority_item=source_item,
    )
    if str(source_item.get("semanticHeaderSha256") or "") != str(
        target_item.get("semanticHeaderSha256") or ""
    ):
        raise TableRecipeProposalError(
            "Semantic signature differs; decision propagation is forbidden."
        )
    source_recipe = compile_structure_recipe(
        validated,
        priority_item=source_item,
    )
    target_preview = target_item.get("representativeTable") or {}
    text_by_location: dict[tuple[int, int], str] = {}
    for row in target_preview.get("previewRows") or []:
        for cell in row.get("cells") or []:
            if cell.get("text") in (None, ""):
                continue
            key = (
                int(cell.get("relativeRow") or 0),
                int(cell.get("relativeColumn") or 0),
            )
            if key in text_by_location:
                raise TableRecipeProposalError(
                    "Target representative has duplicate text coordinates."
                )
            text_by_location[key] = str(cell["text"])

    adapted = json.loads(json.dumps(validated))
    adapted["targetTableStructureId"] = target_item["tableStructureId"]
    adapted["targetFingerprintSha256"] = target_item[
        "fingerprintSha256"
    ]
    adapted["decision"] = "NEW_RECIPE"
    adapted["historicalSourceTableStructureId"] = ""
    target_titles = list(target_preview.get("titleCandidates") or [])
    if target_titles:
        adapted["semanticContract"]["title"] = str(target_titles[0])
    source_groups = list(
        (source_recipe.get("semanticContract") or {}).get("groups") or []
    )
    target_groups: list[dict[str, Any]] = []
    target_label_by_source_label: dict[str, str] = {}
    for group in source_groups:
        selector = group.get("sourceSelector") or {}
        key = (
            int(selector.get("relativeRow") or 0),
            int(selector.get("relativeColumn") or 0),
        )
        target_label = text_by_location.get(key)
        if not target_label:
            raise TableRecipeProposalError(
                "Group selector is not compatible with target representative."
            )
        source_label = str(group.get("canonicalExampleLabel") or "")
        target_label_by_source_label[source_label] = target_label
        target_groups.append(
            {
                "label": target_label,
                "role": str(group.get("role") or "UNASSESSED"),
                "basis": str(group.get("basis") or ""),
            }
        )
    adapted["semanticContract"]["groups"] = target_groups
    adapted["semanticContract"]["comparisonRelations"] = [
        {
            "leftGroup": target_label_by_source_label[
                str(relation["leftGroup"])
            ],
            "rightGroup": target_label_by_source_label[
                str(relation["rightGroup"])
            ],
            "basis": str(relation.get("basis") or ""),
        }
        for relation in validated["semanticContract"][
            "comparisonRelations"
        ]
    ]
    adapted["rationale"] = (
        str(validated.get("rationale") or "")
        + " Contract propagated without a new AI call because the metric-header "
        "signature and all source group selectors are compatible."
    ).strip()
    return validate_table_recipe_decision(
        adapted,
        priority_item=target_item,
    )


def _numeric_columns_by_relative(
    table: dict[str, Any],
) -> dict[int, dict[str, Any]]:
    bounds = table.get("bounds") or {}
    min_column = int(bounds.get("minColumn") or 0)
    return {
        _column_number(str(column.get("column") or "")) - min_column: column
        for column in table.get("numericColumns") or []
    }


def _source_text_at(
    table: dict[str, Any],
    *,
    relative_row: int,
    relative_column: int,
    captured_cells: list[dict[str, Any]] | None = None,
) -> tuple[str, str]:
    bounds = table.get("bounds") or {}
    min_row = int(bounds.get("minRow") or 0)
    min_column = int(bounds.get("minColumn") or 0)
    if captured_cells is not None:
        absolute_row = min_row + int(relative_row)
        absolute_column = min_column + int(relative_column)
        captured_matches = [
            text
            for cell in captured_cells
            if int(cell.get("row") or 0) == absolute_row
            and int(cell.get("column") or 0) == absolute_column
            for text in [_captured_text_value(cell)]
            if text is not None
        ]
        if len(captured_matches) != 1:
            raise TableRecipeProposalError(
                "Capture v2 source text selector expected one cell at "
                f"R{relative_row}C{relative_column}, "
                f"found {len(captured_matches)}."
            )
        return captured_matches[0], "CODE_FROM_CAPTURE_V2_DATABASE"

    matches: list[str] = []
    for row in table.get("previewRows") or []:
        for cell in row.get("cells") or []:
            if str(cell.get("kind") or "") != "TEXT":
                continue
            relative = _relative_coordinate(
                cell.get("coordinate"),
                min_row=min_row,
                min_column=min_column,
            )
            if (
                relative["relativeRow"] == int(relative_row)
                and relative["relativeColumn"] == int(relative_column)
                and cell.get("value") not in (None, "")
            ):
                matches.append(str(cell["value"]))
    if len(matches) != 1:
        raise TableRecipeProposalError(
            "Source text selector expected one cell at "
            f"R{relative_row}C{relative_column}, found {len(matches)}."
        )
    return matches[0], "CODE_FROM_CAPTURED_TABLE_PREVIEW"


def _captured_text_value(cell: dict[str, Any]) -> str | None:
    values = (
        (
            cell.get("cachedValue"),
            cell.get("rawValue"),
            cell.get("displayValue"),
        )
        if cell.get("formula") not in (None, "")
        else (
            cell.get("rawValue"),
            cell.get("displayValue"),
            cell.get("cachedValue"),
        )
    )
    for value in values:
        if isinstance(value, str) and value.strip():
            return value
    return None


def _display_role(number_format: Any, source_display: Any) -> str:
    format_text = str(number_format or "").upper()
    display_text = str(source_display or "").strip()
    if "%" in format_text or display_text.endswith("%"):
        return "PERCENT"
    if any(token in format_text for token in ("YY", "MM", "DD")):
        return "DATE"
    return "NUMBER"


def _metric_cell_facts(
    table: dict[str, Any],
    *,
    relative_column: int,
    source_column: dict[str, Any],
    canonical_name: str,
    unit: str,
    captured_cells: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    bounds = table.get("bounds") or {}
    min_row = int(bounds.get("minRow") or 0)
    min_column = int(bounds.get("minColumn") or 0)
    expected_absolute = min_column + int(relative_column)
    result: list[dict[str, Any]] = []
    if captured_cells is not None:
        for cell in captured_cells:
            if int(cell.get("column") or 0) != expected_absolute:
                continue
            raw_value = cell.get("rawValue")
            cached_value = cell.get("cachedValue")
            authoritative = (
                cached_value
                if cell.get("formula") not in (None, "")
                else raw_value
            )
            if (
                not isinstance(authoritative, (int, float))
                or isinstance(authoritative, bool)
            ):
                continue
            display_value = cell.get("displayValue")
            result.append(
                {
                    "name": canonical_name,
                    "unit": unit,
                    "tableId": str(table.get("tableId") or ""),
                    "columnId": str(source_column.get("columnId") or ""),
                    "coordinate": str(cell.get("coordinate") or ""),
                    "relativeRow": int(cell.get("row") or 0) - min_row,
                    "rawValue": raw_value,
                    "cachedValue": cached_value,
                    "formula": cell.get("formula"),
                    "sourceDisplay": (
                        str(display_value)
                        if display_value not in (None, "")
                        else str(authoritative)
                    ),
                    "numberFormat": str(cell.get("numberFormat") or ""),
                    "displayRole": _display_role(
                        cell.get("numberFormat"),
                        display_value
                        if display_value not in (None, "")
                        else authoritative,
                    ),
                    "valueKind": (
                        "FORMULA"
                        if cell.get("formula") not in (None, "")
                        else "NUMBER"
                    ),
                    "calculationAuthority": "CODE_FROM_CAPTURE_V2_DATABASE",
                }
            )
        return sorted(
            result,
            key=lambda item: (
                int(item["relativeRow"]),
                str(item["coordinate"]),
            ),
        )
    for row in table.get("previewRows") or []:
        for cell in row.get("cells") or []:
            relative = _relative_coordinate(
                cell.get("coordinate"),
                min_row=min_row,
                min_column=min_column,
            )
            if (
                min_column + relative["relativeColumn"]
                != expected_absolute
                or str(cell.get("kind") or "") not in {"NUMBER", "FORMULA"}
                or cell.get("value") in (None, "")
            ):
                continue
            result.append(
                {
                    "name": canonical_name,
                    "unit": unit,
                    "tableId": str(table.get("tableId") or ""),
                    "columnId": str(source_column.get("columnId") or ""),
                    "coordinate": str(cell.get("coordinate") or ""),
                    "relativeRow": relative["relativeRow"],
                    "sourceDisplay": str(cell.get("value")),
                    "numberFormat": str(cell.get("numberFormat") or ""),
                    "displayRole": _display_role(
                        cell.get("numberFormat"),
                        cell.get("value"),
                    ),
                    "valueKind": str(cell.get("kind") or ""),
                    "calculationAuthority": (
                        "CODE_FROM_CAPTURED_TABLE_PREVIEW"
                    ),
                }
            )
    return sorted(
        result,
        key=lambda item: (
            int(item["relativeRow"]),
            str(item["coordinate"]),
        ),
    )


def execute_structure_recipe(
    recipe: dict[str, Any],
    table: dict[str, Any],
    *,
    captured_cells: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    if recipe.get("status") == "QUARANTINED_BY_STRUCTURE_DECISION":
        raise TableRecipeProposalError("Quarantined recipe cannot execute.")
    fingerprint = table_structure_fingerprint(table)
    match = recipe.get("match") or {}
    if fingerprint.get("fingerprintSha256") != match.get(
        "fingerprintSha256"
    ):
        raise TableRecipeProposalError(
            "Table fingerprint does not authorize recipe replay."
        )
    header_signature = semantic_header_signature(table)
    if header_signature != list(
        match.get("semanticHeaderSignature") or []
    ):
        raise TableRecipeProposalError(
            "Table semantic header signature does not authorize recipe replay."
        )
    source_columns = _numeric_columns_by_relative(table)
    facts: list[dict[str, Any]] = []
    cell_facts: list[dict[str, Any]] = []
    metrics: list[dict[str, Any]] = []
    for metric in (recipe.get("semanticContract") or {}).get(
        "metricColumns"
    ) or []:
        relative = int(metric["relativeColumn"])
        source = source_columns.get(relative)
        if source is None:
            raise TableRecipeProposalError(
                f"Missing source numeric column at relative {relative}."
            )
        if str(source.get("columnRole") or "") != str(
            metric.get("expectedColumnRole") or ""
        ):
            raise TableRecipeProposalError(
                f"Column role changed at relative {relative}."
            )
        metrics.append(
            {
                "name": metric["canonicalName"],
                "unit": metric["unit"],
                "axisRef": str(source.get("columnId") or ""),
                "relativeColumn": relative,
            }
        )
        facts.append(
            {
                "name": metric["canonicalName"],
                "unit": metric["unit"],
                "tableId": str(table.get("tableId") or ""),
                "columnId": str(source.get("columnId") or ""),
                "sourceRange": str(source.get("sourceRange") or ""),
                "columnRole": str(source.get("columnRole") or ""),
                "numericCount": int(source.get("numericCount") or 0),
                "min": source.get("min"),
                "max": source.get("max"),
                "average": source.get("average"),
                "calculationAuthority": "CODE_FROM_CAPTURED_RAW_VALUES",
            }
        )
        metric_cells = _metric_cell_facts(
            table,
            relative_column=relative,
            source_column=source,
            canonical_name=str(metric["canonicalName"]),
            unit=str(metric["unit"]),
            captured_cells=captured_cells,
        )
        if len(metric_cells) != int(source.get("numericCount") or 0):
            raise TableRecipeProposalError(
                f"Preview coverage is incomplete at relative {relative}: "
                f"{len(metric_cells)}/"
                f"{int(source.get('numericCount') or 0)} cells."
            )
        cell_facts.extend(metric_cells)
    contract = recipe.get("semanticContract") or {}
    extracted_groups: list[dict[str, Any]] = []
    group_label_by_selector: dict[str, str] = {}
    for group in contract.get("groups") or []:
        selector = group.get("sourceSelector") or {}
        label, label_authority = _source_text_at(
            table,
            relative_row=int(selector.get("relativeRow") or 0),
            relative_column=int(selector.get("relativeColumn") or 0),
            captured_cells=captured_cells,
        )
        selector_id = str(group.get("selectorId") or "")
        group_label_by_selector[selector_id] = label
        extracted_groups.append(
            {
                "label": label,
                "role": str(group.get("role") or "UNASSESSED"),
                "basis": str(group.get("basis") or ""),
                "sourceSelector": selector,
                "calculationAuthority": label_authority,
            }
        )
    extracted_relations = [
        {
            "leftGroup": group_label_by_selector[
                str(relation["leftSelectorId"])
            ],
            "rightGroup": group_label_by_selector[
                str(relation["rightSelectorId"])
            ],
            "basis": str(relation.get("basis") or ""),
        }
        for relation in contract.get("comparisonRelations") or []
    ]
    title_candidates = list(table.get("titleCandidates") or [])
    return {
        "recipeId": recipe["recipeId"],
        "tableStructureId": match["tableStructureId"],
        "fingerprintSha256": fingerprint["fingerprintSha256"],
        "tableId": str(table.get("tableId") or ""),
        "sheet": str(table.get("sheet") or ""),
        "range": str(table.get("range") or ""),
        "semantic": {
            "title": (
                str(title_candidates[0])
                if title_candidates
                else str(contract.get("titleExample") or "")
            ),
            "tableType": contract.get("tableType"),
            "studyGroup": contract.get("studyGroup"),
            "groups": extracted_groups,
            "metrics": metrics,
            "comparisonRelations": extracted_relations,
            "limitations": contract.get("limitations") or [],
            "confidence": (recipe.get("decision") or {}).get("confidence"),
        },
        "deterministicNumericFacts": facts,
        "deterministicCellFacts": cell_facts,
        "evidence": {
            "tableId": str(table.get("tableId") or ""),
            "sheet": str(table.get("sheet") or ""),
            "range": str(table.get("range") or ""),
            "authority": "CODE_FROM_TABLE_FIRST_REQUEST",
        },
    }


def _source_database_path(
    priority_report: dict[str, Any],
) -> Path | None:
    inputs = priority_report.get("inputs") or {}
    configured = str(inputs.get("sourceDatabase") or "")
    if configured:
        path = Path(configured).expanduser().resolve()
        return path if path.is_file() else None
    batch_root = Path(str(inputs.get("newBatchRoot") or ""))
    report_path = batch_root / "report.json"
    if not report_path.is_file():
        return None
    report = _read_json(report_path)
    database = str((report.get("inputs") or {}).get("databasePath") or "")
    if not database:
        return None
    path = Path(database).expanduser().resolve()
    return path if path.is_file() else None


def _captured_cells_for_table(
    connection: sqlite3.Connection,
    *,
    request: dict[str, Any],
    table: dict[str, Any],
) -> list[dict[str, Any]]:
    revision_uid = str((request.get("source") or {}).get("revisionUid") or "")
    sheet_index = int(table.get("sheetIndex") or 0)
    bounds = table.get("bounds") or {}
    if not revision_uid or sheet_index < 1:
        raise TableRecipeProposalError(
            "Request lacks Capture v2 revision/sheet identity."
        )
    rows = connection.execute(
        """
        SELECT
            c.row_index, c.column_index, c.coordinate,
            c.raw_value_json, c.formula_text, c.cached_value_json,
            c.display_value_json, c.number_format
        FROM capture_v2_revisions AS r
        JOIN capture_v2_sheets AS s
          ON s.revision_id=r.revision_id
        JOIN capture_v2_cells AS c
          ON c.sheet_id=s.sheet_id
        WHERE r.revision_uid=?
          AND s.sheet_index=?
          AND c.row_index BETWEEN ? AND ?
          AND c.column_index BETWEEN ? AND ?
        ORDER BY c.row_index, c.column_index
        """,
        (
            revision_uid,
            sheet_index,
            int(bounds.get("minRow") or 0),
            int(bounds.get("maxRow") or 0),
            int(bounds.get("minColumn") or 0),
            int(bounds.get("maxColumn") or 0),
        ),
    ).fetchall()
    return [
        {
            "row": int(row[0]),
            "column": int(row[1]),
            "coordinate": str(row[2]),
            "rawValue": _json_value(row[3]),
            "formula": row[4],
            "cachedValue": _json_value(row[5]),
            "displayValue": _json_value(row[6]),
            "numberFormat": str(row[7] or ""),
        }
        for row in rows
    ]


def load_representative_captured_cells(
    priority_report: dict[str, Any],
    priority_item: dict[str, Any],
) -> list[dict[str, Any]] | None:
    """Load the representative table's full Capture v2 cell range read-only."""

    database_path = _source_database_path(priority_report)
    if database_path is None:
        return None
    member = priority_item.get("representativeMember") or {}
    request_file = str(member.get("requestFile") or "")
    table_id = str(member.get("tableId") or "")
    if not request_file or not table_id:
        raise TableRecipeProposalError(
            "Priority item lacks representative request/table identity."
        )
    batch_root = Path(
        str((priority_report.get("inputs") or {}).get("newBatchRoot") or "")
    )
    request = _read_json(batch_root / "requests" / request_file)
    table = _table_by_id(request, table_id)
    connection = sqlite3.connect(
        database_path.as_uri() + "?mode=ro",
        uri=True,
    )
    try:
        connection.execute("PRAGMA query_only=ON")
        return _captured_cells_for_table(
            connection,
            request=request,
            table=table,
        )
    finally:
        connection.close()


def replay_structure_recipe(
    *,
    recipe: dict[str, Any],
    priority_report: dict[str, Any],
    generated_at: str | None = None,
) -> dict[str, Any]:
    if recipe.get("schemaVersion") != STRUCTURE_RECIPE_SCHEMA_VERSION:
        raise TableRecipeProposalError("Invalid structure recipe schema.")
    match = recipe.get("match") or {}
    structure_id = str(match.get("tableStructureId") or "")
    item = _priority_item(priority_report, structure_id)
    if recipe.get("status") == "QUARANTINED_BY_STRUCTURE_DECISION":
        return {
            "schemaVersion": STRUCTURE_REPLAY_SCHEMA_VERSION,
            "engineVersion": PROPOSAL_ENGINE_VERSION,
            "generatedAt": generated_at
            or datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "recipeId": recipe["recipeId"],
            "tableStructureId": structure_id,
            "status": "QUARANTINED",
            "summary": {
                "memberCount": int(item.get("tableCount") or 0),
                "passed": 0,
                "failed": 0,
                "aiCalls": 0,
                "recipeDecisionAiCalls": int(
                    (recipe.get("decision") or {}).get("aiCallCount") or 0
                ),
            },
            "items": [],
        }

    catalog = _read_json((priority_report.get("inputs") or {})["newCatalog"])
    base_structure_id = str(
        match.get("baseTableStructureId") or structure_id
    )
    structures = [
        structure
        for structure in catalog.get("structures") or []
        if str(structure.get("tableStructureId") or "") == base_structure_id
    ]
    if len(structures) != 1:
        raise TableRecipeProposalError(
            f"Expected one catalog structure {structure_id}."
        )
    batch_root = Path(
        (priority_report.get("inputs") or {})["newBatchRoot"]
    )
    database_path = _source_database_path(priority_report)
    connection: sqlite3.Connection | None = None
    if database_path is not None:
        connection = sqlite3.connect(
            database_path.as_uri() + "?mode=ro",
            uri=True,
        )
        connection.execute("PRAGMA query_only=ON")
    replay_items: list[dict[str, Any]] = []
    try:
        for member in structures[0].get("members") or []:
            try:
                request = _read_json(
                    batch_root / "requests" / str(member["requestFile"])
                )
                table = _table_by_id(request, str(member["tableId"]))
                if semantic_header_signature(table) != list(
                    match.get("semanticHeaderSignature") or []
                ):
                    continue
                captured_cells = (
                    _captured_cells_for_table(
                        connection,
                        request=request,
                        table=table,
                    )
                    if connection is not None
                    else None
                )
                extraction = execute_structure_recipe(
                    recipe,
                    table,
                    captured_cells=captured_cells,
                )
                failures: list[str] = []
                if not extraction["deterministicNumericFacts"]:
                    failures.append("NO_DETERMINISTIC_FACTS")
                if not extraction["deterministicCellFacts"]:
                    failures.append("NO_DETERMINISTIC_CELL_FACTS")
                if any(
                    not fact["columnId"] or not fact["sourceRange"]
                    for fact in extraction["deterministicNumericFacts"]
                ):
                    failures.append("MISSING_COLUMN_EVIDENCE")
                if any(
                    fact["calculationAuthority"]
                    != "CODE_FROM_CAPTURED_RAW_VALUES"
                    for fact in extraction["deterministicNumericFacts"]
                ):
                    failures.append("NON_CODE_OWNED_FACT")
                if any(
                    fact["calculationAuthority"]
                    not in {
                        "CODE_FROM_CAPTURED_TABLE_PREVIEW",
                        "CODE_FROM_CAPTURE_V2_DATABASE",
                    }
                    for fact in extraction["deterministicCellFacts"]
                ):
                    failures.append("NON_CODE_OWNED_CELL_FACT")
                replay_items.append(
                    {
                        "fileName": str(member.get("fileName") or ""),
                        "requestFile": str(member.get("requestFile") or ""),
                        "tableId": str(member.get("tableId") or ""),
                        "sheet": str(member.get("sheet") or ""),
                        "range": str(member.get("range") or ""),
                        "passed": not failures,
                        "failureCodes": failures,
                        "extraction": extraction,
                    }
                )
            except Exception as exc:
                replay_items.append(
                    {
                        "fileName": str(member.get("fileName") or ""),
                        "requestFile": str(member.get("requestFile") or ""),
                        "tableId": str(member.get("tableId") or ""),
                        "sheet": str(member.get("sheet") or ""),
                        "range": str(member.get("range") or ""),
                        "passed": False,
                        "failureCodes": [type(exc).__name__],
                        "error": str(exc),
                    }
                )
    finally:
        if connection is not None:
            connection.close()
    passed = sum(bool(value["passed"]) for value in replay_items)
    all_passed = len(replay_items) >= 2 and passed == len(replay_items)
    if all_passed:
        recipe["status"] = (
            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_NEEDS_CANONICAL_REVIEW"
        )
    else:
        recipe["status"] = "REPLAY_FAILED_QUARANTINED"
    return {
        "schemaVersion": STRUCTURE_REPLAY_SCHEMA_VERSION,
        "engineVersion": PROPOSAL_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "recipeId": recipe["recipeId"],
        "tableStructureId": structure_id,
        "status": recipe["status"],
        "summary": {
            "memberCount": len(replay_items),
            "passed": passed,
            "failed": len(replay_items) - passed,
            "deterministicFactCount": sum(
                len(
                    (value.get("extraction") or {}).get(
                        "deterministicNumericFacts"
                    )
                    or []
                )
                for value in replay_items
            ),
            "deterministicCellFactCount": sum(
                len(
                    (value.get("extraction") or {}).get(
                        "deterministicCellFacts"
                    )
                    or []
                )
                for value in replay_items
            ),
            "aiCalls": 0,
            "recipeDecisionAiCalls": int(
                (recipe.get("decision") or {}).get("aiCallCount") or 0
            ),
        },
        "items": replay_items,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Prioritize and replay bounded table structure recipes."
    )
    commands = parser.add_subparsers(dest="command", required=True)
    prioritize = commands.add_parser("prioritize")
    prioritize.add_argument("--new-catalog", required=True)
    prioritize.add_argument("--historical-catalog", required=True)
    prioritize.add_argument("--new-batch-root", required=True)
    prioritize.add_argument("--historical-recipe-root")
    prioritize.add_argument("--top-k", type=int, default=3)
    prioritize.add_argument("--minimum-measure-columns", type=int, default=2)
    prioritize.add_argument("--output", required=True)

    extend = commands.add_parser("extend")
    extend.add_argument("--expanded-priority-report", required=True)
    extend.add_argument("--baseline-priority-report", required=True)
    extend.add_argument("--output", required=True)

    decide = commands.add_parser("decide")
    decide.add_argument("--priority-report", required=True)
    decide.add_argument("--table-structure-id", required=True)
    decide.add_argument("--output", required=True)
    decide.add_argument("--telemetry-output", required=True)
    decide.add_argument("--reasoning-effort", default="low")
    decide.add_argument("--model")
    decide.add_argument("--timeout-seconds", type=int, default=600)

    replay = commands.add_parser("replay")
    replay.add_argument("--priority-report", required=True)
    replay.add_argument("--decision", required=True)
    replay.add_argument("--recipe-output", required=True)
    replay.add_argument("--replay-output", required=True)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    if arguments.command == "prioritize":
        report = build_table_recipe_priority_report(
            new_catalog_path=arguments.new_catalog,
            historical_catalog_path=arguments.historical_catalog,
            new_batch_root=arguments.new_batch_root,
            historical_recipe_root=arguments.historical_recipe_root,
            top_k=arguments.top_k,
            minimum_measure_columns=arguments.minimum_measure_columns,
        )
        _write_json(arguments.output, report)
        print(json.dumps(report["summary"], ensure_ascii=False))
        return 0
    if arguments.command == "extend":
        report = build_priority_extension_report(
            expanded_priority_report=_read_json(
                arguments.expanded_priority_report
            ),
            baseline_priority_report=_read_json(
                arguments.baseline_priority_report
            ),
        )
        _write_json(arguments.output, report)
        print(json.dumps(report["summary"], ensure_ascii=False))
        return 0
    if arguments.command == "decide":
        report = _read_json(arguments.priority_report)
        decision = run_codex_table_recipe_decision(
            priority_report=report,
            table_structure_id=arguments.table_structure_id,
            output_path=arguments.output,
            telemetry_path=arguments.telemetry_output,
            reasoning_effort=arguments.reasoning_effort,
            model=arguments.model,
            timeout_seconds=arguments.timeout_seconds,
        )
        print(
            json.dumps(
                {
                    "tableStructureId": decision[
                        "targetTableStructureId"
                    ],
                    "decision": decision["decision"],
                    "confidence": decision["confidence"],
                    "aiCalls": 1,
                },
                ensure_ascii=False,
            )
        )
        return 0
    report = _read_json(arguments.priority_report)
    decision = _read_json(arguments.decision)
    item = _priority_item(report, decision["targetTableStructureId"])
    recipe = compile_structure_recipe(
        decision,
        priority_item=item,
        representative_captured_cells=load_representative_captured_cells(
            report,
            item,
        ),
    )
    replay = replay_structure_recipe(
        recipe=recipe,
        priority_report=report,
    )
    _write_json(arguments.recipe_output, recipe)
    _write_json(arguments.replay_output, replay)
    print(json.dumps(replay["summary"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
