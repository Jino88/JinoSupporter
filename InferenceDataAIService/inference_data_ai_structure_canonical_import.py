"""Import verified structure replays into the canonical evidence database."""

from __future__ import annotations

import argparse
import json
import math
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable

from inference_data_ai_schema import validate_analysis_integrity
from inference_data_ai_study_contract import validate_study_manifest
from inference_data_ai_study_import import import_study_manifest
from inference_data_ai_table_recipe_proposal import _read_json, _write_json


IMPORT_SCHEMA_VERSION = "excel-structure-canonical-import-v1"
IMPORT_ENGINE_VERSION = "structure-canonical-import-v1.0"
ANALYSIS_KEY = "structure-reuse-incremental-v1"


class StructureCanonicalImportError(RuntimeError):
    """Raised when the incremental target cannot be closed safely."""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise StructureCanonicalImportError(message)


def _finite(value: object) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def _metric_unit(
    fact: dict[str, Any],
    cell_facts: list[dict[str, Any]],
) -> str:
    unit = str(fact.get("unit") or "").strip()
    if unit:
        return unit
    column_id = str(fact.get("columnId") or "")
    roles = {
        str(item.get("displayRole") or "").upper()
        for item in cell_facts
        if str(item.get("columnId") or "") == column_id
    }
    return "%" if "PERCENT" in roles else ""


def _evidence(
    *,
    sheet: str,
    address: str,
    role: str,
) -> list[dict[str, str]]:
    return [
        {
            "sheet": sheet,
            "range": address,
            "role": role,
        }
    ]


def _study_from_replay_item(item: dict[str, Any]) -> dict[str, Any]:
    extraction = dict(item.get("extraction") or {})
    semantic = dict(extraction.get("semantic") or {})
    table_id = str(extraction.get("tableId") or item.get("tableId") or "")
    sheet = str(extraction.get("sheet") or item.get("sheet") or "")
    table_range = str(extraction.get("range") or item.get("range") or "")
    recipe_id = str(extraction.get("recipeId") or "")
    numeric_facts = list(
        extraction.get("deterministicNumericFacts") or []
    )
    cell_facts = list(extraction.get("deterministicCellFacts") or [])
    _require(table_id != "", "Replay item is missing tableId.")
    _require(sheet != "", f"Replay table {table_id} is missing sheet.")
    _require(table_range != "", f"Replay table {table_id} is missing range.")
    _require(
        bool(numeric_facts),
        f"Replay table {table_id} has no deterministic numeric facts.",
    )

    table_evidence = _evidence(
        sheet=sheet,
        address=table_range,
        role="SOURCE_TABLE",
    )
    outcomes: list[dict[str, Any]] = []
    seen_keys: set[str] = set()
    for index, fact in enumerate(numeric_facts, start=1):
        column_id = str(fact.get("columnId") or "")
        source_range = str(fact.get("sourceRange") or table_range)
        key_suffix = column_id.rsplit("_", 1)[-1] if column_id else str(index)
        outcome_key = f"metric-{index}-{key_suffix}"
        if outcome_key in seen_keys:
            outcome_key = f"{outcome_key}-{index}"
        seen_keys.add(outcome_key)
        numeric_count = int(fact.get("numericCount") or 0)
        minimum = _finite(fact.get("min"))
        maximum = _finite(fact.get("max"))
        average = _finite(fact.get("average"))
        _require(
            numeric_count > 0
            and minimum is not None
            and maximum is not None
            and average is not None,
            f"Replay metric {table_id}/{outcome_key} is incomplete.",
        )
        observation: dict[str, Any] = {
            "key": "code-summary",
            "arm": "source-table",
            "min": minimum,
            "max": maximum,
            "average": average,
            "sampleSize": numeric_count,
            "status": "CODE_EXTRACTED",
            "verificationStatus": "NEEDS_REVIEW",
            "evidence": _evidence(
                sheet=sheet,
                address=source_range,
                role="OBSERVATION_SOURCE",
            ),
            "details": {
                "calculationAuthority": str(
                    fact.get("calculationAuthority") or ""
                ),
                "columnId": column_id,
                "numericCount": numeric_count,
                "recipeId": recipe_id,
                "tableId": table_id,
            },
        }
        if numeric_count == 1:
            observation["valueNumber"] = average
        metric_name = str(fact.get("name") or "").strip()
        if not metric_name:
            metric_name = f"Source metric {index}"
        outcomes.append(
            {
                "key": outcome_key,
                "originalLabel": metric_name,
                "canonicalName": metric_name,
                "metricType": "DESCRIPTIVE_STATISTIC",
                "unit": _metric_unit(fact, cell_facts),
                "favorableDirection": "UNKNOWN",
                "verificationStatus": "NEEDS_REVIEW",
                "evidence": _evidence(
                    sheet=sheet,
                    address=source_range,
                    role="OUTCOME_SOURCE",
                ),
                "observations": [observation],
                "attributes": {
                    "columnRole": str(fact.get("columnRole") or ""),
                    "recipeId": recipe_id,
                    "tableId": table_id,
                },
            }
        )

    semantic_title = str(semantic.get("title") or "").strip()
    table_type = str(semantic.get("tableType") or "DESCRIPTIVE").upper()
    return {
        "key": f"table-{table_id}",
        "title": semantic_title or f"{sheet}!{table_range}",
        "purpose": "Program-extracted quantitative source table",
        "designType": (
            table_type
            if table_type in {"DESCRIPTIVE", "COMPARISON"}
            else "DESCRIPTIVE"
        ),
        "verificationStatus": "NEEDS_REVIEW",
        "comparabilityStatus": "UNASSESSED",
        "confoundingStatus": "UNASSESSED",
        "status": "NEEDS_REVIEW",
        "summary": (
            f"{len(outcomes)} metrics were extracted by deterministic "
            f"recipe {recipe_id}; semantic interpretation needs review."
        ),
        "limitations": [
            "Values, statistics and evidence are code-owned.",
            "No comparison effect is eligible until canonical review.",
        ],
        "evidence": table_evidence,
        "contexts": [],
        "factors": [],
        "arms": [
            {
                "key": "source-table",
                "role": "OTHER",
                "label": "Source table",
                "verificationStatus": "NEEDS_REVIEW",
                "evidence": table_evidence,
                "factorValues": [],
            }
        ],
        "outcomes": outcomes,
        "measurementSeries": [],
        "comparisons": [],
        "conclusions": [],
    }


def build_recipe_manifest(
    *,
    source: dict[str, Any],
    replay_items: Iterable[dict[str, Any]],
) -> dict[str, Any]:
    items_by_table: dict[str, dict[str, Any]] = {}
    for item in replay_items:
        _require(bool(item.get("passed")), "Only passed replay items may import.")
        table_id = str(
            (item.get("extraction") or {}).get("tableId")
            or item.get("tableId")
            or ""
        )
        _require(table_id != "", "Replay item is missing tableId.")
        _require(
            table_id not in items_by_table,
            f"Duplicate replay tableId: {table_id}",
        )
        items_by_table[table_id] = item
    studies = [
        _study_from_replay_item(items_by_table[table_id])
        for table_id in sorted(items_by_table)
    ]
    _require(bool(studies), "Recipe manifest requires at least one study.")
    file_name = str(source.get("fileName") or Path(
        str(source["sourcePath"])
    ).name)
    manifest = {
        "schemaVersion": "canonical-study-manifest-v1",
        "source": {
            "dataset": str(source["dataset"]),
            "sourcePath": str(source["sourcePath"]),
            "revisionUid": str(source["revisionUid"]),
            "contentSha256": str(source.get("contentSha256") or ""),
            "contentComplete": True,
        },
        "workbookAnalysis": {
            "key": ANALYSIS_KEY,
            "title": file_name,
            "type": "STRUCTURE_REUSE",
            "purpose": "Canonical import of deterministic recipe replays",
            "scope": "Verified repeated quantitative table structures",
            "status": "NEEDS_REVIEW",
            "verificationStatus": "NEEDS_REVIEW",
            "summary": (
                f"{len(studies)} replay-verified quantitative tables were "
                "imported with program-owned observations and evidence."
            ),
            "limitations": [
                "Canonical semantic names and comparison eligibility need review.",
                "No AI-authored numeric value was imported.",
            ],
            "evidence": [],
        },
        "studies": studies,
    }
    return validate_study_manifest(manifest)


def build_terminal_manifest(
    *,
    source: dict[str, Any],
    verification_status: str,
    analysis_status: str,
    summary: str,
    limitations: Iterable[str],
) -> dict[str, Any]:
    file_name = str(source.get("fileName") or Path(
        str(source["sourcePath"])
    ).name)
    manifest = {
        "schemaVersion": "canonical-study-manifest-v1",
        "source": {
            "dataset": str(source["dataset"]),
            "sourcePath": str(source["sourcePath"]),
            "revisionUid": str(source["revisionUid"]),
            "contentSha256": str(source.get("contentSha256") or ""),
            "contentComplete": verification_status != "FAILED",
        },
        "workbookAnalysis": {
            "key": ANALYSIS_KEY,
            "title": file_name,
            "type": "STRUCTURE_REUSE",
            "purpose": "Incremental workbook terminal coverage",
            "scope": "Structure-reuse incremental batch",
            "status": analysis_status,
            "verificationStatus": verification_status,
            "summary": summary,
            "limitations": list(limitations),
            "evidence": [],
        },
        "studies": [],
    }
    return validate_study_manifest(manifest)


def _replay_items_by_revision(
    registry: dict[str, Any],
) -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = defaultdict(list)
    seen_tables: set[str] = set()
    for recipe in registry.get("recipes") or []:
        replay_path = Path(str(recipe.get("replayFile") or ""))
        _require(
            replay_path.is_file(),
            f"Registered replay is missing: {replay_path}",
        )
        replay = _read_json(replay_path)
        _require(
            int((replay.get("summary") or {}).get("failed") or 0) == 0,
            f"Registered replay contains failures: {replay_path}",
        )
        for item in replay.get("items") or []:
            table_id = str(item.get("tableId") or "")
            _require(
                table_id and table_id not in seen_tables,
                f"Replay table is duplicated or empty: {table_id}",
            )
            seen_tables.add(table_id)
            revision_uid = Path(
                str(item.get("requestFile") or "")
            ).stem
            _require(
                revision_uid.startswith("capture_revision_"),
                f"Replay requestFile lacks revision UID: {table_id}",
            )
            result[revision_uid].append(item)
    _require(
        len(seen_tables)
        == int((registry.get("summary") or {}).get(
            "registeredTableCount"
        ) or 0),
        "Replay table total differs from the registry.",
    )
    return dict(result)


def _target_capture_ids(
    table_match_report: dict[str, Any],
) -> tuple[set[int], dict[int, dict[str, Any]]]:
    success = {
        int(item["captureRevisionId"])
        for item in table_match_report.get("workbooks") or []
    }
    failures = {
        int(item["captureRevisionId"]): item
        for item in table_match_report.get("failures") or []
    }
    target = success | set(failures)
    expected = int(
        (table_match_report.get("summary") or {}).get(
            "eligibleWorkbookCount"
        )
        or 0
    )
    _require(
        len(target) == expected,
        f"Target capture IDs {len(target)} do not equal eligible {expected}.",
    )
    return target, failures


def _source_rows(
    connection: sqlite3.Connection,
    capture_ids: set[int],
) -> list[dict[str, Any]]:
    placeholders = ",".join("?" for _ in capture_ids)
    rows = connection.execute(
        f"""
        SELECT
            sr.revision_id, sr.revision_uid, sr.content_sha256,
            sr.capture_v2_revision_id, sd.dataset, sd.source_path,
            sd.original_file_name, cw.workbook_status
        FROM source_revisions sr
        JOIN source_documents sd ON sd.document_id=sr.document_id
        JOIN capture_v2_workbooks cw
          ON cw.revision_id=sr.capture_v2_revision_id
        WHERE sr.capture_v2_revision_id IN ({placeholders})
        ORDER BY sr.capture_v2_revision_id
        """,
        tuple(sorted(capture_ids)),
    ).fetchall()
    _require(
        len(rows) == len(capture_ids),
        "Not every target capture revision has one canonical source bridge.",
    )
    _require(
        len({int(row["capture_v2_revision_id"]) for row in rows})
        == len(capture_ids),
        "Target capture revision bridges are not one-to-one.",
    )
    return [
        {
            "revisionId": int(row["revision_id"]),
            "revisionUid": str(row["revision_uid"]),
            "contentSha256": str(row["content_sha256"]),
            "captureRevisionId": int(row["capture_v2_revision_id"]),
            "dataset": str(row["dataset"]),
            "sourcePath": str(row["source_path"]),
            "fileName": str(row["original_file_name"]),
            "workbookStatus": str(row["workbook_status"]),
        }
        for row in rows
    ]


def _active_analysis(
    connection: sqlite3.Connection,
    revision_id: int,
) -> list[dict[str, Any]]:
    return [
        {
            "workbookAnalysisId": int(row["workbook_analysis_id"]),
            "analysisKey": str(row["analysis_key"]),
            "analysisStatus": str(row["analysis_status"]),
            "verificationStatus": str(row["verification_status"]),
            "analyzerName": str(row["analyzer_name"]),
        }
        for row in connection.execute(
            """
            SELECT
                workbook_analysis_id, analysis_key, analysis_status,
                verification_status, analyzer_name
            FROM workbook_analyses
            WHERE revision_id=? AND verification_status<>'STALE'
            ORDER BY workbook_analysis_id
            """,
            (revision_id,),
        )
    ]


def build_import_plan(
    connection: sqlite3.Connection,
    *,
    table_match_report: dict[str, Any],
    registry: dict[str, Any],
) -> dict[str, Any]:
    capture_ids, failures = _target_capture_ids(table_match_report)
    replay_by_revision = _replay_items_by_revision(registry)
    sources = _source_rows(connection, capture_ids)
    source_uids = {str(source["revisionUid"]) for source in sources}
    _require(
        set(replay_by_revision) <= source_uids,
        "Registered replay contains a revision outside the target batch.",
    )

    actions: list[dict[str, Any]] = []
    for source in sources:
        active = _active_analysis(connection, int(source["revisionId"]))
        own = [
            value
            for value in active
            if value["analysisKey"] == ANALYSIS_KEY
        ]
        other = [
            value
            for value in active
            if value["analysisKey"] != ANALYSIS_KEY
        ]
        revision_uid = str(source["revisionUid"])
        replay_items = replay_by_revision.get(revision_uid, [])
        failure = failures.get(int(source["captureRevisionId"]))
        if other:
            action = "PRESERVE_EXISTING_CANONICAL"
            manifest = None
            reason = (
                f"{len(other)} non-stale canonical analysis record(s) exist."
            )
        elif replay_items:
            action = "IMPORT_RECIPE_REPLAY"
            manifest = build_recipe_manifest(
                source=source,
                replay_items=replay_items,
            )
            reason = (
                f"{len(replay_items)} replay-verified table(s) are ready."
            )
        elif failure is not None:
            action = "IMPORT_FAILED_TERMINAL"
            message = str(failure.get("message") or "")
            manifest = build_terminal_manifest(
                source=source,
                verification_status="FAILED",
                analysis_status="FAILED_TABLE_REQUEST",
                summary="Table-first semantic request could not be built.",
                limitations=[
                    str(failure.get("errorType") or "SemanticPacketError"),
                    message,
                ],
            )
            reason = message
        elif str(source["workbookStatus"]).upper() in {
            "EMPTY_WORKBOOK",
            "NO_TABULAR_EVIDENCE",
        }:
            action = "IMPORT_EXCLUDED_TERMINAL"
            status = str(source["workbookStatus"]).upper()
            manifest = build_terminal_manifest(
                source=source,
                verification_status="EXCLUDED",
                analysis_status=status,
                summary=f"Source workbook terminal status is {status}.",
                limitations=["No queryable tabular study was produced."],
            )
            reason = status
        else:
            action = "IMPORT_NEEDS_REVIEW_TERMINAL"
            manifest = build_terminal_manifest(
                source=source,
                verification_status="NEEDS_REVIEW",
                analysis_status="NEEDS_REVIEW_NO_VERIFIED_RECIPE",
                summary=(
                    "Source is captured but no replay-verified recipe result "
                    "is available for canonical observations."
                ),
                limitations=[
                    "The workbook remains visible in canonical coverage.",
                    "No numeric claim was imported without a verified recipe.",
                ],
            )
            reason = "No replay-verified recipe result."
        actions.append(
            {
                "action": action,
                "source": source,
                "existingAnalyses": active,
                "hadOwnAnalysis": bool(own),
                "reason": reason,
                "manifest": manifest,
            }
        )

    action_counts = Counter(item["action"] for item in actions)
    return {
        "schemaVersion": IMPORT_SCHEMA_VERSION,
        "engineVersion": IMPORT_ENGINE_VERSION,
        "status": "PLANNED",
        "summary": {
            "targetWorkbookCount": len(sources),
            "actionCounts": dict(sorted(action_counts.items())),
            "recipeReplayFileCount": len(replay_by_revision),
            "recipeReplayTableCount": sum(
                len(items) for items in replay_by_revision.values()
            ),
        },
        "actions": actions,
    }


def _target_database_coverage(
    connection: sqlite3.Connection,
    capture_ids: set[int],
) -> dict[str, Any]:
    placeholders = ",".join("?" for _ in capture_ids)
    params = tuple(sorted(capture_ids))
    rows = connection.execute(
        f"""
        SELECT
            sr.capture_v2_revision_id,
            COUNT(DISTINCT wa.workbook_analysis_id)
                AS active_analysis_count,
            COUNT(DISTINCT ks.study_id) AS study_count,
            COUNT(DISTINCT ko.observation_id) AS observation_count
        FROM source_revisions sr
        LEFT JOIN workbook_analyses wa
          ON wa.revision_id=sr.revision_id
         AND wa.verification_status<>'STALE'
        LEFT JOIN knowledge_studies ks
          ON ks.workbook_analysis_id=wa.workbook_analysis_id
        LEFT JOIN knowledge_outcomes kout ON kout.study_id=ks.study_id
        LEFT JOIN knowledge_observations ko
          ON ko.outcome_id=kout.outcome_id
        WHERE sr.capture_v2_revision_id IN ({placeholders})
        GROUP BY sr.capture_v2_revision_id
        ORDER BY sr.capture_v2_revision_id
        """,
        params,
    ).fetchall()
    active = sum(int(row["active_analysis_count"] or 0) > 0 for row in rows)
    active_analysis_count = sum(
        int(row["active_analysis_count"] or 0) for row in rows
    )
    maximum_active = max(
        (int(row["active_analysis_count"] or 0) for row in rows),
        default=0,
    )
    status_rows = connection.execute(
        f"""
        SELECT
            wa.analysis_status,
            wa.verification_status,
            COUNT(*) AS item_count
        FROM source_revisions sr
        JOIN workbook_analyses wa
          ON wa.revision_id=sr.revision_id
         AND wa.verification_status<>'STALE'
        WHERE sr.capture_v2_revision_id IN ({placeholders})
        GROUP BY wa.analysis_status, wa.verification_status
        ORDER BY wa.analysis_status, wa.verification_status
        """,
        params,
    ).fetchall()
    return {
        "targetWorkbookCount": len(capture_ids),
        "sourceRevisionCount": len(rows),
        "activeCanonicalAnalysisCount": active_analysis_count,
        "workbooksWithActiveCanonicalAnalysis": active,
        "workbooksWithoutActiveCanonicalAnalysis": len(capture_ids) - active,
        "workbooksWithMultipleActiveCanonicalAnalyses": sum(
            int(row["active_analysis_count"] or 0) > 1
            for row in rows
        ),
        "maximumActiveCanonicalAnalysesPerWorkbook": maximum_active,
        "studyCount": sum(int(row["study_count"] or 0) for row in rows),
        "observationCount": sum(
            int(row["observation_count"] or 0) for row in rows
        ),
        "analysisStatusCounts": [
            {
                "analysisStatus": str(row["analysis_status"]),
                "verificationStatus": str(row["verification_status"]),
                "count": int(row["item_count"]),
            }
            for row in status_rows
        ],
    }


def _importer_database_coverage(
    connection: sqlite3.Connection,
    capture_ids: set[int],
) -> dict[str, Any]:
    placeholders = ",".join("?" for _ in capture_ids)
    params = tuple(sorted(capture_ids))
    row = connection.execute(
        f"""
        WITH target AS (
            SELECT revision_id
            FROM source_revisions
            WHERE capture_v2_revision_id IN ({placeholders})
        ),
        own AS (
            SELECT wa.workbook_analysis_id, wa.analysis_uid
            FROM target t
            JOIN workbook_analyses wa ON wa.revision_id=t.revision_id
            WHERE wa.analysis_key=?
              AND wa.verification_status<>'STALE'
        ),
        entities(entity_type, entity_uid) AS (
            SELECT 'WORKBOOK_ANALYSIS', analysis_uid FROM own
            UNION ALL
            SELECT 'STUDY', ks.study_uid
            FROM own
            JOIN knowledge_studies ks USING(workbook_analysis_id)
            UNION ALL
            SELECT 'ARM', ka.arm_uid
            FROM own
            JOIN knowledge_studies ks USING(workbook_analysis_id)
            JOIN knowledge_arms ka USING(study_id)
            UNION ALL
            SELECT 'OUTCOME', kout.outcome_uid
            FROM own
            JOIN knowledge_studies ks USING(workbook_analysis_id)
            JOIN knowledge_outcomes kout USING(study_id)
            UNION ALL
            SELECT 'OBSERVATION', ko.observation_uid
            FROM own
            JOIN knowledge_studies ks USING(workbook_analysis_id)
            JOIN knowledge_outcomes kout USING(study_id)
            JOIN knowledge_observations ko USING(outcome_id)
        )
        SELECT
            (SELECT COUNT(*) FROM own) AS analysis_count,
            (
                SELECT COUNT(*)
                FROM own
                JOIN knowledge_studies ks USING(workbook_analysis_id)
            ) AS study_count,
            (
                SELECT COUNT(*)
                FROM own
                JOIN knowledge_studies ks USING(workbook_analysis_id)
                JOIN knowledge_arms ka USING(study_id)
            ) AS arm_count,
            (
                SELECT COUNT(*)
                FROM own
                JOIN knowledge_studies ks USING(workbook_analysis_id)
                JOIN knowledge_outcomes kout USING(study_id)
            ) AS outcome_count,
            (
                SELECT COUNT(*)
                FROM own
                JOIN knowledge_studies ks USING(workbook_analysis_id)
                JOIN knowledge_outcomes kout USING(study_id)
                JOIN knowledge_observations ko USING(outcome_id)
            ) AS observation_count,
            (
                SELECT COUNT(*)
                FROM entities e
                JOIN entity_evidence_links eel
                  ON eel.entity_type=e.entity_type
                 AND eel.entity_uid=e.entity_uid
            ) AS evidence_link_count,
            (
                SELECT COUNT(DISTINCT eel.evidence_id)
                FROM entities e
                JOIN entity_evidence_links eel
                  ON eel.entity_type=e.entity_type
                 AND eel.entity_uid=e.entity_uid
            ) AS distinct_evidence_count
        """,
        (*params, ANALYSIS_KEY),
    ).fetchone()
    status_rows = connection.execute(
        f"""
        SELECT
            wa.analysis_status,
            wa.verification_status,
            COUNT(*) AS item_count
        FROM source_revisions sr
        JOIN workbook_analyses wa ON wa.revision_id=sr.revision_id
        WHERE sr.capture_v2_revision_id IN ({placeholders})
          AND wa.analysis_key=?
          AND wa.verification_status<>'STALE'
        GROUP BY wa.analysis_status, wa.verification_status
        ORDER BY wa.analysis_status, wa.verification_status
        """,
        (*params, ANALYSIS_KEY),
    ).fetchall()
    return {
        "analysisCount": int(row["analysis_count"]),
        "studyCount": int(row["study_count"]),
        "armCount": int(row["arm_count"]),
        "outcomeCount": int(row["outcome_count"]),
        "observationCount": int(row["observation_count"]),
        "entityEvidenceLinkCount": int(row["evidence_link_count"]),
        "distinctEvidenceCount": int(row["distinct_evidence_count"]),
        "analysisStatusCounts": [
            {
                "analysisStatus": str(status["analysis_status"]),
                "verificationStatus": str(
                    status["verification_status"]
                ),
                "count": int(status["item_count"]),
            }
            for status in status_rows
        ],
    }


def apply_import_plan(
    connection: sqlite3.Connection,
    plan: dict[str, Any],
    *,
    now_iso: Callable[[], str] = _now,
) -> dict[str, Any]:
    imported: list[dict[str, Any]] = []
    connection.execute("BEGIN IMMEDIATE")
    try:
        for action in plan.get("actions") or []:
            if action["action"] == "PRESERVE_EXISTING_CANONICAL":
                continue
            manifest = action.get("manifest")
            _require(
                isinstance(manifest, dict),
                f"{action['action']} is missing its manifest.",
            )
            result = import_study_manifest(
                connection,
                manifest,
                now_iso=now_iso,
                source_claims_prevalidated=True,
            )
            integrity = validate_analysis_integrity(
                connection,
                workbook_analysis_id=int(result["workbookAnalysisId"]),
            )
            _require(
                bool(integrity.get("ok")),
                "Imported analysis integrity failed for "
                + str(action["source"]["revisionUid"]),
            )
            imported.append(
                {
                    "revisionUid": action["source"]["revisionUid"],
                    "action": action["action"],
                    "result": result,
                    "integrity": integrity,
                }
            )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    return {
        "importedWorkbookCount": len(imported),
        "imported": imported,
    }


def run_structure_canonical_import(
    *,
    database_path: str | Path,
    batch_root: str | Path,
    artifact_root: str | Path,
    output_path: str | Path,
    apply: bool,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    batch = Path(batch_root).expanduser().resolve()
    artifacts = Path(artifact_root).expanduser().resolve()
    output = Path(output_path).expanduser().resolve()
    report = _read_json(batch / "report.json")
    registry = _read_json(batch / "recipe-registry.json")

    if apply:
        connection = sqlite3.connect(database)
    else:
        connection = sqlite3.connect(
            f"file:{database.as_posix()}?mode=ro",
            uri=True,
        )
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys=ON")
    try:
        plan = build_import_plan(
            connection,
            table_match_report=report,
            registry=registry,
        )
        capture_ids, _failures = _target_capture_ids(report)
        before = _target_database_coverage(connection, capture_ids)
        importer_before = _importer_database_coverage(
            connection,
            capture_ids,
        )
        import_result: dict[str, Any] = {
            "importedWorkbookCount": 0,
            "imported": [],
        }
        if apply:
            artifacts.mkdir(parents=True, exist_ok=True)
            manifests = artifacts / "manifests"
            manifests.mkdir(parents=True, exist_ok=True)
            for action in plan["actions"]:
                if action.get("manifest") is None:
                    continue
                _write_json(
                    manifests
                    / f"{action['source']['revisionUid']}.json",
                    action["manifest"],
                )
            import_result = apply_import_plan(connection, plan)
        after = _target_database_coverage(connection, capture_ids)
        importer_after = _importer_database_coverage(
            connection,
            capture_ids,
        )
        if apply:
            _require(
                after["workbooksWithActiveCanonicalAnalysis"]
                == after["targetWorkbookCount"],
                "Applied import did not close canonical workbook coverage.",
            )
            _require(
                after["workbooksWithoutActiveCanonicalAnalysis"] == 0,
                "Applied import left canonical workbook coverage gaps.",
            )
            _require(
                after[
                    "workbooksWithMultipleActiveCanonicalAnalyses"
                ]
                == 0,
                "Applied import created multiple active analyses per workbook.",
            )
        public_actions = [
            {
                key: value
                for key, value in action.items()
                if key != "manifest"
            }
            for action in plan["actions"]
        ]
        result = {
            "schemaVersion": IMPORT_SCHEMA_VERSION,
            "engineVersion": IMPORT_ENGINE_VERSION,
            "generatedAt": _now(),
            "status": "APPLIED" if apply else "DRY_RUN_VALIDATED",
            "databasePath": str(database),
            "batchRoot": str(batch),
            "aiUsage": {
                "aiCallCount": 0,
                "fileLevelAiCallCount": 0,
                "retryCount": 0,
                "numericAuthority": (
                    "CODE_FROM_CAPTURED_RAW_VALUES"
                ),
            },
            "summary": {
                **plan["summary"],
                "sourceAndCaptureRevisionCount": len(capture_ids),
                "tableRequestBuiltCount": int(
                    (report.get("summary") or {}).get(
                        "requestBuiltCount"
                    )
                    or 0
                ),
                "tableRequestFailedCount": len(_failures),
                "aiCallCount": 0,
                "databaseCoverageBefore": before,
                "databaseCoverageAfter": after,
                "importerCoverageBefore": importer_before,
                "importerCoverageAfter": importer_after,
                "importedWorkbookCount": import_result[
                    "importedWorkbookCount"
                ],
            },
            "actions": public_actions,
            "imports": import_result["imported"],
        }
        _write_json(output, result)
        return result
    finally:
        connection.close()


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Import structure-replay observations and terminal coverage into "
            "the canonical evidence database."
        )
    )
    parser.add_argument("--db", required=True)
    parser.add_argument("--batch-root", required=True)
    parser.add_argument("--artifact-root", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--apply", action="store_true")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    result = run_structure_canonical_import(
        database_path=arguments.db,
        batch_root=arguments.batch_root,
        artifact_root=arguments.artifact_root,
        output_path=arguments.output,
        apply=bool(arguments.apply),
    )
    print(json.dumps(result["summary"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
