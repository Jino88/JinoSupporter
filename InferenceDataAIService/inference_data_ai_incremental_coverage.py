"""Build a fail-closed coverage report for an incremental structure run."""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_table_recipe_proposal import _read_json, _write_json


COVERAGE_SCHEMA_VERSION = "excel-incremental-structure-coverage-v1"
COVERAGE_ENGINE_VERSION = "incremental-structure-coverage-v1.0"


class IncrementalCoverageError(RuntimeError):
    """Raised when source artifacts do not form a closed accounting set."""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise IncrementalCoverageError(message)


def _outcome_breakdown(
    outcomes: dict[str, dict[str, Any]],
) -> dict[str, dict[str, int]]:
    grouped: dict[str, dict[str, int]] = defaultdict(
        lambda: {
            "structureCount": 0,
            "tableCount": 0,
            "workbookReferences": 0,
        }
    )
    for outcome in outcomes.values():
        status = str(outcome.get("status") or "")
        grouped[status]["structureCount"] += 1
        grouped[status]["tableCount"] += int(
            outcome.get("tableCount") or 0
        )
        grouped[status]["workbookReferences"] += int(
            outcome.get("workbookCount") or 0
        )
    return dict(sorted(grouped.items()))


def build_incremental_coverage_report(
    *,
    table_match_report: dict[str, Any],
    table_structure_catalog: dict[str, Any],
    priority_report: dict[str, Any],
    completion_state: dict[str, Any],
    recipe_registry: dict[str, Any],
    telemetry_values: Iterable[dict[str, Any]],
    generated_at: str | None = None,
) -> dict[str, Any]:
    match_summary = table_match_report.get("summary") or {}
    catalog_summary = table_structure_catalog.get("summary") or {}
    priority_summary = priority_report.get("summary") or {}
    completion_summary = completion_state.get("summary") or {}
    registry_summary = recipe_registry.get("summary") or {}

    eligible_workbooks = int(
        match_summary.get("eligibleWorkbookCount") or 0
    )
    request_completed_workbooks = int(
        match_summary.get("completedWorkbookCount") or 0
    )
    request_failed_workbooks = int(
        match_summary.get("failedWorkbookCount") or 0
    )
    request_failures = list(table_match_report.get("failures") or [])
    cataloged_tables = int(match_summary.get("tableCount") or 0)

    _require(
        eligible_workbooks
        == request_completed_workbooks + request_failed_workbooks,
        "Table-request workbook accounting does not close.",
    )
    _require(
        request_failed_workbooks == len(request_failures),
        "Table-request failure detail count does not match the summary.",
    )
    _require(
        cataloged_tables == int(catalog_summary.get("tableCount") or 0),
        "Table-match and structure-catalog totals differ.",
    )

    structures = list(table_structure_catalog.get("structures") or [])
    quantitative_tables = sum(
        int(item.get("tableCount") or 0)
        for item in structures
        if bool(item.get("quantitative"))
    )
    non_quantitative_tables = cataloged_tables - quantitative_tables
    reusable_quantitative_tables = sum(
        int(item.get("tableCount") or 0)
        for item in structures
        if bool(item.get("quantitative"))
        and int(item.get("workbookCount") or 0) >= 2
    )
    non_reusable_quantitative_tables = (
        quantitative_tables - reusable_quantitative_tables
    )
    _require(
        reusable_quantitative_tables
        == int(
            catalog_summary.get(
                "quantitativeTablesInReusableStructures"
            )
            or 0
        ),
        "Reusable quantitative table count is inconsistent.",
    )

    queue = list(priority_report.get("queue") or [])
    queue_ids = [str(item.get("tableStructureId") or "") for item in queue]
    _require(
        len(queue_ids) == len(set(queue_ids)) and all(queue_ids),
        "Priority queue structure IDs must be unique and non-empty.",
    )
    queued_tables = sum(int(item.get("tableCount") or 0) for item in queue)
    queue_workbook_references = sum(
        int(item.get("workbookCount") or 0) for item in queue
    )
    _require(
        queued_tables
        == int(priority_summary.get("coveredTableCount") or 0),
        "Priority queue table total does not match its summary.",
    )
    _require(
        queue_workbook_references
        == int(priority_summary.get("coveredWorkbookReferences") or 0),
        "Priority queue workbook-reference total does not match.",
    )

    outcomes = dict(completion_state.get("outcomes") or {})
    _require(
        completion_state.get("status") == "COMPLETED",
        "Completion state is not COMPLETED.",
    )
    _require(
        set(outcomes) == set(queue_ids),
        "Completion outcomes do not exactly cover the priority queue.",
    )
    _require(
        int(completion_summary.get("unresolvedStructureCount") or 0) == 0,
        "Completion summary still contains unresolved structures.",
    )
    breakdown = _outcome_breakdown(outcomes)

    unknown_statuses = [
        status
        for status in breakdown
        if not (
            status.startswith("REGISTERED")
            or status.startswith("QUARANTINED")
        )
    ]
    _require(
        not unknown_statuses,
        "Unsupported terminal outcome statuses: "
        + ", ".join(unknown_statuses),
    )
    registered_tables = sum(
        value["tableCount"]
        for status, value in breakdown.items()
        if status.startswith("REGISTERED")
    )
    quarantined_tables = sum(
        value["tableCount"]
        for status, value in breakdown.items()
        if status.startswith("QUARANTINED")
    )
    registered_structures = sum(
        value["structureCount"]
        for status, value in breakdown.items()
        if status.startswith("REGISTERED")
    )
    quarantined_structures = sum(
        value["structureCount"]
        for status, value in breakdown.items()
        if status.startswith("QUARANTINED")
    )
    _require(
        registered_tables + quarantined_tables == queued_tables,
        "Registered and quarantined table counts do not close the queue.",
    )
    _require(
        registered_structures + quarantined_structures == len(queue),
        "Registered and quarantined structures do not close the queue.",
    )
    _require(
        registered_tables
        == int(registry_summary.get("registeredTableCount") or 0),
        "Registered table total differs from the recipe registry.",
    )
    _require(
        registered_structures
        == int(registry_summary.get("registeredRecipeCount") or 0),
        "Registered structure total differs from the recipe registry.",
    )

    telemetry = list(telemetry_values)
    telemetry_ids = [
        str(value.get("tableStructureId") or "") for value in telemetry
    ]
    _require(
        len(telemetry_ids) == len(set(telemetry_ids))
        and all(telemetry_ids),
        "AI telemetry must have one non-empty record per structure.",
    )
    _require(
        all(
            int(value.get("aiCallBudget") or 0) <= 1
            and int(value.get("aiCallsAttempted") or 0) <= 1
            and int(value.get("retryCount") or 0) == 0
            for value in telemetry
        ),
        "AI telemetry violates the one-call/no-retry policy.",
    )
    ai_attempted = sum(
        int(value.get("aiCallsAttempted") or 0) for value in telemetry
    )
    ai_succeeded = sum(
        int(value.get("aiCallsSucceeded") or 0) for value in telemetry
    )
    ai_failed = sum(value.get("status") == "FAILED" for value in telemetry)
    retry_count = sum(
        int(value.get("retryCount") or 0) for value in telemetry
    )
    prompt_bytes = sum(
        int(value.get("promptBytes") or 0) for value in telemetry
    )
    output_bytes = sum(
        int(value.get("outputBytes") or 0) for value in telemetry
    )
    duration_ms = sum(
        int(value.get("durationMs") or 0) for value in telemetry
    )
    _require(
        ai_attempted
        == int(completion_summary.get("aiCallsAttempted") or 0),
        "AI attempted count differs from completion state.",
    )
    _require(
        ai_succeeded
        == int(completion_summary.get("aiCallsSucceeded") or 0),
        "AI success count differs from completion state.",
    )
    _require(
        ai_failed == int(completion_summary.get("aiCallsFailed") or 0),
        "AI failure count differs from completion state.",
    )
    _require(
        retry_count == 0
        and int(completion_summary.get("retryCount") or 0) == 0,
        "AI retry count must remain zero.",
    )
    _require(
        int(completion_summary.get("fileLevelAiCalls") or 0) == 0,
        "File-level AI calls must remain disabled.",
    )

    outside_tables = cataloged_tables - queued_tables
    outside_quantitative_tables = quantitative_tables - queued_tables
    outside_reusable_quantitative_tables = (
        reusable_quantitative_tables - queued_tables
    )
    _require(
        min(
            outside_tables,
            outside_quantitative_tables,
            outside_reusable_quantitative_tables,
        )
        >= 0,
        "Priority queue is not a subset of the catalog totals.",
    )
    _require(
        outside_tables
        == outside_quantitative_tables + non_quantitative_tables,
        "Outside-queue table accounting does not close.",
    )
    _require(
        outside_quantitative_tables
        == outside_reusable_quantitative_tables
        + non_reusable_quantitative_tables,
        "Outside-queue quantitative accounting does not close.",
    )

    avoided_table_level_calls = queued_tables - ai_attempted
    avoided_structure_level_calls = len(queue) - ai_attempted
    target_calls = eligible_workbooks * 0.1
    return {
        "schemaVersion": COVERAGE_SCHEMA_VERSION,
        "engineVersion": COVERAGE_ENGINE_VERSION,
        "generatedAt": generated_at or _now(),
        "status": (
            "ACCOUNTED_WITH_EXPLICIT_TABLE_REQUEST_FAILURES"
            if request_failed_workbooks
            else "ACCOUNTED"
        ),
        "workbookCoverage": {
            "eligible": eligible_workbooks,
            "sourceAndCapturePresent": eligible_workbooks,
            "tableRequestCompleted": request_completed_workbooks,
            "tableRequestFailed": request_failed_workbooks,
            "tableRequestFailures": request_failures,
        },
        "tableCoverage": {
            "catalogedTableCount": cataloged_tables,
            "quantitativeCandidateTableCount": quantitative_tables,
            "nonQuantitativeTableCount": non_quantitative_tables,
            "repeatedQuantitativeQueueTableCount": queued_tables,
            "registeredProgramExtractionTableCount": registered_tables,
            "quarantinedTableCount": quarantined_tables,
            "outsideRepeatedQueueTableCount": outside_tables,
            "outsideRepeatedQueue": {
                "status": "ACCOUNTED_NOT_PARAMETER_EXTRACTED_IN_THIS_PASS",
                "quantitativeTableCount": outside_quantitative_tables,
                "reusableQuantitativeTableCount": (
                    outside_reusable_quantitative_tables
                ),
                "nonReusableQuantitativeTableCount": (
                    non_reusable_quantitative_tables
                ),
                "nonQuantitativeTableCount": non_quantitative_tables,
            },
        },
        "queueCoverage": {
            "structureCount": len(queue),
            "workbookReferences": queue_workbook_references,
            "registeredStructureCount": registered_structures,
            "quarantinedStructureCount": quarantined_structures,
            "unresolvedStructureCount": 0,
            "outcomes": breakdown,
        },
        "aiUsage": {
            "attempted": ai_attempted,
            "succeeded": ai_succeeded,
            "failed": ai_failed,
            "retryCount": retry_count,
            "fileLevelAiCalls": 0,
            "promptBytes": prompt_bytes,
            "outputBytes": output_bytes,
            "serialDurationMs": duration_ms,
            "serialDurationMinutes": round(duration_ms / 60_000, 3),
            "tokenUsage": "UNAVAILABLE_IN_TELEMETRY",
            "averageCallsPerEligibleWorkbook": round(
                ai_attempted / eligible_workbooks, 6
            )
            if eligible_workbooks
            else 0.0,
            "targetAtMostPointOneCallPerWorkbook": {
                "maximumCalls": target_calls,
                "met": ai_attempted <= target_calls,
            },
        },
        "aiAvoidance": {
            "versusOneCallPerQueuedTable": {
                "baselineCalls": queued_tables,
                "avoidedCalls": avoided_table_level_calls,
                "reductionPercent": round(
                    avoided_table_level_calls * 100 / queued_tables,
                    3,
                )
                if queued_tables
                else 0.0,
            },
            "versusOneCallPerQueuedStructure": {
                "baselineCalls": len(queue),
                "avoidedCalls": avoided_structure_level_calls,
                "reductionPercent": round(
                    avoided_structure_level_calls * 100 / len(queue),
                    3,
                )
                if queue
                else 0.0,
            },
        },
        "policy": {
            "aiScope": "STRUCTURE_ONLY",
            "maxAiCallsPerStructure": 1,
            "retryCount": 0,
            "numericAndEvidenceExtraction": "PROGRAM_FROM_CAPTURE",
            "failedValidation": "FAIL_CLOSED_QUARANTINE",
        },
        "invariants": {
            "eligibleWorkbooksAccounted": True,
            "catalogedTablesAccounted": True,
            "queueStructuresResolved": True,
            "queueTablesRegisteredOrQuarantined": True,
            "registryTotalsMatch": True,
            "oneAiCallMaximumPerStructure": True,
            "retryCountZero": True,
            "fileLevelAiCallsZero": True,
        },
    }


def build_incremental_coverage_from_paths(
    *,
    table_match_report_path: str | Path,
    table_structure_catalog_path: str | Path,
    priority_report_path: str | Path,
    completion_state_path: str | Path,
    recipe_registry_path: str | Path,
    telemetry_root: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    telemetry_path = Path(telemetry_root).expanduser().resolve()
    report = build_incremental_coverage_report(
        table_match_report=_read_json(
            Path(table_match_report_path).expanduser().resolve()
        ),
        table_structure_catalog=_read_json(
            Path(table_structure_catalog_path).expanduser().resolve()
        ),
        priority_report=_read_json(
            Path(priority_report_path).expanduser().resolve()
        ),
        completion_state=_read_json(
            Path(completion_state_path).expanduser().resolve()
        ),
        recipe_registry=_read_json(
            Path(recipe_registry_path).expanduser().resolve()
        ),
        telemetry_values=[
            _read_json(path)
            for path in sorted(telemetry_path.glob("*.json"))
        ],
    )
    _write_json(Path(output_path).expanduser().resolve(), report)
    return report


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Close and audit incremental structure coverage."
    )
    parser.add_argument("--table-match-report", required=True)
    parser.add_argument("--table-structure-catalog", required=True)
    parser.add_argument("--priority-report", required=True)
    parser.add_argument("--completion-state", required=True)
    parser.add_argument("--recipe-registry", required=True)
    parser.add_argument("--telemetry-root", required=True)
    parser.add_argument("--output", required=True)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    report = build_incremental_coverage_from_paths(
        table_match_report_path=arguments.table_match_report,
        table_structure_catalog_path=arguments.table_structure_catalog,
        priority_report_path=arguments.priority_report,
        completion_state_path=arguments.completion_state,
        recipe_registry_path=arguments.recipe_registry,
        telemetry_root=arguments.telemetry_root,
        output_path=arguments.output,
    )
    print(
        json.dumps(
            {
                "status": report["status"],
                "workbookCoverage": report["workbookCoverage"],
                "tableCoverage": report["tableCoverage"],
                "queueCoverage": {
                    key: value
                    for key, value in report["queueCoverage"].items()
                    if key != "outcomes"
                },
                "aiUsage": report["aiUsage"],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
