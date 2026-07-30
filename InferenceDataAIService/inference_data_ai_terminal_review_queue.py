"""Build a resumable human-review queue for canonical terminal workbooks.

The queue is entirely deterministic. It joins canonical import actions with
table-match, structure-catalog, priority, completion, and optional source-owned
decision artifacts without reading or interpreting numeric cell values.
"""

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
    _semantic_header_sha256,
    semantic_header_signature,
)


TERMINAL_REVIEW_QUEUE_SCHEMA_VERSION = "excel-terminal-review-queue-v1"
TERMINAL_REVIEW_QUEUE_ENGINE_VERSION = "terminal-review-queue-v1.0"
TERMINAL_ACTION = "IMPORT_NEEDS_REVIEW_TERMINAL"

REPEATED_REVIEW = "REPEATED_QUANTITATIVE_CONTRACT_REVIEW"
ONE_OFF_REVIEW = "ONE_OFF_OR_NON_REUSABLE_QUANTITATIVE_REVIEW"
NON_QUANTITATIVE_REVIEW = "NON_QUANTITATIVE_OR_UNSTRUCTURED_REVIEW"

REPEATED_TABLE = "REPEATED_QUANTITATIVE_QUARANTINED"
ONE_OFF_TABLE = "ONE_OFF_OR_NON_REUSABLE_QUANTITATIVE"
NON_QUANTITATIVE_TABLE = "NON_QUANTITATIVE_OR_SUPPORTING"
REGISTERED_CONFLICT = "REGISTERED_RECIPE_PRESENT_CONFLICT"
PENDING_CONFLICT = "REPEATED_QUANTITATIVE_PENDING"


class TerminalReviewQueueError(RuntimeError):
    """Raised when terminal artifacts cannot be joined safely."""


def _read_json(path: str | Path) -> dict[str, Any]:
    source = Path(path).expanduser().resolve()
    with source.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise TerminalReviewQueueError(
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


def _unique_index(
    values: Iterable[dict[str, Any]],
    *,
    key,
    label: str,
) -> dict[Any, dict[str, Any]]:
    result: dict[Any, dict[str, Any]] = {}
    for value in values:
        identity = key(value)
        if identity in result:
            raise TerminalReviewQueueError(
                f"Duplicate {label}: {identity}"
            )
        result[identity] = value
    return result


def _model_family(file_name: str) -> str:
    normalized = file_name.upper()
    patterns = (
        (
            "X526B",
            r"X526(?:B(?:[\s_-]*(?:TOP|BOTTOM))?|"
            r"[\s_-]*BOTTOM|BT)\b",
        ),
        (
            "X626B",
            r"X626(?:B(?:[\s_-]*(?:TOP|BOTTOM))?|"
            r"[\s_-]*BOTTOM|BT)\b",
        ),
        (
            "S931B",
            r"S931(?:B(?:[\s_-]*(?:TOP|BOTTOM))?|"
            r"[\s_-]*BOTTOM|BT)\b",
        ),
        ("X526", r"X526(?:[\s_-]*(?:TOP|BOTTOM))?\b"),
        ("X626", r"X626(?:[\s_-]*(?:TOP|BOTTOM))?\b"),
        ("S931", r"S931(?:[\s_-]*(?:TOP|BOTTOM))?\b"),
    )
    for family, pattern in patterns:
        if re.search(pattern, normalized):
            return family
    if re.search(r"\bTIU\b", normalized):
        return "TIU"
    return "OTHER"


def _file_review_class(table_counts: Counter[str]) -> str:
    if table_counts[REGISTERED_CONFLICT]:
        return REGISTERED_CONFLICT
    if table_counts[PENDING_CONFLICT]:
        return PENDING_CONFLICT
    if table_counts[REPEATED_TABLE]:
        return REPEATED_REVIEW
    if table_counts[ONE_OFF_TABLE]:
        return ONE_OFF_REVIEW
    if table_counts[NON_QUANTITATIVE_TABLE]:
        return NON_QUANTITATIVE_REVIEW
    return "UNCLASSIFIED_REVIEW"


def _review_priority(
    review_class: str,
    outcome_statuses: Iterable[str],
) -> int:
    if review_class == ONE_OFF_REVIEW:
        return 6
    if review_class == NON_QUANTITATIVE_REVIEW:
        return 7
    statuses = set(outcome_statuses)
    if "QUARANTINED_RECIPE_CONTRACT_FAILURE" in statuses:
        return 1
    if "QUARANTINED_AI_FAILURE_NO_RETRY" in statuses:
        return 2
    if "QUARANTINED_PRECHECK" in statuses:
        return 3
    if "QUARANTINED_SOURCE_PRECHECK" in statuses:
        return 4
    if "QUARANTINED_AI_DECISION" in statuses:
        return 5
    return 0


def _recommended_action(
    review_class: str,
    outcome_statuses: Iterable[str],
) -> str:
    if review_class == ONE_OFF_REVIEW:
        return (
            "Review the quantitative table as a one-off item; create a reusable "
            "recipe only after recurrence and stable source semantics are shown."
        )
    if review_class == NON_QUANTITATIVE_REVIEW:
        return (
            "Review as non-quantitative or supporting material; do not create "
            "numeric claims from this table."
        )
    statuses = set(outcome_statuses)
    if "QUARANTINED_RECIPE_CONTRACT_FAILURE" in statuses:
        return (
            "Inspect the deterministic compiler/selector contract and replay "
            "the affected repeated structure without a file-level AI call."
        )
    if "QUARANTINED_AI_FAILURE_NO_RETRY" in statuses:
        return (
            "Inspect preserved structure telemetry and source labels; require "
            "an explicit recovery decision before any new bounded AI call."
        )
    if "QUARANTINED_PRECHECK" in statuses:
        return (
            "Resolve the structural precheck failure before recipe proposal."
        )
    if "QUARANTINED_SOURCE_PRECHECK" in statuses:
        return (
            "Review generic or ambiguous source headers and keep the structure "
            "terminal unless a source-owned metric meaning is explicit."
        )
    if "QUARANTINED_AI_DECISION" in statuses:
        return (
            "Review the explicit quarantine rationale; retain terminal status "
            "unless a human-approved semantic contract is supplied."
        )
    return (
        "Review as non-quantitative or supporting material; do not create "
        "numeric claims from this table."
    )


def _source_owned_non_metric(
    decision: dict[str, Any] | None,
) -> bool:
    if not decision or str(decision.get("decision") or "") != "QUARANTINE":
        return False
    contract = decision.get("semanticContract") or {}
    if not isinstance(contract, dict):
        return False
    metrics = contract.get("metricColumns")
    return isinstance(metrics, list) and not metrics


def _date_only_metadata_table(captured_table: dict[str, Any]) -> bool:
    numeric_columns = captured_table.get("numericColumns") or []
    if not numeric_columns:
        return False
    formats = [
        str(number_format or "").casefold()
        for column in numeric_columns
        for number_format in column.get("numberFormats") or []
    ]
    if not formats or not all(
        re.search(r"d{1,4}", number_format)
        and re.search(r"m{1,4}", number_format)
        for number_format in formats
    ):
        return False
    label_values = [
        str(header)
        for column in numeric_columns
        for header in column.get("headerTexts") or []
    ]
    label_values.extend(
        str(label)
        for row in captured_table.get("rowLabels") or []
        for label in row.get("labels") or []
    )
    label_values.extend(
        str(title)
        for title in captured_table.get("titleCandidates") or []
    )
    label_tokens = {
        token
        for value in label_values
        for token in re.findall(r"[A-Z]+", value.upper())
    }
    administrative_tokens = {
        "APPROVAL",
        "CHECKER",
        "DATE",
        "FINISH",
        "MAKER",
        "NAME",
        "PERIOD",
        "REQUEST",
        "SIGNATURE",
    }
    return len(label_tokens & administrative_tokens) >= 2


def build_terminal_review_queue(
    *,
    canonical_audit: dict[str, Any],
    table_match_report: dict[str, Any],
    table_structure_catalog: dict[str, Any],
    priority_report: dict[str, Any],
    completion_state: dict[str, Any],
    source_owned_decisions: Iterable[dict[str, Any]] = (),
    generated_at: str | None = None,
) -> dict[str, Any]:
    """Join terminal workbook actions to deterministic review categories."""

    terminal_actions = [
        action
        for action in canonical_audit.get("actions") or []
        if str(action.get("action") or "") == TERMINAL_ACTION
    ]
    actions_by_revision = _unique_index(
        terminal_actions,
        key=lambda value: int(
            (value.get("source") or {}).get("captureRevisionId") or 0
        ),
        label="terminal capture revision",
    )
    workbooks_by_revision = _unique_index(
        table_match_report.get("workbooks") or [],
        key=lambda value: int(value.get("captureRevisionId") or 0),
        label="table-match capture revision",
    )
    structures_by_fingerprint = _unique_index(
        table_structure_catalog.get("structures") or [],
        key=lambda value: str(value.get("fingerprintSha256") or ""),
        label="table structure fingerprint",
    )
    priority_by_variant = _unique_index(
        priority_report.get("queue") or [],
        key=lambda value: (
            str(value.get("baseTableStructureId") or ""),
            str(value.get("semanticHeaderSha256") or ""),
        ),
        label="priority semantic variant",
    )
    outcomes = completion_state.get("outcomes") or {}
    if not isinstance(outcomes, dict):
        raise TerminalReviewQueueError("Completion outcomes must be an object.")
    source_owned_by_structure = _unique_index(
        source_owned_decisions,
        key=lambda value: str(
            value.get("targetTableStructureId") or ""
        ),
        label="source-owned target structure",
    )

    items: list[dict[str, Any]] = []
    repeated_groups: dict[str, dict[str, Any]] = {}
    table_class_counts: Counter[str] = Counter()
    file_class_counts: Counter[str] = Counter()
    outcome_table_counts: Counter[str] = Counter()
    model_family_counts: Counter[str] = Counter()
    classified_table_count = 0
    source_owned_non_metric_table_count = 0
    date_only_metadata_table_count = 0

    for revision_id, action in actions_by_revision.items():
        workbook = workbooks_by_revision.get(revision_id)
        if workbook is None:
            raise TerminalReviewQueueError(
                f"Terminal revision is absent from table-match report: "
                f"{revision_id}"
            )
        request_path = Path(str(workbook.get("requestPath") or "")).resolve()
        request = _read_json(request_path)
        request_tables = _unique_index(
            request.get("tables") or [],
            key=lambda value: str(value.get("tableId") or ""),
            label=f"request table for revision {revision_id}",
        )
        table_counts: Counter[str] = Counter()
        outcome_statuses: set[str] = set()
        table_items: list[dict[str, Any]] = []

        for table_match in workbook.get("tables") or []:
            table_id = str(table_match.get("tableId") or "")
            captured_table = request_tables.get(table_id)
            if captured_table is None:
                raise TerminalReviewQueueError(
                    f"Request table missing: revision={revision_id}, "
                    f"table={table_id}"
                )
            fingerprint = str(
                table_match.get("fingerprintSha256") or ""
            )
            structure = structures_by_fingerprint.get(fingerprint)
            if structure is None:
                raise TerminalReviewQueueError(
                    f"Catalog structure missing: {fingerprint}"
                )
            base_structure_id = str(
                structure.get("tableStructureId") or ""
            )
            semantic_sha256 = _semantic_header_sha256(
                semantic_header_signature(captured_table)
            )
            priority_item = priority_by_variant.get(
                (base_structure_id, semantic_sha256)
            )
            priority_structure_id: str | None = None
            outcome_status: str | None = None
            classification_basis: str
            date_only_metadata = _date_only_metadata_table(captured_table)
            if priority_item is not None:
                priority_structure_id = str(
                    priority_item.get("tableStructureId") or ""
                )
                outcome_status = str(
                    (outcomes.get(priority_structure_id) or {}).get(
                        "status"
                    )
                    or ""
                )
                outcome_statuses.add(outcome_status)
                outcome_table_counts[outcome_status] += 1
                if outcome_status.startswith("REGISTERED"):
                    table_class = REGISTERED_CONFLICT
                    classification_basis = "PRIORITY_REGISTERED_CONFLICT"
                elif _source_owned_non_metric(
                    source_owned_by_structure.get(priority_structure_id)
                ):
                    table_class = NON_QUANTITATIVE_TABLE
                    classification_basis = (
                        "SOURCE_OWNED_NON_METRIC_QUARANTINE"
                    )
                    source_owned_non_metric_table_count += 1
                elif date_only_metadata:
                    table_class = NON_QUANTITATIVE_TABLE
                    classification_basis = "DATE_ONLY_METADATA"
                    date_only_metadata_table_count += 1
                elif outcome_status.startswith("QUARANTINED"):
                    table_class = REPEATED_TABLE
                    classification_basis = "PRIORITY_QUARANTINE"
                else:
                    table_class = PENDING_CONFLICT
                    classification_basis = "PRIORITY_PENDING"
            elif date_only_metadata:
                table_class = NON_QUANTITATIVE_TABLE
                classification_basis = "DATE_ONLY_METADATA"
                date_only_metadata_table_count += 1
            elif bool(structure.get("quantitative")):
                table_class = ONE_OFF_TABLE
                classification_basis = "CATALOG_QUANTITATIVE_ONE_OFF"
            else:
                table_class = NON_QUANTITATIVE_TABLE
                classification_basis = "CATALOG_NON_QUANTITATIVE"

            table_counts[table_class] += 1
            table_class_counts[table_class] += 1
            classified_table_count += 1
            table_item = {
                "tableId": table_id,
                "sheet": str(table_match.get("sheet") or ""),
                "range": str(table_match.get("range") or ""),
                "classification": table_class,
                "classificationBasis": classification_basis,
                "numericCellCount": int(
                    table_match.get("numericCellCount") or 0
                ),
                "fingerprintSha256": fingerprint,
                "baseTableStructureId": base_structure_id,
                "semanticHeaderSha256": semantic_sha256,
                "priorityStructureId": priority_structure_id,
                "priorityOutcomeStatus": outcome_status,
            }
            table_items.append(table_item)

            if table_class == REPEATED_TABLE and priority_item is not None:
                group = repeated_groups.setdefault(
                    priority_structure_id or "",
                    {
                        "priorityStructureId": priority_structure_id,
                        "baseTableStructureId": base_structure_id,
                        "semanticHeaderSha256": semantic_sha256,
                        "outcomeStatus": outcome_status,
                        "queueRank": int(priority_item.get("rank") or 0),
                        "queuePriorityScore": float(
                            priority_item.get("priorityScore") or 0
                        ),
                        "safetyReasons": list(
                            priority_item.get("safetyReasons") or []
                        ),
                        "members": [],
                    },
                )
                group["members"].append(
                    {
                        "captureRevisionId": revision_id,
                        "fileName": str(workbook.get("fileName") or ""),
                        "tableId": table_id,
                        "sheet": str(table_match.get("sheet") or ""),
                        "range": str(table_match.get("range") or ""),
                    }
                )

        review_class = _file_review_class(table_counts)
        file_class_counts[review_class] += 1
        model_family = _model_family(str(workbook.get("fileName") or ""))
        model_family_counts[model_family] += 1
        review_priority = _review_priority(
            review_class,
            outcome_statuses,
        )
        source = action.get("source") or {}
        items.append(
            {
                "reviewRank": 0,
                "reviewPriority": review_priority,
                "reviewClass": review_class,
                "recommendedAction": _recommended_action(
                    review_class,
                    outcome_statuses,
                ),
                "modelFamily": model_family,
                "captureRevisionId": revision_id,
                "revisionUid": str(source.get("revisionUid") or ""),
                "contentSha256": str(source.get("contentSha256") or ""),
                "fileName": str(source.get("fileName") or ""),
                "sourcePath": str(source.get("sourcePath") or ""),
                "terminalReason": str(action.get("reason") or ""),
                "tableCount": len(table_items),
                "tableClassCounts": dict(sorted(table_counts.items())),
                "priorityOutcomeStatuses": sorted(outcome_statuses),
                "tables": table_items,
            }
        )

    items.sort(
        key=lambda value: (
            int(value["reviewPriority"]),
            str(value["modelFamily"]),
            str(value["fileName"]).casefold(),
            int(value["captureRevisionId"]),
        )
    )
    for rank, item in enumerate(items, start=1):
        item["reviewRank"] = rank

    repeated_structure_groups = []
    for group in repeated_groups.values():
        members = group["members"]
        group["tableCount"] = len(members)
        group["workbookCount"] = len(
            {int(member["captureRevisionId"]) for member in members}
        )
        repeated_structure_groups.append(group)
    repeated_structure_groups.sort(
        key=lambda value: (
            _review_priority(
                REPEATED_REVIEW,
                [str(value.get("outcomeStatus") or "")],
            ),
            -int(value["workbookCount"]),
            -int(value["tableCount"]),
            str(value["priorityStructureId"]),
        )
    )

    expected_terminal_count = int(
        (
            (canonical_audit.get("summary") or {}).get("actionCounts")
            or {}
        ).get(TERMINAL_ACTION)
        or 0
    )
    matched_table_count = sum(
        int(workbook.get("tableCount") or 0)
        for revision_id, workbook in workbooks_by_revision.items()
        if revision_id in actions_by_revision
    )
    invariants = {
        "terminalCountMatchesCanonicalSummary": (
            len(items) == expected_terminal_count
        ),
        "terminalCaptureRevisionsUnique": (
            len(actions_by_revision) == len(terminal_actions)
        ),
        "allTerminalWorkbooksMapped": (
            len(items) == len(terminal_actions)
        ),
        "allTerminalTablesClassified": (
            classified_table_count == matched_table_count
        ),
        "noRegisteredRecipeConflict": (
            table_class_counts[REGISTERED_CONFLICT] == 0
        ),
        "noPendingRepeatedStructure": (
            table_class_counts[PENDING_CONFLICT] == 0
        ),
        "completionQueueResolved": (
            str(completion_state.get("status") or "") == "COMPLETED"
            and int(
                (completion_state.get("summary") or {}).get(
                    "unresolvedStructureCount"
                )
                or 0
            )
            == 0
        ),
    }
    return {
        "schemaVersion": TERMINAL_REVIEW_QUEUE_SCHEMA_VERSION,
        "engineVersion": TERMINAL_REVIEW_QUEUE_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "status": (
            "READY_FOR_HUMAN_REVIEW"
            if all(invariants.values())
            else "INCOMPLETE"
        ),
        "policy": {
            "aiCalls": 0,
            "numericValuesRead": False,
            "fileLevelAiEnabled": False,
            "registeredRecipeResultsExcluded": True,
            "sourceOwnedNonMetricQuarantinesAreSupporting": True,
            "dateOnlyAdministrativeTablesAreSupporting": True,
        },
        "summary": {
            "terminalWorkbookCount": len(items),
            "terminalTableCount": classified_table_count,
            "reviewClassCounts": dict(sorted(file_class_counts.items())),
            "tableClassCounts": dict(sorted(table_class_counts.items())),
            "priorityOutcomeTableCounts": dict(
                sorted(outcome_table_counts.items())
            ),
            "modelFamilyCounts": dict(sorted(model_family_counts.items())),
            "repeatedStructureGroupCount": len(
                repeated_structure_groups
            ),
            "sourceOwnedNonMetricTableCount": (
                source_owned_non_metric_table_count
            ),
            "dateOnlyMetadataTableCount": date_only_metadata_table_count,
        },
        "invariants": invariants,
        "repeatedStructureGroups": repeated_structure_groups,
        "items": items,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Build a deterministic human-review queue for canonical terminal "
            "workbooks."
        )
    )
    parser.add_argument("--canonical-audit", required=True)
    parser.add_argument("--table-match-report", required=True)
    parser.add_argument("--table-structure-catalog", required=True)
    parser.add_argument("--priority-report", required=True)
    parser.add_argument("--completion-state", required=True)
    parser.add_argument("--source-owned-decision-dir")
    parser.add_argument("--output", required=True)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(argv)
    source_owned_decisions = []
    if arguments.source_owned_decision_dir:
        decision_dir = (
            Path(arguments.source_owned_decision_dir).expanduser().resolve()
        )
        if not decision_dir.is_dir():
            raise TerminalReviewQueueError(
                f"Source-owned decision directory not found: {decision_dir}"
            )
        source_owned_decisions = [
            _read_json(path)
            for path in sorted(decision_dir.glob("*.source-owned.json"))
        ]
    report = build_terminal_review_queue(
        canonical_audit=_read_json(arguments.canonical_audit),
        table_match_report=_read_json(arguments.table_match_report),
        table_structure_catalog=_read_json(
            arguments.table_structure_catalog
        ),
        priority_report=_read_json(arguments.priority_report),
        completion_state=_read_json(arguments.completion_state),
        source_owned_decisions=source_owned_decisions,
    )
    _write_json(arguments.output, report)
    print(json.dumps(report["summary"], ensure_ascii=False))
    return 0 if report["status"] == "READY_FOR_HUMAN_REVIEW" else 2


if __name__ == "__main__":
    raise SystemExit(main())
