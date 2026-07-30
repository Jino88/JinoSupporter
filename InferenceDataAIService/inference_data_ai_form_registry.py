"""Review-gated registry for Excel layout families.

The registry groups already captured workbooks by stable layout features,
analyzes one representative plus up to two validation samples, and requires a
human decision before a family can enter the full AI ingestion manifest.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
from collections import Counter
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence


REGISTRY_SCHEMA_VERSION = "excel-form-family-registry-v1"
GROUP_REVIEW_SCHEMA_VERSION = "excel-form-group-review-v1"
CONTRACT_SCHEMA_VERSION = "excel-form-family-contract-v1"
APPROVED_DECISIONS = {"APPROVED_NEW", "LINKED_EXISTING"}
DECISION_STATUSES = {
    "PENDING",
    "ANALYZED_PENDING_APPROVAL",
    "APPROVED_NEW",
    "LINKED_EXISTING",
    "EXCLUDED",
}
STRUCTURAL_TERMS = {
    "AVG",
    "BAKO",
    "CONTROL",
    "DATE",
    "DEFECT",
    "DIMENSION",
    "FAIL",
    "FUNCTION",
    "GROUP",
    "HEARING",
    "INPUT",
    "INSPECTION",
    "ITEM",
    "LOT",
    "MAX",
    "MEASUREMENT",
    "MIN",
    "MODEL",
    "NG",
    "NOISE",
    "NORMAL",
    "OK",
    "PASS",
    "POSITION",
    "RATE",
    "RESULT",
    "SAMPLE",
    "SIGMA",
    "SPEC",
    "SPL",
    "STANDARD",
    "TEST",
    "THD",
    "TENSION",
    "TOTAL",
    "TOUCH",
    "VALUE",
    "VISION",
    "VISUAL",
}
TOKEN_PATTERN = re.compile(r"[A-Z][A-Z0-9+/%._-]{1,}")


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def ensure_form_registry_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS excel_form_families (
            family_id TEXT PRIMARY KEY,
            schema_version TEXT NOT NULL,
            display_name TEXT NOT NULL DEFAULT '',
            decision_status TEXT NOT NULL
                CHECK(decision_status IN (
                    'PENDING',
                    'ANALYZED_PENDING_APPROVAL',
                    'APPROVED_NEW',
                    'LINKED_EXISTING',
                    'EXCLUDED'
                )),
            family_signature_json TEXT NOT NULL,
            representative_source_path TEXT NOT NULL DEFAULT '',
            representative_revision_id INTEGER NOT NULL DEFAULT 0,
            linked_form_signature_id TEXT NOT NULL DEFAULT '',
            extraction_contract_json TEXT,
            validation_status TEXT NOT NULL DEFAULT '',
            validation_sample_count INTEGER NOT NULL DEFAULT 0,
            member_count INTEGER NOT NULL DEFAULT 0,
            reviewer TEXT NOT NULL DEFAULT '',
            notes TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_excel_form_families_status
            ON excel_form_families(decision_status, family_id);
        """
    )


def _bucket(value: int, limits: Sequence[int]) -> str:
    for limit in limits:
        if value <= limit:
            return f"<=${limit}".replace("$", "")
    return f">{limits[-1]}"


def _structural_tokens(values: Iterable[str]) -> list[str]:
    tokens: set[str] = set()
    for value in values:
        for match in TOKEN_PATTERN.finditer(str(value).upper()):
            token = match.group(0).strip("._-")
            if token in STRUCTURAL_TERMS:
                tokens.add(token)
    return sorted(tokens)


def family_descriptor(signature: dict[str, Any]) -> dict[str, Any]:
    """Build a row-growth-tolerant family identity from a strict signature."""

    sheets: list[dict[str, Any]] = []
    for profile in signature.get("sheetProfiles") or []:
        columns = int(profile.get("columns") or 0)
        merges = str(profile.get("mergeBucket") or "0")
        formulas = str(profile.get("formulaBucket") or "0")
        sheets.append(
            {
                "tabular": bool(profile.get("tabular")),
                "columnBand": _bucket(columns, (3, 6, 10, 16, 24, 40)),
                "mergeBand": merges,
                "hasFormula": formulas != "0",
                "structuralTokens": _structural_tokens(
                    profile.get("tokens") or []
                ),
            }
        )
    canonical_value = {
        "workbookStatus": str(
            signature.get("workbookStatus") or ""
        ),
        "sheetCount": int(signature.get("sheetCount") or len(sheets)),
        "tabularSheetCount": int(
            signature.get("tabularSheetCount") or 0
        ),
        "sheets": sheets,
    }
    canonical = json.dumps(
        canonical_value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return {
        **canonical_value,
        "familyId": "family-"
        + hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:16],
    }


def _optional_json_object(value: Any) -> dict[str, Any] | None:
    if value is None or str(value).strip() == "":
        return None
    parsed = json.loads(str(value))
    if not isinstance(parsed, dict):
        raise ValueError("Stored extraction contract must be a JSON object.")
    return parsed


def load_family_decisions(
    connection: sqlite3.Connection,
) -> dict[str, dict[str, Any]]:
    ensure_form_registry_schema(connection)
    rows = connection.execute(
        """
        SELECT *
        FROM excel_form_families
        ORDER BY family_id
        """
    ).fetchall()
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        contract = _optional_json_object(
            row["extraction_contract_json"]
        )
        result[str(row["family_id"])] = {
            "familyId": str(row["family_id"]),
            "displayName": str(row["display_name"]),
            "decisionStatus": str(row["decision_status"]),
            "linkedFormSignatureId": str(
                row["linked_form_signature_id"]
            ),
            "contractAvailable": bool(
                row["extraction_contract_json"]
            ),
            "extractionContract": contract,
            "recommendation": str(
                (contract or {}).get("recommendation") or ""
            ),
            "confidence": float(
                (contract or {}).get("confidence") or 0
            ),
            "validationStatus": str(row["validation_status"]),
            "validationSampleCount": int(
                row["validation_sample_count"]
            ),
            "reviewer": str(row["reviewer"]),
            "notes": str(row["notes"]),
        }
    return result


def apply_registry_decision(
    classification: dict[str, Any],
    *,
    family: dict[str, Any],
    decisions: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    decision = decisions.get(str(family["familyId"]))
    if decision is None:
        return classification
    status = decision["decisionStatus"]
    registry_fields = {
        "formFamilyId": str(family["familyId"]),
        "familyDisplayName": str(decision["displayName"]),
        "registryDecision": status,
    }
    if status in APPROVED_DECISIONS:
        return {
            **classification,
            **registry_fields,
            "status": "KNOWN_FORM",
            "nearestKnownFormSignatureId": (
                decision["linkedFormSignatureId"]
                or classification.get(
                    "nearestKnownFormSignatureId",
                    "",
                )
            ),
            "reason": (
                "사람이 승인한 신규 양식군입니다."
                if status == "APPROVED_NEW"
                else "사람이 기존 양식에 연결한 양식군입니다."
            ),
            **(
                {
                    "extractionContract": decision[
                        "extractionContract"
                    ]
                }
                if status == "APPROVED_NEW"
                and decision["extractionContract"] is not None
                else {}
            ),
        }
    if status == "EXCLUDED":
        return {
            **classification,
            **registry_fields,
            "status": "EXCLUDED_FORM",
            "reason": "사람이 전체 처리 제외로 판정한 양식군입니다.",
        }
    return {
        **classification,
        **registry_fields,
    }


def _load_json(path: str | Path) -> dict[str, Any]:
    value = json.loads(
        Path(path).expanduser().resolve().read_text(encoding="utf-8")
    )
    if not isinstance(value, dict):
        raise ValueError("Form preflight report must be a JSON object.")
    return value


def build_form_group_review(
    *,
    connection: sqlite3.Connection,
    report: dict[str, Any],
) -> dict[str, Any]:
    from inference_data_ai_form_preflight import (
        signature_from_database,
    )

    ensure_form_registry_schema(connection)
    decisions = load_family_decisions(connection)
    grouped: dict[str, list[dict[str, Any]]] = {}
    descriptor_by_id: dict[str, dict[str, Any]] = {}
    signature_by_id: dict[str, dict[str, Any]] = {}
    for item in report.get("items") or []:
        if not isinstance(item, dict):
            continue
        item_status = str(item.get("status") or "")
        if (
            item_status
            not in {
                "SIMILAR_FORM_REVIEW",
                "NEW_FORM",
                "EXCLUDED_FORM",
            }
            and not str(item.get("registryDecision") or "")
        ):
            continue
        revision_id = int(item.get("captureRevisionId") or 0)
        if revision_id <= 0:
            continue
        signature = signature_from_database(connection, revision_id)
        family = family_descriptor(signature)
        family_id = str(family["familyId"])
        grouped.setdefault(family_id, []).append(item)
        descriptor_by_id[family_id] = family
        signature_by_id[family_id] = signature

    groups: list[dict[str, Any]] = []
    for family_id, members in grouped.items():
        ordered = sorted(
            members,
            key=lambda item: (
                -float(item.get("similarity") or 0),
                str(item.get("sourcePath") or "").casefold(),
            ),
        )
        representative = ordered[0]
        decision = decisions.get(family_id) or {
            "displayName": "",
            "decisionStatus": "PENDING",
            "linkedFormSignatureId": "",
            "contractAvailable": False,
            "validationStatus": "",
            "validationSampleCount": 0,
            "recommendation": "",
            "confidence": 0.0,
            "reviewer": "",
            "notes": "",
        }
        status_counts = Counter(
            str(item.get("status") or "") for item in members
        )
        nearest = Counter(
            str(item.get("nearestKnownFormSignatureId") or "")
            for item in members
            if str(item.get("nearestKnownFormSignatureId") or "")
        )
        nearest_signature = (
            nearest.most_common(1)[0][0] if nearest else ""
        )
        nearest_source = next(
            (
                str(item.get("nearestKnownSource") or "")
                for item in ordered
                if str(item.get("nearestKnownFormSignatureId") or "")
                == nearest_signature
            ),
            "",
        )
        groups.append(
            {
                "familyId": family_id,
                "displayName": (
                    decision["displayName"]
                    or Path(
                        str(
                            representative.get("sourcePath")
                            or ""
                        )
                    ).stem
                ),
                "decisionStatus": decision["decisionStatus"],
                "memberCount": len(members),
                "representativeSource": str(
                    representative.get("sourcePath") or ""
                ),
                "representativeRevisionId": int(
                    representative.get("captureRevisionId") or 0
                ),
                "sampleSources": [
                    str(item.get("sourcePath") or "")
                    for item in ordered[:3]
                ],
                "averageSimilarity": round(
                    sum(
                        float(item.get("similarity") or 0)
                        for item in members
                    )
                    / len(members),
                    4,
                ),
                "nearestKnownFormSignatureId": (
                    decision["linkedFormSignatureId"]
                    or nearest_signature
                ),
                "nearestKnownSource": nearest_source,
                "candidateStatuses": dict(
                    sorted(status_counts.items())
                ),
                "contractAvailable": decision[
                    "contractAvailable"
                ],
                "validationStatus": decision["validationStatus"],
                "validationSampleCount": decision[
                    "validationSampleCount"
                ],
                "recommendation": decision["recommendation"],
                "confidence": decision["confidence"],
                "reviewer": decision["reviewer"],
                "notes": decision["notes"],
                "familySignature": descriptor_by_id[family_id],
                "representativeFormSignature": signature_by_id[
                    family_id
                ],
            }
        )
    groups.sort(
        key=lambda group: (
            0
            if group["decisionStatus"] in {
                "PENDING",
                "ANALYZED_PENDING_APPROVAL",
            }
            else 1,
            -int(group["memberCount"]),
            str(group["familyId"]),
        )
    )
    return {
        "schemaVersion": GROUP_REVIEW_SCHEMA_VERSION,
        "generatedAt": utc_now_iso(),
        "preflightStatus": str(report.get("status") or ""),
        "sourceRoot": str(report.get("sourceRoot") or ""),
        "summary": {
            "groupCount": len(groups),
            "pendingCount": sum(
                group["decisionStatus"]
                in {"PENDING", "ANALYZED_PENDING_APPROVAL"}
                for group in groups
            ),
            "approvedCount": sum(
                group["decisionStatus"] in APPROVED_DECISIONS
                for group in groups
            ),
            "excludedCount": sum(
                group["decisionStatus"] == "EXCLUDED"
                for group in groups
            ),
            "workbookCount": sum(
                int(group["memberCount"]) for group in groups
            ),
        },
        "groups": groups,
    }


def write_form_group_review(
    *,
    database_path: str | Path,
    report_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    report = _load_json(report_path)
    with closing(sqlite3.connect(database)) as connection:
        connection.row_factory = sqlite3.Row
        review = build_form_group_review(
            connection=connection,
            report=report,
        )
    output = Path(output_path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(review, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return review


def _find_group(
    review: dict[str, Any],
    family_id: str,
) -> dict[str, Any]:
    return next(
        (
            group
            for group in review.get("groups") or []
            if group.get("familyId") == family_id
        ),
        None,
    ) or (_raise_unknown_family(family_id))


def _raise_unknown_family(family_id: str) -> Any:
    raise ValueError(f"Unknown form family: {family_id}")


def decide_form_family(
    *,
    database_path: str | Path,
    report_path: str | Path,
    family_id: str,
    decision: str,
    reviewer: str,
    display_name: str = "",
    linked_form_signature_id: str = "",
    notes: str = "",
    group_snapshot: dict[str, Any] | None = None,
) -> dict[str, Any]:
    reviewer_name = reviewer.strip()
    if not reviewer_name:
        raise ValueError("reviewer is required for a form-family decision.")
    normalized = decision.strip().upper()
    mapped = {
        "REGISTER_NEW": "APPROVED_NEW",
        "LINK_EXISTING": "LINKED_EXISTING",
        "EXCLUDE": "EXCLUDED",
    }.get(normalized)
    if mapped is None:
        raise ValueError(
            "decision must be REGISTER_NEW, LINK_EXISTING, or EXCLUDE."
        )
    database = Path(database_path).expanduser().resolve()
    with closing(sqlite3.connect(database)) as connection:
        connection.row_factory = sqlite3.Row
        if group_snapshot is None:
            report = _load_json(report_path)
            review = build_form_group_review(
                connection=connection,
                report=report,
            )
            group = _find_group(review, family_id)
        else:
            group = group_snapshot
            if str(group.get("familyId") or "") != family_id:
                raise ValueError(
                    "Form-family snapshot does not match family_id."
                )
        current = connection.execute(
            """
            SELECT extraction_contract_json, validation_status,
                   validation_sample_count, created_at
            FROM excel_form_families
            WHERE family_id=?
            """,
            (family_id,),
        ).fetchone()
        if mapped == "APPROVED_NEW" and (
            current is None
            or not current["extraction_contract_json"]
            or str(current["validation_status"]) != "PASSED"
        ):
            raise ValueError(
                "신규 양식 등록 전에 대표본 AI 분석과 표본 검증을 완료하세요."
            )
        linked = (
            linked_form_signature_id.strip()
            or str(group["nearestKnownFormSignatureId"] or "")
        )
        if mapped == "LINKED_EXISTING" and not linked:
            raise ValueError("연결할 기존 양식 서명이 없습니다.")
        if mapped == "LINKED_EXISTING":
            from inference_data_ai_form_preflight import (
                load_known_forms,
            )

            known_signature_ids = {
                str(item["signature"]["formSignatureId"])
                for item in load_known_forms(connection)
            }
            if linked not in known_signature_ids:
                raise ValueError(
                    "연결 대상이 현재 분석 완료된 기존 양식 서명에 "
                    "존재하지 않습니다."
                )
        required_validation_count = len(group["sampleSources"])
        if mapped == "APPROVED_NEW" and int(
            current["validation_sample_count"]
        ) != required_validation_count:
            raise ValueError(
                "신규 양식 등록에 필요한 대표본·표본 검증 수가 "
                "현재 양식군과 일치하지 않습니다."
            )
        now = utc_now_iso()
        contract_json = (
            str(current["extraction_contract_json"])
            if current is not None
            and current["extraction_contract_json"]
            else None
        )
        validation_status = (
            str(current["validation_status"])
            if current is not None
            else ""
        )
        validation_count = (
            int(current["validation_sample_count"])
            if current is not None
            else 0
        )
        created_at = (
            str(current["created_at"])
            if current is not None
            else now
        )
        connection.execute(
            """
            INSERT INTO excel_form_families(
                family_id, schema_version, display_name,
                decision_status, family_signature_json,
                representative_source_path,
                representative_revision_id,
                linked_form_signature_id,
                extraction_contract_json, validation_status,
                validation_sample_count, member_count,
                reviewer, notes, created_at, updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(family_id) DO UPDATE SET
                display_name=excluded.display_name,
                decision_status=excluded.decision_status,
                family_signature_json=excluded.family_signature_json,
                representative_source_path=
                    excluded.representative_source_path,
                representative_revision_id=
                    excluded.representative_revision_id,
                linked_form_signature_id=
                    excluded.linked_form_signature_id,
                extraction_contract_json=
                    excluded.extraction_contract_json,
                validation_status=excluded.validation_status,
                validation_sample_count=
                    excluded.validation_sample_count,
                member_count=excluded.member_count,
                reviewer=excluded.reviewer,
                notes=excluded.notes,
                updated_at=excluded.updated_at
            """,
            (
                family_id,
                REGISTRY_SCHEMA_VERSION,
                display_name.strip() or str(group["displayName"]),
                mapped,
                json.dumps(
                    group["familySignature"],
                    ensure_ascii=False,
                    sort_keys=True,
                ),
                str(group["representativeSource"]),
                int(group["representativeRevisionId"]),
                linked if mapped == "LINKED_EXISTING" else "",
                contract_json,
                validation_status,
                validation_count,
                int(group["memberCount"]),
                reviewer_name,
                notes.strip(),
                created_at,
                now,
            ),
        )
        connection.commit()
    return {
        "status": mapped,
        "familyId": family_id,
        "memberCount": int(group["memberCount"]),
        "linkedFormSignatureId": (
            linked if mapped == "LINKED_EXISTING" else ""
        ),
    }


def reclassify_form_preflight_report(
    *,
    database_path: str | Path,
    report_path: str | Path,
) -> dict[str, Any]:
    """Rebuild report decisions and its full-processing manifest from DB."""

    from inference_data_ai_form_preflight import (
        _atomic_write_json,
        classify_form,
        load_known_forms,
        signature_from_database,
    )

    database = Path(database_path).expanduser().resolve()
    report_file = Path(report_path).expanduser().resolve()
    report = _load_json(report_file)
    items = report.get("items")
    if not isinstance(items, list):
        raise ValueError("Form preflight report must contain an items array.")
    with closing(sqlite3.connect(database)) as connection:
        connection.row_factory = sqlite3.Row
        decisions = load_family_decisions(connection)
        known_forms = load_known_forms(connection)
        rebuilt: list[dict[str, Any]] = []
        for original in items:
            if not isinstance(original, dict):
                continue
            item = dict(original)
            revision_id = int(item.get("captureRevisionId") or 0)
            if revision_id <= 0:
                rebuilt.append(item)
                continue
            signature = signature_from_database(
                connection,
                revision_id,
            )
            family = family_descriptor(signature)
            classification = apply_registry_decision(
                classify_form(signature, known_forms),
                family=family,
                decisions=decisions,
            )
            for key in (
                "status",
                "similarity",
                "nearestKnownSource",
                "nearestKnownFormSignatureId",
                "reason",
                "registryDecision",
                "familyDisplayName",
                "extractionContract",
            ):
                item.pop(key, None)
            item.update(classification)
            item["formSignatureId"] = signature["formSignatureId"]
            item["formFamilyId"] = family["familyId"]
            rebuilt.append(item)

    counts = {
        status: sum(
            str(item.get("status") or "") == status
            for item in rebuilt
        )
        for status in (
            "KNOWN_FORM",
            "SIMILAR_FORM_REVIEW",
            "NEW_FORM",
            "EXCLUDED_FORM",
            "CAPTURE_FAILED",
        )
    }
    manifest_path = Path(
        str(
            report.get("knownFormManifestPath")
            or report_file.with_name(
                report_file.stem + ".known-forms.manifest.json"
            )
        )
    ).expanduser().resolve()
    manifest = {
        "schemaVersion": "excel-form-preflight-manifest-v1",
        "sourceRoot": str(report.get("sourceRoot") or ""),
        "workbooks": [
            {
                "relativePath": str(item.get("relativePath") or ""),
                "contentSha256": str(
                    item.get("contentSha256") or ""
                ),
                "formSignatureId": str(
                    item.get("formSignatureId") or ""
                ),
                "preflightStatus": "KNOWN_FORM",
                "formFamilyId": str(
                    item.get("formFamilyId") or ""
                ),
                "registryDecision": str(
                    item.get("registryDecision") or ""
                ),
                **(
                    {
                        "extractionContract": item[
                            "extractionContract"
                        ]
                    }
                    if item.get("extractionContract")
                    else {}
                ),
            }
            for item in rebuilt
            if str(item.get("status") or "") == "KNOWN_FORM"
        ],
    }
    summary = dict(report.get("summary") or {})
    summary.update(
        {
            "total": len(rebuilt),
            "knownForms": counts["KNOWN_FORM"],
            "similarReview": counts["SIMILAR_FORM_REVIEW"],
            "newForms": counts["NEW_FORM"],
            "excludedForms": counts["EXCLUDED_FORM"],
            "captureFailed": counts["CAPTURE_FAILED"],
            "fullProcessingAllowed": (
                str(report.get("status") or "") == "COMPLETED"
                and counts["KNOWN_FORM"] > 0
            ),
        }
    )
    report.update(
        {
            "generatedAt": utc_now_iso(),
            "registryAppliedAt": utc_now_iso(),
            "knownCatalogCount": len(known_forms),
            "knownFormManifestPath": str(manifest_path),
            "summary": summary,
            "items": rebuilt,
        }
    )
    _atomic_write_json(manifest_path, manifest)
    _atomic_write_json(report_file, report)
    return report


def _capture_preview(
    connection: sqlite3.Connection,
    revision_id: int,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    sheets = connection.execute(
        """
        SELECT sheet_id, title, used_bounds_json, content_bounds_json,
               merge_count, formula_cell_count
        FROM capture_v2_sheets
        WHERE revision_id=?
        ORDER BY sheet_index
        """,
        (revision_id,),
    ).fetchall()
    for sheet in sheets:
        cells: list[dict[str, Any]] = []
        for cell in connection.execute(
            """
            SELECT coordinate, display_value_json, raw_value_json,
                   merge_role, number_format
            FROM capture_v2_cells
            WHERE sheet_id=?
              AND (display_value_json IS NOT NULL
                   OR raw_value_json IS NOT NULL)
            ORDER BY row_index, column_index
            LIMIT 180
            """,
            (int(sheet["sheet_id"]),),
        ):
            raw = (
                cell["display_value_json"]
                if cell["display_value_json"] is not None
                else cell["raw_value_json"]
            )
            try:
                value = json.loads(raw) if raw is not None else None
            except json.JSONDecodeError:
                value = str(raw)
            cells.append(
                {
                    "coordinate": str(cell["coordinate"]),
                    "value": value,
                    "mergeRole": str(cell["merge_role"]),
                    "numberFormat": str(cell["number_format"]),
                }
            )
        result.append(
            {
                "title": str(sheet["title"]),
                "usedBounds": json.loads(
                    sheet["used_bounds_json"] or "null"
                ),
                "contentBounds": json.loads(
                    sheet["content_bounds_json"] or "null"
                ),
                "mergeCount": int(sheet["merge_count"]),
                "formulaCellCount": int(sheet["formula_cell_count"]),
                "cells": cells,
            }
        )
    return result


def _contract_schema() -> dict[str, Any]:
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "additionalProperties": False,
        "required": [
            "schemaVersion",
            "familyId",
            "familyName",
            "documentType",
            "extractionContract",
            "sampleValidation",
            "confidence",
            "recommendation",
        ],
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": CONTRACT_SCHEMA_VERSION,
            },
            "familyId": {"type": "string"},
            "familyName": {"type": "string"},
            "documentType": {"type": "string"},
            "extractionContract": {
                "type": "object",
                "additionalProperties": False,
                "required": [
                    "targetSheets",
                    "headerPatterns",
                    "tableRules",
                    "requiredFields",
                    "cautions",
                ],
                "properties": {
                    "targetSheets": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "headerPatterns": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "tableRules": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "requiredFields": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "cautions": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                },
            },
            "sampleValidation": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": [
                        "sourcePath",
                        "compatible",
                        "reason",
                    ],
                    "properties": {
                        "sourcePath": {"type": "string"},
                        "compatible": {"type": "boolean"},
                        "reason": {"type": "string"},
                    },
                },
            },
            "confidence": {
                "type": "number",
                "minimum": 0,
                "maximum": 1,
            },
            "recommendation": {
                "type": "string",
                "enum": [
                    "REGISTER_NEW",
                    "LINK_EXISTING",
                    "EXCLUDE",
                ]
            },
        },
    }


def _codex_command(value: str | None) -> list[str]:
    if value and value.strip():
        return [value.strip()]
    executable = shutil.which(
        "codex.cmd" if os.name == "nt" else "codex"
    )
    if not executable:
        raise RuntimeError("Codex CLI 실행 파일을 찾을 수 없습니다.")
    return [executable]


def analyze_form_family(
    *,
    database_path: str | Path,
    report_path: str | Path,
    family_id: str,
    output_path: str | Path,
    codex_executable: str | None = None,
    reasoning_effort: str = "medium",
    timeout_seconds: int = 900,
    group_snapshot: dict[str, Any] | None = None,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    report = _load_json(report_path)
    with closing(sqlite3.connect(database)) as connection:
        connection.row_factory = sqlite3.Row
        if group_snapshot is None:
            review = build_form_group_review(
                connection=connection,
                report=report,
            )
            group = _find_group(review, family_id)
        else:
            group = group_snapshot
            if str(group.get("familyId") or "") != family_id:
                raise ValueError(
                    "Form-family snapshot does not match family_id."
                )
        sample_sources = list(group["sampleSources"])
        request = {
            "schemaVersion": "excel-form-family-analysis-request-v1",
            "familyId": family_id,
            "familySignature": group["familySignature"],
            "memberCount": int(group["memberCount"]),
            "representativeSource": group["representativeSource"],
            "nearestKnownSource": group["nearestKnownSource"],
            "samples": [
                {
                    "sourcePath": source,
                    "isRepresentative": index == 0,
                    "capture": _capture_preview(
                        connection,
                        next(
                            int(item.get("captureRevisionId") or 0)
                            for item in report.get("items") or []
                            if str(item.get("sourcePath") or "")
                            == source
                        ),
                    ),
                }
                for index, source in enumerate(sample_sources)
            ],
        }

    output = Path(output_path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(
        prefix="form-family-ai-",
        dir=output.parent,
    ) as temporary:
        temporary_root = Path(temporary)
        schema_path = temporary_root / "contract.schema.json"
        response_path = temporary_root / "response.json"
        schema_path.write_text(
            json.dumps(
                _contract_schema(),
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        prompt = (
            "You analyze one already-captured Excel layout family. "
            "Do not open files, do not modify data, and do not infer result "
            "values. Build a reusable extraction contract from the bounded "
            "coordinate previews. Validate whether every supplied sample can "
            "use the same contract. Recommend REGISTER_NEW only when the "
            "layout is coherent; LINK_EXISTING when it clearly matches the "
            "nearest known form; otherwise EXCLUDE. Return only JSON.\n\n"
            + json.dumps(request, ensure_ascii=False, indent=2)
        )
        command = [
            *_codex_command(codex_executable),
            "exec",
            "--sandbox",
            "read-only",
            "--output-schema",
            str(schema_path),
            "--output-last-message",
            str(response_path),
        ]
        if reasoning_effort:
            command.extend(
                [
                    "-c",
                    f'model_reasoning_effort="{reasoning_effort}"',
                ]
            )
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
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise RuntimeError(
                "대표본 AI 분석 실패: " + detail[-1600:]
            )
        if not response_path.is_file():
            raise RuntimeError("대표본 AI 분석 결과 JSON이 없습니다.")
        contract = json.loads(
            response_path.read_text(encoding="utf-8")
        )
    if contract.get("schemaVersion") != CONTRACT_SCHEMA_VERSION:
        raise ValueError("AI 양식 계약 schemaVersion이 올바르지 않습니다.")
    if contract.get("familyId") != family_id:
        raise ValueError("AI 양식 계약 familyId가 요청과 다릅니다.")
    supplied_samples = {
        str(item.get("sourcePath") or "")
        for item in contract.get("sampleValidation") or []
    }
    if supplied_samples != set(sample_sources):
        raise ValueError("AI 표본 검증 목록이 요청 표본과 다릅니다.")
    all_compatible = all(
        bool(item.get("compatible"))
        for item in contract.get("sampleValidation") or []
    )
    validation_status = "PASSED" if all_compatible else "FAILED"
    output.write_text(
        json.dumps(contract, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    with closing(sqlite3.connect(database)) as connection:
        connection.row_factory = sqlite3.Row
        ensure_form_registry_schema(connection)
        now = utc_now_iso()
        connection.execute(
            """
            INSERT INTO excel_form_families(
                family_id, schema_version, display_name,
                decision_status, family_signature_json,
                representative_source_path,
                representative_revision_id,
                linked_form_signature_id,
                extraction_contract_json, validation_status,
                validation_sample_count, member_count,
                reviewer, notes, created_at, updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(family_id) DO UPDATE SET
                display_name=excluded.display_name,
                decision_status=excluded.decision_status,
                family_signature_json=excluded.family_signature_json,
                representative_source_path=
                    excluded.representative_source_path,
                representative_revision_id=
                    excluded.representative_revision_id,
                extraction_contract_json=
                    excluded.extraction_contract_json,
                validation_status=excluded.validation_status,
                validation_sample_count=
                    excluded.validation_sample_count,
                member_count=excluded.member_count,
                updated_at=excluded.updated_at
            """,
            (
                family_id,
                REGISTRY_SCHEMA_VERSION,
                str(contract.get("familyName") or group["displayName"]),
                "ANALYZED_PENDING_APPROVAL",
                json.dumps(
                    group["familySignature"],
                    ensure_ascii=False,
                    sort_keys=True,
                ),
                str(group["representativeSource"]),
                int(group["representativeRevisionId"]),
                "",
                json.dumps(
                    contract,
                    ensure_ascii=False,
                    sort_keys=True,
                ),
                validation_status,
                len(sample_sources),
                int(group["memberCount"]),
                "",
                "",
                now,
                now,
            ),
        )
        connection.commit()
    return {
        "status": "ANALYZED_PENDING_APPROVAL",
        "familyId": family_id,
        "contractPath": str(output),
        "validationStatus": validation_status,
        "validationSampleCount": len(sample_sources),
        "recommendation": str(contract.get("recommendation") or ""),
        "confidence": float(contract.get("confidence") or 0),
    }


__all__ = [
    "APPROVED_DECISIONS",
    "CONTRACT_SCHEMA_VERSION",
    "GROUP_REVIEW_SCHEMA_VERSION",
    "analyze_form_family",
    "apply_registry_decision",
    "build_form_group_review",
    "decide_form_family",
    "ensure_form_registry_schema",
    "family_descriptor",
    "load_family_decisions",
    "reclassify_form_preflight_report",
    "write_form_group_review",
]
