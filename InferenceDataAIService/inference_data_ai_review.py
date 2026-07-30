from __future__ import annotations

import datetime as dt
import json
import sqlite3
from collections import defaultdict
from contextlib import contextmanager
from typing import Any, Iterator

from inference_data_ai_effects import (
    EffectCalculationError,
    calculate_effect_bundle,
)
from inference_data_ai_schema import public_id, resolve_unit_id, stable_uid


REVIEW_CONTRACT_VERSION = "canonical-human-review-v1"
_DECISIONS = {
    "APPROVE",
    "REJECT",
    "EXCLUDE",
    "RETURN_TO_REVIEW",
}


class ReviewGateError(ValueError):
    """Raised when a human-review decision cannot safely be applied."""

    def __init__(self, code: str, message: str):
        super().__init__(f"{code}: {message}")
        self.code = code


def _utc_now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def _rows(
    connection: sqlite3.Connection,
    sql: str,
    params: tuple[object, ...] = (),
) -> list[dict[str, Any]]:
    cursor = connection.execute(sql, params)
    columns = [str(item[0]) for item in cursor.description]
    return [
        {columns[index]: value for index, value in enumerate(row)}
        for row in cursor.fetchall()
    ]


def _one(
    connection: sqlite3.Connection,
    sql: str,
    params: tuple[object, ...] = (),
) -> dict[str, Any] | None:
    rows = _rows(connection, sql, params)
    return rows[0] if rows else None


def _text(value: object) -> str:
    return str(value or "").strip()


def _status(value: object) -> str:
    return _text(value).upper()


def _normalized_metric_type(value: object) -> str:
    return " ".join(
        _text(value).casefold().replace("-", " ").replace("_", " ").split()
    )


def _is_denominator_only_outcome(outcome: dict[str, Any]) -> bool:
    return _normalized_metric_type(outcome.get("metric_type")) in {
        "cohort size",
        "denominator",
        "input size",
        "sample size",
    }


def _effect_outcome_base_key(outcome: dict[str, Any]) -> str:
    tokens = _normalized_metric_type(outcome.get("outcome_key")).split()
    return " ".join(
        token
        for token in tokens
        if token not in {"count", "rate", "percent", "percentage", "pct"}
    )


def _paired_rate_signature(
    pair: dict[str, Any],
) -> tuple[object, ...] | None:
    values = (
        pair["compared"].get("numerator"),
        pair["compared"].get("denominator"),
        pair["control"].get("numerator"),
        pair["control"].get("denominator"),
    )
    if any(value is None for value in values):
        return None
    return (
        _text(pair.get("stratumKey")),
        _text(pair.get("replicateKey")),
        *(float(value) for value in values),
    )


def _json(value: object) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


@contextmanager
def _atomic(connection: sqlite3.Connection) -> Iterator[None]:
    savepoint = "canonical_human_review"
    connection.execute(f"SAVEPOINT {savepoint}")
    try:
        yield
    except BaseException:
        connection.execute(f"ROLLBACK TO SAVEPOINT {savepoint}")
        connection.execute(f"RELEASE SAVEPOINT {savepoint}")
        raise
    else:
        connection.execute(f"RELEASE SAVEPOINT {savepoint}")


def _comparison_context(
    connection: sqlite3.Connection,
    public_comparison_id: str,
) -> dict[str, Any]:
    comparison_id = _text(public_comparison_id).upper()
    if not comparison_id:
        raise ReviewGateError(
            "PUBLIC_COMPARISON_ID_REQUIRED",
            "A public CMP identifier is required.",
        )
    row = _one(
        connection,
        """
        SELECT
            c.comparison_id, c.comparison_uid, c.public_comparison_id,
            c.comparison_key, c.compared_arm_id, c.control_arm_id,
            c.design_type AS comparison_design_type,
            c.matching_basis AS comparison_matching_basis,
            c.validity_status, c.confounding_status AS comparison_confounding_status,
            c.exclusion_reason, c.summary_text AS comparison_summary,
            c.aggregation_eligible AS comparison_aggregation_eligible,
            c.verification_status AS comparison_verification_status,
            s.study_id, s.study_uid, s.public_data_id, s.study_key,
            s.title AS study_title, s.design_type AS study_design_type,
            s.comparison_basis, s.analysis_status AS study_analysis_status,
            s.verification_status AS study_verification_status,
            s.comparability_status,
            s.confounding_status AS study_confounding_status,
            wa.workbook_analysis_id, wa.analysis_uid, wa.public_analysis_id,
            wa.title AS analysis_title,
            wa.analysis_status AS workbook_analysis_status,
            wa.verification_status AS workbook_verification_status,
            wa.revision_id AS analysis_revision_id,
            wa.document_id AS analysis_document_id,
            sr.revision_uid, sr.source_fingerprint, sr.fingerprint_kind,
            sr.content_sha256, sr.capture_status, sr.source_content_status,
            sr.is_current, sr.document_id AS revision_document_id,
            sd.source_path, sd.original_file_name, sd.lifecycle_status,
            compared.arm_uid AS compared_arm_uid,
            compared.arm_key AS compared_arm_key,
            compared.label AS compared_arm_label,
            compared.verification_status AS compared_arm_verification_status,
            control.arm_uid AS control_arm_uid,
            control.arm_key AS control_arm_key,
            control.label AS control_arm_label,
            control.verification_status AS control_arm_verification_status
        FROM knowledge_comparisons AS c
        JOIN knowledge_studies AS s ON s.study_id=c.study_id
        JOIN workbook_analyses AS wa
          ON wa.workbook_analysis_id=s.workbook_analysis_id
        JOIN source_revisions AS sr ON sr.revision_id=wa.revision_id
        JOIN source_documents AS sd ON sd.document_id=wa.document_id
        JOIN knowledge_arms AS compared
          ON compared.arm_id=c.compared_arm_id
        JOIN knowledge_arms AS control
          ON control.arm_id=c.control_arm_id
        WHERE UPPER(c.public_comparison_id)=?
        LIMIT 1
        """,
        (comparison_id,),
    )
    if row is None:
        raise ReviewGateError(
            "COMPARISON_NOT_FOUND",
            f"Comparison {comparison_id} does not exist.",
        )
    return row


def _latest_decision(
    connection: sqlite3.Connection,
    comparison_uid: str,
) -> dict[str, Any] | None:
    return _one(
        connection,
        """
        SELECT decision_uid, decision, reason, reviewer, decided_at,
               supersedes_decision_uid
        FROM review_decisions
        WHERE entity_type='COMPARISON' AND entity_uid=?
        ORDER BY review_decision_id DESC
        LIMIT 1
        """,
        (comparison_uid,),
    )


def _source_blockers(
    connection: sqlite3.Connection,
    context: dict[str, Any],
) -> list[dict[str, str]]:
    blockers: list[dict[str, str]] = []
    if _status(context["workbook_verification_status"]) == "STALE":
        blockers.append(
            {
                "code": "ANALYSIS_SUPERSEDED",
                "message": "The comparison belongs to a superseded analysis draft.",
            }
        )
    if int(context["analysis_document_id"]) != int(
        context["revision_document_id"]
    ):
        blockers.append(
            {
                "code": "SOURCE_DOCUMENT_MISMATCH",
                "message": "Analysis and revision belong to different documents.",
            }
        )
    if not bool(context["is_current"]):
        blockers.append(
            {
                "code": "SOURCE_NOT_CURRENT",
                "message": "The comparison belongs to a superseded revision.",
            }
        )
    if _status(context["capture_status"]) != "CAPTURED":
        blockers.append(
            {
                "code": "SOURCE_NOT_CAPTURED",
                "message": "The current revision is not in CAPTURED state.",
            }
        )
    if _status(context["lifecycle_status"]) != "ACTIVE":
        blockers.append(
            {
                "code": "SOURCE_NOT_ACTIVE",
                "message": "The source document is not active.",
            }
        )
    if _status(context["fingerprint_kind"]) != "SHA256":
        blockers.append(
            {
                "code": "SOURCE_HASH_KIND_INVALID",
                "message": "The current revision is not fingerprinted with SHA256.",
            }
        )
    content_sha256 = _text(context["content_sha256"]).lower()
    source_fingerprint = _text(context["source_fingerprint"]).lower()
    if (
        len(content_sha256) != 64
        or any(character not in "0123456789abcdef" for character in content_sha256)
    ):
        blockers.append(
            {
                "code": "SOURCE_HASH_INVALID",
                "message": "The current revision has no valid SHA256 content hash.",
            }
        )
    elif source_fingerprint != content_sha256:
        blockers.append(
            {
                "code": "SOURCE_HASH_MISMATCH",
                "message": "Revision fingerprint does not match its content hash.",
            }
        )
    current_count = int(
        connection.execute(
            """
            SELECT COUNT(*)
            FROM source_revisions
            WHERE document_id=? AND is_current=1
            """,
            (int(context["analysis_document_id"]),),
        ).fetchone()[0]
    )
    if current_count != 1:
        blockers.append(
            {
                "code": "SOURCE_CURRENT_REVISION_AMBIGUOUS",
                "message": "The source must have exactly one current revision.",
            }
        )
    return blockers


def _direct_current_evidence(
    connection: sqlite3.Connection,
    *,
    entity_type: str,
    entity_uid: str,
    context: dict[str, Any],
) -> list[dict[str, Any]]:
    expected_revision_id = int(context["analysis_revision_id"])
    expected_sha256 = _text(context["content_sha256"]).lower()
    return _rows(
        connection,
        """
        SELECT
            e.evidence_id, e.evidence_uid, e.public_evidence_id,
            e.sheet_name, e.range_address, e.evidence_role,
            l.evidence_role AS link_role, e.verification_status,
            e.revision_id,
            COALESCE(NULLIF(LOWER(e.content_sha256), ''),
                     LOWER(er.content_sha256)) AS effective_content_sha256
        FROM entity_evidence_links AS l
        JOIN evidence_items AS e ON e.evidence_id=l.evidence_id
        JOIN source_revisions AS er ON er.revision_id=e.revision_id
        WHERE UPPER(l.entity_type)=?
          AND l.entity_uid=?
          AND e.verification_status='VERIFIED'
          AND e.revision_id=?
          AND er.is_current=1
          AND er.capture_status='CAPTURED'
          AND LOWER(er.content_sha256)=?
          AND COALESCE(NULLIF(LOWER(e.content_sha256), ''),
                       LOWER(er.content_sha256))=?
        ORDER BY e.public_evidence_id
        """,
        (
            _status(entity_type),
            entity_uid,
            expected_revision_id,
            expected_sha256,
            expected_sha256,
        ),
    )


def _outcomes_and_pairs(
    connection: sqlite3.Connection,
    context: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    outcomes = _rows(
        connection,
        """
        SELECT
            o.outcome_id, o.outcome_uid, o.outcome_key,
            o.original_label, o.metric_type, o.original_unit,
            o.denominator_basis, o.favorable_direction,
            o.verification_status,
            u.canonical_symbol
        FROM knowledge_outcomes AS o
        LEFT JOIN knowledge_units AS u ON u.unit_id=o.unit_id
        WHERE o.study_id=?
        ORDER BY o.outcome_id
        """,
        (int(context["study_id"]),),
    )
    pairs: list[dict[str, Any]] = []
    for outcome in outcomes:
        observations = _rows(
            connection,
            """
            SELECT
                observation_id, observation_uid, outcome_id, arm_id,
                observation_key, stratum_key, replicate_key, observed_at,
                value_number, value_text, numerator, denominator, rate_ppm,
                min_value, max_value, average_value, sample_size,
                result_status, verification_status, details_json
            FROM knowledge_observations
            WHERE outcome_id=? AND arm_id IN (?, ?)
            ORDER BY observation_id
            """,
            (
                int(outcome["outcome_id"]),
                int(context["compared_arm_id"]),
                int(context["control_arm_id"]),
            ),
        )
        by_arm_and_partition: dict[
            tuple[int, str, str], list[dict[str, Any]]
        ] = defaultdict(list)
        for observation in observations:
            key = (
                int(observation["arm_id"]),
                _text(observation["stratum_key"]),
                _text(observation["replicate_key"]),
            )
            by_arm_and_partition[key].append(observation)
        compared_partitions = {
            (key[1], key[2])
            for key in by_arm_and_partition
            if key[0] == int(context["compared_arm_id"])
        }
        control_partitions = {
            (key[1], key[2])
            for key in by_arm_and_partition
            if key[0] == int(context["control_arm_id"])
        }
        for stratum_key, replicate_key in sorted(
            compared_partitions & control_partitions
        ):
            compared = by_arm_and_partition[
                (
                    int(context["compared_arm_id"]),
                    stratum_key,
                    replicate_key,
                )
            ]
            control = by_arm_and_partition[
                (
                    int(context["control_arm_id"]),
                    stratum_key,
                    replicate_key,
                )
            ]
            if len(compared) != 1 or len(control) != 1:
                raise ReviewGateError(
                    "OBSERVATION_PAIR_AMBIGUOUS",
                    (
                        f"Outcome {outcome['outcome_key']} has more than one "
                        "observation in the same stratum/replicate partition."
                    ),
                )
            pairs.append(
                {
                    "outcome": outcome,
                    "compared": compared[0],
                    "control": control[0],
                    "stratumKey": stratum_key,
                    "replicateKey": replicate_key,
                }
            )
    return outcomes, pairs


def _approval_plan(
    connection: sqlite3.Connection,
    context: dict[str, Any],
) -> dict[str, Any]:
    blockers = _source_blockers(connection, context)
    if int(context["compared_arm_id"]) == int(context["control_arm_id"]):
        blockers.append(
            {
                "code": "ARMS_NOT_DISTINCT",
                "message": "Compared and control arms must be distinct.",
            }
        )
    if not _text(context["comparison_matching_basis"]):
        blockers.append(
            {
                "code": "MATCHING_BASIS_REQUIRED",
                "message": "The comparison requires a non-empty matching basis.",
            }
        )
    if _status(context["comparability_status"]) != "VALID":
        blockers.append(
            {
                "code": "STUDY_COMPARABILITY_NOT_VALID",
                "message": "Study comparability must be explicitly VALID.",
            }
        )
    if _status(context["study_confounding_status"]) != "NONE":
        blockers.append(
            {
                "code": "STUDY_CONFOUNDING_NOT_NONE",
                "message": "Study confounding must be explicitly NONE.",
            }
        )
    if _status(context["validity_status"]) != "VALID":
        blockers.append(
            {
                "code": "COMPARISON_VALIDITY_NOT_VALID",
                "message": "Comparison validity must be explicitly VALID.",
            }
        )
    if _status(context["comparison_confounding_status"]) != "NONE":
        blockers.append(
            {
                "code": "COMPARISON_CONFOUNDING_NOT_NONE",
                "message": "Comparison confounding must be explicitly NONE.",
            }
        )

    comparison_evidence = _direct_current_evidence(
        connection,
        entity_type="COMPARISON",
        entity_uid=str(context["comparison_uid"]),
        context=context,
    )
    if not comparison_evidence:
        blockers.append(
            {
                "code": "COMPARISON_EVIDENCE_REQUIRED",
                "message": (
                    "The comparison requires direct verified evidence from "
                    "the current source revision."
                ),
            }
        )

    try:
        outcomes, pairs = _outcomes_and_pairs(connection, context)
    except ReviewGateError as exc:
        outcomes, pairs = [], []
        blockers.append({"code": exc.code, "message": str(exc).split(": ", 1)[-1]})
    if not pairs:
        blockers.append(
            {
                "code": "PAIRED_OBSERVATIONS_REQUIRED",
                "message": (
                    "At least one outcome must have a compatible compared/control "
                    "observation pair."
                ),
            }
        )

    planned_effects: list[dict[str, Any]] = []
    pair_payloads: list[dict[str, Any]] = []
    explicit_rate_pairs = {
        (
            _effect_outcome_base_key(pair["outcome"]),
            signature,
        )
        for pair in pairs
        if "rate" in _normalized_metric_type(
            pair["outcome"].get("metric_type")
        )
        and (signature := _paired_rate_signature(pair)) is not None
    }
    for pair in pairs:
        compared_evidence = _direct_current_evidence(
            connection,
            entity_type="OBSERVATION",
            entity_uid=str(pair["compared"]["observation_uid"]),
            context=context,
        )
        control_evidence = _direct_current_evidence(
            connection,
            entity_type="OBSERVATION",
            entity_uid=str(pair["control"]["observation_uid"]),
            context=context,
        )
        for side, evidence in (
            ("compared", compared_evidence),
            ("control", control_evidence),
        ):
            if evidence:
                continue
            blockers.append(
                {
                    "code": "OBSERVATION_EVIDENCE_REQUIRED",
                    "message": (
                        f"The {side} observation for outcome "
                        f"{pair['outcome']['outcome_key']} requires direct "
                        "verified current-revision evidence."
                    ),
                }
            )
        pair_payloads.append(
            {
                **pair,
                "comparedEvidence": compared_evidence,
                "controlEvidence": control_evidence,
            }
        )
        if not compared_evidence or not control_evidence:
            continue
        if _is_denominator_only_outcome(pair["outcome"]):
            continue
        metric_type = _normalized_metric_type(
            pair["outcome"].get("metric_type")
        )
        rate_signature = _paired_rate_signature(pair)
        if (
            "count" in metric_type
            and rate_signature is not None
            and (
                _effect_outcome_base_key(pair["outcome"]),
                rate_signature,
            )
            in explicit_rate_pairs
        ):
            # Preserve the exact source count as descriptive evidence, but do
            # not emit a second copy of the same rate effects when the source
            # also supplies an explicitly named rate outcome.
            continue
        comparison_for_calculation = {
            **context,
            "verification_status": "VERIFIED",
            "compared_arm_id": int(context["compared_arm_id"]),
            "control_arm_id": int(context["control_arm_id"]),
            "confounding_status": context[
                "comparison_confounding_status"
            ],
        }
        study_for_calculation = {
            **context,
            "verification_status": "VERIFIED",
            "comparability_status": context["comparability_status"],
            "confounding_status": context["study_confounding_status"],
        }
        outcome_for_calculation = {
            **pair["outcome"],
            "unit": (
                pair["outcome"]["canonical_symbol"]
                or pair["outcome"]["original_unit"]
            ),
        }
        try:
            bundle = calculate_effect_bundle(
                compared_observation=pair["compared"],
                control_observation=pair["control"],
                comparison=comparison_for_calculation,
                outcome=outcome_for_calculation,
                study=study_for_calculation,
            )
        except EffectCalculationError as exc:
            blockers.append(
                {
                    "code": "EFFECT_CALCULATION_FAILED",
                    "message": (
                        f"Outcome {pair['outcome']['outcome_key']}: {exc}"
                    ),
                }
            )
            continue
        for effect in bundle:
            planned_effects.append(
                {
                    **effect,
                    "outcome": pair["outcome"],
                    "comparedObservation": pair["compared"],
                    "controlObservation": pair["control"],
                    "evidence": [
                        *comparison_evidence,
                        *compared_evidence,
                        *control_evidence,
                    ],
                }
            )

    identity_counts: dict[tuple[int, str, str], int] = defaultdict(int)
    for effect in planned_effects:
        identity_counts[
            (
                int(effect["outcome"]["outcome_id"]),
                str(effect["effectType"]),
                str(effect["formulaVersion"]),
            )
        ] += 1
    if any(count > 1 for count in identity_counts.values()):
        blockers.append(
            {
                "code": "MULTIPLE_PAIRS_REQUIRE_SEPARATE_COMPARISONS",
                "message": (
                    "A comparison cannot store multiple stratum/replicate "
                    "effects for the same outcome and formula. Create separate "
                    "comparison records."
                ),
            }
        )
    return {
        "ready": not blockers,
        "blockers": blockers,
        "outcomes": outcomes,
        "pairs": pair_payloads,
        "effects": planned_effects,
        "comparisonEvidence": comparison_evidence,
    }


def list_review_queue(
    connection: sqlite3.Connection,
    *,
    limit: int = 500,
) -> dict[str, Any]:
    """Return current Study/Comparison records that still require review."""

    if limit < 1:
        raise ValueError("limit must be positive")
    items = _rows(
        connection,
        """
        SELECT
            c.public_comparison_id AS publicComparisonId,
            s.public_data_id AS publicDataId,
            wa.public_analysis_id AS publicAnalysisId,
            sd.source_path AS sourcePath,
            sd.original_file_name AS fileName,
            s.title AS studyTitle,
            c.summary_text AS comparisonSummary,
            s.verification_status AS studyVerificationStatus,
            s.comparability_status AS studyComparabilityStatus,
            s.confounding_status AS studyConfoundingStatus,
            c.verification_status AS comparisonVerificationStatus,
            c.validity_status AS comparisonValidityStatus,
            c.confounding_status AS comparisonConfoundingStatus,
            c.matching_basis AS matchingBasis,
            sr.revision_uid AS revisionUid,
            sr.content_sha256 AS contentSha256
        FROM knowledge_comparisons AS c
        JOIN knowledge_studies AS s ON s.study_id=c.study_id
        JOIN workbook_analyses AS wa
          ON wa.workbook_analysis_id=s.workbook_analysis_id
        JOIN source_revisions AS sr ON sr.revision_id=wa.revision_id
        JOIN source_documents AS sd ON sd.document_id=wa.document_id
        WHERE sr.is_current=1
          AND sr.capture_status='CAPTURED'
          AND sd.lifecycle_status='ACTIVE'
          AND wa.verification_status<>'STALE'
          AND (
              s.verification_status='NEEDS_REVIEW'
              OR c.verification_status='NEEDS_REVIEW'
              OR c.validity_status='NEEDS_REVIEW'
          )
        ORDER BY sd.source_path, s.public_data_id, c.public_comparison_id
        LIMIT ?
        """,
        (limit,),
    )
    return {
        "schemaVersion": REVIEW_CONTRACT_VERSION,
        "imagesAnalyzed": False,
        "count": len(items),
        "items": items,
    }


def get_review_detail(
    connection: sqlite3.Connection,
    public_comparison_id: str,
) -> dict[str, Any]:
    """Return current evidence and approval readiness without mutating data."""

    context = _comparison_context(connection, public_comparison_id)
    plan = _approval_plan(connection, context)
    latest = _latest_decision(connection, str(context["comparison_uid"]))
    pair_details = []
    for pair in plan["pairs"]:
        compared = pair["compared"]
        control = pair["control"]
        pair_details.append(
            {
                "outcomeUid": str(pair["outcome"]["outcome_uid"]),
                "outcomeKey": str(pair["outcome"]["outcome_key"]),
                "outcomeLabel": str(pair["outcome"]["original_label"]),
                "outcomeUnit": str(
                    pair["outcome"]["canonical_symbol"]
                    or pair["outcome"]["original_unit"]
                    or ""
                ),
                "stratumKey": str(pair["stratumKey"]),
                "replicateKey": str(pair["replicateKey"]),
                "comparedObservation": {
                    "uid": str(compared["observation_uid"]),
                    "valueNumber": compared["value_number"],
                    "valueText": str(compared["value_text"] or ""),
                    "numerator": compared["numerator"],
                    "denominator": compared["denominator"],
                    "ratePpm": compared["rate_ppm"],
                    "sampleSize": compared["sample_size"],
                    "evidence": [
                        {
                            "publicEvidenceId": str(
                                item["public_evidence_id"]
                            ),
                            "sheet": str(item["sheet_name"]),
                            "range": str(item["range_address"]),
                        }
                        for item in pair["comparedEvidence"]
                    ],
                },
                "controlObservation": {
                    "uid": str(control["observation_uid"]),
                    "valueNumber": control["value_number"],
                    "valueText": str(control["value_text"] or ""),
                    "numerator": control["numerator"],
                    "denominator": control["denominator"],
                    "ratePpm": control["rate_ppm"],
                    "sampleSize": control["sample_size"],
                    "evidence": [
                        {
                            "publicEvidenceId": str(
                                item["public_evidence_id"]
                            ),
                            "sheet": str(item["sheet_name"]),
                            "range": str(item["range_address"]),
                        }
                        for item in pair["controlEvidence"]
                    ],
                },
            }
        )
    return {
        "schemaVersion": REVIEW_CONTRACT_VERSION,
        "imagesAnalyzed": False,
        "publicComparisonId": str(context["public_comparison_id"]),
        "publicDataId": str(context["public_data_id"]),
        "publicAnalysisId": str(context["public_analysis_id"]),
        "source": {
            "path": str(context["source_path"]),
            "fileName": str(context["original_file_name"]),
            "revisionUid": str(context["revision_uid"]),
            "contentSha256": str(context["content_sha256"]),
            "isCurrent": bool(context["is_current"]),
            "captureStatus": str(context["capture_status"]),
        },
        "study": {
            "uid": str(context["study_uid"]),
            "title": str(context["study_title"]),
            "verificationStatus": str(context["study_verification_status"]),
            "comparabilityStatus": str(context["comparability_status"]),
            "confoundingStatus": str(context["study_confounding_status"]),
        },
        "comparison": {
            "uid": str(context["comparison_uid"]),
            "key": str(context["comparison_key"]),
            "verificationStatus": str(
                context["comparison_verification_status"]
            ),
            "validityStatus": str(context["validity_status"]),
            "confoundingStatus": str(
                context["comparison_confounding_status"]
            ),
            "matchingBasis": str(context["comparison_matching_basis"]),
            "comparedArm": {
                "uid": str(context["compared_arm_uid"]),
                "key": str(context["compared_arm_key"]),
                "label": str(context["compared_arm_label"]),
            },
            "controlArm": {
                "uid": str(context["control_arm_uid"]),
                "key": str(context["control_arm_key"]),
                "label": str(context["control_arm_label"]),
            },
            "evidenceIds": [
                str(item["public_evidence_id"])
                for item in plan["comparisonEvidence"]
            ],
            "evidence": [
                {
                    "publicEvidenceId": str(item["public_evidence_id"]),
                    "sheet": str(item["sheet_name"]),
                    "range": str(item["range_address"]),
                }
                for item in plan["comparisonEvidence"]
            ],
        },
        "pairedObservations": pair_details,
        "approvalReadiness": {
            "ready": bool(plan["ready"]),
            "blockers": plan["blockers"],
            "plannedEffectCount": len(plan["effects"]),
        },
        "latestDecision": latest,
    }


def _validate_decision_inputs(
    decision: str,
    reviewer: str,
    reason: str,
) -> tuple[str, str, str]:
    normalized_decision = _status(decision)
    normalized_reviewer = _text(reviewer)
    normalized_reason = _text(reason)
    if normalized_decision not in _DECISIONS:
        raise ReviewGateError(
            "DECISION_INVALID",
            f"Decision must be one of {sorted(_DECISIONS)}.",
        )
    if not normalized_reviewer:
        raise ReviewGateError(
            "REVIEWER_REQUIRED",
            "Reviewer must be non-empty.",
        )
    if not normalized_reason:
        raise ReviewGateError(
            "REASON_REQUIRED",
            "Decision reason must be non-empty.",
        )
    return normalized_decision, normalized_reviewer, normalized_reason


def _apply_explicit_assessment(
    connection: sqlite3.Connection,
    context: dict[str, Any],
    *,
    study_comparability_status: str | None,
    study_confounding_status: str | None,
    comparison_validity_status: str | None,
    comparison_confounding_status: str | None,
    matching_basis: str | None,
) -> tuple[dict[str, str], bool]:
    """Persist only review fields explicitly supplied by the human reviewer."""

    specifications = (
        (
            "studyComparabilityStatus",
            study_comparability_status,
            {"VALID", "PARTIAL", "INVALID", "UNASSESSED"},
            _status(context["comparability_status"]),
        ),
        (
            "studyConfoundingStatus",
            study_confounding_status,
            {"NONE", "POSSIBLE", "CONFOUNDED", "UNASSESSED"},
            _status(context["study_confounding_status"]),
        ),
        (
            "comparisonValidityStatus",
            comparison_validity_status,
            {"VALID", "NEEDS_REVIEW", "INVALID", "EXCLUDED"},
            _status(context["validity_status"]),
        ),
        (
            "comparisonConfoundingStatus",
            comparison_confounding_status,
            {"NONE", "POSSIBLE", "CONFOUNDED", "UNASSESSED"},
            _status(context["comparison_confounding_status"]),
        ),
    )
    normalized: dict[str, str] = {}
    changed = False
    for name, supplied, allowed, current in specifications:
        if supplied is None:
            normalized[name] = current
            continue
        value = _status(supplied)
        if value not in allowed:
            raise ReviewGateError(
                "ASSESSMENT_STATUS_INVALID",
                f"{name} must be one of {sorted(allowed)}.",
            )
        normalized[name] = value
        changed = changed or value != current

    current_matching_basis = _text(context["comparison_matching_basis"])
    if matching_basis is None:
        normalized_matching_basis = current_matching_basis
    else:
        normalized_matching_basis = _text(matching_basis)
        if not normalized_matching_basis:
            raise ReviewGateError(
                "MATCHING_BASIS_REQUIRED",
                "An explicitly supplied matching basis cannot be empty.",
            )
        changed = changed or normalized_matching_basis != current_matching_basis
    normalized["matchingBasis"] = normalized_matching_basis

    if changed:
        connection.execute(
            """
            UPDATE knowledge_studies
            SET comparability_status=?, confounding_status=?
            WHERE study_id=?
            """,
            (
                normalized["studyComparabilityStatus"],
                normalized["studyConfoundingStatus"],
                int(context["study_id"]),
            ),
        )
        connection.execute(
            """
            UPDATE knowledge_comparisons
            SET validity_status=?, confounding_status=?, matching_basis=?
            WHERE comparison_id=?
            """,
            (
                normalized["comparisonValidityStatus"],
                normalized["comparisonConfoundingStatus"],
                normalized["matchingBasis"],
                int(context["comparison_id"]),
            ),
        )
    return normalized, changed


def _disable_effects(
    connection: sqlite3.Connection,
    comparison_id: int,
    verification_status: str,
) -> None:
    connection.execute(
        """
        UPDATE knowledge_effects
        SET aggregation_eligible=0, verification_status=?
        WHERE comparison_id=?
        """,
        (verification_status, comparison_id),
    )


def _apply_nonapproval(
    connection: sqlite3.Connection,
    context: dict[str, Any],
    decision: str,
    reason: str,
) -> None:
    comparison_id = int(context["comparison_id"])
    if decision == "REJECT":
        _disable_effects(connection, comparison_id, "INVALID")
        connection.execute(
            """
            UPDATE knowledge_comparisons
            SET validity_status='INVALID', aggregation_eligible=0,
                verification_status='VERIFIED', exclusion_reason=?
            WHERE comparison_id=?
            """,
            (reason, comparison_id),
        )
    elif decision == "EXCLUDE":
        _disable_effects(connection, comparison_id, "EXCLUDED")
        connection.execute(
            """
            UPDATE knowledge_comparisons
            SET validity_status='EXCLUDED', aggregation_eligible=0,
                verification_status='VERIFIED', exclusion_reason=?
            WHERE comparison_id=?
            """,
            (reason, comparison_id),
        )
    else:
        _disable_effects(connection, comparison_id, "NEEDS_REVIEW")
        connection.execute(
            """
            UPDATE knowledge_comparisons
            SET validity_status='NEEDS_REVIEW', aggregation_eligible=0,
                verification_status='NEEDS_REVIEW', exclusion_reason=?
            WHERE comparison_id=?
            """,
            (reason, comparison_id),
        )


def _apply_approval(
    connection: sqlite3.Connection,
    context: dict[str, Any],
    plan: dict[str, Any],
    *,
    reviewer: str,
    reason: str,
    decision_uid: str,
) -> list[str]:
    if not plan["ready"]:
        first = plan["blockers"][0]
        raise ReviewGateError(str(first["code"]), str(first["message"]))

    connection.execute(
        """
        UPDATE workbook_analyses
        SET verification_status='VERIFIED'
        WHERE workbook_analysis_id=?
        """,
        (int(context["workbook_analysis_id"]),),
    )
    connection.execute(
        """
        UPDATE knowledge_studies
        SET verification_status='VERIFIED'
        WHERE study_id=?
        """,
        (int(context["study_id"]),),
    )
    connection.execute(
        """
        UPDATE knowledge_comparisons
        SET verification_status='VERIFIED', validity_status='VALID',
            confounding_status='NONE', aggregation_eligible=1,
            exclusion_reason=''
        WHERE comparison_id=?
        """,
        (int(context["comparison_id"]),),
    )

    used_observation_ids = {
        int(effect["comparedObservation"]["observation_id"])
        for effect in plan["effects"]
    } | {
        int(effect["controlObservation"]["observation_id"])
        for effect in plan["effects"]
    }
    used_outcome_ids = {
        int(effect["outcome"]["outcome_id"]) for effect in plan["effects"]
    }
    connection.executemany(
        """
        UPDATE knowledge_observations
        SET verification_status='VERIFIED'
        WHERE observation_id=?
        """,
        [(observation_id,) for observation_id in sorted(used_observation_ids)],
    )
    connection.executemany(
        """
        UPDATE knowledge_outcomes
        SET verification_status='VERIFIED'
        WHERE outcome_id=?
        """,
        [(outcome_id,) for outcome_id in sorted(used_outcome_ids)],
    )

    planned_identities = {
        (
            int(effect["outcome"]["outcome_id"]),
            str(effect["effectType"]),
            str(effect["formulaVersion"]),
        )
        for effect in plan["effects"]
    }
    existing = _rows(
        connection,
        """
        SELECT effect_id, effect_uid, outcome_id, effect_type, formula_version
        FROM knowledge_effects
        WHERE comparison_id=?
        """,
        (int(context["comparison_id"]),),
    )
    for effect in existing:
        identity = (
            int(effect["outcome_id"]),
            str(effect["effect_type"]),
            str(effect["formula_version"]),
        )
        if identity in planned_identities:
            continue
        connection.execute(
            """
            UPDATE knowledge_effects
            SET aggregation_eligible=0, verification_status='EXCLUDED'
            WHERE effect_id=?
            """,
            (int(effect["effect_id"]),),
        )

    effect_public_ids: list[str] = []
    for effect in plan["effects"]:
        outcome_uid = str(effect["outcome"]["outcome_uid"])
        effect_uid = stable_uid(
            "effect",
            str(context["comparison_uid"]),
            outcome_uid,
            str(effect["effectType"]),
            str(effect["formulaVersion"]),
        )
        effect_public_id = public_id("EFF", effect_uid)
        conflicting = _one(
            connection,
            """
            SELECT effect_id, effect_uid
            FROM knowledge_effects
            WHERE comparison_id=? AND outcome_id=? AND effect_type=?
              AND formula_version=?
            """,
            (
                int(context["comparison_id"]),
                int(effect["outcome"]["outcome_id"]),
                str(effect["effectType"]),
                str(effect["formulaVersion"]),
            ),
        )
        if conflicting and str(conflicting["effect_uid"]) != effect_uid:
            connection.execute(
                """
                DELETE FROM entity_evidence_links
                WHERE entity_type='EFFECT' AND entity_uid=?
                """,
                (str(conflicting["effect_uid"]),),
            )
            connection.execute(
                "DELETE FROM knowledge_effects WHERE effect_id=?",
                (int(conflicting["effect_id"]),),
            )

        details = {
            "contractVersion": REVIEW_CONTRACT_VERSION,
            "decisionUid": decision_uid,
            "reviewer": reviewer,
            "reason": reason,
            "imagesAnalyzed": False,
            "comparedObservationUid": str(
                effect["comparedObservation"]["observation_uid"]
            ),
            "controlObservationUid": str(
                effect["controlObservation"]["observation_uid"]
            ),
        }
        calculation_text = (
            f"{effect['effectType']} calculated by "
            f"{effect['formulaVersion']} from compared observation "
            f"{effect['comparedObservation']['observation_uid']} and control "
            f"observation {effect['controlObservation']['observation_uid']}."
        )
        unit = str(effect.get("unit") or "")
        connection.execute(
            """
            INSERT INTO knowledge_effects(
                effect_uid, public_effect_id, comparison_id, outcome_id,
                effect_type, estimate, unit_id, original_unit,
                ci_lower, ci_upper, formula_version, calculation_text,
                direction, aggregation_eligible, verification_status,
                details_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, 1,
                      'VERIFIED', ?)
            ON CONFLICT(effect_uid) DO UPDATE SET
                comparison_id=excluded.comparison_id,
                outcome_id=excluded.outcome_id,
                effect_type=excluded.effect_type,
                estimate=excluded.estimate,
                unit_id=excluded.unit_id,
                original_unit=excluded.original_unit,
                ci_lower=NULL,
                ci_upper=NULL,
                formula_version=excluded.formula_version,
                calculation_text=excluded.calculation_text,
                direction=excluded.direction,
                aggregation_eligible=1,
                verification_status='VERIFIED',
                details_json=excluded.details_json
            """,
            (
                effect_uid,
                effect_public_id,
                int(context["comparison_id"]),
                int(effect["outcome"]["outcome_id"]),
                str(effect["effectType"]),
                float(effect["estimate"]),
                resolve_unit_id(connection, unit),
                unit,
                str(effect["formulaVersion"]),
                calculation_text,
                str(effect["direction"]),
                _json(details),
            ),
        )
        connection.execute(
            """
            DELETE FROM entity_evidence_links
            WHERE entity_type='EFFECT' AND entity_uid=?
            """,
            (effect_uid,),
        )
        evidence_by_id = {
            int(item["evidence_id"]): item for item in effect["evidence"]
        }
        connection.executemany(
            """
            INSERT INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id,
                evidence_role, claim_scope
            ) VALUES ('EFFECT', ?, ?, 'REVIEW_APPROVAL', ?)
            """,
            [
                (
                    effect_uid,
                    evidence_id,
                    f"{REVIEW_CONTRACT_VERSION}:{decision_uid}",
                )
                for evidence_id in sorted(evidence_by_id)
            ],
        )
        effect_public_ids.append(effect_public_id)
    return sorted(effect_public_ids)


def decide_comparison(
    connection: sqlite3.Connection,
    public_comparison_id: str,
    *,
    decision: str,
    reviewer: str,
    reason: str,
    decided_at: str | None = None,
    study_comparability_status: str | None = None,
    study_confounding_status: str | None = None,
    comparison_validity_status: str | None = None,
    comparison_confounding_status: str | None = None,
    matching_basis: str | None = None,
) -> dict[str, Any]:
    """Apply an explicit human decision to one public CMP identifier.

    No path auto-approves.  APPROVE first reconstructs a deterministic,
    source-verified calculation plan and only then changes eligibility.
    """

    decision, reviewer, reason = _validate_decision_inputs(
        decision,
        reviewer,
        reason,
    )
    with _atomic(connection):
        context = _comparison_context(connection, public_comparison_id)
        source_blockers = _source_blockers(connection, context)
        if source_blockers:
            first = source_blockers[0]
            raise ReviewGateError(str(first["code"]), str(first["message"]))
        supplied_assessment = any(
            value is not None
            for value in (
                study_comparability_status,
                study_confounding_status,
                comparison_validity_status,
                comparison_confounding_status,
                matching_basis,
            )
        )
        if supplied_assessment and decision != "APPROVE":
            raise ReviewGateError(
                "ASSESSMENT_ONLY_WITH_APPROVAL",
                "Explicit assessment fields are accepted only with APPROVE.",
            )
        assessment, assessment_changed = _apply_explicit_assessment(
            connection,
            context,
            study_comparability_status=study_comparability_status,
            study_confounding_status=study_confounding_status,
            comparison_validity_status=comparison_validity_status,
            comparison_confounding_status=comparison_confounding_status,
            matching_basis=matching_basis,
        )
        if assessment_changed:
            context = _comparison_context(
                connection,
                public_comparison_id,
            )
        latest = _latest_decision(connection, str(context["comparison_uid"]))
        repeated = bool(
            latest
            and str(latest["decision"]) == decision
            and str(latest["reviewer"]) == reviewer
            and str(latest["reason"]) == reason
            and not assessment_changed
        )
        timestamp = (
            str(latest["decided_at"])
            if repeated
            else _text(decided_at) or _utc_now_iso()
        )
        decision_uid = (
            str(latest["decision_uid"])
            if repeated
            else stable_uid(
                "review-decision",
                str(context["comparison_uid"]),
                decision,
                reviewer,
                reason,
                _json(assessment),
                timestamp,
            )
        )
        effect_public_ids: list[str] = []
        if decision == "APPROVE":
            plan = _approval_plan(connection, context)
            effect_public_ids = _apply_approval(
                connection,
                context,
                plan,
                reviewer=reviewer,
                reason=reason,
                decision_uid=decision_uid,
            )
        else:
            _apply_nonapproval(
                connection,
                context,
                decision,
                reason,
            )
        if not repeated:
            connection.execute(
                """
                INSERT INTO review_decisions(
                    decision_uid, entity_type, entity_uid, decision,
                    reason, reviewer, decided_at, supersedes_decision_uid
                ) VALUES (?, 'COMPARISON', ?, ?, ?, ?, ?, ?)
                """,
                (
                    decision_uid,
                    str(context["comparison_uid"]),
                    decision,
                    reason,
                    reviewer,
                    timestamp,
                    str(latest["decision_uid"]) if latest else "",
                ),
            )
        current = _comparison_context(connection, public_comparison_id)
        return {
            "schemaVersion": REVIEW_CONTRACT_VERSION,
            "imagesAnalyzed": False,
            "publicComparisonId": str(current["public_comparison_id"]),
            "publicDataId": str(current["public_data_id"]),
            "decisionUid": decision_uid,
            "decision": decision,
            "reviewer": reviewer,
            "reason": reason,
            "decidedAt": timestamp,
            "idempotent": repeated,
            "comparisonVerificationStatus": str(
                current["comparison_verification_status"]
            ),
            "comparisonValidityStatus": str(current["validity_status"]),
            "comparisonAggregationEligible": bool(
                current["comparison_aggregation_eligible"]
            ),
            "assessment": assessment,
            "effectPublicIds": effect_public_ids,
        }


__all__ = [
    "REVIEW_CONTRACT_VERSION",
    "ReviewGateError",
    "decide_comparison",
    "get_review_detail",
    "list_review_queue",
]
