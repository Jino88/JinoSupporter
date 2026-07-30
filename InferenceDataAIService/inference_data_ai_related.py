"""Deterministic duplicate and related-study discovery for the canonical DB.

The module deliberately separates two different signals:

* exact duplicates are current source revisions with the same non-empty
  SHA-256 content hash at another source path;
* related studies are ranked by transparent lexical overlap of canonical,
  human/seed alias, original factor/outcome/context, and study-title terms.

Neither signal is evidence of an effect, a relationship, or causality. Images
are never extracted or analyzed.
"""

from __future__ import annotations

import json
import re
import sqlite3
import unicodedata
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator


RELATED_STUDIES_SCHEMA_VERSION = "canonical-related-studies-v2"
TOKEN_PATTERN = re.compile(r"\w+(?:[+./%:-]\w+)*", re.UNICODE)
CATEGORY_WEIGHTS = {
    "factor": 0.35,
    "outcome": 0.35,
    "context": 0.20,
    "studyTitle": 0.10,
}
REQUIRED_TABLES = {
    "source_documents",
    "source_revisions",
    "workbook_analyses",
    "knowledge_studies",
    "knowledge_study_contexts",
    "knowledge_factors",
    "knowledge_outcomes",
    "knowledge_concepts",
    "knowledge_concept_aliases",
}


class RelatedStudyError(RuntimeError):
    """Raised when a deterministic related-study report cannot be built."""


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return re.sub(r"\s+", " ", text).strip()


def unicode_tokens(value: object) -> set[str]:
    """Return open-domain Unicode terms without a product vocabulary."""

    return {
        token
        for token in TOKEN_PATTERN.findall(normalize_text(value))
        if token
    }


def _rows(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> list[dict[str, Any]]:
    cursor = connection.execute(sql, tuple(parameters))
    columns = [str(item[0]) for item in cursor.description or ()]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def _validate_schema(connection: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )
    }
    missing = sorted(REQUIRED_TABLES - tables)
    if missing:
        raise RelatedStudyError(
            "Canonical knowledge DB is missing required tables: "
            + ", ".join(missing)
        )


@contextmanager
def connect_knowledge_readonly(
    database_path: str | Path,
) -> Iterator[sqlite3.Connection]:
    """Open an existing canonical SQLite database in enforced read-only mode."""

    path = Path(database_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    connection = sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True)
    try:
        connection.execute("PRAGMA query_only=ON")
        yield connection
    finally:
        connection.close()


def _revision_rows(
    connection: sqlite3.Connection,
    where_sql: str,
    parameters: Iterable[Any],
) -> list[dict[str, Any]]:
    return _rows(
        connection,
        f"""
        SELECT
            r.revision_id, r.revision_uid, r.content_sha256,
            r.source_fingerprint, r.capture_contract, r.capture_status,
            r.is_current, d.document_id, d.document_uid, d.dataset,
            d.source_path, d.original_file_name, d.lifecycle_status
        FROM source_revisions r
        JOIN source_documents d ON d.document_id=r.document_id
        WHERE {where_sql}
        ORDER BY r.revision_id
        """,
        parameters,
    )


def _study_rows_for_revision(
    connection: sqlite3.Connection,
    revision_id: int,
) -> list[dict[str, Any]]:
    return _rows(
        connection,
        """
        SELECT
            s.study_id, s.study_uid, s.public_data_id, s.title,
            s.verification_status, s.comparability_status,
            s.confounding_status, a.public_analysis_id,
            a.verification_status AS analysis_verification_status
        FROM knowledge_studies s
        JOIN workbook_analyses a
          ON a.workbook_analysis_id=s.workbook_analysis_id
        WHERE a.revision_id=?
        ORDER BY s.public_data_id, s.study_uid
        """,
        (revision_id,),
    )


def _resolve_target(
    connection: sqlite3.Connection,
    target_identifier: str,
) -> tuple[str, dict[str, Any], list[dict[str, Any]]]:
    identifier = str(target_identifier or "").strip()
    if not identifier:
        raise ValueError("target_identifier must not be empty")

    data_rows = _rows(
        connection,
        """
        SELECT
            s.study_id, s.study_uid, s.public_data_id, s.title,
            s.verification_status, s.comparability_status,
            s.confounding_status, a.public_analysis_id,
            a.verification_status AS analysis_verification_status,
            a.revision_id
        FROM knowledge_studies s
        JOIN workbook_analyses a
          ON a.workbook_analysis_id=s.workbook_analysis_id
        WHERE LOWER(s.public_data_id)=LOWER(?)
        ORDER BY s.study_id
        """,
        (identifier,),
    )
    revision_rows = _revision_rows(
        connection,
        "LOWER(r.revision_uid)=LOWER(?)",
        (identifier,),
    )
    if data_rows and revision_rows:
        raise RelatedStudyError(
            "Target identifier is ambiguous between a DATA ID and revision UID."
        )
    if len(data_rows) > 1 or len(revision_rows) > 1:
        raise RelatedStudyError(
            "Target identifier has multiple case-insensitive canonical matches."
        )
    if data_rows:
        study = data_rows[0]
        revision = _revision_rows(
            connection,
            "r.revision_id=?",
            (int(study["revision_id"]),),
        )
        if len(revision) != 1:
            raise RelatedStudyError("Target DATA ID has no canonical source revision.")
        study.pop("revision_id", None)
        return "PUBLIC_DATA_ID", revision[0], [study]
    if revision_rows:
        revision = revision_rows[0]
        return (
            "REVISION_UID",
            revision,
            _study_rows_for_revision(connection, int(revision["revision_id"])),
        )
    raise RelatedStudyError(
        f"Unknown public DATA ID or revision UID: {identifier}"
    )


def _source_payload(revision: dict[str, Any]) -> dict[str, Any]:
    return {
        "documentUid": str(revision["document_uid"]),
        "revisionUid": str(revision["revision_uid"]),
        "dataset": str(revision["dataset"]),
        "sourcePath": str(revision["source_path"]),
        "originalFileName": str(revision["original_file_name"]),
        "contentSha256": str(revision["content_sha256"] or ""),
        "sourceFingerprint": str(revision["source_fingerprint"]),
        "captureContract": str(revision["capture_contract"]),
        "captureStatus": str(revision["capture_status"]),
        "isCurrent": bool(revision["is_current"]),
        "documentLifecycleStatus": str(revision["lifecycle_status"]),
    }


def _study_payload(study: dict[str, Any]) -> dict[str, Any]:
    return {
        "publicDataId": str(study["public_data_id"]),
        "studyUid": str(study["study_uid"]),
        "title": str(study["title"]),
        "publicAnalysisId": str(study["public_analysis_id"]),
        "verificationStatus": str(study["verification_status"]),
        "analysisVerificationStatus": str(
            study["analysis_verification_status"]
        ),
        "comparabilityStatus": str(study["comparability_status"]),
        "confoundingStatus": str(study["confounding_status"]),
    }


def _current_study_rows(
    connection: sqlite3.Connection,
) -> list[dict[str, Any]]:
    return _rows(
        connection,
        """
        SELECT
            s.study_id, s.study_uid, s.public_data_id, s.title,
            s.verification_status, s.comparability_status,
            s.confounding_status, a.public_analysis_id,
            a.verification_status AS analysis_verification_status,
            r.revision_id, r.revision_uid, r.content_sha256,
            r.source_fingerprint, r.capture_contract, r.capture_status,
            r.is_current, d.document_id, d.document_uid, d.dataset,
            d.source_path, d.original_file_name, d.lifecycle_status
        FROM knowledge_studies s
        JOIN workbook_analyses a
          ON a.workbook_analysis_id=s.workbook_analysis_id
        JOIN source_revisions r ON r.revision_id=a.revision_id
        JOIN source_documents d ON d.document_id=a.document_id
        WHERE r.is_current=1
          AND r.capture_status<>'STALE'
          AND d.lifecycle_status='ACTIVE'
          AND a.verification_status<>'STALE'
          AND s.verification_status<>'STALE'
        ORDER BY s.public_data_id, s.study_uid
        """,
    )


def _empty_profile() -> dict[str, set[str]]:
    return {category: set() for category in CATEGORY_WEIGHTS}


def _add_field_terms(
    profile: dict[str, dict[str, set[str]]],
    study_id: int,
    category: str,
    *values: object,
) -> None:
    for value in values:
        profile[study_id][category].update(unicode_tokens(value))


def _concept_aliases(
    connection: sqlite3.Connection,
) -> dict[int, list[str]]:
    aliases: dict[int, list[str]] = defaultdict(list)
    for row in _rows(
        connection,
        """
        SELECT concept_id, alias_text
        FROM knowledge_concept_aliases
        ORDER BY concept_id, alias_id
        """,
    ):
        aliases[int(row["concept_id"])].append(str(row["alias_text"]))
    return dict(aliases)


def _study_term_profiles(
    connection: sqlite3.Connection,
    study_ids: Iterable[int],
    title_by_study: dict[int, str],
) -> dict[int, dict[str, set[str]]]:
    ordered_ids = sorted(set(int(study_id) for study_id in study_ids))
    profile = {
        study_id: _empty_profile()
        for study_id in ordered_ids
    }
    for study_id in ordered_ids:
        _add_field_terms(
            profile,
            study_id,
            "studyTitle",
            title_by_study.get(study_id, ""),
        )
    if not ordered_ids:
        return profile

    aliases = _concept_aliases(connection)
    placeholders = ",".join("?" for _ in ordered_ids)
    factor_rows = _rows(
        connection,
        f"""
        SELECT
            f.study_id, f.original_label, f.concept_id,
            c.canonical_name
        FROM knowledge_factors f
        LEFT JOIN knowledge_concepts c ON c.concept_id=f.concept_id
        WHERE f.study_id IN ({placeholders})
        ORDER BY f.study_id, f.factor_id
        """,
        ordered_ids,
    )
    for row in factor_rows:
        concept_alias_values = (
            aliases.get(int(row["concept_id"]), [])
            if row["concept_id"] is not None
            else []
        )
        _add_field_terms(
            profile,
            int(row["study_id"]),
            "factor",
            row["original_label"],
            row["canonical_name"],
            *concept_alias_values,
        )

    outcome_rows = _rows(
        connection,
        f"""
        SELECT
            o.study_id, o.original_label, o.concept_id,
            c.canonical_name
        FROM knowledge_outcomes o
        LEFT JOIN knowledge_concepts c ON c.concept_id=o.concept_id
        WHERE o.study_id IN ({placeholders})
        ORDER BY o.study_id, o.outcome_id
        """,
        ordered_ids,
    )
    for row in outcome_rows:
        concept_alias_values = (
            aliases.get(int(row["concept_id"]), [])
            if row["concept_id"] is not None
            else []
        )
        _add_field_terms(
            profile,
            int(row["study_id"]),
            "outcome",
            row["original_label"],
            row["canonical_name"],
            *concept_alias_values,
        )

    context_rows = _rows(
        connection,
        f"""
        SELECT
            x.study_id, x.original_value, x.concept_id,
            c.canonical_name
        FROM knowledge_study_contexts x
        LEFT JOIN knowledge_concepts c ON c.concept_id=x.concept_id
        WHERE x.study_id IN ({placeholders})
        ORDER BY x.study_id, x.context_id
        """,
        ordered_ids,
    )
    for row in context_rows:
        concept_alias_values = (
            aliases.get(int(row["concept_id"]), [])
            if row["concept_id"] is not None
            else []
        )
        _add_field_terms(
            profile,
            int(row["study_id"]),
            "context",
            row["original_value"],
            row["canonical_name"],
            *concept_alias_values,
        )
    return profile


def _union_profiles(
    profiles: dict[int, dict[str, set[str]]],
    study_ids: Iterable[int],
) -> dict[str, set[str]]:
    combined = _empty_profile()
    for study_id in study_ids:
        for category in CATEGORY_WEIGHTS:
            combined[category].update(profiles[int(study_id)][category])
    return combined


def _score_profile(
    target: dict[str, set[str]],
    candidate: dict[str, set[str]],
) -> tuple[float, dict[str, dict[str, Any]], list[dict[str, Any]]]:
    available_weight = sum(
        CATEGORY_WEIGHTS[category]
        for category in CATEGORY_WEIGHTS
        if target[category]
    )
    if available_weight == 0:
        return 0.0, {}, []

    score = 0.0
    components: dict[str, dict[str, Any]] = {}
    reasons: list[dict[str, Any]] = []
    for category, configured_weight in CATEGORY_WEIGHTS.items():
        target_terms = target[category]
        if not target_terms:
            continue
        candidate_terms = candidate[category]
        shared = sorted(target_terms & candidate_terms)
        union = target_terms | candidate_terms
        jaccard = len(shared) / len(union) if union else 0.0
        effective_weight = configured_weight / available_weight
        contribution = effective_weight * jaccard
        score += contribution
        components[category] = {
            "configuredWeight": configured_weight,
            "effectiveWeight": round(effective_weight, 6),
            "targetTermCount": len(target_terms),
            "candidateTermCount": len(candidate_terms),
            "sharedTermCount": len(shared),
            "unionTermCount": len(union),
            "jaccard": round(jaccard, 6),
            "weightedContribution": round(contribution, 6),
            "sharedTerms": shared,
        }
        if shared:
            reasons.append(
                {
                    "category": category,
                    "sharedTerms": shared,
                    "reason": (
                        "Shared normalized terms in the same canonical, alias, "
                        f"or original {category} field category."
                    ),
                }
            )
    return round(score, 6), components, reasons


def _exact_content_duplicates(
    connection: sqlite3.Connection,
    target_revision: dict[str, Any],
) -> tuple[list[dict[str, Any]], set[int]]:
    content_hash = str(target_revision["content_sha256"] or "").strip().lower()
    if not content_hash:
        return [], set()
    rows = _rows(
        connection,
        """
        SELECT
            r.revision_id, r.revision_uid, r.content_sha256,
            r.source_fingerprint, r.capture_contract, r.capture_status,
            r.is_current, d.document_id, d.document_uid, d.dataset,
            d.source_path, d.original_file_name, d.lifecycle_status,
            s.public_data_id
        FROM source_revisions r
        JOIN source_documents d ON d.document_id=r.document_id
        LEFT JOIN workbook_analyses a ON a.revision_id=r.revision_id
        LEFT JOIN knowledge_studies s
          ON s.workbook_analysis_id=a.workbook_analysis_id
        WHERE LOWER(r.content_sha256)=?
          AND r.content_sha256<>''
          AND r.is_current=1
          AND r.capture_status<>'STALE'
          AND d.lifecycle_status='ACTIVE'
          AND r.revision_id<>?
        ORDER BY
            LOWER(d.source_path), d.source_path, r.revision_uid,
            s.public_data_id
        """,
        (content_hash, int(target_revision["revision_id"])),
    )
    target_path = normalize_text(target_revision["source_path"])
    grouped: dict[int, dict[str, Any]] = {}
    for row in rows:
        # Same-path re-captures are lifecycle history, not a duplicate across
        # source paths.
        if normalize_text(row["source_path"]) == target_path:
            continue
        revision_id = int(row["revision_id"])
        item = grouped.get(revision_id)
        if item is None:
            item = {
                "matchKind": "EXACT_CONTENT_SHA256",
                "contentSha256": str(row["content_sha256"]),
                "source": _source_payload(row),
                "publicDataIds": [],
            }
            grouped[revision_id] = item
        public_data_id = row["public_data_id"]
        if public_data_id is not None:
            item["publicDataIds"].append(str(public_data_id))
    duplicates = list(grouped.values())
    for duplicate in duplicates:
        duplicate["publicDataIds"] = sorted(set(duplicate["publicDataIds"]))
    duplicates.sort(
        key=lambda item: (
            normalize_text(item["source"]["sourcePath"]),
            item["source"]["revisionUid"],
        )
    )
    return duplicates, set(grouped)


def build_related_studies(
    connection: sqlite3.Connection,
    target_identifier: str,
    *,
    limit: int = 25,
) -> dict[str, Any]:
    """Build a deterministic related-study report from a canonical DB.

    ``target_identifier`` may be a stable public DATA ID or a source revision
    UID. Only current, active, non-stale studies are similarity candidates.
    The supplied connection is only queried; use
    :func:`build_related_studies_from_db` to enforce SQLite read-only mode.
    """

    if isinstance(limit, bool) or not isinstance(limit, int) or limit < 1:
        raise ValueError("limit must be a positive integer")
    _validate_schema(connection)
    identifier_type, target_revision, target_studies = _resolve_target(
        connection,
        target_identifier,
    )
    duplicates, duplicate_revision_ids = _exact_content_duplicates(
        connection,
        target_revision,
    )

    current_studies = _current_study_rows(connection)
    target_study_ids = {
        int(study["study_id"])
        for study in target_studies
    }
    profile_rows = {
        int(study["study_id"]): study
        for study in [*target_studies, *current_studies]
    }
    title_by_study = {
        study_id: str(study["title"])
        for study_id, study in profile_rows.items()
    }
    profiles = _study_term_profiles(
        connection,
        profile_rows,
        title_by_study,
    )
    target_profile = _union_profiles(profiles, target_study_ids)

    related: list[dict[str, Any]] = []
    zero_overlap_count = 0
    exact_duplicate_study_count = 0
    for candidate in current_studies:
        study_id = int(candidate["study_id"])
        if study_id in target_study_ids:
            continue
        if int(candidate["revision_id"]) in duplicate_revision_ids:
            exact_duplicate_study_count += 1
            continue
        score, components, reasons = _score_profile(
            target_profile,
            profiles[study_id],
        )
        if score <= 0:
            zero_overlap_count += 1
            continue
        related.append(
            {
                **_study_payload(candidate),
                "source": _source_payload(candidate),
                "similarityScore": score,
                "scoreComponents": components,
                "sharedTermReasons": reasons,
                "similarityIsRelationshipEvidence": False,
                "similarityIsCausalEvidence": False,
            }
        )
    related.sort(
        key=lambda item: (
            -float(item["similarityScore"]),
            normalize_text(item["publicDataId"]),
            item["studyUid"],
        )
    )
    candidate_count = len(related)
    related = related[:limit]

    resolved_identifier = (
        str(target_studies[0]["public_data_id"])
        if identifier_type == "PUBLIC_DATA_ID"
        else str(target_revision["revision_uid"])
    )
    return {
        "schemaVersion": RELATED_STUDIES_SCHEMA_VERSION,
        "targetIdentifier": resolved_identifier,
        "targetIdentifierType": identifier_type,
        "target": {
            "source": _source_payload(target_revision),
            "studies": [_study_payload(study) for study in target_studies],
            "termProfile": {
                category: sorted(target_profile[category])
                for category in CATEGORY_WEIGHTS
            },
        },
        "exactContentDuplicates": duplicates,
        "relatedStudies": related,
        "summary": {
            "targetStudyCount": len(target_studies),
            "exactContentDuplicateSourceCount": len(duplicates),
            "exactDuplicateStudiesExcludedFromSimilarityRanking": (
                exact_duplicate_study_count
            ),
            "positiveOverlapStudyCount": candidate_count,
            "returnedRelatedStudyCount": len(related),
            "zeroOverlapStudyCount": zero_overlap_count,
            "similarityAvailable": any(target_profile.values()),
        },
        "scoring": {
            "method": "CATEGORY_WEIGHTED_JACCARD",
            "categoryWeights": dict(CATEGORY_WEIGHTS),
            "termNormalization": "Unicode NFKC, case-folding, whitespace normalization",
            "weightNormalization": (
                "Configured weights are renormalized across target categories "
                "that contain at least one term."
            ),
            "fieldScope": {
                "factor": [
                    "originalLabel",
                    "canonicalName",
                    "conceptAliases",
                ],
                "outcome": [
                    "originalLabel",
                    "canonicalName",
                    "conceptAliases",
                ],
                "context": [
                    "originalValue",
                    "canonicalName",
                    "conceptAliases",
                ],
                "studyTitle": ["title"],
            },
            "zeroOverlapStudiesReturned": False,
            "exactContentDuplicatesExcludedFromSimilarityRanking": True,
        },
        "safety": {
            "similarityIsEvidence": False,
            "similarityIsCausality": False,
            "message": (
                "Lexical similarity is a discovery aid only. Inspect verified "
                "control/comparison evidence before drawing a relationship or "
                "causal conclusion."
            ),
            "imagesAnalyzed": False,
        },
    }


def build_related_studies_from_db(
    database_path: str | Path,
    target_identifier: str,
    *,
    limit: int = 25,
) -> dict[str, Any]:
    """Open ``database_path`` read-only and build a related-study report."""

    with connect_knowledge_readonly(database_path) as connection:
        return build_related_studies(
            connection,
            target_identifier,
            limit=limit,
        )


def related_studies_json_bytes(report: dict[str, Any]) -> bytes:
    """Serialize a report canonically for stable files and regression tests."""

    return (
        json.dumps(
            report,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        )
        + "\n"
    ).encode("utf-8")
