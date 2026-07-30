"""Domain-neutral natural-language retrieval over canonical knowledge SQLite.

The retriever returns every study that has a textual or canonical-concept
match.  It never treats retrieval relevance as answer eligibility: only
verified, valid, unconfounded, aggregation-eligible comparisons and effects
with verified source evidence are admitted to ``answerEligibleEffects``.
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


EVIDENCE_PACK_SCHEMA_VERSION = "canonical-evidence-pack-v1"
TOKEN_PATTERN = re.compile(r"\w+(?:[+./%:-]\w+)*", re.UNICODE)
PUBLIC_ID_PATTERN = re.compile(
    r"\b(?:data|cmp|eff|evd)-[a-z0-9][a-z0-9-]*\b",
    re.IGNORECASE,
)
# Domain-neutral question/request words are removed only when at least one
# substantive token remains. They are not product, factor, or outcome rules.
QUERY_FUNCTION_WORDS = {
    "all",
    "any",
    "data",
    "evidence",
    "find",
    "for",
    "me",
    "related",
    "relationship",
    "review",
    "show",
    "summarize",
    "summary",
    "tell",
    "the",
    "what",
    "검토",
    "관계",
    "관련",
    "대해",
    "대한",
    "데이터",
    "모두",
    "보여줘",
    "보여주세요",
    "알려줘",
    "알려주세요",
    "어떤",
    "요약",
    "요약해줘",
    "요약해주세요",
    "자료",
    "전부",
    "무엇",
    "무엇이며",
    "중",
    "가장",
    "낮은",
    "높은",
    "조건",
    "조건은",
    "표본",
    "수",
    "수와",
    "각",
    "세부",
    "함께",
    "비교",
    "비교해줘",
    "비교해주세요",
}
KOREAN_PARTICLE_SUFFIXES = (
    "으로",
    "에서",
    "에게",
    "까지",
    "부터",
    "처럼",
    "보다",
    "은",
    "는",
    "이",
    "가",
    "을",
    "를",
    "과",
    "와",
    "의",
    "에",
    "로",
    "도",
    "만",
)
REQUIRED_TABLES = {
    "source_documents",
    "source_revisions",
    "workbook_analyses",
    "knowledge_concepts",
    "knowledge_concept_aliases",
    "knowledge_studies",
    "knowledge_study_contexts",
    "knowledge_factors",
    "knowledge_arms",
    "knowledge_arm_factor_values",
    "knowledge_outcomes",
    "knowledge_observations",
    "knowledge_measurement_series",
    "knowledge_measurement_points",
    "knowledge_comparisons",
    "knowledge_effects",
    "knowledge_units",
    "evidence_items",
    "entity_evidence_links",
}


class EvidenceQueryError(RuntimeError):
    """Raised when a trustworthy evidence pack cannot be assembled."""


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return re.sub(r"\s+", " ", text).strip()


def unicode_tokens(value: object) -> list[str]:
    """Tokenize arbitrary Unicode without a domain vocabulary or whitelist."""

    return [token for token in TOKEN_PATTERN.findall(normalize_text(value)) if token]


def _unique_tokens(value: object) -> list[str]:
    return list(dict.fromkeys(unicode_tokens(value)))


def _search_tokens(query_tokens: list[str]) -> list[str]:
    substantive: list[str] = []
    for token in query_tokens:
        candidates = [token]
        for suffix in KOREAN_PARTICLE_SUFFIXES:
            if not token.endswith(suffix):
                continue
            stem = token[: -len(suffix)]
            if len(stem) >= 2:
                candidates = [stem]
                break
        if any(candidate in QUERY_FUNCTION_WORDS for candidate in candidates):
            continue
        for candidate in candidates:
            if (
                candidate not in QUERY_FUNCTION_WORDS
                and candidate not in substantive
            ):
                substantive.append(candidate)
    return substantive or query_tokens


def _json_value(value: str | None, fallback: Any) -> Any:
    if value is None or value == "":
        return fallback
    try:
        return json.loads(value)
    except json.JSONDecodeError as error:
        raise EvidenceQueryError("Canonical knowledge DB contains invalid JSON.") from error


def _rows(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> list[dict[str, Any]]:
    cursor = connection.execute(sql, tuple(parameters))
    columns = [description[0] for description in cursor.description or ()]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def _first(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> dict[str, Any] | None:
    rows = _rows(connection, sql, parameters)
    return rows[0] if rows else None


def _validate_schema(connection: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )
    }
    missing = sorted(REQUIRED_TABLES - tables)
    if missing:
        raise EvidenceQueryError(
            "Canonical knowledge DB is missing required tables: "
            + ", ".join(missing)
        )


@contextmanager
def connect_knowledge_readonly(database_path: str | Path) -> Iterator[sqlite3.Connection]:
    path = Path(database_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    connection = sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True)
    try:
        connection.execute("PRAGMA query_only=ON")
        yield connection
    finally:
        connection.close()


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


def _study_base_rows(connection: sqlite3.Connection) -> list[dict[str, Any]]:
    return _rows(
        connection,
        """
        SELECT
            s.study_id, s.study_uid, s.public_data_id, s.study_key,
            s.title, s.purpose, s.hypothesis, s.objective, s.design_type,
            s.comparison_basis, s.analysis_status, s.verification_status,
            s.comparability_status, s.confounding_status, s.decision_text,
            s.summary_text, s.limitations_json,
            a.workbook_analysis_id, a.analysis_uid,
            a.public_analysis_id, a.title AS analysis_title,
            a.purpose AS analysis_purpose, a.scope_text,
            a.decision_text AS analysis_decision_text,
            a.consolidated_summary, a.verification_status AS analysis_verification_status,
            r.revision_id, r.revision_uid, r.content_sha256,
            r.source_fingerprint, r.capture_contract,
            COALESCE(NULLIF(r.source_content_status, ''), r.capture_status)
                AS capture_status,
            r.is_current, d.document_id, d.dataset, d.source_path,
            d.original_file_name
        FROM knowledge_studies s
        JOIN workbook_analyses a
          ON a.workbook_analysis_id=s.workbook_analysis_id
        JOIN source_revisions r ON r.revision_id=a.revision_id
        JOIN source_documents d ON d.document_id=a.document_id
        WHERE a.verification_status<>'STALE'
        ORDER BY s.public_data_id
        """,
    )


def _terminal_source_rows(
    connection: sqlite3.Connection,
) -> list[dict[str, Any]]:
    """Return current source analyses that intentionally have no Study."""

    return _rows(
        connection,
        """
        SELECT
            a.public_analysis_id, a.analysis_uid, a.title,
            a.purpose, a.scope_text, a.analysis_status,
            a.verification_status, a.decision_text,
            a.consolidated_summary, a.limitations_json,
            r.revision_id, r.revision_uid, r.content_sha256,
            r.source_fingerprint, r.capture_contract,
            COALESCE(NULLIF(r.source_content_status, ''), r.capture_status)
                AS source_content_status,
            d.document_id, d.dataset, d.source_path,
            d.original_file_name
        FROM workbook_analyses a
        JOIN source_revisions r ON r.revision_id=a.revision_id
        JOIN source_documents d ON d.document_id=a.document_id
        WHERE r.is_current=1
          AND d.lifecycle_status='ACTIVE'
          AND NOT EXISTS (
              SELECT 1
              FROM knowledge_studies s
              WHERE s.workbook_analysis_id=a.workbook_analysis_id
          )
          AND (
              a.analysis_status IN ('EMPTY_WORKBOOK', 'NO_TABULAR_EVIDENCE')
              OR COALESCE(NULLIF(r.source_content_status, ''), r.capture_status)
                 IN ('EMPTY_WORKBOOK', 'NO_TABULAR_EVIDENCE')
          )
        ORDER BY a.public_analysis_id, r.revision_uid
        """,
    )


def _entity_study_ids(
    connection: sqlite3.Connection,
    entity_type: str,
    entity_uid: str,
) -> set[int]:
    entity_type = entity_type.upper()
    queries = {
        "WORKBOOK_ANALYSIS": """
            SELECT s.study_id
            FROM knowledge_studies s
            JOIN workbook_analyses a
              ON a.workbook_analysis_id=s.workbook_analysis_id
            WHERE a.analysis_uid=?
        """,
        "STUDY": "SELECT study_id FROM knowledge_studies WHERE study_uid=?",
        "CONTEXT": """
            SELECT study_id FROM knowledge_study_contexts WHERE context_uid=?
        """,
        "FACTOR": "SELECT study_id FROM knowledge_factors WHERE factor_uid=?",
        "ARM": "SELECT study_id FROM knowledge_arms WHERE arm_uid=?",
        "OUTCOME": "SELECT study_id FROM knowledge_outcomes WHERE outcome_uid=?",
        "OBSERVATION": """
            SELECT o.study_id
            FROM knowledge_observations v
            JOIN knowledge_outcomes o ON o.outcome_id=v.outcome_id
            WHERE v.observation_uid=?
        """,
        "COMPARISON": """
            SELECT study_id FROM knowledge_comparisons WHERE comparison_uid=?
        """,
        "EFFECT": """
            SELECT c.study_id
            FROM knowledge_effects e
            JOIN knowledge_comparisons c ON c.comparison_id=e.comparison_id
            WHERE e.effect_uid=?
        """,
    }
    sql = queries.get(entity_type)
    if sql is None:
        return set()
    return {
        int(row["study_id"])
        for row in _rows(connection, sql, (entity_uid,))
    }


def _direct_public_id_matches(
    connection: sqlite3.Connection,
    question: str,
) -> dict[int, list[str]]:
    """Resolve stable public IDs without relying on ordinary text relevance."""

    identifiers = list(
        dict.fromkeys(
            match.group(0).upper()
            for match in PUBLIC_ID_PATTERN.finditer(normalize_text(question))
        )
    )
    matches: dict[int, set[str]] = defaultdict(set)
    for identifier in identifiers:
        prefix = identifier.split("-", 1)[0]
        if prefix == "DATA":
            rows = _rows(
                connection,
                """
                SELECT study_id, public_data_id AS public_id
                FROM knowledge_studies
                WHERE LOWER(public_data_id)=LOWER(?)
                """,
                (identifier,),
            )
            for row in rows:
                matches[int(row["study_id"])].add(str(row["public_id"]))
        elif prefix == "CMP":
            rows = _rows(
                connection,
                """
                SELECT study_id, public_comparison_id AS public_id
                FROM knowledge_comparisons
                WHERE LOWER(public_comparison_id)=LOWER(?)
                """,
                (identifier,),
            )
            for row in rows:
                matches[int(row["study_id"])].add(str(row["public_id"]))
        elif prefix == "EFF":
            rows = _rows(
                connection,
                """
                SELECT c.study_id, e.public_effect_id AS public_id
                FROM knowledge_effects e
                JOIN knowledge_comparisons c
                  ON c.comparison_id=e.comparison_id
                WHERE LOWER(e.public_effect_id)=LOWER(?)
                """,
                (identifier,),
            )
            for row in rows:
                matches[int(row["study_id"])].add(str(row["public_id"]))
        elif prefix == "EVD":
            rows = _rows(
                connection,
                """
                SELECT
                    e.public_evidence_id AS public_id,
                    l.entity_type, l.entity_uid
                FROM evidence_items e
                JOIN entity_evidence_links l ON l.evidence_id=e.evidence_id
                WHERE LOWER(e.public_evidence_id)=LOWER(?)
                """,
                (identifier,),
            )
            for row in rows:
                for study_id in _entity_study_ids(
                    connection,
                    str(row["entity_type"]),
                    str(row["entity_uid"]),
                ):
                    matches[study_id].add(str(row["public_id"]))
    return {
        study_id: sorted(public_ids)
        for study_id, public_ids in matches.items()
    }


def _contexts(
    connection: sqlite3.Connection,
    study_id: int,
    aliases: dict[int, list[str]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in _rows(
        connection,
        """
        SELECT
            x.context_id, x.context_uid, x.context_kind, x.concept_id,
            x.original_value, x.normalized_value, x.value_number,
            x.start_value, x.end_value, x.attributes_json,
            x.verification_status, u.canonical_symbol,
            c.canonical_name, c.concept_kind
        FROM knowledge_study_contexts x
        LEFT JOIN knowledge_concepts c ON c.concept_id=x.concept_id
        LEFT JOIN knowledge_units u ON u.unit_id=x.unit_id
        WHERE x.study_id=?
        ORDER BY x.context_id
        """,
        (study_id,),
    ):
        concept_id = int(row["concept_id"]) if row["concept_id"] is not None else None
        result.append(
            {
                "contextUid": str(row["context_uid"]),
                "kind": str(row["context_kind"]),
                "originalValue": str(row["original_value"]),
                "normalizedValue": str(row["normalized_value"]),
                "valueNumber": row["value_number"],
                "unit": row["canonical_symbol"],
                "startValue": str(row["start_value"]),
                "endValue": str(row["end_value"]),
                "concept": (
                    {
                        "conceptId": concept_id,
                        "kind": str(row["concept_kind"]),
                        "canonicalName": str(row["canonical_name"]),
                        "aliases": aliases.get(concept_id, []),
                    }
                    if concept_id is not None
                    else None
                ),
                "verificationStatus": str(row["verification_status"]),
                "attributes": _json_value(row["attributes_json"], {}),
            }
        )
    return result


def _factors(
    connection: sqlite3.Connection,
    study_id: int,
    aliases: dict[int, list[str]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in _rows(
        connection,
        """
        SELECT
            f.factor_id, f.factor_uid, f.concept_id, f.factor_key,
            f.factor_domain, f.original_label, f.baseline_condition,
            f.changed_condition, f.change_direction, f.isolation_status,
            f.verification_status, f.attributes_json,
            c.canonical_name, c.concept_kind
        FROM knowledge_factors f
        LEFT JOIN knowledge_concepts c ON c.concept_id=f.concept_id
        WHERE f.study_id=?
        ORDER BY f.factor_id
        """,
        (study_id,),
    ):
        concept_id = int(row["concept_id"]) if row["concept_id"] is not None else None
        result.append(
            {
                "factorId": int(row["factor_id"]),
                "factorUid": str(row["factor_uid"]),
                "factorKey": str(row["factor_key"]),
                "domain": str(row["factor_domain"]),
                "originalLabel": str(row["original_label"]),
                "baselineCondition": str(row["baseline_condition"]),
                "changedCondition": str(row["changed_condition"]),
                "changeDirection": str(row["change_direction"]),
                "isolationStatus": str(row["isolation_status"]),
                "verificationStatus": str(row["verification_status"]),
                "concept": (
                    {
                        "conceptId": concept_id,
                        "kind": str(row["concept_kind"]),
                        "canonicalName": str(row["canonical_name"]),
                        "aliases": aliases.get(concept_id, []),
                    }
                    if concept_id is not None
                    else None
                ),
                "attributes": _json_value(row["attributes_json"], {}),
            }
        )
    return result


def _arm_factor_values(
    connection: sqlite3.Connection,
    arm_id: int,
) -> list[dict[str, Any]]:
    return [
        {
            "factorId": int(row["factor_id"]),
            "factorUid": str(row["factor_uid"]),
            "factorLabel": str(row["original_label"]),
            "originalValue": str(row["original_value"]),
            "valueNumber": row["value_number"],
            "unit": row["canonical_symbol"],
            "isBaseline": bool(row["is_baseline"]),
            "heldConstant": bool(row["held_constant"]),
        }
        for row in _rows(
            connection,
            """
            SELECT
                v.factor_id, f.factor_uid, f.original_label,
                v.original_value, v.value_number, u.canonical_symbol,
                v.is_baseline, v.held_constant
            FROM knowledge_arm_factor_values v
            JOIN knowledge_factors f ON f.factor_id=v.factor_id
            LEFT JOIN knowledge_units u ON u.unit_id=v.unit_id
            WHERE v.arm_id=?
            ORDER BY v.factor_id
            """,
            (arm_id,),
        )
    ]


def _arms(
    connection: sqlite3.Connection,
    study_id: int,
) -> list[dict[str, Any]]:
    arms: list[dict[str, Any]] = []
    for row in _rows(
        connection,
        """
        SELECT
            arm_id, arm_uid, arm_key, arm_role, label, condition_text,
            sample_size, sample_basis, matching_basis, attributes_json,
            verification_status
        FROM knowledge_arms
        WHERE study_id=?
        ORDER BY arm_id
        """,
        (study_id,),
    ):
        arm_id = int(row["arm_id"])
        arms.append(
            {
                "armId": arm_id,
                "armUid": str(row["arm_uid"]),
                "armKey": str(row["arm_key"]),
                "role": str(row["arm_role"]),
                "label": str(row["label"]),
                "condition": str(row["condition_text"]),
                "sampleSize": row["sample_size"],
                "sampleBasis": str(row["sample_basis"]),
                "matchingBasis": str(row["matching_basis"]),
                "verificationStatus": str(row["verification_status"]),
                "factorValues": _arm_factor_values(connection, arm_id),
                "attributes": _json_value(row["attributes_json"], {}),
            }
        )
    return arms


def _outcomes(
    connection: sqlite3.Connection,
    study_id: int,
    aliases: dict[int, list[str]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in _rows(
        connection,
        """
        SELECT
            o.outcome_id, o.outcome_uid, o.outcome_key, o.concept_id,
            o.original_label, o.outcome_domain, o.metric_type,
            o.original_unit, o.denominator_basis, o.favorable_direction,
            o.definition_text, o.spec_text, o.verification_status,
            o.attributes_json, u.canonical_symbol, c.canonical_name,
            c.concept_kind
        FROM knowledge_outcomes o
        LEFT JOIN knowledge_concepts c ON c.concept_id=o.concept_id
        LEFT JOIN knowledge_units u ON u.unit_id=o.unit_id
        WHERE o.study_id=?
        ORDER BY o.outcome_id
        """,
        (study_id,),
    ):
        concept_id = int(row["concept_id"]) if row["concept_id"] is not None else None
        result.append(
            {
                "outcomeId": int(row["outcome_id"]),
                "outcomeUid": str(row["outcome_uid"]),
                "outcomeKey": str(row["outcome_key"]),
                "originalLabel": str(row["original_label"]),
                "domain": str(row["outcome_domain"]),
                "metricType": str(row["metric_type"]),
                "unit": row["canonical_symbol"] or row["original_unit"],
                "originalUnit": str(row["original_unit"]),
                "denominatorBasis": str(row["denominator_basis"]),
                "favorableDirection": str(row["favorable_direction"]),
                "definition": str(row["definition_text"]),
                "specification": str(row["spec_text"]),
                "verificationStatus": str(row["verification_status"]),
                "concept": (
                    {
                        "conceptId": concept_id,
                        "kind": str(row["concept_kind"]),
                        "canonicalName": str(row["canonical_name"]),
                        "aliases": aliases.get(concept_id, []),
                    }
                    if concept_id is not None
                    else None
                ),
                "attributes": _json_value(row["attributes_json"], {}),
            }
        )
    return result


def _measurement_series(
    connection: sqlite3.Connection,
    study_id: int,
) -> list[dict[str, Any]]:
    """Return compact series metadata and deterministic point summaries.

    Raw measurement points intentionally stay in SQLite.  The evidence pack
    carries only searchable semantic metadata and aggregate descriptions so a
    wide source table cannot expand into thousands of JSON records.
    """

    series_columns = {
        str(row["name"])
        for row in _rows(
            connection,
            "PRAGMA table_info(knowledge_measurement_series)",
        )
    }
    axis_source_sql = (
        "ms.axis_source"
        if "axis_source" in series_columns
        else "'ROW_IDENTITY'"
    )
    point_columns = {
        str(row["name"])
        for row in _rows(
            connection,
            "PRAGMA table_info(knowledge_measurement_points)",
        )
    }
    replicate_role_sql = (
        "p.replicate_role"
        if "replicate_role" in point_columns
        else "'RAW'"
    )
    raw_replicate_filter = (
        "AND replicate_role='RAW'"
        if "replicate_role" in point_columns
        else ""
    )
    aggregate_replicate_filter = (
        "AND replicate_role='AGGREGATE'"
        if "replicate_role" in point_columns
        else "AND 0"
    )
    result: list[dict[str, Any]] = []
    rows = _rows(
        connection,
        f"""
        SELECT
            ms.series_id, ms.series_uid, ms.public_series_id,
            ms.series_key, ms.sheet_name, ms.header_range,
            ms.value_range, ms.row_identity_range, ms.axis_name,
            {axis_source_sql} AS axis_source,
            ms.original_axis_unit, ms.original_value_unit,
            ms.stratum_key, ms.verification_status, ms.details_json,
            axis_unit.canonical_symbol AS canonical_axis_unit,
            value_unit.canonical_symbol AS canonical_value_unit,
            o.outcome_uid, o.outcome_key, o.original_label,
            o.outcome_domain, o.metric_type,
            a.arm_uid, a.arm_key, a.arm_role, a.label AS arm_label,
            a.condition_text,
            COUNT(p.point_id) AS point_count,
            SUM(
                CASE WHEN p.point_id IS NOT NULL
                          AND {replicate_role_sql}='RAW'
                     THEN 1 ELSE 0 END
            ) AS raw_point_count,
            SUM(
                CASE WHEN p.point_id IS NOT NULL
                          AND {replicate_role_sql}='AGGREGATE'
                     THEN 1 ELSE 0 END
            ) AS aggregate_point_count,
            MIN(
                CASE WHEN {replicate_role_sql}='RAW'
                     THEN p.value_number END
            ) AS minimum_value,
            MAX(
                CASE WHEN {replicate_role_sql}='RAW'
                     THEN p.value_number END
            ) AS maximum_value,
            AVG(
                CASE WHEN {replicate_role_sql}='RAW'
                     THEN p.value_number END
            ) AS average_value,
            COUNT(
                DISTINCT CASE WHEN {replicate_role_sql}='RAW'
                              THEN p.axis_source_coordinate END
            ) AS distinct_axis_count,
            COUNT(
                DISTINCT CASE WHEN {replicate_role_sql}='RAW'
                              THEN p.replicate_source_coordinate END
            ) AS distinct_replicate_count,
            COUNT(
                DISTINCT CASE WHEN {replicate_role_sql}='AGGREGATE'
                              THEN p.replicate_source_coordinate END
            ) AS aggregate_replicate_count
        FROM knowledge_measurement_series ms
        JOIN knowledge_outcomes o ON o.outcome_id=ms.outcome_id
        JOIN knowledge_arms a ON a.arm_id=ms.arm_id
        LEFT JOIN knowledge_units axis_unit
            ON axis_unit.unit_id=ms.axis_unit_id
        LEFT JOIN knowledge_units value_unit
            ON value_unit.unit_id=ms.value_unit_id
        LEFT JOIN knowledge_measurement_points p
            ON p.series_id=ms.series_id
        WHERE ms.study_id=?
        GROUP BY ms.series_id
        ORDER BY ms.series_id
        """,
        (study_id,),
    )
    for row in rows:
        series_id = int(row["series_id"])
        replicate_keys = [
            str(item["replicate_key"])
            for item in _rows(
                connection,
                f"""
                SELECT DISTINCT replicate_key
                FROM knowledge_measurement_points
                WHERE series_id=? AND replicate_key<>''
                  {raw_replicate_filter}
                ORDER BY replicate_key
                """,
                (series_id,),
            )
        ]
        aggregate_replicate_keys = [
            str(item["replicate_key"])
            for item in _rows(
                connection,
                f"""
                SELECT DISTINCT replicate_key
                FROM knowledge_measurement_points
                WHERE series_id=? AND replicate_key<>''
                  {aggregate_replicate_filter}
                ORDER BY replicate_key
                """,
                (series_id,),
            )
        ]
        result.append(
            {
                "seriesId": series_id,
                "seriesUid": str(row["series_uid"]),
                "publicSeriesId": str(row["public_series_id"]),
                "seriesKey": str(row["series_key"]),
                "outcome": {
                    "outcomeUid": str(row["outcome_uid"]),
                    "outcomeKey": str(row["outcome_key"]),
                    "originalLabel": str(row["original_label"]),
                    "domain": str(row["outcome_domain"]),
                    "metricType": str(row["metric_type"]),
                },
                "arm": {
                    "armUid": str(row["arm_uid"]),
                    "armKey": str(row["arm_key"]),
                    "role": str(row["arm_role"]),
                    "label": str(row["arm_label"]),
                    "condition": str(row["condition_text"]),
                },
                "sheet": str(row["sheet_name"]),
                "headerRange": str(row["header_range"]),
                "valueRange": str(row["value_range"]),
                "rowIdentityRange": str(row["row_identity_range"]),
                "axisLabel": str(row["axis_name"]),
                "axisSource": str(row["axis_source"]),
                "axisUnit": (
                    row["canonical_axis_unit"]
                    or str(row["original_axis_unit"])
                ),
                "originalAxisUnit": str(row["original_axis_unit"]),
                "valueUnit": (
                    row["canonical_value_unit"]
                    or str(row["original_value_unit"])
                ),
                "originalValueUnit": str(row["original_value_unit"]),
                "stratumKey": str(row["stratum_key"]),
                "replicateKeys": replicate_keys,
                "aggregateReplicateKeys": aggregate_replicate_keys,
                "verificationStatus": str(row["verification_status"]),
                "interpretationStatus": "DESCRIPTIVE_ONLY",
                "pointSummary": {
                    "pointCount": int(row["point_count"]),
                    "rawPointCount": int(row["raw_point_count"]),
                    "aggregatePointCount": int(
                        row["aggregate_point_count"]
                    ),
                    "minimum": row["minimum_value"],
                    "maximum": row["maximum_value"],
                    "average": row["average_value"],
                    "distinctAxisCount": int(row["distinct_axis_count"]),
                    "distinctReplicateCount": int(
                        row["distinct_replicate_count"]
                    ),
                    "aggregateReplicateCount": int(
                        row["aggregate_replicate_count"]
                    ),
                },
                "details": _json_value(row["details_json"], {}),
            }
        )
    return result


def _search_fields(candidate: dict[str, Any]) -> list[tuple[str, object, float]]:
    study = candidate["study"]
    fields: list[tuple[str, object, float]] = [
        ("publicDataId", candidate["publicDataId"], 20.0),
        ("source.fileName", candidate["source"]["fileName"], 4.0),
        ("source.sourcePath", candidate["source"]["sourcePath"], 2.0),
        ("analysis.publicAnalysisId", candidate["analysis"]["publicAnalysisId"], 12.0),
        ("study.title", study["title"], 4.0),
        ("study.purpose", study["purpose"], 2.0),
        ("study.hypothesis", study["hypothesis"], 2.0),
        ("study.objective", study["objective"], 2.0),
        ("study.comparisonBasis", study["comparisonBasis"], 2.0),
        ("study.decision", study["decision"], 1.5),
        ("study.summary", study["summary"], 2.5),
    ]
    for index, context in enumerate(candidate["contexts"]):
        prefix = f"contexts[{index}]"
        fields.extend(
            [
                (f"{prefix}.kind", context["kind"], 2.0),
                (f"{prefix}.originalValue", context["originalValue"], 4.0),
                (f"{prefix}.normalizedValue", context["normalizedValue"], 3.0),
            ]
        )
        if context["concept"]:
            fields.append(
                (
                    f"{prefix}.concept.canonicalName",
                    context["concept"]["canonicalName"],
                    7.0,
                )
            )
            for alias in context["concept"]["aliases"]:
                fields.append((f"{prefix}.concept.alias", alias, 7.0))
    for index, factor in enumerate(candidate["factors"]):
        prefix = f"factors[{index}]"
        fields.extend(
            [
                (f"{prefix}.factorKey", factor["factorKey"], 3.0),
                (f"{prefix}.domain", factor["domain"], 3.0),
                (f"{prefix}.originalLabel", factor["originalLabel"], 6.0),
                (f"{prefix}.baselineCondition", factor["baselineCondition"], 4.0),
                (f"{prefix}.changedCondition", factor["changedCondition"], 4.0),
                (f"{prefix}.changeDirection", factor["changeDirection"], 2.0),
            ]
        )
        if factor["concept"]:
            fields.append(
                (
                    f"{prefix}.concept.canonicalName",
                    factor["concept"]["canonicalName"],
                    8.0,
                )
            )
            for alias in factor["concept"]["aliases"]:
                fields.append((f"{prefix}.concept.alias", alias, 8.0))
    for index, arm in enumerate(candidate["arms"]):
        prefix = f"arms[{index}]"
        fields.extend(
            [
                (f"{prefix}.label", arm["label"], 3.0),
                (f"{prefix}.condition", arm["condition"], 4.0),
                (f"{prefix}.matchingBasis", arm["matchingBasis"], 2.0),
            ]
        )
        for factor_value in arm["factorValues"]:
            fields.extend(
                [
                    (
                        f"{prefix}.factorValue.label",
                        factor_value["factorLabel"],
                        4.0,
                    ),
                    (
                        f"{prefix}.factorValue.originalValue",
                        factor_value["originalValue"],
                        4.0,
                    ),
                ]
            )
    for index, outcome in enumerate(candidate["outcomes"]):
        prefix = f"outcomes[{index}]"
        fields.extend(
            [
                (f"{prefix}.outcomeKey", outcome["outcomeKey"], 3.0),
                (f"{prefix}.originalLabel", outcome["originalLabel"], 6.0),
                (f"{prefix}.domain", outcome["domain"], 3.0),
                (f"{prefix}.metricType", outcome["metricType"], 3.0),
                (f"{prefix}.definition", outcome["definition"], 4.0),
                (f"{prefix}.specification", outcome["specification"], 3.0),
            ]
        )
        if outcome["concept"]:
            fields.append(
                (
                    f"{prefix}.concept.canonicalName",
                    outcome["concept"]["canonicalName"],
                    8.0,
                )
            )
            for alias in outcome["concept"]["aliases"]:
                fields.append((f"{prefix}.concept.alias", alias, 8.0))
    for index, series in enumerate(candidate.get("measurementSeries", [])):
        prefix = f"measurementSeries[{index}]"
        fields.extend(
            [
                (f"{prefix}.seriesKey", series["seriesKey"], 4.0),
                (f"{prefix}.axisLabel", series["axisLabel"], 5.0),
                (f"{prefix}.axisSource", series["axisSource"], 2.0),
                (f"{prefix}.axisUnit", series["axisUnit"], 3.0),
                (
                    f"{prefix}.originalAxisUnit",
                    series["originalAxisUnit"],
                    3.0,
                ),
                (f"{prefix}.valueUnit", series["valueUnit"], 3.0),
                (
                    f"{prefix}.originalValueUnit",
                    series["originalValueUnit"],
                    3.0,
                ),
                (f"{prefix}.stratumKey", series["stratumKey"], 4.0),
            ]
        )
        for replicate_key in series["replicateKeys"]:
            fields.append(
                (f"{prefix}.replicateKey", replicate_key, 3.0)
            )
    return fields


def _score_candidate(
    candidate: dict[str, Any],
    query_tokens: list[str],
) -> dict[str, Any]:
    return _score_fields(_search_fields(candidate), query_tokens)


def _score_fields(
    fields: list[tuple[str, object, float]],
    query_tokens: list[str],
    *,
    allow_single_distinctive_exact_match: bool = False,
) -> dict[str, Any]:
    score = 0.0
    matched_terms: set[str] = set()
    exact_matched_terms: set[str] = set()
    matched_fields: list[dict[str, Any]] = []
    for field_name, value, weight in fields:
        field_tokens = set(unicode_tokens(value))
        exact_matched_terms.update(
            query_token
            for query_token in query_tokens
            if query_token in field_tokens
        )
        overlap = sorted(
            {
                query_token
                for query_token in query_tokens
                for field_token in field_tokens
                if query_token == field_token
                or (
                    min(len(query_token), len(field_token)) >= 3
                    and abs(len(query_token) - len(field_token)) <= 3
                    and (
                        query_token.startswith(field_token)
                        or field_token.startswith(query_token)
                    )
                )
            }
        )
        if not overlap:
            continue
        contribution = weight * len(overlap)
        score += contribution
        matched_terms.update(overlap)
        matched_fields.append(
            {
                "field": field_name,
                "terms": overlap,
                "score": contribution,
            }
        )
    has_distinctive_exact_match = (
        allow_single_distinctive_exact_match
        and any(len(term) >= 5 for term in exact_matched_terms)
    )
    required_matches = (
        1
        if len(query_tokens) == 1 or has_distinctive_exact_match
        else 2
    )
    if len(matched_terms) < required_matches:
        score = 0.0
    else:
        coverage = len(matched_terms) / len(query_tokens)
        score += 100.0 * coverage * coverage
    return {
        "score": round(score, 6),
        "matchedTerms": sorted(matched_terms),
        "matchedTermCount": len(matched_terms),
        "queryTermCoverage": round(
            len(matched_terms) / len(query_tokens),
            6,
        ),
        "matchedFields": matched_fields,
    }


def _terminal_source_exclusion(
    row: dict[str, Any],
    query_tokens: list[str],
) -> dict[str, Any]:
    analysis_status = str(row["analysis_status"] or "")
    status = (
        analysis_status
        if analysis_status in {"EMPTY_WORKBOOK", "NO_TABULAR_EVIDENCE"}
        else str(row["source_content_status"] or "NO_STUDY_RECORD")
    )
    relevance = _score_fields(
        [
            ("publicAnalysisId", row["public_analysis_id"], 12.0),
            ("source.fileName", row["original_file_name"], 4.0),
            ("source.sourcePath", row["source_path"], 2.0),
            ("analysis.title", row["title"], 4.0),
            ("analysis.purpose", row["purpose"], 2.0),
            ("analysis.scope", row["scope_text"], 2.0),
            ("analysis.decision", row["decision_text"], 1.5),
            ("analysis.summary", row["consolidated_summary"], 2.5),
            ("analysis.status", row["analysis_status"], 5.0),
            ("source.contentStatus", status, 5.0),
        ],
        query_tokens,
        allow_single_distinctive_exact_match=True,
    )
    return {
        "publicAnalysisId": str(row["public_analysis_id"]),
        "analysisUid": str(row["analysis_uid"]),
        "revisionUid": str(row["revision_uid"]),
        "sourcePath": str(row["source_path"]),
        "fileName": str(row["original_file_name"]),
        "contentSha256": str(row["content_sha256"]),
        "sourceFingerprint": str(row["source_fingerprint"]),
        "captureContract": str(row["capture_contract"]),
        "sourceContentStatus": status,
        "analysisStatus": str(row["analysis_status"]),
        "verificationStatus": str(row["verification_status"]),
        "summary": str(row["consolidated_summary"]),
        "limitations": _json_value(row["limitations_json"], []),
        "exclusionReasons": [
            {
                "code": status,
                "message": (
                    "No canonical Study was created because the current "
                    f"source content status is {status}."
                ),
            },
            {
                "code": "IMAGES_NOT_ANALYZED",
                "message": "Embedded images are outside the configured scope.",
            },
        ],
        "relevance": relevance,
        "imagesAnalyzed": False,
    }


def _study_candidate(
    connection: sqlite3.Connection,
    row: dict[str, Any],
    aliases: dict[int, list[str]],
) -> dict[str, Any]:
    study_id = int(row["study_id"])
    return {
        "publicDataId": str(row["public_data_id"]),
        "studyUid": str(row["study_uid"]),
        "source": {
            "documentId": int(row["document_id"]),
            "revisionId": int(row["revision_id"]),
            "revisionUid": str(row["revision_uid"]),
            "dataset": str(row["dataset"]),
            "sourcePath": str(row["source_path"]),
            "fileName": str(row["original_file_name"]),
            "contentSha256": str(row["content_sha256"]),
            "sourceFingerprint": str(row["source_fingerprint"]),
            "captureContract": str(row["capture_contract"]),
            "captureStatus": str(row["capture_status"]),
            "isCurrent": bool(row["is_current"]),
        },
        "study": {
            "studyId": study_id,
            "studyKey": str(row["study_key"]),
            "title": str(row["title"]),
            "purpose": str(row["purpose"]),
            "hypothesis": str(row["hypothesis"]),
            "objective": str(row["objective"]),
            "designType": str(row["design_type"]),
            "comparisonBasis": str(row["comparison_basis"]),
            "analysisStatus": str(row["analysis_status"]),
            "verificationStatus": str(row["verification_status"]),
            "comparabilityStatus": str(row["comparability_status"]),
            "confoundingStatus": str(row["confounding_status"]),
            "decision": str(row["decision_text"]),
            "summary": str(row["summary_text"]),
            "limitations": _json_value(row["limitations_json"], []),
        },
        "analysis": {
            "analysisUid": str(row["analysis_uid"]),
            "publicAnalysisId": str(row["public_analysis_id"]),
            "title": str(row["analysis_title"]),
            "purpose": str(row["analysis_purpose"]),
            "scope": str(row["scope_text"]),
            "decision": str(row["analysis_decision_text"]),
            "summary": str(row["consolidated_summary"]),
            "verificationStatus": str(row["analysis_verification_status"]),
        },
        "contexts": _contexts(connection, study_id, aliases),
        "factors": _factors(connection, study_id, aliases),
        "arms": _arms(connection, study_id),
        "outcomes": _outcomes(connection, study_id, aliases),
        "measurementSeries": _measurement_series(connection, study_id),
    }


def _evidence_index(
    connection: sqlite3.Connection,
) -> dict[tuple[str, str], list[dict[str, Any]]]:
    index: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in _rows(
        connection,
        """
        SELECT
            l.entity_type, l.entity_uid, l.evidence_role AS link_role,
            l.claim_scope, e.evidence_id, e.evidence_uid,
            e.public_evidence_id, e.evidence_kind, e.sheet_name,
            e.start_row, e.start_col, e.end_row, e.end_col,
            e.range_address, e.evidence_role, e.source_text, e.note,
            e.content_sha256 AS evidence_content_sha256,
            e.verification_status, r.revision_id, r.revision_uid,
            r.content_sha256 AS revision_content_sha256,
            d.source_path, d.original_file_name
        FROM entity_evidence_links l
        JOIN evidence_items e ON e.evidence_id=l.evidence_id
        JOIN source_revisions r ON r.revision_id=e.revision_id
        JOIN source_documents d ON d.document_id=r.document_id
        ORDER BY e.public_evidence_id, l.entity_type, l.entity_uid
        """,
    ):
        index[(str(row["entity_type"]).upper(), str(row["entity_uid"]))].append(
            {
                "evidenceId": int(row["evidence_id"]),
                "evidenceUid": str(row["evidence_uid"]),
                "publicEvidenceId": str(row["public_evidence_id"]),
                "kind": str(row["evidence_kind"]),
                "sourcePath": str(row["source_path"]),
                "fileName": str(row["original_file_name"]),
                "revisionId": int(row["revision_id"]),
                "revisionUid": str(row["revision_uid"]),
                "contentSha256": (
                    str(row["evidence_content_sha256"])
                    or str(row["revision_content_sha256"])
                ),
                "sheet": str(row["sheet_name"]),
                "range": str(row["range_address"]),
                "start": {
                    "row": int(row["start_row"]),
                    "column": int(row["start_col"]),
                },
                "end": {
                    "row": int(row["end_row"]),
                    "column": int(row["end_col"]),
                },
                "role": str(row["evidence_role"]),
                "linkRole": str(row["link_role"]),
                "claimScope": str(row["claim_scope"]),
                "sourceText": str(row["source_text"]),
                "note": str(row["note"]),
                "verificationStatus": str(row["verification_status"]),
            }
        )
    return dict(index)


def _citations(
    evidence_index: dict[tuple[str, str], list[dict[str, Any]]],
    entities: Iterable[tuple[str, str]],
) -> list[dict[str, Any]]:
    citations: dict[int, dict[str, Any]] = {}
    for entity_type, entity_uid in entities:
        key = (entity_type.upper(), entity_uid)
        for evidence in evidence_index.get(key, []):
            evidence_id = int(evidence["evidenceId"])
            if evidence_id not in citations:
                citations[evidence_id] = {
                    **evidence,
                    "linkedEntities": [],
                }
            linked = {
                "entityType": entity_type.upper(),
                "entityUid": entity_uid,
            }
            if linked not in citations[evidence_id]["linkedEntities"]:
                citations[evidence_id]["linkedEntities"].append(linked)
    return sorted(citations.values(), key=lambda item: item["publicEvidenceId"])


def _observations(
    connection: sqlite3.Connection,
    outcome_id: int,
    arm_id: int,
) -> list[dict[str, Any]]:
    return [
        {
            "observationUid": str(row["observation_uid"]),
            "observationKey": str(row["observation_key"]),
            "stratumKey": str(row["stratum_key"]),
            "replicateKey": str(row["replicate_key"]),
            "observedAt": str(row["observed_at"]),
            "valueNumber": row["value_number"],
            "valueText": str(row["value_text"]),
            "numerator": row["numerator"],
            "denominator": row["denominator"],
            "ratePpm": row["rate_ppm"],
            "min": row["min_value"],
            "max": row["max_value"],
            "average": row["average_value"],
            "sampleSize": row["sample_size"],
            "resultStatus": str(row["result_status"]),
            "verificationStatus": str(row["verification_status"]),
            "details": _json_value(row["details_json"], {}),
        }
        for row in _rows(
            connection,
            """
            SELECT
                observation_uid, observation_key, stratum_key, replicate_key,
                observed_at, value_number, value_text, numerator, denominator,
                rate_ppm, min_value, max_value, average_value, sample_size,
                result_status, verification_status, details_json
            FROM knowledge_observations
            WHERE outcome_id=? AND arm_id=?
            ORDER BY observation_id
            """,
            (outcome_id, arm_id),
        )
    ]


def _comparison_rows(
    connection: sqlite3.Connection,
    study_id: int,
) -> list[dict[str, Any]]:
    return _rows(
        connection,
        """
        SELECT
            comparison_id, comparison_uid, public_comparison_id,
            comparison_key, compared_arm_id, control_arm_id, design_type,
            matching_basis, validity_status, confounding_status,
            exclusion_reason, direction, summary_text,
            aggregation_eligible, verification_status, details_json
        FROM knowledge_comparisons
        WHERE study_id=?
        ORDER BY public_comparison_id
        """,
        (study_id,),
    )


def _effect_rows(
    connection: sqlite3.Connection,
    comparison_id: int,
) -> list[dict[str, Any]]:
    return _rows(
        connection,
        """
        SELECT
            e.effect_id, e.effect_uid, e.public_effect_id, e.outcome_id,
            e.effect_type, e.estimate, e.original_unit, e.ci_lower,
            e.ci_upper, e.formula_version, e.calculation_text, e.direction,
            e.aggregation_eligible, e.verification_status, e.details_json,
            u.canonical_symbol
        FROM knowledge_effects e
        LEFT JOIN knowledge_units u ON u.unit_id=e.unit_id
        WHERE e.comparison_id=?
        ORDER BY e.public_effect_id
        """,
        (comparison_id,),
    )


def _eligibility_reasons(
    candidate: dict[str, Any],
    comparison: dict[str, Any],
    effect: dict[str, Any],
    citations: list[dict[str, Any]],
) -> list[dict[str, str]]:
    reasons = _candidate_trust_reasons(candidate)
    effect_status = str(effect["verification_status"])
    if effect["estimate"] is None:
        reasons.append(
            {
                "code": "EFFECT_ESTIMATE_MISSING",
                "message": "Effect estimate is missing.",
            }
        )
    if effect_status != "VERIFIED":
        reasons.append(
            {
                "code": f"EFFECT_{effect_status}",
                "message": f"Effect verification status is {effect_status}.",
            }
        )
    if not bool(effect["aggregation_eligible"]):
        reasons.append(
            {
                "code": "EFFECT_NOT_AGGREGATION_ELIGIBLE",
                "message": "Effect is not approved for aggregation.",
            }
        )

    comparison_status = str(comparison["verification_status"])
    if comparison_status != "VERIFIED":
        reasons.append(
            {
                "code": f"COMPARISON_{comparison_status}",
                "message": (
                    f"Comparison verification status is {comparison_status}."
                ),
            }
        )
    validity = str(comparison["validity_status"])
    if validity != "VALID":
        reasons.append(
            {
                "code": f"COMPARISON_{validity}",
                "message": f"Comparison validity status is {validity}.",
            }
        )
    confounding = str(comparison["confounding_status"])
    if confounding != "NONE":
        reasons.append(
            {
                "code": f"COMPARISON_{confounding}",
                "message": f"Comparison confounding status is {confounding}.",
            }
        )
        if _is_multi_factor_confounding(candidate, comparison):
            reasons.append(
                {
                    "code": "CONFOUNDED_MULTI_FACTOR",
                    "message": (
                        "Two or more recorded factor values differ across "
                        "the confounded comparison."
                    ),
                }
            )
    if not bool(comparison["aggregation_eligible"]):
        reasons.append(
            {
                "code": "COMPARISON_NOT_AGGREGATION_ELIGIBLE",
                "message": "Comparison is not approved for aggregation.",
            }
        )
    effect_citations = [
        citation
        for citation in citations
        if any(
            linked["entityType"] == "EFFECT"
            and linked["entityUid"] == str(effect["effect_uid"])
            for linked in citation["linkedEntities"]
        )
    ]
    if not effect_citations:
        reasons.append(
            {
                "code": "NO_EFFECT_EVIDENCE",
                "message": "No source evidence is linked directly to this effect.",
            }
        )
    if not any(
        citation["verificationStatus"] == "VERIFIED"
        for citation in effect_citations
    ):
        reasons.append(
            {
                "code": "NO_VERIFIED_EFFECT_EVIDENCE",
                "message": "No verified source evidence is linked directly to this effect.",
            }
        )
    expected_revision_id = int(candidate["source"]["revisionId"])
    if any(
        int(citation["revisionId"]) != expected_revision_id
        for citation in effect_citations
    ):
        reasons.append(
            {
                "code": "EFFECT_EVIDENCE_REVISION_MISMATCH",
                "message": "Effect evidence is not from the study's source revision.",
            }
        )
    expected_sha256 = str(candidate["source"]["contentSha256"]).lower()
    if any(
        str(citation["contentSha256"]).lower() != expected_sha256
        for citation in effect_citations
    ):
        reasons.append(
            {
                "code": "EFFECT_EVIDENCE_CONTENT_MISMATCH",
                "message": "Effect evidence content hash does not match the study source.",
            }
        )
    # A comparison can expose the same state through more than one canonical
    # gate (for example verification and validity can both be NEEDS_REVIEW).
    # Keep one stable reason per code while preserving the first explanation.
    unique_reasons: list[dict[str, str]] = []
    seen_codes: set[str] = set()
    for reason in reasons:
        if reason["code"] in seen_codes:
            continue
        seen_codes.add(reason["code"])
        unique_reasons.append(reason)
    return unique_reasons


def _candidate_trust_reasons(
    candidate: dict[str, Any],
) -> list[dict[str, str]]:
    reasons: list[dict[str, str]] = []
    source = candidate["source"]
    if not bool(source["isCurrent"]):
        reasons.append(
            {
                "code": "SOURCE_NOT_CURRENT",
                "message": "Study belongs to a superseded source revision.",
            }
        )
    capture_status = str(source["captureStatus"]).upper()
    if capture_status != "CAPTURED":
        reasons.append(
            {
                "code": f"SOURCE_{capture_status}",
                "message": f"Source capture status is {capture_status}.",
            }
        )
    analysis_status = str(
        candidate["analysis"]["verificationStatus"]
    ).upper()
    if analysis_status != "VERIFIED":
        reasons.append(
            {
                "code": f"ANALYSIS_{analysis_status}",
                "message": f"Workbook analysis verification status is {analysis_status}.",
            }
        )
    study_status = str(candidate["study"]["verificationStatus"]).upper()
    if study_status != "VERIFIED":
        reasons.append(
            {
                "code": f"STUDY_{study_status}",
                "message": f"Study verification status is {study_status}.",
            }
        )
    comparability = str(
        candidate["study"]["comparabilityStatus"]
    ).upper()
    if comparability != "VALID":
        reasons.append(
            {
                "code": f"STUDY_COMPARABILITY_{comparability}",
                "message": f"Study comparability status is {comparability}.",
            }
        )
    confounding = str(candidate["study"]["confoundingStatus"]).upper()
    if confounding != "NONE":
        reasons.append(
            {
                "code": f"STUDY_CONFOUNDING_{confounding}",
                "message": f"Study confounding status is {confounding}.",
            }
        )
    return reasons


def _comparison_trust_reasons(
    comparison: dict[str, Any],
) -> list[dict[str, str]]:
    """Return the comparison gates even when no Effect row exists."""

    reasons: list[dict[str, str]] = []
    comparison_status = str(comparison["verification_status"]).upper()
    if comparison_status != "VERIFIED":
        reasons.append(
            {
                "code": f"COMPARISON_{comparison_status}",
                "message": (
                    "Comparison verification status is "
                    f"{comparison_status}."
                ),
            }
        )
    validity = str(comparison["validity_status"]).upper()
    if validity != "VALID":
        reasons.append(
            {
                "code": f"COMPARISON_{validity}",
                "message": f"Comparison validity status is {validity}.",
            }
        )
    confounding = str(comparison["confounding_status"]).upper()
    if confounding != "NONE":
        reasons.append(
            {
                "code": f"COMPARISON_{confounding}",
                "message": (
                    "Comparison confounding status is "
                    f"{confounding}."
                ),
            }
        )
    if not bool(comparison["aggregation_eligible"]):
        reasons.append(
            {
                "code": "COMPARISON_NOT_AGGREGATION_ELIGIBLE",
                "message": "Comparison is not approved for aggregation.",
            }
        )
    unique_reasons: list[dict[str, str]] = []
    seen_codes: set[str] = set()
    for reason in reasons:
        if reason["code"] in seen_codes:
            continue
        seen_codes.add(reason["code"])
        unique_reasons.append(reason)
    return unique_reasons


def _arm_by_id(candidate: dict[str, Any], arm_id: int) -> dict[str, Any]:
    for arm in candidate["arms"]:
        if int(arm["armId"]) == arm_id:
            return arm
    raise EvidenceQueryError(f"Comparison references missing arm {arm_id}.")


def _outcome_by_id(candidate: dict[str, Any], outcome_id: int) -> dict[str, Any]:
    for outcome in candidate["outcomes"]:
        if int(outcome["outcomeId"]) == outcome_id:
            return outcome
    raise EvidenceQueryError(f"Effect references missing outcome {outcome_id}.")


def _effect_payload(
    connection: sqlite3.Connection,
    candidate: dict[str, Any],
    comparison: dict[str, Any],
    effect: dict[str, Any],
    evidence_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> tuple[dict[str, Any], list[dict[str, str]]]:
    compared_arm_id = int(comparison["compared_arm_id"])
    control_arm_id = int(comparison["control_arm_id"])
    outcome_id = int(effect["outcome_id"])
    compared_arm = _arm_by_id(candidate, compared_arm_id)
    control_arm = _arm_by_id(candidate, control_arm_id)
    outcome = _outcome_by_id(candidate, outcome_id)
    compared_observations = _observations(
        connection,
        outcome_id,
        compared_arm_id,
    )
    control_observations = _observations(
        connection,
        outcome_id,
        control_arm_id,
    )
    entities = [
        ("EFFECT", str(effect["effect_uid"])),
        ("COMPARISON", str(comparison["comparison_uid"])),
        ("OUTCOME", str(outcome["outcomeUid"])),
        ("ARM", str(compared_arm["armUid"])),
        ("ARM", str(control_arm["armUid"])),
        ("STUDY", str(candidate["studyUid"])),
        ("WORKBOOK_ANALYSIS", str(candidate["analysis"]["analysisUid"])),
    ]
    entities.extend(
        (
            "OBSERVATION",
            str(observation["observationUid"]),
        )
        for observation in compared_observations + control_observations
    )
    citations = _citations(evidence_index, entities)
    payload = {
        "publicDataId": candidate["publicDataId"],
        "publicComparisonId": str(comparison["public_comparison_id"]),
        "publicEffectId": str(effect["public_effect_id"]),
        "publicEvidenceIds": [
            citation["publicEvidenceId"] for citation in citations
        ],
        "sourcePath": candidate["source"]["sourcePath"],
        "source": candidate["source"],
        "analysis": candidate["analysis"],
        "study": {
            "title": candidate["study"]["title"],
            "summary": candidate["study"]["summary"],
            "verificationStatus": candidate["study"]["verificationStatus"],
            "comparabilityStatus": candidate["study"]["comparabilityStatus"],
            "confoundingStatus": candidate["study"]["confoundingStatus"],
        },
        "comparison": {
            "comparisonUid": str(comparison["comparison_uid"]),
            "comparisonKey": str(comparison["comparison_key"]),
            "designType": str(comparison["design_type"]),
            "matchingBasis": str(comparison["matching_basis"]),
            "validityStatus": str(comparison["validity_status"]),
            "confoundingStatus": str(comparison["confounding_status"]),
            "verificationStatus": str(comparison["verification_status"]),
            "aggregationEligible": bool(comparison["aggregation_eligible"]),
            "direction": str(comparison["direction"]),
            "summary": str(comparison["summary_text"]),
            "exclusionReason": str(comparison["exclusion_reason"]),
            "details": _json_value(comparison["details_json"], {}),
            "factorDifferences": _arm_factor_differences(
                compared_arm,
                control_arm,
            ),
            "comparedArm": compared_arm,
            "controlArm": control_arm,
        },
        "outcome": outcome,
        "observations": {
            "comparedArm": compared_observations,
            "controlArm": control_observations,
        },
        "effect": {
            "effectType": str(effect["effect_type"]),
            "estimate": effect["estimate"],
            "unit": effect["canonical_symbol"] or effect["original_unit"],
            "originalUnit": str(effect["original_unit"]),
            "ciLower": effect["ci_lower"],
            "ciUpper": effect["ci_upper"],
            "formulaVersion": str(effect["formula_version"]),
            "calculation": str(effect["calculation_text"]),
            "direction": str(effect["direction"]),
            "aggregationEligible": bool(effect["aggregation_eligible"]),
            "verificationStatus": str(effect["verification_status"]),
            "details": _json_value(effect["details_json"], {}),
        },
        "evidence": citations,
    }
    reasons = _eligibility_reasons(
        candidate,
        comparison,
        effect,
        citations,
    )
    if reasons:
        payload["descriptiveScope"] = "COMPARISON"
        payload["descriptiveOutcomes"] = [
            {
                "outcome": outcome,
                "armObservations": [
                    {
                        "arm": compared_arm,
                        "observations": compared_observations,
                    },
                    {
                        "arm": control_arm,
                        "observations": control_observations,
                    },
                ],
            }
        ]
    return payload, reasons


def _comparison_without_effect(
    connection: sqlite3.Connection,
    candidate: dict[str, Any],
    comparison: dict[str, Any],
    evidence_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> dict[str, Any]:
    compared_arm = _arm_by_id(candidate, int(comparison["compared_arm_id"]))
    control_arm = _arm_by_id(candidate, int(comparison["control_arm_id"]))
    comparison_series = _measurement_series_for_comparison(
        candidate,
        comparison,
    )
    factor_differences = _arm_factor_differences(
        compared_arm,
        control_arm,
    )
    descriptive_outcomes, descriptive_citations = _descriptive_outcomes(
        connection,
        candidate,
        evidence_index,
    )
    citations = _merge_citations(
        _citations(
        evidence_index,
        [
            ("COMPARISON", str(comparison["comparison_uid"])),
            ("STUDY", str(candidate["studyUid"])),
            *[
                ("MEASUREMENT_SERIES", str(series["seriesUid"]))
                for series in comparison_series
            ],
        ],
        ),
        descriptive_citations,
    )
    return {
        "publicDataId": candidate["publicDataId"],
        "publicComparisonId": str(comparison["public_comparison_id"]),
        "publicEffectId": None,
        "publicEvidenceIds": [
            citation["publicEvidenceId"] for citation in citations
        ],
        "sourcePath": candidate["source"]["sourcePath"],
        "comparison": {
            "comparisonUid": str(comparison["comparison_uid"]),
            "comparisonKey": str(comparison["comparison_key"]),
            "designType": str(comparison["design_type"]),
            "matchingBasis": str(comparison["matching_basis"]),
            "validityStatus": str(comparison["validity_status"]),
            "confoundingStatus": str(comparison["confounding_status"]),
            "verificationStatus": str(comparison["verification_status"]),
            "aggregationEligible": bool(comparison["aggregation_eligible"]),
            "direction": str(comparison["direction"]),
            "summary": str(comparison["summary_text"]),
            "exclusionReason": str(comparison["exclusion_reason"]),
            "details": _json_value(comparison["details_json"], {}),
            "factorDifferences": factor_differences,
            "comparedArm": compared_arm,
            "controlArm": control_arm,
        },
        "outcome": None,
        "observations": {"comparedArm": [], "controlArm": []},
        "effect": None,
        "descriptiveScope": "STUDY",
        "descriptiveOutcomes": descriptive_outcomes,
        "descriptiveMeasurementSeries": comparison_series,
        "evidence": citations,
        "exclusionReasons": (
            _candidate_trust_reasons(candidate)
            + _comparison_trust_reasons(comparison)
            + (
                [
                    {
                        "code": "CONFOUNDED_MULTI_FACTOR",
                        "message": (
                            "Two or more recorded factor values differ "
                            "across the confounded comparison."
                        ),
                    }
                ]
                if str(comparison["confounding_status"]).upper()
                == "CONFOUNDED"
                and len(factor_differences) >= 2
                else []
            )
            + [
                {
                    "code": "NO_EFFECT_RECORD",
                    "message": "Comparison has no calculated effect record.",
                }
            ]
        ),
    }


def _factor_identity(value: dict[str, Any]) -> str:
    factor_uid = str(value.get("factorUid") or "").strip()
    if factor_uid:
        return f"UID:{factor_uid}"
    return "LABEL:" + normalize_text(value.get("factorLabel") or "")


def _factor_value_signature(value: dict[str, Any] | None) -> tuple[Any, ...]:
    if value is None:
        return ()
    return (
        str(value.get("originalValue") or "").strip(),
        value.get("valueNumber"),
        str(value.get("unit") or "").strip(),
    )


def _arm_factor_differences(
    compared_arm: dict[str, Any],
    control_arm: dict[str, Any],
) -> list[dict[str, Any]]:
    compared_values = {
        _factor_identity(value): value
        for value in compared_arm.get("factorValues", [])
        if _factor_identity(value) != "LABEL:"
    }
    control_values = {
        _factor_identity(value): value
        for value in control_arm.get("factorValues", [])
        if _factor_identity(value) != "LABEL:"
    }
    differences: list[dict[str, Any]] = []
    for identity in sorted(set(compared_values) | set(control_values)):
        compared = compared_values.get(identity)
        control = control_values.get(identity)
        if _factor_value_signature(compared) == _factor_value_signature(control):
            continue
        exemplar = compared or control or {}
        differences.append(
            {
                "factorUid": str(exemplar.get("factorUid") or ""),
                "factorLabel": str(exemplar.get("factorLabel") or ""),
                "comparedValue": (
                    str(compared.get("originalValue") or "")
                    if compared is not None
                    else ""
                ),
                "controlValue": (
                    str(control.get("originalValue") or "")
                    if control is not None
                    else ""
                ),
                "comparedValueRecorded": compared is not None,
                "controlValueRecorded": control is not None,
                "comparedHeldConstant": (
                    bool(compared.get("heldConstant"))
                    if compared is not None
                    else None
                ),
                "controlHeldConstant": (
                    bool(control.get("heldConstant"))
                    if control is not None
                    else None
                ),
            }
        )
    return differences


def _measurement_series_for_comparison(
    candidate: dict[str, Any],
    comparison: dict[str, Any],
    *,
    outcome_uid: str = "",
) -> list[dict[str, Any]]:
    arm_uids = {
        str(
            _arm_by_id(
                candidate,
                int(comparison[arm_id_field]),
            ).get("armUid")
            or ""
        )
        for arm_id_field in ("compared_arm_id", "control_arm_id")
    }
    return [
        series
        for series in candidate.get("measurementSeries", [])
        if str(series.get("arm", {}).get("armUid") or "") in arm_uids
        and (
            not outcome_uid
            or str(series.get("outcome", {}).get("outcomeUid") or "")
            == outcome_uid
        )
    ]


def _is_multi_factor_confounding(
    candidate: dict[str, Any],
    comparison: dict[str, Any],
) -> bool:
    if str(comparison["confounding_status"]).upper() != "CONFOUNDED":
        return False
    compared_arm = _arm_by_id(
        candidate,
        int(comparison["compared_arm_id"]),
    )
    control_arm = _arm_by_id(
        candidate,
        int(comparison["control_arm_id"]),
    )
    return len(_arm_factor_differences(compared_arm, control_arm)) >= 2


def _descriptive_outcomes(
    connection: sqlite3.Connection,
    candidate: dict[str, Any],
    evidence_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    payloads: list[dict[str, Any]] = []
    entities: list[tuple[str, str]] = [
        ("STUDY", str(candidate["studyUid"])),
    ]
    for outcome in candidate["outcomes"]:
        arm_observations: list[dict[str, Any]] = []
        entities.append(("OUTCOME", str(outcome["outcomeUid"])))
        for arm in candidate["arms"]:
            observations = _observations(
                connection,
                int(outcome["outcomeId"]),
                int(arm["armId"]),
            )
            if not observations:
                continue
            entities.append(("ARM", str(arm["armUid"])))
            entities.extend(
                (
                    "OBSERVATION",
                    str(observation["observationUid"]),
                )
                for observation in observations
            )
            arm_observations.append(
                {
                    "arm": arm,
                    "observations": observations,
                }
            )
        if arm_observations:
            payloads.append(
                {
                    "outcome": outcome,
                    "armObservations": arm_observations,
                }
            )
    return payloads, _citations(evidence_index, entities)


def _series_citations(
    series_records: Iterable[dict[str, Any]],
    evidence_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    return _citations(
        evidence_index,
        [
            ("MEASUREMENT_SERIES", str(item["seriesUid"]))
            for item in series_records
        ],
    )


def _merge_citations(
    *citation_groups: Iterable[dict[str, Any]],
) -> list[dict[str, Any]]:
    by_evidence_id: dict[int, dict[str, Any]] = {}
    for citations in citation_groups:
        for citation in citations:
            evidence_id = int(citation["evidenceId"])
            if evidence_id not in by_evidence_id:
                by_evidence_id[evidence_id] = citation
                continue
            existing_links = by_evidence_id[evidence_id]["linkedEntities"]
            for linked in citation.get("linkedEntities", []):
                if linked not in existing_links:
                    existing_links.append(linked)
    return sorted(
        by_evidence_id.values(),
        key=lambda item: item["publicEvidenceId"],
    )


def _attach_descriptive_measurement_series(
    payload: dict[str, Any],
    candidate: dict[str, Any],
    evidence_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> dict[str, Any]:
    """Attach source-backed series only to a non-quantitative payload."""

    if "descriptiveMeasurementSeries" in payload:
        series = payload["descriptiveMeasurementSeries"]
    else:
        series = candidate.get("measurementSeries", [])
    if not series:
        return payload
    payload["descriptiveMeasurementSeries"] = series
    citations = _merge_citations(
        payload.get("evidence", []),
        _series_citations(series, evidence_index),
    )
    payload["evidence"] = citations
    payload["publicEvidenceIds"] = [
        citation["publicEvidenceId"] for citation in citations
    ]
    return payload


def build_evidence_pack(
    connection: sqlite3.Connection,
    question: str,
) -> dict[str, Any]:
    """Retrieve all relevant studies and partition effects by trust gates."""

    _validate_schema(connection)
    query_tokens = _unique_tokens(question)
    if not query_tokens:
        raise ValueError("question must contain at least one Unicode token.")
    search_tokens = _search_tokens(query_tokens)

    aliases = _concept_aliases(connection)
    direct_matches = _direct_public_id_matches(connection, question)
    scored_candidates: list[dict[str, Any]] = []
    for row in _study_base_rows(connection):
        candidate = _study_candidate(connection, row, aliases)
        relevance = _score_candidate(candidate, search_tokens)
        identifier_matches = direct_matches.get(
            int(candidate["study"]["studyId"]),
            [],
        )
        if identifier_matches:
            relevance["score"] = round(
                float(relevance["score"]) + 1000.0,
                6,
            )
            relevance["directIdentifierMatches"] = identifier_matches
            relevance["matchedFields"].append(
                {
                    "field": "publicIdentifier",
                    "terms": identifier_matches,
                    "score": 1000.0,
                }
            )
        candidate["relevance"] = relevance
        scored_candidates.append(candidate)

    outcome_query_terms = {
        term
        for candidate in scored_candidates
        for field in candidate["relevance"]["matchedFields"]
        if str(field["field"]).startswith("outcomes[")
        for term in field["terms"]
        if term in search_tokens
    }
    context_or_factor_terms = {
        term
        for candidate in scored_candidates
        for field in candidate["relevance"]["matchedFields"]
        if not str(field["field"]).startswith("outcomes[")
        for term in field["terms"]
        if term in search_tokens and term not in outcome_query_terms
    }
    nonnumeric_context_terms = {
        term for term in context_or_factor_terms if not term.isdecimal()
    }
    if nonnumeric_context_terms:
        context_or_factor_terms = nonnumeric_context_terms
    relation_gate_applied = bool(
        outcome_query_terms and context_or_factor_terms
    )

    candidates: list[dict[str, Any]] = []
    for candidate in scored_candidates:
        relevance = candidate["relevance"]
        if relevance["score"] <= 0:
            continue
        if relevance.get("directIdentifierMatches"):
            candidates.append(candidate)
            continue
        if relation_gate_applied:
            candidate_outcome_terms = {
                term
                for field in relevance["matchedFields"]
                if str(field["field"]).startswith("outcomes[")
                for term in field["terms"]
            }
            # Workbook and Study titles often carry the review's broad outcome
            # name while the source table stores only its detailed submetrics.
            # Treat title/file matches as an outcome proxy only for terms that
            # were independently recognized in real outcome fields elsewhere.
            candidate_outcome_terms.update(
                term
                for field in relevance["matchedFields"]
                if str(field["field"])
                in {
                    "source.fileName",
                    "analysis.title",
                    "study.title",
                }
                for term in field["terms"]
                if term in outcome_query_terms
            )
            candidate_context_terms = {
                term
                for field in relevance["matchedFields"]
                if not str(field["field"]).startswith("outcomes[")
                for term in field["terms"]
            }
            if not (
                candidate_outcome_terms & outcome_query_terms
                and candidate_context_terms & context_or_factor_terms
            ):
                continue
        candidates.append(candidate)
    candidates.sort(
        key=lambda candidate: (
            -float(candidate["relevance"]["score"]),
            candidate["publicDataId"],
        )
    )
    for rank, candidate in enumerate(candidates, start=1):
        candidate["rank"] = rank

    source_exclusions = [
        _terminal_source_exclusion(row, search_tokens)
        for row in _terminal_source_rows(connection)
    ]
    source_exclusions = [
        item
        for item in source_exclusions
        if float(item["relevance"]["score"]) > 0
    ]
    if relation_gate_applied:
        terminal_status_requested = bool(
            {"empty_workbook", "no_tabular_evidence"} & set(search_tokens)
        )
        source_exclusions = [
            item
            for item in source_exclusions
            if any(
                str(field["field"]) == "publicAnalysisId"
                for field in item["relevance"]["matchedFields"]
            )
            or set(item["relevance"]["matchedTerms"]) & outcome_query_terms
            or (
                terminal_status_requested
                and set(item["relevance"]["matchedTerms"])
                & context_or_factor_terms
            )
        ]
    source_exclusions.sort(
        key=lambda item: (
            -float(item["relevance"]["score"]),
            item["publicAnalysisId"],
        )
    )
    for rank, item in enumerate(source_exclusions, start=1):
        item["rank"] = rank

    evidence_index = _evidence_index(connection)
    eligible_effects: list[dict[str, Any]] = []
    excluded_candidates: list[dict[str, Any]] = []
    for candidate in candidates:
        comparisons = _comparison_rows(
            connection,
            int(candidate["study"]["studyId"]),
        )
        if not comparisons:
            descriptive_outcomes, citations = _descriptive_outcomes(
                connection,
                candidate,
                evidence_index,
            )
            excluded_candidates.append(
                _attach_descriptive_measurement_series(
                    {
                    "publicDataId": candidate["publicDataId"],
                    "publicComparisonId": None,
                    "publicEffectId": None,
                    "publicEvidenceIds": [
                        citation["publicEvidenceId"] for citation in citations
                    ],
                    "sourcePath": candidate["source"]["sourcePath"],
                    "comparison": None,
                    "outcome": None,
                    "observations": {
                        "comparedArm": [],
                        "controlArm": [],
                    },
                    "effect": None,
                    "descriptiveOutcomes": descriptive_outcomes,
                    "evidence": citations,
                    "exclusionReasons": _candidate_trust_reasons(candidate) + [
                        {
                            "code": "NO_COMPARISON_RECORD",
                            "message": "Study has no control/comparison record.",
                        }
                    ],
                    },
                    candidate,
                    evidence_index,
                )
            )
            continue

        for comparison in comparisons:
            effects = _effect_rows(
                connection,
                int(comparison["comparison_id"]),
            )
            if not effects:
                excluded_candidates.append(
                    _attach_descriptive_measurement_series(
                        _comparison_without_effect(
                            connection,
                            candidate,
                            comparison,
                            evidence_index,
                        ),
                        candidate,
                        evidence_index,
                    )
                )
                continue
            for effect in effects:
                payload, reasons = _effect_payload(
                    connection,
                    candidate,
                    comparison,
                    effect,
                    evidence_index,
                )
                if reasons:
                    payload["exclusionReasons"] = reasons
                    payload["descriptiveMeasurementSeries"] = (
                        _measurement_series_for_comparison(
                            candidate,
                            comparison,
                            outcome_uid=str(
                                payload.get("outcome", {}).get(
                                    "outcomeUid"
                                )
                                or ""
                            ),
                        )
                    )
                    excluded_candidates.append(
                        _attach_descriptive_measurement_series(
                            payload,
                            candidate,
                            evidence_index,
                        )
                    )
                else:
                    eligible_effects.append(payload)

    eligible_summary = summarize_eligible_effects(eligible_effects)
    return {
        "schemaVersion": EVIDENCE_PACK_SCHEMA_VERSION,
        "question": question,
        "normalizedQuestion": normalize_text(question),
        "queryTokens": query_tokens,
        "searchTokens": search_tokens,
        "queryRoleHints": {
            "outcomeTerms": sorted(outcome_query_terms),
            "contextOrFactorTerms": sorted(context_or_factor_terms),
            "relationGateApplied": relation_gate_applied,
        },
        "studyCandidates": candidates,
        "sourceExclusions": source_exclusions,
        "answerEligibleEffects": eligible_effects,
        "excludedCandidates": excluded_candidates,
        "eligibleEffectSummary": eligible_summary,
        "summary": {
            "relevantStudyCount": len(candidates),
            "relevantSourceExclusionCount": len(source_exclusions),
            "answerEligibleEffectCount": len(eligible_effects),
            "excludedCandidateCount": len(excluded_candidates),
            "eligibleSummaryGroupCount": len(eligible_summary),
        },
    }


def summarize_eligible_effects(
    effects: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Deterministically aggregate only already answer-eligible effects."""

    groups: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for item in effects:
        estimate = item.get("effect", {}).get("estimate")
        if estimate is None:
            continue
        key = (
            str(item.get("outcome", {}).get("originalLabel") or ""),
            str(item.get("effect", {}).get("effectType") or ""),
            str(item.get("effect", {}).get("unit") or ""),
        )
        groups[key].append(item)
    summaries: list[dict[str, Any]] = []
    for (outcome, effect_type, unit), items in sorted(groups.items()):
        estimates = [float(item["effect"]["estimate"]) for item in items]
        directions = sorted(
            {
                str(item["effect"].get("direction") or "")
                for item in items
                if str(item["effect"].get("direction") or "")
            }
        )
        summaries.append(
            {
                "outcome": outcome,
                "effectType": effect_type,
                "unit": unit,
                "count": len(items),
                "minimum": min(estimates),
                "maximum": max(estimates),
                "mean": sum(estimates) / len(estimates),
                "directions": directions,
                "directionConflict": len(directions) > 1,
                "publicDataIds": sorted(
                    {str(item["publicDataId"]) for item in items}
                ),
                "publicComparisonIds": sorted(
                    {str(item["publicComparisonId"]) for item in items}
                ),
                "publicEffectIds": sorted(
                    {str(item["publicEffectId"]) for item in items}
                ),
                "publicEvidenceIds": sorted(
                    {
                        str(evidence_id)
                        for item in items
                        for evidence_id in item["publicEvidenceIds"]
                    }
                ),
            }
        )
    return summaries


def build_evidence_pack_from_db(
    database_path: str | Path,
    question: str,
) -> dict[str, Any]:
    with connect_knowledge_readonly(database_path) as connection:
        return build_evidence_pack(connection, question)


__all__ = [
    "EVIDENCE_PACK_SCHEMA_VERSION",
    "EvidenceQueryError",
    "build_evidence_pack",
    "build_evidence_pack_from_db",
    "connect_knowledge_readonly",
    "normalize_text",
    "summarize_eligible_effects",
    "unicode_tokens",
]
