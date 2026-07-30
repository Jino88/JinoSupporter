"""Fail-closed human curation for canonical concepts and aliases."""

from __future__ import annotations

import hashlib
import json
import sqlite3
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Iterator

from inference_data_ai_schema import normalized_term, stable_uid


CONCEPT_CURATION_MIGRATION = "canonical-concept-curation-v1"
CANDIDATE_LIST_SCHEMA_VERSION = "concept-candidate-list-v1"
CONCEPT_LIST_SCHEMA_VERSION = "canonical-concept-list-v1"
CONCEPT_RESOLUTION_SCHEMA_VERSION = "concept-resolution-v1"
CONCEPT_ALIAS_APPROVAL_SCHEMA_VERSION = "concept-alias-approval-v1"

_CANDIDATE_STATUSES = {"OPEN", "APPROVED", "REJECTED", "MERGED"}
_CONCEPT_STATUSES = {"ACTIVE", "DEPRECATED"}
_RESOLUTION_ACTIONS = {"CREATE", "MERGE", "REJECT"}
_ACTION_STATUS = {
    "CREATE": "APPROVED",
    "MERGE": "MERGED",
    "REJECT": "REJECTED",
}
_REQUIRED_TABLES = {
    "knowledge_concepts",
    "knowledge_concept_aliases",
    "knowledge_schema_candidates",
    "knowledge_concept_resolution_history",
    "knowledge_concept_alias_approval_history",
}


class ConceptCurationError(RuntimeError):
    """Raised when a human curation request cannot be applied safely."""


def _text(value: object, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ConceptCurationError(f"{field} must not be empty")
    return text


def _rows(
    connection: sqlite3.Connection,
    sql: str,
    parameters: tuple[object, ...] = (),
) -> list[dict[str, Any]]:
    cursor = connection.execute(sql, parameters)
    columns = [str(item[0]) for item in cursor.description or ()]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def _validate_limit(limit: int) -> int:
    if (
        isinstance(limit, bool)
        or not isinstance(limit, int)
        or limit < 1
        or limit > 10_000
    ):
        raise ValueError("limit must be an integer from 1 through 10000")
    return limit


def _validate_schema(connection: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )
    }
    missing = sorted(_REQUIRED_TABLES - tables)
    if missing:
        raise ConceptCurationError(
            "Canonical DB is missing the concept curation migration: "
            + ", ".join(missing)
        )


def ensure_concept_curation_schema(
    connection: sqlite3.Connection,
    now_iso: Callable[[], str],
) -> None:
    """Install immutable concept-resolution and alias-approval history."""

    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS knowledge_concept_resolution_history (
            resolution_id INTEGER PRIMARY KEY AUTOINCREMENT,
            resolution_uid TEXT NOT NULL UNIQUE,
            schema_candidate_id INTEGER NOT NULL UNIQUE,
            candidate_uid TEXT NOT NULL,
            candidate_kind TEXT NOT NULL,
            candidate_normalized_value TEXT NOT NULL,
            candidate_original_value TEXT NOT NULL,
            candidate_suggested_canonical_name TEXT NOT NULL,
            action TEXT NOT NULL
                CHECK(action IN ('CREATE','MERGE','REJECT')),
            concept_id INTEGER,
            concept_uid TEXT NOT NULL DEFAULT '',
            concept_kind TEXT NOT NULL DEFAULT '',
            canonical_name TEXT NOT NULL DEFAULT '',
            normalized_canonical_name TEXT NOT NULL DEFAULT '',
            alias_id INTEGER,
            alias_uid TEXT NOT NULL DEFAULT '',
            alias_text TEXT NOT NULL DEFAULT '',
            normalized_alias TEXT NOT NULL DEFAULT '',
            reviewer TEXT NOT NULL,
            note TEXT NOT NULL,
            request_json TEXT NOT NULL,
            request_sha256 TEXT NOT NULL UNIQUE,
            resolved_at TEXT NOT NULL,
            FOREIGN KEY(schema_candidate_id)
                REFERENCES knowledge_schema_candidates(schema_candidate_id)
                ON DELETE RESTRICT,
            FOREIGN KEY(concept_id)
                REFERENCES knowledge_concepts(concept_id)
                ON DELETE RESTRICT,
            FOREIGN KEY(alias_id)
                REFERENCES knowledge_concept_aliases(alias_id)
                ON DELETE RESTRICT,
            CHECK(
                (
                    action='REJECT'
                    AND concept_id IS NULL
                    AND alias_id IS NULL
                    AND concept_uid=''
                    AND alias_uid=''
                )
                OR
                (
                    action IN ('CREATE','MERGE')
                    AND concept_id IS NOT NULL
                    AND alias_id IS NOT NULL
                    AND concept_uid<>''
                    AND alias_uid<>''
                )
            )
        );

        CREATE TABLE IF NOT EXISTS knowledge_concept_alias_approval_history (
            alias_approval_id INTEGER PRIMARY KEY AUTOINCREMENT,
            approval_uid TEXT NOT NULL UNIQUE,
            concept_id INTEGER NOT NULL,
            concept_uid TEXT NOT NULL,
            concept_kind TEXT NOT NULL,
            canonical_name TEXT NOT NULL,
            alias_id INTEGER NOT NULL,
            alias_uid TEXT NOT NULL,
            alias_text TEXT NOT NULL,
            normalized_alias TEXT NOT NULL,
            reviewer TEXT NOT NULL,
            note TEXT NOT NULL,
            request_json TEXT NOT NULL,
            request_sha256 TEXT NOT NULL UNIQUE,
            approved_at TEXT NOT NULL,
            FOREIGN KEY(concept_id)
                REFERENCES knowledge_concepts(concept_id)
                ON DELETE RESTRICT,
            FOREIGN KEY(alias_id)
                REFERENCES knowledge_concept_aliases(alias_id)
                ON DELETE RESTRICT
        );

        CREATE INDEX IF NOT EXISTS idx_concept_resolution_action
            ON knowledge_concept_resolution_history(action, resolved_at);
        CREATE INDEX IF NOT EXISTS idx_alias_approval_concept
            ON knowledge_concept_alias_approval_history(
                concept_id, normalized_alias, approved_at
            );

        CREATE TRIGGER IF NOT EXISTS trg_concept_resolution_no_update
        BEFORE UPDATE ON knowledge_concept_resolution_history
        BEGIN
            SELECT RAISE(ABORT, 'concept resolution history is immutable');
        END;

        CREATE TRIGGER IF NOT EXISTS trg_concept_resolution_no_delete
        BEFORE DELETE ON knowledge_concept_resolution_history
        BEGIN
            SELECT RAISE(ABORT, 'concept resolution history is immutable');
        END;

        CREATE TRIGGER IF NOT EXISTS trg_alias_approval_no_update
        BEFORE UPDATE ON knowledge_concept_alias_approval_history
        BEGIN
            SELECT RAISE(ABORT, 'concept alias approval history is immutable');
        END;

        CREATE TRIGGER IF NOT EXISTS trg_alias_approval_no_delete
        BEFORE DELETE ON knowledge_concept_alias_approval_history
        BEGIN
            SELECT RAISE(ABORT, 'concept alias approval history is immutable');
        END;
        """
    )
    connection.execute(
        """
        INSERT OR IGNORE INTO schema_migrations(
            migration_name, applied_at
        ) VALUES (?, ?)
        """,
        (CONCEPT_CURATION_MIGRATION, now_iso()),
    )


def list_schema_candidates(
    connection: sqlite3.Connection,
    *,
    status: str | None = "OPEN",
    candidate_kind: str | None = None,
    query: str | None = None,
    limit: int = 500,
) -> dict[str, Any]:
    """List deterministically filtered schema candidates without mutation."""

    _validate_schema(connection)
    _validate_limit(limit)
    normalized_status = (
        str(status or "").strip().upper() if status is not None else None
    )
    if normalized_status == "ALL":
        normalized_status = None
    if (
        normalized_status is not None
        and normalized_status not in _CANDIDATE_STATUSES
    ):
        raise ValueError("invalid schema candidate status")
    normalized_kind = str(candidate_kind or "").strip().upper()
    needle = normalized_term(query)
    rows = _rows(
        connection,
        """
        SELECT
            schema_candidate_id, candidate_uid, candidate_kind,
            normalized_value, original_value,
            suggested_canonical_name, occurrence_count, status,
            first_seen_at, last_seen_at
        FROM knowledge_schema_candidates
        ORDER BY
            CASE status
                WHEN 'OPEN' THEN 0
                WHEN 'APPROVED' THEN 1
                WHEN 'MERGED' THEN 2
                ELSE 3
            END,
            candidate_kind, normalized_value, candidate_uid
        """,
    )
    values: list[dict[str, Any]] = []
    for row in rows:
        if normalized_status and str(row["status"]) != normalized_status:
            continue
        if (
            normalized_kind
            and str(row["candidate_kind"]).upper() != normalized_kind
        ):
            continue
        if needle and needle not in normalized_term(
            " ".join(
                (
                    str(row["normalized_value"]),
                    str(row["original_value"]),
                    str(row["suggested_canonical_name"]),
                )
            )
        ):
            continue
        values.append(
            {
                "schemaCandidateId": int(row["schema_candidate_id"]),
                "candidateUid": str(row["candidate_uid"]),
                "candidateKind": str(row["candidate_kind"]),
                "normalizedValue": str(row["normalized_value"]),
                "originalValue": str(row["original_value"]),
                "suggestedCanonicalName": str(
                    row["suggested_canonical_name"]
                ),
                "occurrenceCount": int(row["occurrence_count"]),
                "status": str(row["status"]),
                "firstSeenAt": str(row["first_seen_at"]),
                "lastSeenAt": str(row["last_seen_at"]),
            }
        )
        if len(values) >= limit:
            break
    return {
        "schemaVersion": CANDIDATE_LIST_SCHEMA_VERSION,
        "filters": {
            "status": normalized_status or "ALL",
            "candidateKind": normalized_kind,
            "query": str(query or "").strip(),
            "limit": limit,
        },
        "count": len(values),
        "candidates": values,
    }


def _aliases_by_concept(
    connection: sqlite3.Connection,
) -> dict[int, list[dict[str, Any]]]:
    aliases: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in _rows(
        connection,
        """
        SELECT
            alias_id, alias_uid, concept_id, alias_text,
            normalized_alias, language, source, confidence, created_at
        FROM knowledge_concept_aliases
        ORDER BY concept_id, normalized_alias, alias_id
        """,
    ):
        aliases[int(row["concept_id"])].append(
            {
                "aliasId": int(row["alias_id"]),
                "aliasUid": str(row["alias_uid"]),
                "aliasText": str(row["alias_text"]),
                "normalizedAlias": str(row["normalized_alias"]),
                "language": str(row["language"]),
                "source": str(row["source"]),
                "confidence": float(row["confidence"]),
                "createdAt": str(row["created_at"]),
            }
        )
    return dict(aliases)


def list_canonical_concepts(
    connection: sqlite3.Connection,
    *,
    concept_kind: str | None = None,
    lifecycle_status: str | None = "ACTIVE",
    query: str | None = None,
    limit: int = 500,
) -> dict[str, Any]:
    """List canonical concepts and the same aliases used by retrieval."""

    _validate_schema(connection)
    _validate_limit(limit)
    normalized_status = (
        str(lifecycle_status or "").strip().upper()
        if lifecycle_status is not None
        else None
    )
    if normalized_status == "ALL":
        normalized_status = None
    if (
        normalized_status is not None
        and normalized_status not in _CONCEPT_STATUSES
    ):
        raise ValueError("invalid concept lifecycle status")
    normalized_kind = str(concept_kind or "").strip().upper()
    needle = normalized_term(query)
    aliases = _aliases_by_concept(connection)
    rows = _rows(
        connection,
        """
        SELECT
            concept_id, concept_uid, concept_kind, canonical_name,
            normalized_name, description, lifecycle_status,
            created_at, updated_at
        FROM knowledge_concepts
        ORDER BY concept_kind, normalized_name, concept_uid
        """,
    )
    values: list[dict[str, Any]] = []
    for row in rows:
        concept_id = int(row["concept_id"])
        concept_aliases = aliases.get(concept_id, [])
        if (
            normalized_status
            and str(row["lifecycle_status"]) != normalized_status
        ):
            continue
        if (
            normalized_kind
            and str(row["concept_kind"]).upper() != normalized_kind
        ):
            continue
        if needle:
            searchable = normalized_term(
                " ".join(
                    [
                        str(row["canonical_name"]),
                        str(row["description"]),
                        *[
                            str(alias["aliasText"])
                            for alias in concept_aliases
                        ],
                    ]
                )
            )
            if needle not in searchable:
                continue
        values.append(
            {
                "conceptId": concept_id,
                "conceptUid": str(row["concept_uid"]),
                "conceptKind": str(row["concept_kind"]),
                "canonicalName": str(row["canonical_name"]),
                "normalizedName": str(row["normalized_name"]),
                "description": str(row["description"]),
                "lifecycleStatus": str(row["lifecycle_status"]),
                "createdAt": str(row["created_at"]),
                "updatedAt": str(row["updated_at"]),
                "aliases": concept_aliases,
            }
        )
        if len(values) >= limit:
            break
    return {
        "schemaVersion": CONCEPT_LIST_SCHEMA_VERSION,
        "filters": {
            "conceptKind": normalized_kind,
            "lifecycleStatus": normalized_status or "ALL",
            "query": str(query or "").strip(),
            "limit": limit,
        },
        "count": len(values),
        "concepts": values,
    }


def _candidate_concept_kind(candidate_kind: object) -> str:
    value = str(candidate_kind or "").strip().upper()
    if value == "UNIT" or value.startswith("UNIT:"):
        raise ConceptCurationError(
            "UNIT candidates must use the unit curation path"
        )
    if not value.startswith("CONCEPT:"):
        raise ConceptCurationError(
            "candidate_kind must use the CONCEPT:<kind> form"
        )
    concept_kind = value.split(":", 1)[1].strip()
    if not concept_kind:
        raise ConceptCurationError("CONCEPT candidate kind is empty")
    return concept_kind


def _request_value(value: dict[str, Any]) -> tuple[str, str]:
    encoded = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )
    return encoded, hashlib.sha256(encoded.encode("utf-8")).hexdigest()


@contextmanager
def _atomic(
    connection: sqlite3.Connection,
    name: str,
) -> Iterator[None]:
    connection.execute(f"SAVEPOINT {name}")
    try:
        yield
    except Exception:
        connection.execute(f"ROLLBACK TO SAVEPOINT {name}")
        connection.execute(f"RELEASE SAVEPOINT {name}")
        raise
    connection.execute(f"RELEASE SAVEPOINT {name}")


def _concept_by_uid(
    connection: sqlite3.Connection,
    concept_uid: str,
) -> dict[str, Any] | None:
    rows = _rows(
        connection,
        """
        SELECT
            concept_id, concept_uid, concept_kind, canonical_name,
            normalized_name, description, lifecycle_status,
            created_at, updated_at
        FROM knowledge_concepts
        WHERE concept_uid=?
        LIMIT 2
        """,
        (concept_uid,),
    )
    if len(rows) > 1:
        raise ConceptCurationError("concept UID is not unique")
    return rows[0] if rows else None


def _assert_alias_unowned(
    connection: sqlite3.Connection,
    normalized_alias: str,
    *,
    target_concept_id: int | None,
) -> None:
    parameters: list[object] = [normalized_alias, normalized_alias]
    excluded = ""
    if target_concept_id is not None:
        excluded = "AND c.concept_id<>?"
        parameters.append(target_concept_id)
    owners = _rows(
        connection,
        f"""
        SELECT DISTINCT
            c.concept_id, c.concept_uid, c.concept_kind, c.canonical_name
        FROM knowledge_concepts c
        LEFT JOIN knowledge_concept_aliases a
          ON a.concept_id=c.concept_id
        WHERE c.lifecycle_status='ACTIVE'
          AND (
              c.normalized_name=?
              OR a.normalized_alias=?
          )
          {excluded}
        ORDER BY c.concept_id
        LIMIT 2
        """,
        tuple(parameters),
    )
    if owners:
        raise ConceptCurationError(
            "normalized alias is already owned by another active concept: "
            + str(owners[0]["concept_uid"])
        )


def _upsert_alias(
    connection: sqlite3.Connection,
    *,
    concept: dict[str, Any],
    alias_text: str,
    created_at: str,
) -> dict[str, Any]:
    normalized_alias = normalized_term(alias_text)
    if not normalized_alias:
        raise ConceptCurationError("alias must not normalize to empty")
    concept_id = int(concept["concept_id"])
    _assert_alias_unowned(
        connection,
        normalized_alias,
        target_concept_id=concept_id,
    )
    existing = _rows(
        connection,
        """
        SELECT alias_id, alias_uid
        FROM knowledge_concept_aliases
        WHERE concept_id=? AND normalized_alias=?
        LIMIT 2
        """,
        (concept_id, normalized_alias),
    )
    if len(existing) > 1:
        raise ConceptCurationError(
            "canonical alias identity is not unique"
        )
    if existing:
        alias_id = int(existing[0]["alias_id"])
        alias_uid = str(existing[0]["alias_uid"])
        connection.execute(
            """
            UPDATE knowledge_concept_aliases
            SET alias_text=?, source='HUMAN_APPROVED', confidence=1
            WHERE alias_id=?
            """,
            (alias_text, alias_id),
        )
    else:
        alias_uid = stable_uid(
            "alias",
            concept["concept_uid"],
            normalized_alias,
        )
        cursor = connection.execute(
            """
            INSERT INTO knowledge_concept_aliases(
                alias_uid, concept_id, alias_text, normalized_alias,
                language, source, confidence, created_at
            ) VALUES (?, ?, ?, ?, '', 'HUMAN_APPROVED', 1, ?)
            """,
            (
                alias_uid,
                concept_id,
                alias_text,
                normalized_alias,
                created_at,
            ),
        )
        alias_id = int(cursor.lastrowid)
    return {
        "aliasId": alias_id,
        "aliasUid": alias_uid,
        "aliasText": alias_text,
        "normalizedAlias": normalized_alias,
        "source": "HUMAN_APPROVED",
        "confidence": 1.0,
    }


def _resolution_payload(
    row: dict[str, Any],
    *,
    idempotent_replay: bool,
) -> dict[str, Any]:
    concept = None
    alias = None
    if row["concept_id"] is not None:
        concept = {
            "conceptId": int(row["concept_id"]),
            "conceptUid": str(row["concept_uid"]),
            "conceptKind": str(row["concept_kind"]),
            "canonicalName": str(row["canonical_name"]),
            "normalizedName": str(row["normalized_canonical_name"]),
        }
    if row["alias_id"] is not None:
        alias = {
            "aliasId": int(row["alias_id"]),
            "aliasUid": str(row["alias_uid"]),
            "aliasText": str(row["alias_text"]),
            "normalizedAlias": str(row["normalized_alias"]),
            "source": "HUMAN_APPROVED",
            "confidence": 1.0,
        }
    return {
        "schemaVersion": CONCEPT_RESOLUTION_SCHEMA_VERSION,
        "resolutionUid": str(row["resolution_uid"]),
        "candidate": {
            "schemaCandidateId": int(row["schema_candidate_id"]),
            "candidateUid": str(row["candidate_uid"]),
            "candidateKind": str(row["candidate_kind"]),
            "normalizedValue": str(
                row["candidate_normalized_value"]
            ),
            "originalValue": str(row["candidate_original_value"]),
            "suggestedCanonicalName": str(
                row["candidate_suggested_canonical_name"]
            ),
            "status": _ACTION_STATUS[str(row["action"])],
        },
        "action": str(row["action"]),
        "concept": concept,
        "alias": alias,
        "reviewer": str(row["reviewer"]),
        "note": str(row["note"]),
        "resolvedAt": str(row["resolved_at"]),
        "idempotentReplay": idempotent_replay,
    }


def _resolution_row(
    connection: sqlite3.Connection,
    candidate_id: int,
) -> dict[str, Any] | None:
    rows = _rows(
        connection,
        """
        SELECT *
        FROM knowledge_concept_resolution_history
        WHERE schema_candidate_id=?
        LIMIT 2
        """,
        (candidate_id,),
    )
    if len(rows) > 1:
        raise ConceptCurationError(
            "candidate has conflicting resolution history"
        )
    return rows[0] if rows else None


def resolve_schema_candidate(
    connection: sqlite3.Connection,
    *,
    candidate_uid: str,
    action: str,
    reviewer: str,
    note: str,
    now_iso: Callable[[], str],
    canonical_name: str = "",
    concept_uid: str = "",
    alias: str = "",
) -> dict[str, Any]:
    """Atomically resolve exactly one OPEN ``CONCEPT:*`` candidate."""

    _validate_schema(connection)
    candidate_identifier = _text(candidate_uid, "candidate_uid")
    normalized_action = _text(action, "action").upper()
    if normalized_action not in _RESOLUTION_ACTIONS:
        raise ConceptCurationError(
            "action must be CREATE, MERGE, or REJECT"
        )
    reviewer_value = _text(reviewer, "reviewer")
    note_value = _text(note, "note")
    canonical_value = str(canonical_name or "").strip()
    target_uid = str(concept_uid or "").strip()
    alias_value = str(alias or "").strip()
    candidates = _rows(
        connection,
        """
        SELECT *
        FROM knowledge_schema_candidates
        WHERE candidate_uid=?
        LIMIT 2
        """,
        (candidate_identifier,),
    )
    if not candidates:
        raise ConceptCurationError("unknown schema candidate UID")
    if len(candidates) > 1:
        raise ConceptCurationError("schema candidate UID is not unique")
    candidate = candidates[0]
    candidate_id = int(candidate["schema_candidate_id"])
    concept_kind = _candidate_concept_kind(candidate["candidate_kind"])

    if normalized_action == "CREATE":
        canonical_value = _text(canonical_value, "canonical_name")
        alias_value = _text(alias_value, "alias")
        if target_uid:
            raise ConceptCurationError(
                "CREATE must not specify an existing concept_uid"
            )
    elif normalized_action == "MERGE":
        target_uid = _text(target_uid, "concept_uid")
        alias_value = _text(alias_value, "alias")
        if canonical_value:
            raise ConceptCurationError(
                "MERGE uses the existing canonical name"
            )
    else:
        if canonical_value or target_uid or alias_value:
            raise ConceptCurationError(
                "REJECT must not specify concept or alias fields"
            )

    request = {
        "candidateUid": candidate_identifier,
        "action": normalized_action,
        "canonicalName": canonical_value,
        "conceptUid": target_uid,
        "alias": alias_value,
        "reviewer": reviewer_value,
        "note": note_value,
    }
    request_json, request_sha256 = _request_value(request)
    expected_status = _ACTION_STATUS[normalized_action]

    with _atomic(connection, "human_concept_resolution"):
        existing_resolution = _resolution_row(
            connection,
            candidate_id,
        )
        if existing_resolution is not None:
            if (
                str(existing_resolution["request_sha256"])
                == request_sha256
                and str(existing_resolution["request_json"])
                == request_json
                and str(existing_resolution["action"])
                == normalized_action
                and str(candidate["status"]) == expected_status
            ):
                return _resolution_payload(
                    existing_resolution,
                    idempotent_replay=True,
                )
            raise ConceptCurationError(
                "conflicting repeated resolution for schema candidate"
            )
        if str(candidate["status"]) != "OPEN":
            raise ConceptCurationError(
                "schema candidate status must be OPEN"
            )

        concept: dict[str, Any] | None = None
        alias_value_result: dict[str, Any] | None = None
        resolved_at = _text(now_iso(), "resolved_at")
        if normalized_action == "CREATE":
            normalized_canonical = normalized_term(canonical_value)
            normalized_alias = normalized_term(alias_value)
            if not normalized_canonical:
                raise ConceptCurationError(
                    "canonical_name must not normalize to empty"
                )
            if not normalized_alias:
                raise ConceptCurationError(
                    "alias must not normalize to empty"
                )
            duplicate = connection.execute(
                """
                SELECT concept_uid
                FROM knowledge_concepts
                WHERE concept_kind=? AND normalized_name=?
                LIMIT 1
                """,
                (concept_kind, normalized_canonical),
            ).fetchone()
            if duplicate is not None:
                raise ConceptCurationError(
                    "canonical concept already exists; use MERGE: "
                    + str(duplicate[0])
                )
            _assert_alias_unowned(
                connection,
                normalized_canonical,
                target_concept_id=None,
            )
            _assert_alias_unowned(
                connection,
                normalized_alias,
                target_concept_id=None,
            )
            created_uid = stable_uid(
                "concept",
                concept_kind,
                normalized_canonical,
            )
            cursor = connection.execute(
                """
                INSERT INTO knowledge_concepts(
                    concept_uid, concept_kind, canonical_name,
                    normalized_name, description, lifecycle_status,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, '', 'ACTIVE', ?, ?)
                """,
                (
                    created_uid,
                    concept_kind,
                    canonical_value,
                    normalized_canonical,
                    resolved_at,
                    resolved_at,
                ),
            )
            concept = {
                "concept_id": int(cursor.lastrowid),
                "concept_uid": created_uid,
                "concept_kind": concept_kind,
                "canonical_name": canonical_value,
                "normalized_name": normalized_canonical,
                "lifecycle_status": "ACTIVE",
            }
            alias_value_result = _upsert_alias(
                connection,
                concept=concept,
                alias_text=alias_value,
                created_at=resolved_at,
            )
        elif normalized_action == "MERGE":
            concept = _concept_by_uid(connection, target_uid)
            if concept is None:
                raise ConceptCurationError("unknown canonical concept UID")
            if str(concept["lifecycle_status"]) != "ACTIVE":
                raise ConceptCurationError(
                    "target canonical concept must be ACTIVE"
                )
            _text(
                concept["canonical_name"],
                "target canonical concept name",
            )
            if str(concept["concept_kind"]).upper() != concept_kind:
                raise ConceptCurationError(
                    "candidate and concept kind mismatch"
                )
            alias_value_result = _upsert_alias(
                connection,
                concept=concept,
                alias_text=alias_value,
                created_at=resolved_at,
            )

        cursor = connection.execute(
            """
            UPDATE knowledge_schema_candidates
            SET status=?, last_seen_at=?
            WHERE schema_candidate_id=? AND status='OPEN'
            """,
            (expected_status, resolved_at, candidate_id),
        )
        if int(cursor.rowcount) != 1:
            raise ConceptCurationError(
                "schema candidate changed during resolution"
            )
        resolution_uid = stable_uid(
            "concept-resolution",
            candidate_identifier,
            request_sha256,
        )
        connection.execute(
            """
            INSERT INTO knowledge_concept_resolution_history(
                resolution_uid, schema_candidate_id, candidate_uid,
                candidate_kind, candidate_normalized_value,
                candidate_original_value,
                candidate_suggested_canonical_name, action,
                concept_id, concept_uid, concept_kind, canonical_name,
                normalized_canonical_name, alias_id, alias_uid,
                alias_text, normalized_alias, reviewer, note,
                request_json, request_sha256, resolved_at
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?
            )
            """,
            (
                resolution_uid,
                candidate_id,
                str(candidate["candidate_uid"]),
                str(candidate["candidate_kind"]),
                str(candidate["normalized_value"]),
                str(candidate["original_value"]),
                str(candidate["suggested_canonical_name"]),
                normalized_action,
                None if concept is None else int(concept["concept_id"]),
                "" if concept is None else str(concept["concept_uid"]),
                "" if concept is None else str(concept["concept_kind"]),
                "" if concept is None else str(concept["canonical_name"]),
                "" if concept is None else str(concept["normalized_name"]),
                (
                    None
                    if alias_value_result is None
                    else int(alias_value_result["aliasId"])
                ),
                (
                    ""
                    if alias_value_result is None
                    else str(alias_value_result["aliasUid"])
                ),
                (
                    ""
                    if alias_value_result is None
                    else str(alias_value_result["aliasText"])
                ),
                (
                    ""
                    if alias_value_result is None
                    else str(alias_value_result["normalizedAlias"])
                ),
                reviewer_value,
                note_value,
                request_json,
                request_sha256,
                resolved_at,
            ),
        )
        stored = _resolution_row(connection, candidate_id)
        if stored is None:
            raise ConceptCurationError(
                "concept resolution history was not persisted"
            )
        return _resolution_payload(stored, idempotent_replay=False)


def _alias_approval_payload(
    row: dict[str, Any],
    *,
    idempotent_replay: bool,
) -> dict[str, Any]:
    return {
        "schemaVersion": CONCEPT_ALIAS_APPROVAL_SCHEMA_VERSION,
        "approvalUid": str(row["approval_uid"]),
        "concept": {
            "conceptId": int(row["concept_id"]),
            "conceptUid": str(row["concept_uid"]),
            "conceptKind": str(row["concept_kind"]),
            "canonicalName": str(row["canonical_name"]),
        },
        "alias": {
            "aliasId": int(row["alias_id"]),
            "aliasUid": str(row["alias_uid"]),
            "aliasText": str(row["alias_text"]),
            "normalizedAlias": str(row["normalized_alias"]),
            "source": "HUMAN_APPROVED",
            "confidence": 1.0,
        },
        "reviewer": str(row["reviewer"]),
        "note": str(row["note"]),
        "approvedAt": str(row["approved_at"]),
        "idempotentReplay": idempotent_replay,
    }


def upsert_human_concept_alias(
    connection: sqlite3.Connection,
    *,
    concept_uid: str,
    alias: str,
    reviewer: str,
    note: str,
    now_iso: Callable[[], str],
) -> dict[str, Any]:
    """Upsert one human-approved alias with immutable approval history."""

    _validate_schema(connection)
    concept_identifier = _text(concept_uid, "concept_uid")
    alias_value = _text(alias, "alias")
    reviewer_value = _text(reviewer, "reviewer")
    note_value = _text(note, "note")
    normalized_alias = normalized_term(alias_value)
    if not normalized_alias:
        raise ConceptCurationError("alias must not normalize to empty")
    request = {
        "conceptUid": concept_identifier,
        "alias": alias_value,
        "reviewer": reviewer_value,
        "note": note_value,
    }
    request_json, request_sha256 = _request_value(request)

    with _atomic(connection, "human_concept_alias"):
        replay_rows = _rows(
            connection,
            """
            SELECT *
            FROM knowledge_concept_alias_approval_history
            WHERE request_sha256=?
            LIMIT 2
            """,
            (request_sha256,),
        )
        if replay_rows:
            if (
                len(replay_rows) == 1
                and str(replay_rows[0]["request_json"]) == request_json
            ):
                return _alias_approval_payload(
                    replay_rows[0],
                    idempotent_replay=True,
                )
            raise ConceptCurationError(
                "conflicting alias approval request fingerprint"
            )
        concept = _concept_by_uid(connection, concept_identifier)
        if concept is None:
            raise ConceptCurationError("unknown canonical concept UID")
        if str(concept["lifecycle_status"]) != "ACTIVE":
            raise ConceptCurationError(
                "target canonical concept must be ACTIVE"
            )
        _text(
            concept["canonical_name"],
            "target canonical concept name",
        )
        approved_at = _text(now_iso(), "approved_at")
        alias_value_result = _upsert_alias(
            connection,
            concept=concept,
            alias_text=alias_value,
            created_at=approved_at,
        )
        approval_uid = stable_uid(
            "concept-alias-approval",
            request_sha256,
        )
        connection.execute(
            """
            INSERT INTO knowledge_concept_alias_approval_history(
                approval_uid, concept_id, concept_uid, concept_kind,
                canonical_name, alias_id, alias_uid, alias_text,
                normalized_alias, reviewer, note, request_json,
                request_sha256, approved_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                approval_uid,
                int(concept["concept_id"]),
                str(concept["concept_uid"]),
                str(concept["concept_kind"]),
                str(concept["canonical_name"]),
                int(alias_value_result["aliasId"]),
                str(alias_value_result["aliasUid"]),
                str(alias_value_result["aliasText"]),
                str(alias_value_result["normalizedAlias"]),
                reviewer_value,
                note_value,
                request_json,
                request_sha256,
                approved_at,
            ),
        )
        stored = _rows(
            connection,
            """
            SELECT *
            FROM knowledge_concept_alias_approval_history
            WHERE request_sha256=?
            LIMIT 1
            """,
            (request_sha256,),
        )[0]
        return _alias_approval_payload(
            stored,
            idempotent_replay=False,
        )


@contextmanager
def connect_curation_readonly(
    database_path: str | Path,
) -> Iterator[sqlite3.Connection]:
    path = Path(database_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    connection = sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    try:
        connection.execute("PRAGMA query_only=ON")
        yield connection
    finally:
        connection.close()


__all__ = [
    "CANDIDATE_LIST_SCHEMA_VERSION",
    "CONCEPT_ALIAS_APPROVAL_SCHEMA_VERSION",
    "CONCEPT_CURATION_MIGRATION",
    "CONCEPT_LIST_SCHEMA_VERSION",
    "CONCEPT_RESOLUTION_SCHEMA_VERSION",
    "ConceptCurationError",
    "connect_curation_readonly",
    "ensure_concept_curation_schema",
    "list_canonical_concepts",
    "list_schema_candidates",
    "resolve_schema_candidate",
    "upsert_human_concept_alias",
]
