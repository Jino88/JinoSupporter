"""Read-only, revision-safe evidence detail retrieval.

The public ``EVD-*`` identifier is resolved against the canonical evidence
layer and then followed through its explicit ``capture_v2_revision_id``
bridge.  The implementation never falls back to another source revision or
sheet.  Images are intentionally outside this contract; only Capture v2
tabular cells and structural metadata are returned.
"""

from __future__ import annotations

import json
import re
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator


EVIDENCE_DETAIL_SCHEMA_VERSION = "canonical-evidence-detail-v1"
PUBLIC_EVIDENCE_ID_PATTERN = re.compile(r"^EVD-[A-Z0-9][A-Z0-9-]*$", re.IGNORECASE)
REQUIRED_TABLES = {
    "source_documents",
    "source_revisions",
    "evidence_items",
    "entity_evidence_links",
    "capture_v2_documents",
    "capture_v2_revisions",
    "capture_v2_workbooks",
    "capture_v2_sheets",
    "capture_v2_cells",
    "capture_v2_merged_ranges",
    "capture_v2_row_dimensions",
    "capture_v2_column_dimensions",
    "workbook_analyses",
    "knowledge_studies",
    "knowledge_study_contexts",
    "knowledge_factors",
    "knowledge_arms",
    "knowledge_outcomes",
    "knowledge_observations",
    "knowledge_comparisons",
    "knowledge_effects",
    "knowledge_claims",
}


class EvidenceDetailError(RuntimeError):
    """Base error for evidence-detail requests."""


class InvalidEvidenceIdError(EvidenceDetailError):
    """The supplied value is not a stable public EVD identifier."""


class EvidenceNotFoundError(EvidenceDetailError):
    """No evidence item has the requested public identifier."""


class AmbiguousEvidenceIdError(EvidenceDetailError):
    """More than one case-insensitive EVD match exists."""


class EvidenceTrustError(EvidenceDetailError):
    """The evidence cannot be bound to one current Capture v2 revision."""


def _rows(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> list[dict[str, Any]]:
    cursor = connection.execute(sql, tuple(parameters))
    names = [str(column[0]) for column in cursor.description or ()]
    return [dict(zip(names, row)) for row in cursor.fetchall()]


def _first(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> dict[str, Any] | None:
    rows = _rows(connection, sql, parameters)
    return rows[0] if rows else None


def _json_value(value: object, default: Any) -> Any:
    if value is None:
        return default
    try:
        return json.loads(str(value))
    except (TypeError, ValueError) as exc:
        raise EvidenceTrustError("Capture v2 contains invalid stored JSON.") from exc


def _column_name(column_index: int) -> str:
    if column_index < 1:
        raise ValueError("Excel column indexes are one-based.")
    letters: list[str] = []
    current = column_index
    while current:
        current, remainder = divmod(current - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def _column_index(column_name: str) -> int:
    index = 0
    for character in column_name.upper():
        if not "A" <= character <= "Z":
            raise ValueError("Invalid Excel column name.")
        index = index * 26 + ord(character) - ord("A") + 1
    return index


def _a1_range(
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> str:
    start = f"{_column_name(start_col)}{start_row}"
    end = f"{_column_name(end_col)}{end_row}"
    return start if start == end else f"{start}:{end}"


def _canonical_a1(value: object) -> str:
    match = re.fullmatch(
        r"\s*\$?([A-Z]+)\$?([1-9]\d*)"
        r"(?::\$?([A-Z]+)\$?([1-9]\d*))?\s*",
        str(value or ""),
        re.IGNORECASE,
    )
    if match is None:
        raise EvidenceTrustError("Evidence contains an invalid stored A1 range.")
    start_col = _column_index(match.group(1))
    start_row = int(match.group(2))
    end_col = _column_index(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_row < start_row or end_col < start_col:
        raise EvidenceTrustError("Evidence contains a reversed stored A1 range.")
    return _a1_range(start_row, start_col, end_row, end_col)


@contextmanager
def connect_evidence_readonly(
    database_path: str | Path,
) -> Iterator[sqlite3.Connection]:
    """Open SQLite with both URI-level and connection-level write prevention."""

    path = Path(database_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    connection = sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True)
    try:
        connection.execute("PRAGMA query_only=ON")
        yield connection
    finally:
        connection.close()


def _validate_schema(connection: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )
    }
    missing = sorted(REQUIRED_TABLES - tables)
    if missing:
        raise EvidenceTrustError(
            "Evidence detail DB is missing required tables: " + ", ".join(missing)
        )
    revision_columns = {
        str(row[1])
        for row in connection.execute("PRAGMA table_info(source_revisions)")
    }
    if "capture_v2_revision_id" not in revision_columns:
        raise EvidenceTrustError(
            "Canonical source_revisions is missing capture_v2_revision_id."
        )


def _resolve_evidence(
    connection: sqlite3.Connection,
    public_evidence_id: str,
) -> dict[str, Any]:
    identifier = str(public_evidence_id or "").strip()
    if not PUBLIC_EVIDENCE_ID_PATTERN.fullmatch(identifier):
        raise InvalidEvidenceIdError(
            "Evidence ID must be one stable public EVD-* identifier."
        )
    rows = _rows(
        connection,
        """
        SELECT
            e.evidence_id, e.evidence_uid, e.public_evidence_id,
            e.revision_id AS evidence_revision_id, e.legacy_evidence_id,
            e.evidence_kind, e.sheet_name, e.start_row, e.start_col,
            e.end_row, e.end_col, e.range_address, e.evidence_role,
            e.source_text, e.note,
            e.content_sha256 AS evidence_content_sha256,
            e.verification_status AS evidence_verification_status,
            e.created_at,
            r.revision_uid, r.document_id AS canonical_document_id,
            r.source_fingerprint, r.fingerprint_kind,
            r.content_sha256 AS canonical_content_sha256,
            r.size_bytes, r.mtime_ns, r.extractor_name, r.extractor_version,
            r.capture_contract AS canonical_capture_contract,
            r.capture_status AS canonical_capture_status,
            r.is_current AS canonical_is_current, r.captured_at,
            r.capture_v2_revision_id,
            d.document_uid, d.dataset, d.source_path, d.original_file_name,
            d.source_kind, d.lifecycle_status
        FROM evidence_items e
        JOIN source_revisions r ON r.revision_id=e.revision_id
        JOIN source_documents d ON d.document_id=r.document_id
        WHERE LOWER(e.public_evidence_id)=LOWER(?)
        ORDER BY e.evidence_id
        """,
        (identifier,),
    )
    if not rows:
        raise EvidenceNotFoundError(f"Evidence ID does not exist: {identifier}")
    if len(rows) != 1:
        raise AmbiguousEvidenceIdError(
            f"Evidence ID is ambiguous under case-insensitive matching: {identifier}"
        )
    return rows[0]


def _resolve_capture_revision(
    connection: sqlite3.Connection,
    evidence: dict[str, Any],
) -> dict[str, Any]:
    canonical_revision_id = int(evidence["evidence_revision_id"])
    current_rows = _rows(
        connection,
        """
        SELECT revision_id, revision_uid
        FROM source_revisions
        WHERE document_id=? AND is_current=1
        ORDER BY revision_id
        """,
        (int(evidence["canonical_document_id"]),),
    )
    if len(current_rows) != 1:
        raise EvidenceTrustError(
            "Canonical source document does not have exactly one current revision."
        )
    if (
        not bool(evidence["canonical_is_current"])
        or int(current_rows[0]["revision_id"]) != canonical_revision_id
    ):
        raise EvidenceTrustError(
            "Evidence belongs to a stale canonical revision; no current-revision "
            "fallback is permitted."
        )
    if str(evidence["lifecycle_status"]) != "ACTIVE":
        raise EvidenceTrustError("Evidence source document is not active.")
    if str(evidence["canonical_capture_status"]) != "CAPTURED":
        raise EvidenceTrustError(
            "Evidence canonical revision is not in CAPTURED state."
        )
    capture_revision_id = evidence["capture_v2_revision_id"]
    if capture_revision_id is None:
        raise EvidenceTrustError(
            "Evidence revision has no explicit Capture v2 revision bridge."
        )

    capture = _first(
        connection,
        """
        SELECT
            r.revision_id, r.revision_uid, r.document_id,
            r.content_sha256, r.capture_contract, r.extractor_name,
            r.extractor_version, r.size_bytes, r.mtime_ns,
            r.capture_status, r.is_current, r.capture_json_sha256,
            r.captured_at, r.stale_at,
            d.source_path, d.file_name, d.source_kind,
            w.workbook_status, w.is_truly_empty, w.sheet_count,
            w.nonempty_sheet_count, w.tabular_sheet_count
        FROM capture_v2_revisions r
        JOIN capture_v2_documents d ON d.document_id=r.document_id
        JOIN capture_v2_workbooks w ON w.revision_id=r.revision_id
        WHERE r.revision_id=?
        """,
        (int(capture_revision_id),),
    )
    if capture is None:
        raise EvidenceTrustError("Explicit Capture v2 revision bridge is missing.")
    capture_current_rows = _rows(
        connection,
        """
        SELECT revision_id, revision_uid
        FROM capture_v2_revisions
        WHERE document_id=? AND is_current=1
        ORDER BY revision_id
        """,
        (int(capture["document_id"]),),
    )
    if len(capture_current_rows) != 1:
        raise EvidenceTrustError(
            "Capture v2 source document does not have exactly one current revision."
        )
    if (
        not bool(capture["is_current"])
        or int(capture_current_rows[0]["revision_id"]) != int(capture["revision_id"])
    ):
        raise EvidenceTrustError(
            "Evidence belongs to a stale Capture v2 revision; no current-revision "
            "fallback is permitted."
        )
    if str(capture["capture_status"]) != "CAPTURED":
        raise EvidenceTrustError("Capture v2 revision is not in CAPTURED state.")

    equality_checks = {
        "revisionUid": (
            str(evidence["revision_uid"]),
            str(capture["revision_uid"]),
        ),
        "contentSha256": (
            str(evidence["canonical_content_sha256"]),
            str(capture["content_sha256"]),
        ),
        "captureContract": (
            str(evidence["canonical_capture_contract"]),
            str(capture["capture_contract"]),
        ),
        "sourcePath": (
            str(Path(str(evidence["source_path"])).expanduser().resolve()),
            str(Path(str(capture["source_path"])).expanduser().resolve()),
        ),
    }
    mismatches = [
        name
        for name, (canonical_value, capture_value) in equality_checks.items()
        if canonical_value != capture_value
    ]
    evidence_hash = str(evidence["evidence_content_sha256"] or "")
    if (
        evidence_hash
        and evidence_hash != str(evidence["canonical_content_sha256"])
    ):
        mismatches.append("evidenceContentSha256")
    if mismatches:
        raise EvidenceTrustError(
            "Canonical/Capture v2 revision binding mismatch: "
            + ", ".join(mismatches)
        )
    return capture


def _resolve_sheet(
    connection: sqlite3.Connection,
    capture_revision_id: int,
    sheet_name: str,
) -> dict[str, Any]:
    sheets = _rows(
        connection,
        """
        SELECT
            sheet_id, sheet_index, title, sheet_state, capture_status,
            is_truly_empty, has_tabular_evidence, nonempty_cell_count,
            structural_cell_count, captured_cell_count, formula_cell_count,
            merge_count, used_bounds_json, content_bounds_json,
            freeze_panes, auto_filter, metadata_json
        FROM capture_v2_sheets
        WHERE revision_id=? AND title=?
        ORDER BY sheet_index
        """,
        (capture_revision_id, sheet_name),
    )
    if not sheets:
        raise EvidenceTrustError(
            "Evidence sheet does not exist in its exact Capture v2 revision."
        )
    if len(sheets) != 1:
        raise EvidenceTrustError(
            "Evidence sheet is ambiguous in its exact Capture v2 revision."
        )
    return sheets[0]


def _row_dimensions(
    connection: sqlite3.Connection,
    sheet_id: int,
    start_row: int,
    end_row: int,
) -> list[dict[str, Any]]:
    return [
        {
            "row": int(row["row_index"]),
            "height": row["height"],
            "hidden": bool(row["hidden"]),
        }
        for row in _rows(
            connection,
            """
            SELECT row_index, height, hidden
            FROM capture_v2_row_dimensions
            WHERE sheet_id=? AND row_index BETWEEN ? AND ?
            ORDER BY row_index
            """,
            (sheet_id, start_row, end_row),
        )
    ]


def _column_dimensions(
    connection: sqlite3.Connection,
    sheet_id: int,
    start_col: int,
    end_col: int,
) -> list[dict[str, Any]]:
    return [
        {
            "key": str(row["dimension_key"]),
            "minColumn": int(row["min_column"]),
            "maxColumn": int(row["max_column"]),
            "width": row["width"],
            "hidden": bool(row["hidden"]),
        }
        for row in _rows(
            connection,
            """
            SELECT dimension_key, min_column, max_column, width, hidden
            FROM capture_v2_column_dimensions
            WHERE sheet_id=? AND max_column>=? AND min_column<=?
            ORDER BY min_column, max_column, dimension_key
            """,
            (sheet_id, start_col, end_col),
        )
    ]


def _cell_record(
    row: dict[str, Any],
    *,
    row_hidden: bool,
    column_hidden: bool,
) -> dict[str, Any]:
    row_index = int(row["row_index"])
    column_index = int(row["column_index"])
    coordinate = str(row["coordinate"])
    expected_coordinate = _a1_range(
        row_index, column_index, row_index, column_index
    )
    if coordinate.replace("$", "").upper() != expected_coordinate:
        raise EvidenceTrustError(
            "Capture v2 cell coordinate does not match its row/column indexes."
        )
    return {
        "row": row_index,
        "column": column_index,
        "coordinate": coordinate,
        "rawValue": _json_value(row["raw_value_json"], None),
        "formula": row["formula_text"],
        "cachedValue": _json_value(row["cached_value_json"], None),
        "displayValue": _json_value(row["display_value_json"], None),
        "dataType": str(row["data_type"]),
        "cachedDataType": row["cached_data_type"],
        "numberFormat": str(row["number_format"]),
        "styleId": int(row["style_id"]),
        "style": _json_value(row["style_json"], {}),
        "mergeRange": row["merge_range"],
        "mergeRole": str(row["merge_role"]),
        "rowHidden": row_hidden,
        "columnHidden": column_hidden,
    }


def _cell_at(
    connection: sqlite3.Connection,
    sheet_id: int,
    row_index: int,
    column_index: int,
) -> dict[str, Any] | None:
    row = _first(
        connection,
        """
        SELECT *
        FROM capture_v2_cells
        WHERE sheet_id=? AND row_index=? AND column_index=?
        """,
        (sheet_id, row_index, column_index),
    )
    if row is None:
        return None
    row_dimension = _first(
        connection,
        """
        SELECT hidden FROM capture_v2_row_dimensions
        WHERE sheet_id=? AND row_index=?
        """,
        (sheet_id, row_index),
    )
    column_dimension = _first(
        connection,
        """
        SELECT MAX(hidden) AS hidden FROM capture_v2_column_dimensions
        WHERE sheet_id=? AND min_column<=? AND max_column>=?
        """,
        (sheet_id, column_index, column_index),
    )
    return _cell_record(
        row,
        row_hidden=bool(row_dimension and row_dimension["hidden"]),
        column_hidden=bool(column_dimension and column_dimension["hidden"]),
    )


def _preview(
    connection: sqlite3.Connection,
    evidence: dict[str, Any],
    sheet: dict[str, Any],
) -> dict[str, Any]:
    sheet_id = int(sheet["sheet_id"])
    start_row = int(evidence["start_row"])
    start_col = int(evidence["start_col"])
    end_row = int(evidence["end_row"])
    end_col = int(evidence["end_col"])
    row_dimensions = _row_dimensions(
        connection, sheet_id, start_row, end_row
    )
    column_dimensions = _column_dimensions(
        connection, sheet_id, start_col, end_col
    )
    hidden_rows = {
        int(item["row"]) for item in row_dimensions if bool(item["hidden"])
    }
    hidden_column_ranges = [
        (int(item["minColumn"]), int(item["maxColumn"]))
        for item in column_dimensions
        if bool(item["hidden"])
    ]

    cell_rows = _rows(
        connection,
        """
        SELECT *
        FROM capture_v2_cells
        WHERE sheet_id=?
          AND row_index BETWEEN ? AND ?
          AND column_index BETWEEN ? AND ?
        ORDER BY row_index, column_index
        """,
        (sheet_id, start_row, end_row, start_col, end_col),
    )
    cells = [
        _cell_record(
            row,
            row_hidden=int(row["row_index"]) in hidden_rows,
            column_hidden=any(
                minimum <= int(row["column_index"]) <= maximum
                for minimum, maximum in hidden_column_ranges
            ),
        )
        for row in cell_rows
    ]

    merge_rows = _rows(
        connection,
        """
        SELECT
            address, min_row, min_column, max_row, max_column,
            anchor_coordinate
        FROM capture_v2_merged_ranges
        WHERE sheet_id=?
          AND max_row>=? AND min_row<=?
          AND max_column>=? AND min_column<=?
        ORDER BY min_row, min_column, max_row, max_column, address
        """,
        (sheet_id, start_row, end_row, start_col, end_col),
    )
    merged_ranges: list[dict[str, Any]] = []
    for merge in merge_rows:
        anchor_row = int(merge["min_row"])
        anchor_column = int(merge["min_column"])
        canonical_merge_address = _a1_range(
            anchor_row,
            anchor_column,
            int(merge["max_row"]),
            int(merge["max_column"]),
        )
        if _canonical_a1(merge["address"]) != canonical_merge_address:
            raise EvidenceTrustError(
                "Capture v2 merge address does not match its numeric coordinates."
            )
        expected_anchor = _a1_range(
            anchor_row, anchor_column, anchor_row, anchor_column
        )
        if (
            str(merge["anchor_coordinate"]).replace("$", "").upper()
            != expected_anchor
        ):
            raise EvidenceTrustError(
                "Capture v2 merge anchor does not match its numeric coordinates."
            )
        anchor_outside = not (
            start_row <= anchor_row <= end_row
            and start_col <= anchor_column <= end_col
        )
        merged_ranges.append(
            {
                "address": str(merge["address"]),
                "minRow": anchor_row,
                "minColumn": anchor_column,
                "maxRow": int(merge["max_row"]),
                "maxColumn": int(merge["max_column"]),
                "anchorCoordinate": str(merge["anchor_coordinate"]),
                "anchorOutsideEvidenceRange": anchor_outside,
                # An external anchor is returned only as merge context.  It is
                # never added to ``cells`` or represented as cited evidence.
                "anchorCell": (
                    _cell_at(
                        connection,
                        sheet_id,
                        anchor_row,
                        anchor_column,
                    )
                    if anchor_outside
                    else None
                ),
            }
        )

    return {
        "sheet": {
            "sheetIndex": int(sheet["sheet_index"]),
            "title": str(sheet["title"]),
            "state": str(sheet["sheet_state"]),
            "captureStatus": str(sheet["capture_status"]),
            "isTrulyEmpty": bool(sheet["is_truly_empty"]),
            "hasTabularEvidence": bool(sheet["has_tabular_evidence"]),
            "nonEmptyCellCount": int(sheet["nonempty_cell_count"]),
            "structuralCellCount": int(sheet["structural_cell_count"]),
            "capturedCellCount": int(sheet["captured_cell_count"]),
            "formulaCellCount": int(sheet["formula_cell_count"]),
            "mergeCount": int(sheet["merge_count"]),
            "usedBounds": _json_value(sheet["used_bounds_json"], None),
            "contentBounds": _json_value(sheet["content_bounds_json"], None),
            "freezePanes": sheet["freeze_panes"],
            "autoFilter": sheet["auto_filter"],
            "metadata": _json_value(sheet["metadata_json"], {}),
        },
        "range": {
            "sheet": str(evidence["sheet_name"]),
            "a1": _a1_range(start_row, start_col, end_row, end_col),
            "storedA1": str(evidence["range_address"]),
            "start": {"row": start_row, "column": start_col},
            "end": {"row": end_row, "column": end_col},
        },
        "cells": cells,
        "capturedCellCountInRange": len(cells),
        "mergedRanges": merged_ranges,
        "rowDimensions": row_dimensions,
        "columnDimensions": column_dimensions,
        "imagesAnalyzed": False,
    }


ENTITY_QUERIES: dict[str, str] = {
    "WORKBOOK_ANALYSIS": """
        SELECT public_analysis_id AS public_id, title AS label,
               verification_status
        FROM workbook_analyses WHERE analysis_uid=?
    """,
    "STUDY": """
        SELECT public_data_id AS public_id, title AS label,
               verification_status
        FROM knowledge_studies WHERE study_uid=?
    """,
    "CONTEXT": """
        SELECT '' AS public_id, original_value AS label, verification_status
        FROM knowledge_study_contexts WHERE context_uid=?
    """,
    "FACTOR": """
        SELECT '' AS public_id, original_label AS label, verification_status
        FROM knowledge_factors WHERE factor_uid=?
    """,
    "ARM": """
        SELECT '' AS public_id, label, verification_status
        FROM knowledge_arms WHERE arm_uid=?
    """,
    "OUTCOME": """
        SELECT '' AS public_id, original_label AS label, verification_status
        FROM knowledge_outcomes WHERE outcome_uid=?
    """,
    "OBSERVATION": """
        SELECT '' AS public_id,
               COALESCE(NULLIF(value_text, ''), CAST(value_number AS TEXT)) AS label,
               verification_status
        FROM knowledge_observations WHERE observation_uid=?
    """,
    "COMPARISON": """
        SELECT public_comparison_id AS public_id, summary_text AS label,
               verification_status
        FROM knowledge_comparisons WHERE comparison_uid=?
    """,
    "EFFECT": """
        SELECT public_effect_id AS public_id, effect_type AS label,
               verification_status
        FROM knowledge_effects WHERE effect_uid=?
    """,
    "CLAIM": """
        SELECT public_claim_id AS public_id, claim_text AS label,
               verification_status
        FROM knowledge_claims WHERE claim_uid=?
    """,
}


def _linked_entities(
    connection: sqlite3.Connection,
    evidence_id: int,
) -> list[dict[str, Any]]:
    links = _rows(
        connection,
        """
        SELECT
            entity_type, entity_uid, evidence_role, claim_scope,
            entity_evidence_link_id
        FROM entity_evidence_links
        WHERE evidence_id=?
        ORDER BY entity_type, entity_uid, evidence_role,
                 entity_evidence_link_id
        """,
        (evidence_id,),
    )
    result: list[dict[str, Any]] = []
    for link in links:
        entity_type = str(link["entity_type"]).upper()
        entity_uid = str(link["entity_uid"])
        entity = None
        query = ENTITY_QUERIES.get(entity_type)
        if query is not None:
            entity = _first(connection, query, (entity_uid,))
        result.append(
            {
                "entityType": entity_type,
                "entityUid": entity_uid,
                "publicId": str(entity["public_id"]) if entity else "",
                "label": str(entity["label"] or "") if entity else "",
                "verificationStatus": (
                    str(entity["verification_status"]) if entity else ""
                ),
                "exists": entity is not None,
                "evidenceRole": str(link["evidence_role"]),
                "claimScope": str(link["claim_scope"]),
            }
        )
    return result


def build_evidence_detail(
    connection: sqlite3.Connection,
    public_evidence_id: str,
) -> dict[str, Any]:
    """Return one current, explicitly bridged Capture v2 evidence preview."""

    _validate_schema(connection)
    evidence = _resolve_evidence(connection, public_evidence_id)
    capture = _resolve_capture_revision(connection, evidence)
    exact_a1 = _a1_range(
        int(evidence["start_row"]),
        int(evidence["start_col"]),
        int(evidence["end_row"]),
        int(evidence["end_col"]),
    )
    stored_a1 = _canonical_a1(evidence["range_address"])
    if stored_a1 != exact_a1:
        raise EvidenceTrustError(
            "Evidence stored A1 range does not match its numeric coordinates."
        )
    sheet = _resolve_sheet(
        connection,
        int(capture["revision_id"]),
        str(evidence["sheet_name"]),
    )
    verification_status = str(evidence["evidence_verification_status"])
    if verification_status == "STALE":
        raise EvidenceTrustError("Evidence item is marked STALE.")
    trusted = verification_status == "VERIFIED"
    return {
        "schemaVersion": EVIDENCE_DETAIL_SCHEMA_VERSION,
        "publicEvidenceId": str(evidence["public_evidence_id"]),
        "evidenceUid": str(evidence["evidence_uid"]),
        "trust": {
            "status": (
                "CURRENT_CAPTURE_VERIFIED"
                if trusted
                else "CURRENT_CAPTURE_UNVERIFIED"
            ),
            "trusted": trusted,
            "sourceActive": True,
            "canonicalRevisionCurrent": True,
            "captureRevisionCurrent": True,
            "revisionUidMatches": True,
            "contentSha256Matches": True,
            "captureContractMatches": True,
            "exactRevisionFallbackUsed": False,
            "evidenceVerificationStatus": verification_status,
        },
        "source": {
            "documentId": int(evidence["canonical_document_id"]),
            "documentUid": str(evidence["document_uid"]),
            "dataset": str(evidence["dataset"]),
            "sourcePath": str(evidence["source_path"]),
            "fileName": str(evidence["original_file_name"]),
            "sourceKind": str(evidence["source_kind"]),
            "lifecycleStatus": str(evidence["lifecycle_status"]),
            "contentSha256": str(evidence["canonical_content_sha256"]),
            "sizeBytes": int(evidence["size_bytes"]),
            "mtimeNs": int(evidence["mtime_ns"]),
        },
        "revision": {
            "canonicalRevisionId": int(evidence["evidence_revision_id"]),
            "revisionUid": str(evidence["revision_uid"]),
            "sourceFingerprint": str(evidence["source_fingerprint"]),
            "fingerprintKind": str(evidence["fingerprint_kind"]),
            "captureContract": str(evidence["canonical_capture_contract"]),
            "captureStatus": str(evidence["canonical_capture_status"]),
            "isCurrent": bool(evidence["canonical_is_current"]),
            "extractorName": str(evidence["extractor_name"]),
            "extractorVersion": str(evidence["extractor_version"]),
            "capturedAt": str(evidence["captured_at"]),
            "captureV2": {
                "revisionId": int(capture["revision_id"]),
                "revisionUid": str(capture["revision_uid"]),
                "captureStatus": str(capture["capture_status"]),
                "isCurrent": bool(capture["is_current"]),
                "captureJsonSha256": str(capture["capture_json_sha256"]),
                "capturedAt": str(capture["captured_at"]),
                "staleAt": capture["stale_at"],
                "workbookStatus": str(capture["workbook_status"]),
                "isTrulyEmpty": bool(capture["is_truly_empty"]),
                "sheetCount": int(capture["sheet_count"]),
                "nonEmptySheetCount": int(capture["nonempty_sheet_count"]),
                "tabularSheetCount": int(capture["tabular_sheet_count"]),
            },
        },
        "evidence": {
            "kind": str(evidence["evidence_kind"]),
            "sheet": str(evidence["sheet_name"]),
            "range": exact_a1,
            "storedRange": str(evidence["range_address"]),
            "start": {
                "row": int(evidence["start_row"]),
                "column": int(evidence["start_col"]),
            },
            "end": {
                "row": int(evidence["end_row"]),
                "column": int(evidence["end_col"]),
            },
            "role": str(evidence["evidence_role"]),
            "sourceText": str(evidence["source_text"]),
            "note": str(evidence["note"]),
            "contentSha256": str(
                evidence["evidence_content_sha256"]
                or evidence["canonical_content_sha256"]
            ),
            "verificationStatus": verification_status,
            "createdAt": str(evidence["created_at"]),
        },
        "linkedEntities": _linked_entities(
            connection, int(evidence["evidence_id"])
        ),
        "preview": _preview(connection, evidence, sheet),
        "scope": {
            "tabularCaptureOnly": True,
            "imagesAnalyzed": False,
        },
    }


def build_evidence_detail_from_db(
    database_path: str | Path,
    public_evidence_id: str,
) -> dict[str, Any]:
    """Open ``database_path`` read-only and return one evidence detail."""

    with connect_evidence_readonly(database_path) as connection:
        return build_evidence_detail(connection, public_evidence_id)
