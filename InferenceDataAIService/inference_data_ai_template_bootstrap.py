"""Build a read-only structure catalog from completed table-first assets.

This command never calls AI and never mutates the Capture v2 database.  It
maps completed table-first requests back to their exact capture revisions,
creates v2 fingerprints, groups exact structures, and proposes similarity
families for later recipe replay.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sqlite3
from collections import defaultdict
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_recipe_matcher import fingerprint_similarity
from inference_data_ai_structure_fingerprint import (
    build_structure_fingerprint_from_database,
)


BOOTSTRAP_CATALOG_SCHEMA_VERSION = "excel-template-bootstrap-catalog-v1"
BOOTSTRAP_ENGINE_VERSION = "table-first-structure-bootstrap-v1.0"
DEFAULT_FAMILY_SIMILARITY_THRESHOLD = 0.90


class TemplateBootstrapError(RuntimeError):
    """Raised when completed requests cannot be tied to Capture v2 safely."""


def _canonical_json_bytes(value: Any) -> bytes:
    return (
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n"
    ).encode("utf-8")


def _stable_id(prefix: str, value: Any) -> str:
    digest = hashlib.sha256(
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()
    return f"{prefix}-{digest[:20]}"


def _atomic_write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_bytes(_canonical_json_bytes(value))
    os.replace(temporary, path)


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise TemplateBootstrapError(f"JSON root must be an object: {path}")
    return value


def _request_records(
    batch_root: Path,
    *,
    limit: int | None = None,
) -> list[dict[str, Any]]:
    request_root = batch_root / "requests"
    analysis_root = batch_root / "analyses"
    projection_root = batch_root / "projections"
    if not request_root.is_dir():
        raise TemplateBootstrapError(f"Request directory is missing: {request_root}")
    records: list[dict[str, Any]] = []
    request_paths = sorted(request_root.glob("*.json"))
    if limit is not None:
        request_paths = request_paths[: max(int(limit), 0)]
    for path in request_paths:
        request = _read_json(path)
        source = request.get("source")
        if not isinstance(source, dict):
            raise TemplateBootstrapError(f"Request source is missing: {path}")
        content_sha = str(source.get("contentSha256") or "")
        revision_uid = str(source.get("revisionUid") or "")
        if len(content_sha) != 64:
            raise TemplateBootstrapError(f"Request SHA-256 is invalid: {path}")
        analysis_path = analysis_root / path.name
        projection_path = projection_root / path.name
        records.append(
            {
                "requestPath": str(path.resolve()),
                "requestFile": path.name,
                "fileName": str(source.get("fileName") or ""),
                "sourcePath": str(source.get("sourcePath") or ""),
                "contentSha256": content_sha,
                "revisionUid": revision_uid,
                "analysisExists": analysis_path.is_file(),
                "projectionExists": projection_path.is_file(),
            }
        )
    return records


def _open_read_only_database(path: Path) -> sqlite3.Connection:
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise FileNotFoundError(resolved)
    uri = resolved.as_uri() + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True, timeout=30)
    connection.execute("PRAGMA query_only=ON")
    return connection


def _capture_revision(
    connection: sqlite3.Connection,
    record: dict[str, Any],
) -> dict[str, Any] | None:
    revision_uid = str(record.get("revisionUid") or "")
    if revision_uid:
        row = connection.execute(
            """
            SELECT revision_id, revision_uid, content_sha256, is_current
            FROM capture_v2_revisions
            WHERE revision_uid=?
            """,
            (revision_uid,),
        ).fetchone()
        if row is not None:
            if str(row[2]) != record["contentSha256"]:
                raise TemplateBootstrapError(
                    f"Revision UID/SHA mismatch: {revision_uid}"
                )
            return {
                "revisionId": int(row[0]),
                "revisionUid": str(row[1]),
                "contentSha256": str(row[2]),
                "isCurrent": bool(row[3]),
                "matchKind": "REVISION_UID",
            }
    rows = connection.execute(
        """
        SELECT revision_id, revision_uid, content_sha256, is_current
        FROM capture_v2_revisions
        WHERE content_sha256=?
        ORDER BY is_current DESC, revision_id
        """,
        (record["contentSha256"],),
    ).fetchall()
    if not rows:
        return None
    if len(rows) > 1 and not revision_uid:
        raise TemplateBootstrapError(
            "SHA matches multiple revisions and request has no revision UID: "
            + record["contentSha256"]
        )
    row = rows[0]
    return {
        "revisionId": int(row[0]),
        "revisionUid": str(row[1]),
        "contentSha256": str(row[2]),
        "isCurrent": bool(row[3]),
        "matchKind": "CONTENT_SHA256",
    }


def _coarse_partition(fingerprint: dict[str, Any]) -> str:
    value = {
        "sheetCount": fingerprint["workbook"]["sheetCount"],
        "tabularSheetCount": fingerprint["workbook"]["tabularSheetCount"],
        "sheets": [
            {
                "state": sheet["sheetState"],
                "tabular": sheet["tabular"],
                "columnBucket": sheet["usedRangeBucket"]["columns"],
                "mergeCount": len(sheet["mergedGeometry"]),
                "formulaPatternCount": len(sheet["formulaPatternHashes"]),
            }
            for sheet in fingerprint["sheets"]
        ],
    }
    return _stable_id("partition", value)


def _exact_structures(
    fingerprinted: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    fingerprints: dict[str, dict[str, Any]] = {}
    for item in fingerprinted:
        digest = item["fingerprint"]["fingerprintSha256"]
        groups[digest].append(item)
        fingerprints.setdefault(digest, item["fingerprint"])
    result: list[dict[str, Any]] = []
    for digest, members in groups.items():
        ordered_members = sorted(
            members,
            key=lambda item: (
                item["record"]["fileName"],
                item["record"]["contentSha256"],
            ),
        )
        result.append(
            {
                "structureId": f"structure-{digest[:20]}",
                "fingerprintSha256": digest,
                "fileCount": len(ordered_members),
                "fingerprint": fingerprints[digest],
                "members": [
                    {
                        "fileName": item["record"]["fileName"],
                        "sourcePath": item["record"]["sourcePath"],
                        "contentSha256": item["record"]["contentSha256"],
                        "revisionId": item["capture"]["revisionId"],
                        "revisionUid": item["capture"]["revisionUid"],
                    }
                    for item in ordered_members
                ],
            }
        )
    result.sort(
        key=lambda item: (
            -int(item["fileCount"]),
            str(item["fingerprintSha256"]),
        )
    )
    return result


def _candidate_families(
    structures: list[dict[str, Any]],
    *,
    similarity_threshold: float,
) -> list[dict[str, Any]]:
    partitions: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for structure in structures:
        partitions[_coarse_partition(structure["fingerprint"])].append(structure)

    families: list[dict[str, Any]] = []
    for partition_id in sorted(partitions):
        candidates = sorted(
            partitions[partition_id],
            key=lambda item: (
                -int(item["fileCount"]),
                str(item["fingerprintSha256"]),
            ),
        )
        clusters: list[dict[str, Any]] = []
        for structure in candidates:
            ranked: list[tuple[float, int, dict[str, Any]]] = []
            for index, cluster in enumerate(clusters):
                similarity = fingerprint_similarity(
                    structure["fingerprint"],
                    cluster["representative"]["fingerprint"],
                )
                ranked.append((float(similarity["score"]), index, similarity))
            if ranked:
                best_score, best_index, details = max(
                    ranked,
                    key=lambda item: (item[0], -item[1]),
                )
            else:
                best_score, best_index, details = 0.0, -1, {}
            if best_index >= 0 and best_score >= similarity_threshold:
                clusters[best_index]["members"].append(
                    {
                        "structure": structure,
                        "scoreToRepresentative": round(best_score, 6),
                        "components": details["components"],
                    }
                )
            else:
                clusters.append(
                    {
                        "representative": structure,
                        "members": [
                            {
                                "structure": structure,
                                "scoreToRepresentative": 1.0,
                                "components": {},
                            }
                        ],
                    }
                )
        for cluster in clusters:
            member_ids = sorted(
                item["structure"]["structureId"] for item in cluster["members"]
            )
            file_count = sum(
                int(item["structure"]["fileCount"])
                for item in cluster["members"]
            )
            families.append(
                {
                    "familyId": _stable_id("family-candidate", member_ids),
                    "status": "CANDIDATE",
                    "partitionId": partition_id,
                    "representativeStructureId": cluster["representative"][
                        "structureId"
                    ],
                    "fileCount": file_count,
                    "exactStructureCount": len(cluster["members"]),
                    "minimumScoreToRepresentative": min(
                        float(item["scoreToRepresentative"])
                        for item in cluster["members"]
                    ),
                    "members": [
                        {
                            "structureId": item["structure"]["structureId"],
                            "fileCount": item["structure"]["fileCount"],
                            "scoreToRepresentative": item[
                                "scoreToRepresentative"
                            ],
                            "components": item["components"],
                        }
                        for item in sorted(
                            cluster["members"],
                            key=lambda item: (
                                -int(item["structure"]["fileCount"]),
                                str(item["structure"]["structureId"]),
                            ),
                        )
                    ],
                }
            )
    families.sort(
        key=lambda item: (
            -int(item["fileCount"]),
            -int(item["exactStructureCount"]),
            str(item["familyId"]),
        )
    )
    return families


def build_template_bootstrap_catalog(
    *,
    database_path: str | Path,
    table_first_batch_root: str | Path,
    limit: int | None = None,
    family_similarity_threshold: float = DEFAULT_FAMILY_SIMILARITY_THRESHOLD,
    generated_at: str | None = None,
) -> dict[str, Any]:
    """Build a catalog in memory using read-only Capture v2 queries."""

    database = Path(database_path).expanduser().resolve()
    batch_root = Path(table_first_batch_root).expanduser().resolve()
    records = _request_records(batch_root, limit=limit)
    fingerprinted: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    incomplete_assets: list[dict[str, Any]] = []
    with closing(_open_read_only_database(database)) as connection:
        for record in records:
            if not record["analysisExists"] or not record["projectionExists"]:
                incomplete_assets.append(record)
                continue
            capture = _capture_revision(connection, record)
            if capture is None:
                missing.append(record)
                continue
            fingerprinted.append(
                {
                    "record": record,
                    "capture": capture,
                    "fingerprint": build_structure_fingerprint_from_database(
                        connection,
                        capture["revisionId"],
                    ),
                }
            )
    structures = _exact_structures(fingerprinted)
    families = _candidate_families(
        structures,
        similarity_threshold=float(family_similarity_threshold),
    )
    exact_reusable_files = sum(
        int(item["fileCount"])
        for item in structures
        if int(item["fileCount"]) > 1
    )
    family_reusable_files = sum(
        int(item["fileCount"])
        for item in families
        if int(item["fileCount"]) > 1
    )
    return {
        "schemaVersion": BOOTSTRAP_CATALOG_SCHEMA_VERSION,
        "engineVersion": BOOTSTRAP_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            "databasePath": str(database),
            "tableFirstBatchRoot": str(batch_root),
            "requestCount": len(records),
            "limit": limit,
            "familySimilarityThreshold": float(family_similarity_threshold),
            "databaseMode": "READ_ONLY",
            "aiCalls": 0,
        },
        "summary": {
            "requestCount": len(records),
            "fingerprintedFileCount": len(fingerprinted),
            "missingCaptureCount": len(missing),
            "incompleteAssetCount": len(incomplete_assets),
            "exactStructureCount": len(structures),
            "exactReusableStructureCount": sum(
                int(item["fileCount"]) > 1 for item in structures
            ),
            "exactReusableFileCount": exact_reusable_files,
            "candidateFamilyCount": len(families),
            "candidateReusableFamilyCount": sum(
                int(item["fileCount"]) > 1 for item in families
            ),
            "candidateReusableFileCount": family_reusable_files,
            "largestExactStructureFileCount": max(
                (int(item["fileCount"]) for item in structures),
                default=0,
            ),
            "largestCandidateFamilyFileCount": max(
                (int(item["fileCount"]) for item in families),
                default=0,
            ),
        },
        "missingCaptures": missing,
        "incompleteAssets": incomplete_assets,
        "families": families,
        "structures": structures,
    }


def write_template_bootstrap_catalog(
    catalog: dict[str, Any],
    output_path: str | Path,
) -> Path:
    target = Path(output_path).expanduser().resolve()
    _atomic_write_json(target, catalog)
    return target


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Build a read-only, AI-free structure catalog from completed "
            "table-first requests and Capture v2 SQLite."
        )
    )
    parser.add_argument("--db", required=True, help="Capture v2 SQLite database")
    parser.add_argument(
        "--table-first-batch-root",
        required=True,
        help="Batch directory containing requests/analyses/projections",
    )
    parser.add_argument("--out", required=True, help="Output catalog JSON")
    parser.add_argument("--limit", type=int)
    parser.add_argument(
        "--family-similarity-threshold",
        type=float,
        default=DEFAULT_FAMILY_SIMILARITY_THRESHOLD,
    )
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    catalog = build_template_bootstrap_catalog(
        database_path=arguments.db,
        table_first_batch_root=arguments.table_first_batch_root,
        limit=arguments.limit,
        family_similarity_threshold=arguments.family_similarity_threshold,
    )
    output = write_template_bootstrap_catalog(catalog, arguments.out)
    print(
        json.dumps(
            {
                "output": str(output),
                "summary": catalog["summary"],
                "aiCalls": 0,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = [
    "BOOTSTRAP_CATALOG_SCHEMA_VERSION",
    "BOOTSTRAP_ENGINE_VERSION",
    "DEFAULT_FAMILY_SIMILARITY_THRESHOLD",
    "TemplateBootstrapError",
    "build_template_bootstrap_catalog",
    "write_template_bootstrap_catalog",
]
