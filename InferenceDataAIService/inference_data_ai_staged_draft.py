"""Deterministic staged Study-draft planning and consolidation.

The v2 source-identity implementation is exported from this compatibility
module.  The older canonical-part helpers remain only so existing artifacts
and callers can be diagnosed; the workflow no longer uses them.
"""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
from typing import Any, Sequence

from inference_data_ai_schema import stable_uid
from inference_data_ai_staged_draft_v2 import (
    CONSOLIDATOR_CONTRACT_VERSION,
    FRAGMENT_CONTRACT_VERSION,
    FRAGMENT_PROMPT_VERSION,
    FRAGMENT_VALIDATOR_VERSION,
    StagedDraftV2Error,
    assess_one_call_budget,
    audit_no_candidate_source_inventory,
    audit_unselected_source_inventory,
    build_deterministic_acoustic_matrix_fragment_v2,
    build_deterministic_error_axis_tail_fragment_v2,
    build_deterministic_fo_fragment_v2,
    build_deterministic_function_fragment_v2,
    build_deterministic_function_grid_fragment_v2,
    build_deterministic_mask_fragment_v2,
    build_deterministic_nti_f0_fragment_v2,
    build_deterministic_nti_horizontal_matrix_fragment_v2,
    build_deterministic_result_table_fragment_v2,
    build_fragment_envelope,
    build_fragment_prompt,
    build_monolithic_request,
    build_study_registry_v2,
    chunks_for_part_v2,
    final_provenance_v2,
    final_provenance_v2_matches,
    finalize_fragment_envelope,
    fragment_artifact_paths,
    locators_for_part_v2,
    merge_fragment_records,
    normalize_fragment_evidence_dispositions,
    part_provenance_v2,
    part_provenance_v2_matches,
    plan_study_draft_v2,
    promote_required_source_locator_sections,
    project_canonical_manifest,
    registry_for_part,
    select_draft_universe,
    stable_record_id,
    validate_fragment_v2,
)


STAGED_DRAFT_PLAN_SCHEMA_VERSION = "study-draft-plan-v1"
STAGED_DRAFT_PART_PROVENANCE_SCHEMA_VERSION = (
    "study-draft-part-provenance-v1"
)


class StagedDraftError(RuntimeError):
    """Raised when staged planning or consolidation cannot remain exact."""


def _compact_json_bytes(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _sheet_identity(chunk: dict[str, Any]) -> tuple[int, str]:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return int(sheet.get("sheetIndex") or 0), str(
            sheet.get("title") or ""
        )
    return int(chunk.get("sheetIndex") or 0), str(sheet or "")


def _section_identity(
    chunk: dict[str, Any],
) -> tuple[int, str, object]:
    sheet_index, sheet_title = _sheet_identity(chunk)
    section_index = chunk.get("sectionIndex")
    if section_index is None:
        # Missing section metadata must never join unrelated chunks.
        section_index = f"chunk:{chunk.get('chunkId')}"
    return sheet_index, sheet_title, section_index


def _source_cell_key(
    chunk: dict[str, Any],
    cell: dict[str, Any],
) -> str:
    explicit = str(cell.get("sourceCellKey") or "").strip()
    if explicit:
        return explicit
    sheet_index, sheet_title = _sheet_identity(chunk)
    coordinate = str(cell.get("coordinate") or cell.get("c") or "").strip()
    if not coordinate:
        raise StagedDraftError(
            f"Chunk {chunk.get('chunkId')} has a primary cell without "
            "sourceCellKey or coordinate"
        )
    return f"{sheet_index}:{sheet_title}:{coordinate}"


def _is_candidate_result(result: dict[str, Any]) -> bool:
    return bool(
        str(result.get("status") or "").upper()
        in {"CANDIDATES", "NEEDS_REVIEW"}
        and result.get("candidates")
    )


def plan_study_draft_parts(
    *,
    packet_set: dict[str, Any],
    locator_results: Sequence[dict[str, Any]],
    prompt_version: str,
    max_chunks: int,
    max_cells: int,
    max_serialized_bytes: int,
) -> dict[str, Any]:
    """Create a deterministic, exact-ownership part plan.

    Every chunk in a candidate-bearing sheet section is selected, including
    continuation chunks whose own locator result has no candidate. Parts never
    cross section boundaries and each part is a contiguous section fragment.
    """

    if max_chunks < 1 or max_cells < 1 or max_serialized_bytes < 1:
        raise StagedDraftError("Study-draft part limits must be positive")
    chunks = list(packet_set.get("chunks", []))
    chunk_by_id: dict[str, dict[str, Any]] = {}
    for chunk in chunks:
        chunk_id = str(chunk.get("chunkId") or "")
        if not chunk_id or chunk_id in chunk_by_id:
            raise StagedDraftError(
                "Semantic packet chunks require unique nonempty chunkId"
            )
        chunk_by_id[chunk_id] = chunk
    locator_by_id: dict[str, dict[str, Any]] = {}
    for result in locator_results:
        chunk_id = str(result.get("chunkId") or "")
        if not chunk_id or chunk_id in locator_by_id:
            raise StagedDraftError(
                "Locator results require unique nonempty chunkId"
            )
        locator_by_id[chunk_id] = result
    if set(locator_by_id) != set(chunk_by_id):
        missing = sorted(set(chunk_by_id) - set(locator_by_id))
        unexpected = sorted(set(locator_by_id) - set(chunk_by_id))
        raise StagedDraftError(
            "Locator/chunk coverage mismatch: "
            f"missing={missing}, unexpected={unexpected}"
        )

    candidate_chunk_ids = [
        chunk_id
        for chunk_id in chunk_by_id
        if _is_candidate_result(locator_by_id[chunk_id])
    ]
    candidate_sections = {
        _section_identity(chunk_by_id[chunk_id])
        for chunk_id in candidate_chunk_ids
    }
    selected_chunks = [
        chunk
        for chunk in chunks
        if _section_identity(chunk) in candidate_sections
    ]
    selected_chunk_ids = [
        str(chunk["chunkId"]) for chunk in selected_chunks
    ]

    all_owned_cells: list[str] = []
    seen_owned_cells: set[str] = set()
    for chunk in selected_chunks:
        for cell in chunk.get("cells", []):
            key = _source_cell_key(chunk, cell)
            if key in seen_owned_cells:
                raise StagedDraftError(
                    f"Selected primary source cell has duplicate ownership: {key}"
                )
            seen_owned_cells.add(key)
            all_owned_cells.append(key)

    source_revision = packet_set.get("inventory", {}).get(
        "sourceRevision",
        {},
    )
    revision_uid = str(source_revision.get("revisionUid") or "")
    content_sha256 = str(source_revision.get("contentSha256") or "").lower()
    if not revision_uid or not content_sha256:
        raise StagedDraftError(
            "Semantic packet inventory lacks source revision identity"
        )

    raw_parts: list[dict[str, Any]] = []
    section_chunks: list[dict[str, Any]] = []
    current_section: tuple[int, str, object] | None = None

    def fragment_serialized_bytes(
        fragment_chunks: Sequence[dict[str, Any]],
    ) -> int:
        return len(
            _compact_json_bytes(
                {
                    "focusedChunks": list(fragment_chunks),
                    "locatorResults": [
                        locator_by_id[str(chunk["chunkId"])]
                        for chunk in fragment_chunks
                    ],
                }
            )
        )

    def flush_section() -> None:
        nonlocal section_chunks
        if not section_chunks:
            return
        current: list[dict[str, Any]] = []
        current_cells = 0
        for chunk in section_chunks:
            chunk_cells = len(chunk.get("cells", []))
            chunk_bytes = fragment_serialized_bytes([chunk])
            if (
                chunk_cells > max_cells
                or chunk_bytes > max_serialized_bytes
            ):
                raise StagedDraftError(
                    f"Chunk {chunk['chunkId']} exceeds a Study-draft part "
                    "cell or serialized-byte limit and cannot be split safely"
                )
            proposed = [*current, chunk]
            proposed_bytes = fragment_serialized_bytes(proposed)
            if current and (
                len(current) >= max_chunks
                or current_cells + chunk_cells > max_cells
                or proposed_bytes > max_serialized_bytes
            ):
                raw_parts.append({"chunks": current})
                current = []
                current_cells = 0
                proposed = [chunk]
                proposed_bytes = fragment_serialized_bytes(proposed)
            if proposed_bytes > max_serialized_bytes:
                raise StagedDraftError(
                    f"Chunk {chunk['chunkId']} exceeds the exact serialized "
                    "Study-draft part limit"
                )
            current.append(chunk)
            current_cells += chunk_cells
        if current:
            raw_parts.append({"chunks": current})
        section_chunks = []

    for chunk in selected_chunks:
        section = _section_identity(chunk)
        if current_section is not None and section != current_section:
            flush_section()
        current_section = section
        section_chunks.append(chunk)
    flush_section()

    parts: list[dict[str, Any]] = []
    owned_chunks: list[str] = []
    owned_cells: list[str] = []
    for part_index, raw_part in enumerate(raw_parts, start=1):
        part_chunks = raw_part["chunks"]
        chunk_ids = [str(chunk["chunkId"]) for chunk in part_chunks]
        part_cells = [
            _source_cell_key(chunk, cell)
            for chunk in part_chunks
            for cell in chunk.get("cells", [])
        ]
        sheet_index, sheet_title, section_index = _section_identity(
            part_chunks[0]
        )
        part_id = stable_uid(
            "study-draft-part",
            revision_uid,
            prompt_version,
            sheet_index,
            sheet_title,
            section_index,
            *chunk_ids,
        )
        part = {
            "partIndex": part_index,
            "partId": part_id,
            "sheetIndex": sheet_index,
            "sheetTitle": sheet_title,
            "sectionIndex": section_index,
            "chunkIds": chunk_ids,
            "ownedSourceCellKeys": part_cells,
            "chunkCount": len(part_chunks),
            "cellCount": len(part_cells),
            "serializedBytes": fragment_serialized_bytes(part_chunks),
        }
        parts.append(part)
        owned_chunks.extend(chunk_ids)
        owned_cells.extend(part_cells)

    if owned_chunks != selected_chunk_ids:
        raise StagedDraftError(
            "Study-draft part chunk ownership is not exact or source ordered"
        )
    if owned_cells != all_owned_cells:
        raise StagedDraftError(
            "Study-draft part source-cell ownership is not exact"
        )
    if len({part["partId"] for part in parts}) != len(parts):
        raise StagedDraftError("Study-draft part IDs collided")

    plan_id = stable_uid(
        "study-draft-plan",
        revision_uid,
        prompt_version,
        max_chunks,
        max_cells,
        max_serialized_bytes,
        *(part["partId"] for part in parts),
    )
    return {
        "schemaVersion": STAGED_DRAFT_PLAN_SCHEMA_VERSION,
        "planId": plan_id,
        "promptVersion": prompt_version,
        "source": {
            "revisionUid": revision_uid,
            "contentSha256": content_sha256,
        },
        "limits": {
            "maxChunks": max_chunks,
            "maxCells": max_cells,
            "maxSerializedBytes": max_serialized_bytes,
        },
        "locatorAssessedChunkIds": list(chunk_by_id),
        "candidateChunkIds": candidate_chunk_ids,
        "selectedChunkIds": selected_chunk_ids,
        "continuationChunkIds": [
            chunk_id
            for chunk_id in selected_chunk_ids
            if chunk_id not in set(candidate_chunk_ids)
        ],
        "ownedSourceCellCount": len(all_owned_cells),
        "ownedSourceCellSha256": hashlib.sha256(
            "\n".join(all_owned_cells).encode("utf-8")
        ).hexdigest(),
        "parts": parts,
    }


def chunks_for_part(
    packet_set: dict[str, Any],
    part: dict[str, Any],
) -> list[dict[str, Any]]:
    chunk_by_id = {
        str(chunk["chunkId"]): chunk
        for chunk in packet_set.get("chunks", [])
    }
    try:
        return [chunk_by_id[str(value)] for value in part["chunkIds"]]
    except KeyError as exc:
        raise StagedDraftError(
            f"Study-draft part references missing chunk {exc.args[0]}"
        ) from exc


def locator_results_for_part(
    locator_results: Sequence[dict[str, Any]],
    part: dict[str, Any],
) -> list[dict[str, Any]]:
    result_by_id = {
        str(result["chunkId"]): result for result in locator_results
    }
    try:
        return [result_by_id[str(value)] for value in part["chunkIds"]]
    except KeyError as exc:
        raise StagedDraftError(
            f"Study-draft part references missing locator {exc.args[0]}"
        ) from exc


def part_artifact_paths(
    artifact_directory: Path,
    part: dict[str, Any],
) -> tuple[Path, Path]:
    part_directory = artifact_directory / "draft-parts"
    safe_part = "".join(
        char if char.isalnum() or char in "._-" else "_"
        for char in str(part["partId"])
    )
    return (
        part_directory / f"{safe_part}.manifest.json",
        part_directory / f"{safe_part}.provenance.json",
    )


def part_provenance_value(
    *,
    plan: dict[str, Any],
    part: dict[str, Any],
    manifest_path: Path,
    manifest_sha256: str,
    generated_at: str,
) -> dict[str, Any]:
    return {
        "schemaVersion": STAGED_DRAFT_PART_PROVENANCE_SCHEMA_VERSION,
        "promptVersion": plan["promptVersion"],
        "planId": plan["planId"],
        "partId": part["partId"],
        "source": copy.deepcopy(plan["source"]),
        "chunkIds": list(part["chunkIds"]),
        "ownedSourceCellKeys": list(part["ownedSourceCellKeys"]),
        "manifestPath": str(manifest_path),
        "manifestSha256": manifest_sha256,
        "generatedAt": generated_at,
        "imagesAnalyzed": False,
    }


def part_provenance_matches(
    *,
    provenance: dict[str, Any],
    plan: dict[str, Any],
    part: dict[str, Any],
    manifest_sha256: str,
) -> bool:
    return bool(
        provenance.get("schemaVersion")
        == STAGED_DRAFT_PART_PROVENANCE_SCHEMA_VERSION
        and provenance.get("promptVersion") == plan.get("promptVersion")
        and provenance.get("planId") == plan.get("planId")
        and provenance.get("partId") == part.get("partId")
        and provenance.get("source") == plan.get("source")
        and provenance.get("chunkIds") == part.get("chunkIds")
        and provenance.get("ownedSourceCellKeys")
        == part.get("ownedSourceCellKeys")
        and provenance.get("manifestSha256") == manifest_sha256
    )


def _ordered_union(values: Sequence[object]) -> list[Any]:
    result: list[Any] = []
    seen: set[bytes] = set()
    for value in values:
        encoded = _compact_json_bytes(value)
        if encoded in seen:
            continue
        seen.add(encoded)
        result.append(copy.deepcopy(value))
    return result


def consolidate_study_draft_parts(
    *,
    plan: dict[str, Any],
    part_manifests: Sequence[tuple[dict[str, Any], dict[str, Any]]],
    source: dict[str, Any],
    content_complete: bool,
) -> dict[str, Any]:
    """Consolidate validated parts without semantic cross-part inference."""

    expected_parts = list(plan.get("parts", []))
    if len(part_manifests) != len(expected_parts):
        raise StagedDraftError(
            "Cannot consolidate until every Study-draft part is present"
        )
    analyses: list[dict[str, Any]] = []
    studies: list[dict[str, Any]] = []
    rewritten_keys: set[str] = set()
    for expected, (part, manifest) in zip(
        expected_parts,
        part_manifests,
        strict=True,
    ):
        if part.get("partId") != expected.get("partId"):
            raise StagedDraftError(
                "Study-draft parts are missing or out of source order"
            )
        manifest_source = manifest.get("source", {})
        for field in (
            "dataset",
            "sourcePath",
            "revisionUid",
            "contentSha256",
        ):
            expected_value = str(source.get(field) or "")
            actual_value = str(manifest_source.get(field) or "")
            if field == "contentSha256":
                expected_value = expected_value.lower()
                actual_value = actual_value.lower()
            if actual_value != expected_value:
                raise StagedDraftError(
                    f"Study-draft part source mismatch: {field}"
                )
        if bool(manifest_source.get("contentComplete")):
            raise StagedDraftError(
                "Study-draft fragment must remain contentComplete=false"
            )
        analysis = manifest.get("workbookAnalysis")
        if not isinstance(analysis, dict):
            raise StagedDraftError(
                "Study-draft fragment lacks workbookAnalysis"
            )
        analyses.append(analysis)
        for study_index, original_study in enumerate(
            manifest.get("studies", []),
            start=1,
        ):
            if not isinstance(original_study, dict):
                raise StagedDraftError(
                    "Study-draft fragment contains a non-object Study"
                )
            study = copy.deepcopy(original_study)
            original_key = str(study.get("key") or "").strip()
            rewritten_key = (
                f"{original_key}--"
                f"{str(part['partId']).rsplit('_', 1)[-1][:12]}"
                f"-{study_index}"
            )
            if rewritten_key in rewritten_keys:
                raise StagedDraftError(
                    f"Consolidated Study key collision: {rewritten_key}"
                )
            rewritten_keys.add(rewritten_key)
            study["key"] = rewritten_key
            studies.append(study)

    if not analyses:
        raise StagedDraftError(
            "Cannot consolidate an empty Study-draft part plan"
        )
    limitations = _ordered_union(
        [
            limitation
            for analysis in analyses
            for limitation in analysis.get("limitations", [])
        ]
    )
    evidence = _ordered_union(
        [
            item
            for analysis in analyses
            for item in analysis.get("evidence", [])
        ]
    )
    summaries = _ordered_union(
        [
            str(analysis.get("summary") or "").strip()
            for analysis in analyses
            if str(analysis.get("summary") or "").strip()
        ]
    )
    first_analysis = analyses[0]
    return {
        "schemaVersion": "canonical-study-manifest-v1",
        "source": {
            **copy.deepcopy(source),
            "contentComplete": bool(content_complete),
        },
        "workbookAnalysis": {
            "key": stable_uid(
                "staged-workbook-analysis",
                source["revisionUid"],
                plan["planId"],
            ),
            "title": str(first_analysis.get("title") or ""),
            "summary": " | ".join(summaries),
            "status": "NEEDS_REVIEW",
            "verificationStatus": "NEEDS_REVIEW",
            "limitations": limitations,
            "evidence": evidence,
        },
        "studies": studies,
    }


__all__ = [
    "CONSOLIDATOR_CONTRACT_VERSION",
    "FRAGMENT_CONTRACT_VERSION",
    "FRAGMENT_PROMPT_VERSION",
    "FRAGMENT_VALIDATOR_VERSION",
    "STAGED_DRAFT_PART_PROVENANCE_SCHEMA_VERSION",
    "STAGED_DRAFT_PLAN_SCHEMA_VERSION",
    "StagedDraftError",
    "StagedDraftV2Error",
    "assess_one_call_budget",
    "audit_no_candidate_source_inventory",
    "audit_unselected_source_inventory",
    "build_deterministic_acoustic_matrix_fragment_v2",
    "build_deterministic_fo_fragment_v2",
    "build_deterministic_mask_fragment_v2",
    "build_fragment_envelope",
    "build_fragment_prompt",
    "build_monolithic_request",
    "build_study_registry_v2",
    "chunks_for_part_v2",
    "chunks_for_part",
    "consolidate_study_draft_parts",
    "final_provenance_v2",
    "final_provenance_v2_matches",
    "finalize_fragment_envelope",
    "fragment_artifact_paths",
    "locator_results_for_part",
    "locators_for_part_v2",
    "merge_fragment_records",
    "normalize_fragment_evidence_dispositions",
    "part_artifact_paths",
    "part_provenance_matches",
    "part_provenance_value",
    "part_provenance_v2",
    "part_provenance_v2_matches",
    "plan_study_draft_v2",
    "plan_study_draft_parts",
    "promote_required_source_locator_sections",
    "project_canonical_manifest",
    "registry_for_part",
    "select_draft_universe",
    "stable_record_id",
    "validate_fragment_v2",
]
