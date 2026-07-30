"""Source-identity staged Study drafting for oversized workbooks.

Version 2 never merges partial canonical manifests.  It plans exact source
ownership, validates append-only fragment records against a stable Study
registry and cell allowlist, then deterministically projects one canonical
manifest after every fragment is complete.
"""

from __future__ import annotations

from bisect import bisect_left, bisect_right
import copy
import hashlib
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable, Sequence

from inference_data_ai_content_coverage import (
    build_content_coverage_inventory,
)
from inference_data_ai_schema import stable_uid


STAGED_DRAFT_PLAN_V2_SCHEMA_VERSION = "study-draft-plan-v2"
STUDY_REGISTRY_V2_SCHEMA_VERSION = "study-registry-v2"
STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION = "study-draft-fragment-v2"
STAGED_PART_PROVENANCE_V2_SCHEMA_VERSION = (
    "study-draft-part-provenance-v2"
)
STAGED_FINAL_PROVENANCE_V2_SCHEMA_VERSION = (
    "study-draft-final-provenance-v2"
)
FRAGMENT_CONTRACT_VERSION = "study-draft-fragment-contract-v7"
FRAGMENT_VALIDATOR_VERSION = "study-draft-fragment-validator-v11"
CONSOLIDATOR_CONTRACT_VERSION = "study-draft-consolidator-v5"
FRAGMENT_PROMPT_VERSION = "study-draft-fragment-prompt-v6"
SOURCE_CHUNK_SEGMENT_SCHEMA_VERSION = "source-chunk-segment-v1"

RECORD_TYPES = {
    "STUDY_PATCH",
    "ENTITY_DECLARATION",
    "OBSERVATION_APPEND",
    "SERIES_SEGMENT_APPEND",
    "COMPARISON_LINK_INTENT",
    "CONCLUSION_APPEND",
    "LIMITATION_APPEND",
}
ENTITY_TYPES = {"CONTEXT", "FACTOR", "ARM", "OUTCOME"}
DISPOSITIONS = {
    "RECORD_EVIDENCE",
    "CONTEXT_ONLY",
    "NO_SEMANTIC_RECORD",
}
_A1 = re.compile(
    r"\$?([A-Za-z]{1,4})\$?([1-9]\d*)"
    r"(?::\$?([A-Za-z]{1,4})\$?([1-9]\d*))?"
)
_SEQUENCE_LABEL = re.compile(
    r"(?:^|[\s_/#-])(?:no\.?|number|index|seq(?:uence)?|순번|번호|회차)"
    r"(?:$|[\s_/#-])",
    re.IGNORECASE,
)


class StagedDraftV2Error(RuntimeError):
    """Raised when staged drafting cannot preserve source identity."""


def compact_json_bytes(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def json_sha256(value: object) -> str:
    return hashlib.sha256(compact_json_bytes(value)).hexdigest()


def bytes_sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _sheet(chunk: dict[str, Any]) -> tuple[int, str]:
    value = chunk.get("sheet")
    if isinstance(value, dict):
        return int(value.get("sheetIndex") or 0), str(
            value.get("title") or ""
        )
    return int(chunk.get("sheetIndex") or 0), str(value or "")


def _section(chunk: dict[str, Any]) -> tuple[int, str, object]:
    sheet_index, title = _sheet(chunk)
    section_index = chunk.get("sectionIndex")
    if section_index is None:
        section_index = f"chunk:{chunk.get('chunkId')}"
    return sheet_index, title, section_index


def _coordinate(cell: dict[str, Any]) -> str:
    return str(
        cell.get("coordinate") or cell.get("c") or ""
    ).strip().upper()


def _cell_key(chunk: dict[str, Any], cell: dict[str, Any]) -> str:
    explicit = str(cell.get("sourceCellKey") or "").strip()
    if explicit:
        return explicit
    coordinate = _coordinate(cell)
    if not coordinate:
        raise StagedDraftV2Error(
            f"Chunk {chunk.get('chunkId')} contains an unaddressed cell"
        )
    sheet_index, title = _sheet(chunk)
    return f"{sheet_index}:{title}:{coordinate}"


def _column(label: str) -> int:
    result = 0
    for char in label.upper():
        result = result * 26 + ord(char) - ord("A") + 1
    return result


def range_bounds(address: object) -> tuple[int, int, int, int]:
    match = _A1.fullmatch(str(address or "").strip())
    if not match:
        raise StagedDraftV2Error(f"Invalid A1 range: {address}")
    start_column = _column(match.group(1))
    start_row = int(match.group(2))
    end_column = _column(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_row < start_row or end_column < start_column:
        raise StagedDraftV2Error(f"Reversed A1 range: {address}")
    return start_row, start_column, end_row, end_column


def _position(cell: dict[str, Any]) -> tuple[int, int]:
    if cell.get("row") not in (None, "") and cell.get("column") not in (
        None,
        "",
    ):
        return int(cell["row"]), int(cell["column"])
    start_row, start_column, _end_row, _end_column = range_bounds(
        _coordinate(cell)
    )
    return start_row, start_column


def _is_candidate(result: dict[str, Any]) -> bool:
    return bool(
        str(result.get("status") or "").upper()
        in {"CANDIDATES", "NEEDS_REVIEW"}
        and result.get("candidates")
    )


def select_draft_universe(
    *,
    packet_set: dict[str, Any],
    locator_results: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Select every chunk in every candidate-bearing source section."""

    chunks = list(packet_set.get("chunks", []))
    chunk_by_id: dict[str, dict[str, Any]] = {}
    for chunk in chunks:
        chunk_id = str(chunk.get("chunkId") or "")
        if not chunk_id or chunk_id in chunk_by_id:
            raise StagedDraftV2Error(
                "Source chunks require unique nonempty chunkId"
            )
        chunk_by_id[chunk_id] = chunk
    locator_by_id: dict[str, dict[str, Any]] = {}
    for result in locator_results:
        chunk_id = str(result.get("chunkId") or "")
        if not chunk_id or chunk_id in locator_by_id:
            raise StagedDraftV2Error(
                "Locator results require unique nonempty chunkId"
            )
        locator_by_id[chunk_id] = result
    if set(chunk_by_id) != set(locator_by_id):
        raise StagedDraftV2Error(
            "Locator results must exactly cover source chunks"
        )
    candidate_chunk_ids = [
        chunk_id
        for chunk_id, chunk in chunk_by_id.items()
        if _is_candidate(locator_by_id[chunk_id])
    ]
    candidate_sections = {
        _section(chunk_by_id[chunk_id])
        for chunk_id in candidate_chunk_ids
    }
    selected_chunks = [
        chunk for chunk in chunks if _section(chunk) in candidate_sections
    ]
    selected_ids = [str(chunk["chunkId"]) for chunk in selected_chunks]
    selected_locators = [
        locator_by_id[chunk_id] for chunk_id in selected_ids
    ]
    owned_keys: list[str] = []
    seen: set[str] = set()
    for chunk in selected_chunks:
        for cell in chunk.get("cells", []):
            key = _cell_key(chunk, cell)
            if key in seen:
                raise StagedDraftV2Error(
                    f"Duplicate selected source-cell ownership: {key}"
                )
            seen.add(key)
            owned_keys.append(key)
    return {
        "candidateChunkIds": candidate_chunk_ids,
        "candidateSections": [
            list(value)
            for value in sorted(
                candidate_sections,
                key=lambda item: (
                    int(item[0]),
                    str(item[1]),
                    str(item[2]),
                ),
            )
        ],
        "selectedChunks": selected_chunks,
        "selectedLocatorResults": selected_locators,
        "selectedChunkIds": selected_ids,
        "continuationChunkIds": [
            value
            for value in selected_ids
            if value not in set(candidate_chunk_ids)
        ],
        "ownedSourceCellKeys": owned_keys,
    }


def audit_no_candidate_source_inventory(
    *,
    packet_set: dict[str, Any],
    locator_results: Sequence[dict[str, Any]] = (),
) -> dict[str, Any]:
    """Reuse the general source-content classifier for an all-NO result.

    Formula cells without a usable cached numeric value are added explicitly:
    they are still source results and cannot be silently terminalized merely
    because the numeric classifier cannot inspect their value.
    """

    chunks = list(packet_set.get("chunks", []))
    owned_keys = [
        _cell_key(chunk, cell)
        for chunk in chunks
        for cell in chunk.get("cells", [])
    ]
    inventory = build_content_coverage_inventory(
        chunks=chunks,
        locator_results=list(locator_results),
        expected_source_cell_keys=owned_keys,
    )
    required_by_key: dict[str, dict[str, Any]] = {}
    for collection, content_class in (
        ("requiredCells", "QUANTITATIVE_RESULT"),
        ("categoricalStatusCells", "CATEGORICAL_RESULT"),
        ("narrativeConclusionCells", "SOURCE_CONCLUSION"),
        ("semanticLabelCells", "SEMANTIC_LABEL"),
    ):
        for raw_item in inventory.get(collection, []):
            item = copy.deepcopy(raw_item)
            key = str(item.get("sourceCellKey") or "")
            if not key:
                continue
            item.setdefault("contentClass", content_class)
            required_by_key.setdefault(key, item)
    required = list(required_by_key.values())
    excluded = [
        copy.deepcopy(item)
        for item in inventory.get("excludedCells", [])
    ]
    classified = {
        str(item.get("sourceCellKey") or "")
        for item in [*required, *excluded]
    }
    for chunk in chunks:
        _sheet_index, title = _sheet(chunk)
        for cell in chunk.get("cells", []):
            formula = str(cell.get("formula") or "").strip()
            key = _cell_key(chunk, cell)
            if not formula or key in classified:
                continue
            required.append(
                {
                    "sourceCellKey": key,
                    "chunkId": str(chunk.get("chunkId") or ""),
                    "sheet": title,
                    "coordinate": _coordinate(cell),
                    "numericValue": _numeric_value(cell),
                    "formula": formula,
                    "classification": "REQUIRED_FORMULA",
                    "contentClass": "FORMULA_RESULT",
                    "exclusionReason": "",
                }
            )
    return {
        "schemaVersion": "no-candidate-source-inventory-v1",
        "requiredCells": required,
        "excludedCells": excluded,
        "requiredCellCount": len(required),
        "excludedCellCount": len(excluded),
    }


def audit_unselected_source_inventory(
    *,
    packet_set: dict[str, Any],
    locator_results: Sequence[dict[str, Any]],
    selected_source_cell_keys: Sequence[str],
) -> dict[str, Any]:
    """Find required source results outside the selected draft universe."""

    inventory = audit_no_candidate_source_inventory(
        packet_set=packet_set,
        locator_results=locator_results,
    )
    selected = {str(value) for value in selected_source_cell_keys}
    unselected = [
        item
        for item in inventory["requiredCells"]
        if str(item.get("sourceCellKey") or "") not in selected
    ]
    return {
        "schemaVersion": "unselected-source-inventory-v1",
        "requiredCells": unselected,
        "requiredCellCount": len(unselected),
        "selectedSourceCellCount": len(selected),
    }


def promote_required_source_locator_sections(
    *,
    locator_results: Sequence[dict[str, Any]],
    required_cells: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Promote fail-closed locator misses using exact required source cells."""

    required_by_chunk: dict[str, list[dict[str, Any]]] = {}
    for item in required_cells:
        chunk_id = str(item.get("chunkId") or "")
        source_key = str(item.get("sourceCellKey") or "")
        sheet = str(item.get("sheet") or "")
        coordinate = str(item.get("coordinate") or "")
        if not all((chunk_id, source_key, sheet, coordinate)):
            raise StagedDraftV2Error(
                "Required source promotion needs chunk, key, sheet, and coordinate"
            )
        required_by_chunk.setdefault(chunk_id, []).append(item)

    result = copy.deepcopy(list(locator_results))
    locator_by_id = {
        str(locator.get("chunkId") or ""): locator
        for locator in result
    }
    missing_chunk_ids = sorted(
        set(required_by_chunk).difference(locator_by_id)
    )
    if missing_chunk_ids:
        raise StagedDraftV2Error(
            "Required source promotion references unknown locator chunks: "
            + ", ".join(missing_chunk_ids)
        )

    for chunk_id, items in required_by_chunk.items():
        locator = locator_by_id[chunk_id]
        if _is_candidate(locator):
            continue
        ordered_items = sorted(
            items,
            key=lambda item: (
                str(item.get("sheet") or "").casefold(),
                _position(
                    {
                        "coordinate": str(
                            item.get("coordinate") or ""
                        )
                    }
                ),
                str(item.get("sourceCellKey") or ""),
            ),
        )
        unique_items: list[dict[str, Any]] = []
        seen_keys: set[str] = set()
        for item in ordered_items:
            source_key = str(item["sourceCellKey"])
            if source_key in seen_keys:
                continue
            seen_keys.add(source_key)
            unique_items.append(item)
        evidence = [
            {
                "sheet": str(item["sheet"]),
                "range": str(item["coordinate"]),
                "role": "REQUIRED_SOURCE_COVERAGE",
            }
            for item in unique_items
        ]
        promotion_id = stable_uid(
            "required-source-locator-promotion-v1",
            chunk_id,
            *(str(item["sourceCellKey"]) for item in unique_items),
        )
        locator["status"] = "NEEDS_REVIEW"
        locator["candidates"] = [
            {
                "key": promotion_id,
                "title": "Required source content review",
                "summary": (
                    "Deterministic coverage safeguard for source results "
                    "missed by the semantic locator."
                ),
                "designHint": (
                    "Review the complete captured section containing these "
                    "exact required source cells."
                ),
                "contexts": [],
                "changedFactors": [],
                "outcomes": [],
                "comparisonHints": [],
                "evidence": evidence,
                "limitations": [
                    "Candidate status was promoted deterministically from "
                    "required source-content coverage, not inferred as a "
                    "comparison or causal Study."
                ],
                "confidence": 0.0,
            }
        ]
        locator["deterministicCoveragePromotion"] = {
            "schemaVersion": "required-source-locator-promotion-v1",
            "promotionId": promotion_id,
            "requiredSourceCellKeys": [
                str(item["sourceCellKey"]) for item in unique_items
            ],
            "contentClasses": sorted(
                {
                    str(item.get("contentClass") or "")
                    for item in unique_items
                    if str(item.get("contentClass") or "")
                }
            ),
        }
    return result


def build_monolithic_request(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    universe: dict[str, Any],
    content_complete: bool,
    prompt_text: str,
) -> dict[str, Any]:
    """Describe and hash the exact monolithic runner request."""

    envelope = {
        "source": copy.deepcopy(source),
        "workbook": copy.deepcopy(workbook),
        "locatorResults": copy.deepcopy(
            universe["selectedLocatorResults"]
        ),
        "focusedChunks": copy.deepcopy(universe["selectedChunks"]),
        "contentComplete": bool(content_complete),
    }
    prompt_bytes = prompt_text.encode("utf-8")
    return {
        "envelope": envelope,
        "envelopeBytes": len(compact_json_bytes(envelope)),
        "envelopeSha256": json_sha256(envelope),
        "promptText": prompt_text,
        "promptBytes": len(prompt_bytes),
        "promptSha256": bytes_sha256(prompt_bytes),
    }


def assess_one_call_budget(
    *,
    request: dict[str, Any],
    max_prompt_bytes: int,
    max_source_cells: int | None = None,
) -> dict[str, Any]:
    if max_prompt_bytes < 1:
        raise StagedDraftV2Error(
            "Monolithic Study-draft byte limit must be positive"
        )
    if max_source_cells is not None and max_source_cells < 1:
        raise StagedDraftV2Error(
            "Monolithic Study-draft source-cell limit must be positive"
        )
    prompt_bytes = int(request.get("promptBytes") or 0)
    envelope = request.get("envelope")
    focused_chunks = (
        envelope.get("focusedChunks", [])
        if isinstance(envelope, dict)
        else []
    )
    source_cell_count = sum(
        len(chunk.get("cells", []))
        for chunk in focused_chunks
        if isinstance(chunk, dict)
    )
    prompt_within_budget = prompt_bytes <= max_prompt_bytes
    cells_within_budget = (
        max_source_cells is None
        or source_cell_count <= max_source_cells
    )
    return {
        "mode": (
            "MONOLITHIC"
            if prompt_within_budget and cells_within_budget
            else "STAGED_V2"
        ),
        "promptBytes": prompt_bytes,
        "maxPromptBytes": max_prompt_bytes,
        "sourceCellCount": source_cell_count,
        "maxSourceCells": max_source_cells,
        "promptSha256": str(request.get("promptSha256") or ""),
        "envelopeSha256": str(request.get("envelopeSha256") or ""),
    }


def _source_cell_maps(
    chunks: Sequence[dict[str, Any]],
) -> tuple[
    dict[tuple[str, str], tuple[str, dict[str, Any]]],
    dict[str, tuple[str, str, dict[str, Any]]],
    dict[str, int],
]:
    by_coordinate: dict[
        tuple[str, str], tuple[str, dict[str, Any]]
    ] = {}
    by_key: dict[str, tuple[str, str, dict[str, Any]]] = {}
    order: dict[str, int] = {}
    for chunk in chunks:
        _sheet_index, title = _sheet(chunk)
        for collection in ("cells", "contextCells"):
            for cell in chunk.get(collection, []):
                coordinate = _coordinate(cell)
                if not coordinate:
                    continue
                key = _cell_key(chunk, cell)
                by_coordinate.setdefault(
                    (title.casefold(), coordinate),
                    (key, cell),
                )
                by_key.setdefault(key, (title, coordinate, cell))
        for cell in chunk.get("cells", []):
            key = _cell_key(chunk, cell)
            order.setdefault(key, len(order))
    return by_coordinate, by_key, order


class _EvidenceCellIndex:
    """Sparse sheet/row index for repeated evidence-range expansion."""

    def __init__(
        self,
        *,
        by_coordinate: dict[
            tuple[str, str],
            tuple[str, dict[str, Any]],
        ],
        source_order: dict[str, int],
    ) -> None:
        cells_by_sheet_row: dict[
            str,
            dict[int, list[tuple[int, str, dict[str, Any]]]],
        ] = {}
        for (
            sheet,
            _coordinate_value,
        ), (key, cell) in by_coordinate.items():
            row, _column = _position(cell)
            cells_by_sheet_row.setdefault(sheet, {}).setdefault(
                row,
                [],
            ).append(
                (
                    source_order.get(key, 10**12),
                    key,
                    cell,
                )
            )
        self._cells_by_sheet_row = cells_by_sheet_row
        self._rows_by_sheet = {
            sheet: sorted(rows)
            for sheet, rows in cells_by_sheet_row.items()
        }

    def keys_in_range(
        self,
        *,
        sheet: str,
        bounds: tuple[int, int, int, int],
    ) -> list[tuple[int, str]]:
        sheet_key = sheet.casefold()
        rows = self._rows_by_sheet.get(sheet_key, [])
        start_row, start_column, end_row, end_column = bounds
        start_index = bisect_left(rows, start_row)
        end_index = bisect_right(rows, end_row)
        return [
            (order, key)
            for row in rows[start_index:end_index]
            for order, key, cell in self._cells_by_sheet_row[
                sheet_key
            ][row]
            if start_column <= _position(cell)[1] <= end_column
        ]


def evidence_cell_keys(
    evidence: Sequence[dict[str, Any]],
    *,
    chunks: Sequence[dict[str, Any]],
    _cell_index: _EvidenceCellIndex | None = None,
) -> list[str]:
    """Expand evidence ranges to captured source-cell keys in source order."""

    if _cell_index is None:
        by_coordinate, _by_key, order = _source_cell_maps(chunks)
        cell_index = _EvidenceCellIndex(
            by_coordinate=by_coordinate,
            source_order=order,
        )
    else:
        cell_index = _cell_index
    result: list[str] = []
    seen: set[str] = set()
    for item in evidence:
        if not isinstance(item, dict):
            raise StagedDraftV2Error("Evidence entries must be objects")
        sheet = str(item.get("sheet") or "")
        matches = cell_index.keys_in_range(
            sheet=sheet,
            bounds=range_bounds(item.get("range")),
        )
        for _position_value, key in sorted(matches):
            if key not in seen:
                seen.add(key)
                result.append(key)
    return result


def build_study_registry_v2(
    *,
    source: dict[str, Any],
    universe: dict[str, Any],
    link_proposals: Sequence[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Create exact anchors and validate source-backed Study link proposals.

    Automatic links require a repeated captured context/header cell. Optional
    caller proposals are accepted only when exact link evidence touches every
    proposed candidate. Semantic title similarity alone never links Studies.
    """

    chunks = universe["selectedChunks"]
    chunk_by_id = {
        str(chunk["chunkId"]): chunk for chunk in chunks
    }
    by_coordinate, by_key, source_order = _source_cell_maps(chunks)
    cell_index = _EvidenceCellIndex(
        by_coordinate=by_coordinate,
        source_order=source_order,
    )
    candidate_anchors: list[dict[str, Any]] = []
    assigned_candidates: set[tuple[str, int]] = set()
    for locator in universe["selectedLocatorResults"]:
        chunk_id = str(locator.get("chunkId") or "")
        if chunk_id not in chunk_by_id:
            continue
        chunk = chunk_by_id[chunk_id]
        for ordinal, candidate in enumerate(
            locator.get("candidates", []),
            start=1,
        ):
            if not isinstance(candidate, dict):
                raise StagedDraftV2Error(
                    "Locator candidate must be an object"
                )
            assignment = (chunk_id, ordinal)
            if assignment in assigned_candidates:
                raise StagedDraftV2Error(
                    "Locator candidate was assigned more than once"
                )
            assigned_candidates.add(assignment)
            cell_keys = evidence_cell_keys(
                candidate.get("evidence", []),
                chunks=chunks,
                _cell_index=cell_index,
            )
            if not cell_keys:
                raise StagedDraftV2Error(
                    f"Candidate {chunk_id}[{ordinal}] lacks exact evidence cells"
                )
            anchor_id = stable_uid(
                "candidate-anchor-v2",
                source["revisionUid"],
                chunk_id,
                ordinal,
                *cell_keys,
            )
            context_keys = [
                _cell_key(chunk, cell)
                for cell in chunk.get("contextCells", [])
                if _cell_key(chunk, cell) in by_key
            ]
            candidate_anchors.append(
                {
                    "candidateAnchorId": anchor_id,
                    "chunkId": chunk_id,
                    "candidateOrdinal": ordinal,
                    "candidateKey": str(candidate.get("key") or ""),
                    "titleHint": str(candidate.get("title") or ""),
                    "section": list(_section(chunk)),
                    "evidenceCellKeys": cell_keys,
                    "structuralContextCellKeys": context_keys,
                    "evidence": copy.deepcopy(
                        candidate.get("evidence", [])
                    ),
                }
            )

    anchor_by_id = {
        str(anchor["candidateAnchorId"]): anchor
        for anchor in candidate_anchors
    }

    def exact_evidence_for_keys(
        keys: Sequence[str],
    ) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for key in sorted(
            {str(value) for value in keys},
            key=lambda value: source_order.get(value, 10**12),
        ):
            title, coordinate, _cell = by_key[key]
            result.append(
                {
                    "sheet": title,
                    "range": coordinate,
                    "role": "REGISTRY_LINK",
                }
            )
        return result

    proposed_values: list[dict[str, Any]]
    if link_proposals is None:
        proposed_values = []
        automatic_parent = {
            str(anchor["candidateAnchorId"]): str(
                anchor["candidateAnchorId"]
            )
            for anchor in candidate_anchors
        }

        def automatic_root(anchor_id: str) -> str:
            parent = automatic_parent[anchor_id]
            while parent != automatic_parent[parent]:
                parent = automatic_parent[parent]
            while anchor_id != parent:
                next_anchor_id = automatic_parent[anchor_id]
                automatic_parent[anchor_id] = parent
                anchor_id = next_anchor_id
            return parent

        for left_index, left in enumerate(candidate_anchors):
            left_support = set(left["evidenceCellKeys"]) | set(
                left["structuralContextCellKeys"]
            )
            for right in candidate_anchors[left_index + 1 :]:
                if left["section"] != right["section"]:
                    continue
                right_support = set(right["evidenceCellKeys"]) | set(
                    right["structuralContextCellKeys"]
                )
                direct_overlap = set(left["evidenceCellKeys"]).intersection(
                    right["evidenceCellKeys"]
                )
                repeated_context = (
                    set(left["structuralContextCellKeys"])
                    .intersection(right_support)
                    | set(right["structuralContextCellKeys"]).intersection(
                        left_support
                    )
                )
                link_keys = direct_overlap | repeated_context
                if not link_keys:
                    continue
                left_id = str(left["candidateAnchorId"])
                right_id = str(right["candidateAnchorId"])
                left_root = automatic_root(left_id)
                right_root = automatic_root(right_id)
                if left_root == right_root:
                    # One deterministic spanning edge per newly connected
                    # component preserves the exact linkage without emitting
                    # every redundant pair in dense candidate sections.
                    continue
                automatic_parent[right_root] = left_root
                link_key = min(
                    link_keys,
                    key=lambda key: source_order.get(key, 10**12),
                )
                proposed_values.append(
                    {
                        "memberCandidateAnchorIds": [
                            left_id,
                            right_id,
                        ],
                        "linkEvidence": exact_evidence_for_keys(
                            [link_key]
                        ),
                        "reason": (
                            "OVERLAPPING_CANDIDATE_EVIDENCE"
                            if direct_overlap
                            else "REPEATED_STRUCTURAL_CONTEXT"
                        ),
                    }
                )
    else:
        proposed_values = [
            copy.deepcopy(value) for value in link_proposals
        ]

    validated_proposals: list[dict[str, Any]] = []
    for index, proposal in enumerate(proposed_values):
        if not isinstance(proposal, dict):
            raise StagedDraftV2Error(
                f"Registry link proposal {index} must be an object"
            )
        members = [
            str(value)
            for value in proposal.get(
                "memberCandidateAnchorIds",
                [],
            )
        ]
        if len(set(members)) < 2 or any(
            member not in anchor_by_id for member in members
        ):
            raise StagedDraftV2Error(
                f"Registry link proposal {index} has unknown/duplicate members"
            )
        member_anchors = [anchor_by_id[member] for member in members]
        if len(
            {
                tuple(str(value) for value in anchor["section"])
                for anchor in member_anchors
            }
        ) != 1:
            raise StagedDraftV2Error(
                "Registry links cannot cross captured source sections"
            )
        link_evidence = proposal.get("linkEvidence")
        if not isinstance(link_evidence, list) or not link_evidence:
            raise StagedDraftV2Error(
                f"Registry link proposal {index} requires exact linkEvidence"
            )
        link_keys = evidence_cell_keys(
            link_evidence,
            chunks=chunks,
            _cell_index=cell_index,
        )
        if not link_keys:
            raise StagedDraftV2Error(
                f"Registry link proposal {index} has no captured link cells"
            )
        for anchor in member_anchors:
            support = set(anchor["evidenceCellKeys"]) | set(
                anchor["structuralContextCellKeys"]
            )
            if not support.intersection(link_keys):
                raise StagedDraftV2Error(
                    f"Registry link evidence does not touch candidate "
                    f"{anchor['candidateAnchorId']}"
                )
        normalized_members = sorted(set(members))
        normalized_link_keys = sorted(
            set(link_keys),
            key=lambda key: source_order.get(key, 10**12),
        )
        proposal_id = stable_uid(
            "registry-link-proposal-v2",
            source["revisionUid"],
            *normalized_members,
            *normalized_link_keys,
        )
        if proposal.get("proposalId") not in (None, "", proposal_id):
            raise StagedDraftV2Error(
                f"Registry link proposal {index} has a stale proposalId"
            )
        validated_proposals.append(
            {
                "proposalId": proposal_id,
                "memberCandidateAnchorIds": normalized_members,
                "linkEvidenceCellKeys": normalized_link_keys,
                "linkEvidence": copy.deepcopy(link_evidence),
                "reason": str(proposal.get("reason") or ""),
            }
        )

    # Build logical components only from validated source-backed proposals.
    adjacency: dict[str, set[str]] = {
        anchor_id: set() for anchor_id in anchor_by_id
    }
    proposal_by_member: dict[str, list[dict[str, Any]]] = {}
    for proposal in validated_proposals:
        members = proposal["memberCandidateAnchorIds"]
        for member in members:
            adjacency[member].update(
                value for value in members if value != member
            )
            proposal_by_member.setdefault(member, []).append(proposal)
    components: list[list[dict[str, Any]]] = []
    remaining_ids = set(anchor_by_id)
    for anchor in candidate_anchors:
        anchor_id = str(anchor["candidateAnchorId"])
        if anchor_id not in remaining_ids:
            continue
        stack = [anchor_id]
        component_ids: list[str] = []
        while stack:
            current = stack.pop()
            if current not in remaining_ids:
                continue
            remaining_ids.remove(current)
            component_ids.append(current)
            stack.extend(sorted(adjacency[current], reverse=True))
        components.append(
            [anchor_by_id[value] for value in component_ids]
        )

    studies: list[dict[str, Any]] = []
    for component in components:
        member_ids = sorted(
            str(item["candidateAnchorId"]) for item in component
        )
        link_keys = sorted(
            {
                key
                for item in component
                for key in item["evidenceCellKeys"]
            },
            key=lambda key: source_order.get(key, 10**12),
        )
        component_proposals = {
            proposal["proposalId"]: proposal
            for member_id in member_ids
            for proposal in proposal_by_member.get(member_id, [])
            if set(proposal["memberCandidateAnchorIds"]).issubset(
                set(member_ids)
            )
        }
        registry_link_keys = sorted(
            {
                key
                for proposal in component_proposals.values()
                for key in proposal["linkEvidenceCellKeys"]
            },
            key=lambda key: source_order.get(key, 10**12),
        )
        logical_id = stable_uid(
            "logical-study-v2",
            source["revisionUid"],
            *member_ids,
            *link_keys,
            *registry_link_keys,
        )
        section = component[0]["section"]
        section_anchor_id = stable_uid(
            "section-anchor-v2",
            source["revisionUid"],
            *section,
            *member_ids,
        )
        studies.append(
            {
                "logicalStudyId": logical_id,
                "sectionAnchorId": section_anchor_id,
                "section": copy.deepcopy(section),
                "memberCandidateAnchorIds": member_ids,
                "anchorEvidenceCellKeys": link_keys,
                "registryLinkProposalIds": sorted(
                    component_proposals
                ),
                "registryLinkEvidenceCellKeys": registry_link_keys,
                "titleHint": next(
                    (
                        str(item["titleHint"])
                        for item in component
                        if str(item["titleHint"]).strip()
                    ),
                    "",
                ),
            }
        )
    anchor_members = [
        member
        for study in studies
        for member in study["memberCandidateAnchorIds"]
    ]
    expected_members = [
        str(item["candidateAnchorId"]) for item in candidate_anchors
    ]
    if sorted(anchor_members) != sorted(expected_members):
        raise StagedDraftV2Error(
            "Registry must assign every candidate anchor exactly once"
        )
    registry = {
        "schemaVersion": STUDY_REGISTRY_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": source["revisionUid"],
            "contentSha256": str(source["contentSha256"]).lower(),
        },
        "candidateAnchors": candidate_anchors,
        "linkProposals": validated_proposals,
        "studies": studies,
    }
    registry["registrySha256"] = json_sha256(registry)
    return registry


def _fragment_payload_bytes(
    *,
    chunks: Sequence[dict[str, Any]],
    locator_results: Sequence[dict[str, Any]],
) -> int:
    return len(
        compact_json_bytes(
            {
                "focusedChunks": list(chunks),
                "locatorResults": list(locator_results),
            }
        )
    )


_PLAN_ID_BUDGET_PLACEHOLDER = stable_uid(
    "study-draft-plan-v2",
    "exact-prompt-budget-placeholder",
)


def _fragment_identity_v2() -> dict[str, str]:
    """Return every live contract version that defines a fragment run."""

    return {
        "promptVersion": FRAGMENT_PROMPT_VERSION,
        "fragmentContractVersion": FRAGMENT_CONTRACT_VERSION,
        "validatorContractVersion": FRAGMENT_VALIDATOR_VERSION,
        "consolidatorContractVersion": CONSOLIDATOR_CONTRACT_VERSION,
    }


def _require_current_fragment_identity(
    plan: dict[str, Any],
    part: dict[str, Any] | None = None,
) -> tuple[dict[str, str], str]:
    identity = _fragment_identity_v2()
    identity_sha256 = json_sha256(identity)
    if (
        plan.get("fragmentIdentity") != identity
        or plan.get("fragmentIdentitySha256") != identity_sha256
        or (
            part is not None
            and part.get("fragmentIdentitySha256") != identity_sha256
        )
    ):
        raise StagedDraftV2Error(
            "Fragment plan identity does not match the live prompt "
            "and contracts"
        )
    return identity, identity_sha256


def _registry_studies_for_source_scope(
    *,
    registry: dict[str, Any],
    section_value: tuple[int, str, object],
    source_chunk_ids: Sequence[str] = (),
) -> list[dict[str, Any]]:
    section_studies = [
        study
        for study in registry["studies"]
        if (
            int(study["section"][0]),
            str(study["section"][1]),
            str(study["section"][2]),
        )
        == (
            int(section_value[0]),
            str(section_value[1]),
            str(section_value[2]),
        )
    ]
    scoped_chunk_ids = {str(value) for value in source_chunk_ids}
    if not scoped_chunk_ids:
        return section_studies
    scoped_anchor_ids = {
        str(anchor["candidateAnchorId"])
        for anchor in registry.get("candidateAnchors", [])
        if str(anchor.get("chunkId") or "") in scoped_chunk_ids
    }
    if not scoped_anchor_ids:
        # A continuation-only chunk has no unambiguous local Study anchor.
        # Retaining the full section registry is the only lossless choice.
        return section_studies
    scoped_studies = [
        study
        for study in section_studies
        if scoped_anchor_ids.intersection(
            str(value)
            for value in study.get("memberCandidateAnchorIds", [])
        )
    ]
    if not scoped_studies:
        raise StagedDraftV2Error(
            "Source segment candidate anchors are absent from the registry"
        )
    return scoped_studies


def _planned_fragment_part_v2(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    registry: dict[str, Any],
    prompt_version: str,
    fragment_identity: dict[str, str],
    fragment_identity_sha256: str,
    part_index: int,
    part_chunks: Sequence[dict[str, Any]],
    locator_by_id: dict[str, dict[str, Any]],
    selected_by_key: dict[
        str,
        tuple[str, str, dict[str, Any]],
    ],
    source_order: dict[str, int],
    plan_id: str = _PLAN_ID_BUDGET_PLACEHOLDER,
    source_segments: Sequence[dict[str, Any]] = (),
) -> dict[str, Any]:
    """Build one part and measure its exact finalized prompt."""

    chunk_ids = [str(chunk["chunkId"]) for chunk in part_chunks]
    owned = [
        _cell_key(chunk, cell)
        for chunk in part_chunks
        for cell in chunk.get("cells", [])
    ]
    owned_set = set(owned)
    section_value = _section(part_chunks[0])
    registry_slice = _registry_studies_for_source_scope(
        registry=registry,
        section_value=section_value,
        source_chunk_ids=[
            str(value.get("sourceChunkId") or "")
            for value in source_segments
        ],
    )
    registry_member_ids = {
        str(value)
        for study in registry_slice
        for value in study.get("memberCandidateAnchorIds", [])
    }
    source_chunk_scope = {
        str(value.get("sourceChunkId") or "")
        for value in source_segments
    }
    focused_anchors = [
        anchor
        for anchor in registry.get("candidateAnchors", [])
        if str(anchor.get("candidateAnchorId") or "")
        in registry_member_ids
        and (
            not source_chunk_scope
            or str(anchor.get("chunkId") or "") in source_chunk_scope
        )
    ]
    if source_chunk_scope and not focused_anchors:
        focused_anchors = [
            anchor
            for anchor in registry.get("candidateAnchors", [])
            if str(anchor.get("candidateAnchorId") or "")
            in registry_member_ids
        ]
    focused_anchor_ids = [
        str(anchor["candidateAnchorId"]) for anchor in focused_anchors
    ]
    if source_segments:
        registry_anchor_keys = {
            str(key)
            for anchor in focused_anchors
            for key in [
                *anchor.get("evidenceCellKeys", []),
                *anchor.get("structuralContextCellKeys", []),
            ]
        }
        registry_anchor_keys.update(
            str(key)
            for study in registry_slice
            for key in study.get("registryLinkEvidenceCellKeys", [])
        )
    else:
        registry_anchor_keys = {
            str(key)
            for study in registry_slice
            for key in [
                *study["anchorEvidenceCellKeys"],
                *study.get("registryLinkEvidenceCellKeys", []),
            ]
        }
    shared_candidates = {
        key for key in registry_anchor_keys if key not in owned_set
    }
    for chunk in part_chunks:
        for cell in chunk.get("contextCells", []):
            key = _cell_key(chunk, cell)
            if key not in owned_set and key in selected_by_key:
                shared_candidates.add(key)
    shared = sorted(
        shared_candidates,
        key=lambda key: source_order[key],
    )
    part_id = stable_uid(
        "study-draft-part-v2",
        source["revisionUid"],
        prompt_version,
        fragment_identity_sha256,
        *(
            [json_sha256(list(source_segments))]
            if source_segments
            else []
        ),
        *chunk_ids,
        *owned,
        *shared,
    )
    locator_values = [
        locator_by_id[chunk_id] for chunk_id in chunk_ids
    ]
    part = {
        "partIndex": part_index,
        "partId": part_id,
        "fragmentIdentitySha256": fragment_identity_sha256,
        "section": list(section_value),
        "chunkIds": chunk_ids,
        "ownedSourceCellKeys": owned,
        "sharedAnchorCellKeys": shared,
        "logicalStudyIds": [
            study["logicalStudyId"] for study in registry_slice
        ],
        "candidateAnchorIds": focused_anchor_ids,
        "registryAnchorCellKeys": sorted(
            registry_anchor_keys,
            key=lambda key: source_order[key],
        ),
        "sectionAnchorIds": [
            study["sectionAnchorId"] for study in registry_slice
        ],
        "chunkCount": len(part_chunks),
        "cellCount": len(owned),
        "sourceSegments": copy.deepcopy(list(source_segments)),
        "serializedBytes": _fragment_payload_bytes(
            chunks=part_chunks,
            locator_results=locator_values,
        ),
    }
    envelope = finalize_fragment_envelope(
        build_fragment_envelope(
            source=source,
            workbook=workbook,
            plan={
                "planId": plan_id,
                "fragmentIdentity": copy.deepcopy(fragment_identity),
                "fragmentIdentitySha256": fragment_identity_sha256,
            },
            part=part,
            focused_chunks=part_chunks,
            locator_results=locator_values,
            registry_slice=registry_for_part(
                registry,
                part,
            ),
        )
    )
    part["promptBytes"] = int(envelope["promptBytes"])
    return part


def _is_merged_source_cell(cell: dict[str, Any]) -> bool:
    merge_role = str(cell.get("mergeRole") or "").strip().lower()
    merge_value = cell.get("merge")
    nested_role = (
        str(merge_value.get("role") or "").strip().lower()
        if isinstance(merge_value, dict)
        else ""
    )
    return bool(
        cell.get("mergeRange")
        or merge_role not in {"", "none"}
        or nested_role not in {"", "none"}
    )


def _as_shared_context_cell(cell: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(cell)
    result["contextOnly"] = True
    result["primary"] = False
    return result


def _segment_shared_context_keys(
    *,
    chunk: dict[str, Any],
    registry: dict[str, Any],
    selected_by_key: dict[
        str,
        tuple[str, str, dict[str, Any]],
    ],
    source_order: dict[str, int],
) -> list[str]:
    """Return exact structural cells repeated around every chunk segment."""

    keys: set[str] = set()
    for cell in chunk.get("contextCells", []):
        key = _cell_key(chunk, cell)
        if key in selected_by_key:
            keys.add(key)
    for cell in chunk.get("cells", []):
        key = _cell_key(chunk, cell)
        if key in selected_by_key and _is_merged_source_cell(cell):
            keys.add(key)

    studies = _registry_studies_for_source_scope(
        registry=registry,
        section_value=_section(chunk),
        source_chunk_ids=[str(chunk.get("chunkId") or "")],
    )
    member_ids = {
        str(value)
        for study in studies
        for value in study.get("memberCandidateAnchorIds", [])
    }
    for study in studies:
        keys.update(
            str(value)
            for value in study.get(
                "registryLinkEvidenceCellKeys",
                [],
            )
            if str(value) in selected_by_key
        )
    for anchor in registry.get("candidateAnchors", []):
        if (
            str(anchor.get("candidateAnchorId") or "") not in member_ids
            or str(anchor.get("chunkId") or "")
            != str(chunk.get("chunkId") or "")
        ):
            continue
        evidence_keys = [
            str(value)
            for value in anchor.get("evidenceCellKeys", [])
            if str(value) in selected_by_key
        ]
        if evidence_keys:
            # The complete exact anchor key set remains in the scoped
            # registry. Repeating one literal cell supplies the label/context
            # without duplicating the whole owned candidate region.
            keys.add(evidence_keys[0])
        keys.update(
            str(value)
            for value in anchor.get(
                "structuralContextCellKeys",
                [],
            )
            if str(value) in selected_by_key
        )
    return sorted(keys, key=lambda key: source_order[key])


def _source_segment_descriptor(
    *,
    chunk: dict[str, Any],
    cells: Sequence[dict[str, Any]],
    shared_context_keys: Sequence[str],
) -> dict[str, Any]:
    if not cells:
        raise StagedDraftV2Error(
            "Source chunk segmentation requires at least one source cell"
        )
    owned = [_cell_key(chunk, cell) for cell in cells]
    return {
        "schemaVersion": SOURCE_CHUNK_SEGMENT_SCHEMA_VERSION,
        "sourceChunkId": str(chunk.get("chunkId") or ""),
        "sourceChunkSha256": json_sha256(chunk),
        "firstSourceCellKey": owned[0],
        "lastSourceCellKey": owned[-1],
        "sourceCellCount": len(owned),
        "sharedContextCellKeys": list(shared_context_keys),
    }


def _source_segment_chunk(
    *,
    source_chunk: dict[str, Any],
    cells: Sequence[dict[str, Any]],
    descriptor: dict[str, Any],
    selected_by_key: dict[
        str,
        tuple[str, str, dict[str, Any]],
    ],
) -> dict[str, Any]:
    owned = {_cell_key(source_chunk, cell) for cell in cells}
    shared_keys = [
        str(value)
        for value in descriptor.get("sharedContextCellKeys", [])
    ]
    if len(set(shared_keys)) != len(shared_keys):
        raise StagedDraftV2Error(
            "Source segment shared context contains duplicate cell keys"
        )
    if owned.intersection(shared_keys):
        raise StagedDraftV2Error(
            "Source segment cannot repeat an owned cell as shared context"
        )
    try:
        context_cells = [
            _as_shared_context_cell(selected_by_key[key][2])
            for key in shared_keys
        ]
    except KeyError as exc:
        raise StagedDraftV2Error(
            f"Source segment references unknown context cell {exc.args[0]}"
        ) from exc
    result = copy.deepcopy(source_chunk)
    result["cells"] = copy.deepcopy(list(cells))
    result["contextCells"] = context_cells
    result["sourceSegment"] = {
        key: copy.deepcopy(value)
        for key, value in descriptor.items()
        if key != "sharedContextCellKeys"
    }
    return result


def _split_chunk_for_safe_parts(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    registry: dict[str, Any],
    prompt_version: str,
    fragment_identity: dict[str, str],
    fragment_identity_sha256: str,
    chunk: dict[str, Any],
    locator_by_id: dict[str, dict[str, Any]],
    selected_by_key: dict[
        str,
        tuple[str, str, dict[str, Any]],
    ],
    source_order: dict[str, int],
    first_part_index: int,
    max_cells: int,
    max_prompt_bytes: int,
) -> list[tuple[dict[str, Any], dict[str, Any]]]:
    """Split one chunk without truncation, using exact finalized prompts."""

    cells = list(chunk.get("cells", []))
    if not cells:
        raise StagedDraftV2Error(
            f"Chunk {chunk.get('chunkId')} has an oversized atomic "
            "context envelope with no source cell boundary"
        )
    source_keys = [_cell_key(chunk, cell) for cell in cells]
    positions = [_position(cell) for cell in cells]
    if (
        len(source_keys) != len(set(source_keys))
        or positions != sorted(positions)
    ):
        raise StagedDraftV2Error(
            f"Chunk {chunk.get('chunkId')} has no unique monotonic "
            "source-cell order for lossless segmentation"
        )
    base_context_keys = _segment_shared_context_keys(
        chunk=chunk,
        registry=registry,
        selected_by_key=selected_by_key,
        source_order=source_order,
    )
    results: list[tuple[dict[str, Any], dict[str, Any]]] = []
    offset = 0
    while offset < len(cells):
        candidate_count = min(max_cells, len(cells) - offset)
        accepted: tuple[dict[str, Any], dict[str, Any]] | None = None
        measured_prompt_bytes = 0
        while candidate_count >= 1:
            candidate_cells = cells[offset : offset + candidate_count]
            owned = {
                _cell_key(chunk, cell) for cell in candidate_cells
            }
            shared_context_keys = [
                key for key in base_context_keys if key not in owned
            ]
            descriptor = _source_segment_descriptor(
                chunk=chunk,
                cells=candidate_cells,
                shared_context_keys=shared_context_keys,
            )
            segment = _source_segment_chunk(
                source_chunk=chunk,
                cells=candidate_cells,
                descriptor=descriptor,
                selected_by_key=selected_by_key,
            )
            measured = _planned_fragment_part_v2(
                source=source,
                workbook=workbook,
                registry=registry,
                prompt_version=prompt_version,
                fragment_identity=fragment_identity,
                fragment_identity_sha256=fragment_identity_sha256,
                part_index=first_part_index + len(results),
                part_chunks=[segment],
                locator_by_id=locator_by_id,
                selected_by_key=selected_by_key,
                source_order=source_order,
                source_segments=[descriptor],
            )
            measured_prompt_bytes = int(measured["promptBytes"])
            if measured_prompt_bytes <= max_prompt_bytes:
                accepted = (segment, descriptor)
                break
            if candidate_count == 1:
                break
            candidate_count = max(1, candidate_count // 2)
        if accepted is None:
            cell = cells[offset]
            raise StagedDraftV2Error(
                f"Chunk {chunk.get('chunkId')} atomic source cell "
                f"{_cell_key(chunk, cell)} plus its context envelope "
                f"requires {measured_prompt_bytes} prompt bytes, above "
                f"the safe limit {max_prompt_bytes}"
            )
        results.append(accepted)
        offset += candidate_count

    flattened = [
        _cell_key(segment, cell)
        for segment, _descriptor in results
        for cell in segment.get("cells", [])
    ]
    expected = [_cell_key(chunk, cell) for cell in cells]
    if flattened != expected or len(flattened) != len(set(flattened)):
        raise StagedDraftV2Error(
            "Source chunk segmentation changed source order or ownership"
        )
    return results


def plan_study_draft_v2(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    universe: dict[str, Any],
    registry: dict[str, Any],
    prompt_version: str,
    max_chunks: int,
    max_cells: int,
    max_serialized_bytes: int,
) -> dict[str, Any]:
    """Pack source-contiguous parts by their exact finalized prompt bytes."""

    if min(max_chunks, max_cells, max_serialized_bytes) < 1:
        raise StagedDraftV2Error("Part limits must be positive")
    selected_chunks = list(universe["selectedChunks"])
    fragment_identity = _fragment_identity_v2()
    fragment_identity_sha256 = json_sha256(fragment_identity)
    locator_by_id = {
        str(result["chunkId"]): result
        for result in universe["selectedLocatorResults"]
    }
    (
        _selected_by_coordinate,
        selected_by_key,
        _selected_source_order,
    ) = _source_cell_maps(selected_chunks)
    source_order = {
        key: index
        for index, key in enumerate(selected_by_key)
    }
    raw_parts: list[
        tuple[list[dict[str, Any]], list[dict[str, Any]]]
    ] = []
    section_chunks: list[dict[str, Any]] = []
    current_section: tuple[int, str, object] | None = None

    def flush() -> None:
        nonlocal section_chunks
        current: list[dict[str, Any]] = []
        current_cells = 0
        for chunk in section_chunks:
            chunk_cells = len(chunk.get("cells", []))
            proposed = [*current, chunk]
            exceeds_structural_limit = (
                len(proposed) > max_chunks
                or current_cells + chunk_cells > max_cells
            )
            proposed_prompt_bytes = max_serialized_bytes + 1
            if not exceeds_structural_limit:
                proposed_prompt_bytes = int(
                    _planned_fragment_part_v2(
                        source=source,
                        workbook=workbook,
                        registry=registry,
                        prompt_version=prompt_version,
                        fragment_identity=fragment_identity,
                        fragment_identity_sha256=(
                            fragment_identity_sha256
                        ),
                        part_index=len(raw_parts) + 1,
                        part_chunks=proposed,
                        locator_by_id=locator_by_id,
                        selected_by_key=selected_by_key,
                        source_order=source_order,
                    )["promptBytes"]
                )
            if (
                not exceeds_structural_limit
                and proposed_prompt_bytes <= max_serialized_bytes
            ):
                current = proposed
                current_cells += chunk_cells
                continue

            if current:
                raw_parts.append((current, []))
                current = []
                current_cells = 0
            single_prompt_bytes = int(
                _planned_fragment_part_v2(
                    source=source,
                    workbook=workbook,
                    registry=registry,
                    prompt_version=prompt_version,
                    fragment_identity=fragment_identity,
                    fragment_identity_sha256=(
                        fragment_identity_sha256
                    ),
                    part_index=len(raw_parts) + 1,
                    part_chunks=[chunk],
                    locator_by_id=locator_by_id,
                    selected_by_key=selected_by_key,
                    source_order=source_order,
                )["promptBytes"]
            )
            if (
                chunk_cells <= max_cells
                and single_prompt_bytes <= max_serialized_bytes
            ):
                current = [chunk]
                current_cells = chunk_cells
                continue

            segments = _split_chunk_for_safe_parts(
                source=source,
                workbook=workbook,
                registry=registry,
                prompt_version=prompt_version,
                fragment_identity=fragment_identity,
                fragment_identity_sha256=fragment_identity_sha256,
                chunk=chunk,
                locator_by_id=locator_by_id,
                selected_by_key=selected_by_key,
                source_order=source_order,
                first_part_index=len(raw_parts) + 1,
                max_cells=max_cells,
                max_prompt_bytes=max_serialized_bytes,
            )
            raw_parts.extend(
                ([segment], [descriptor])
                for segment, descriptor in segments
            )
        if current:
            raw_parts.append((current, []))
        section_chunks = []

    for chunk in selected_chunks:
        section = _section(chunk)
        if current_section is not None and section != current_section:
            flush()
        current_section = section
        section_chunks.append(chunk)
    flush()

    parts = [
        _planned_fragment_part_v2(
            source=source,
            workbook=workbook,
            registry=registry,
            prompt_version=prompt_version,
            fragment_identity=fragment_identity,
            fragment_identity_sha256=fragment_identity_sha256,
            part_index=index,
            part_chunks=part_chunks,
            locator_by_id=locator_by_id,
            selected_by_key=selected_by_key,
            source_order=source_order,
            source_segments=source_segments,
        )
        for index, (part_chunks, source_segments) in enumerate(
            raw_parts,
            start=1,
        )
    ]
    all_owned: list[str] = []
    for part in parts:
        all_owned.extend(part["ownedSourceCellKeys"])
    if all_owned != universe["ownedSourceCellKeys"]:
        raise StagedDraftV2Error(
            "Part ownership is not an exact source-ordered union"
        )
    plan = {
        "schemaVersion": STAGED_DRAFT_PLAN_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": source["revisionUid"],
            "contentSha256": str(source["contentSha256"]).lower(),
        },
        "promptVersion": prompt_version,
        "fragmentIdentity": copy.deepcopy(fragment_identity),
        "fragmentIdentitySha256": fragment_identity_sha256,
        "registrySha256": registry["registrySha256"],
        "selectedChunkIds": list(universe["selectedChunkIds"]),
        "continuationChunkIds": list(
            universe["continuationChunkIds"]
        ),
        "ownedSourceCellKeys": list(universe["ownedSourceCellKeys"]),
        "limits": {
            "maxChunks": max_chunks,
            "maxCells": max_cells,
            "maxSerializedBytes": max_serialized_bytes,
            "maxPromptBytes": max_serialized_bytes,
        },
        "sourceSegmentation": {
            "schemaVersion": SOURCE_CHUNK_SEGMENT_SCHEMA_VERSION,
            "segmentedChunkIds": list(
                dict.fromkeys(
                    str(segment["sourceChunkId"])
                    for part in parts
                    for segment in part.get("sourceSegments", [])
                )
            ),
            "segmentPartCount": sum(
                bool(part.get("sourceSegments")) for part in parts
            ),
        },
        "parts": parts,
    }
    plan["planId"] = stable_uid(
        "study-draft-plan-v2",
        source["revisionUid"],
        registry["registrySha256"],
        prompt_version,
        fragment_identity_sha256,
        max_chunks,
        max_cells,
        max_serialized_bytes,
        *(part["partId"] for part in parts),
    )
    for part in parts:
        exact = _planned_fragment_part_v2(
            source=source,
            workbook=workbook,
            registry=registry,
            prompt_version=prompt_version,
            fragment_identity=fragment_identity,
            fragment_identity_sha256=fragment_identity_sha256,
            part_index=int(part["partIndex"]),
            part_chunks=chunks_for_part_v2(universe, part),
            locator_by_id=locator_by_id,
            selected_by_key=selected_by_key,
            source_order=source_order,
            plan_id=plan["planId"],
            source_segments=part.get("sourceSegments", []),
        )
        if (
            exact["partId"] != part["partId"]
            or exact["promptBytes"] != part["promptBytes"]
            or exact["sourceSegments"] != part["sourceSegments"]
            or exact["promptBytes"] > max_serialized_bytes
        ):
            raise StagedDraftV2Error(
                f"Part {part['partId']} exact finalized prompt preflight "
                "is not deterministic or exceeds its limit"
            )
    return plan


def chunks_for_part_v2(
    universe: dict[str, Any],
    part: dict[str, Any],
) -> list[dict[str, Any]]:
    by_id = {
        str(chunk["chunkId"]): chunk
        for chunk in universe["selectedChunks"]
    }
    source_segments = part.get("sourceSegments", [])
    if source_segments:
        if (
            not isinstance(source_segments, list)
            or len(source_segments) != len(part.get("chunkIds", []))
            or [
                str(value.get("sourceChunkId") or "")
                for value in source_segments
                if isinstance(value, dict)
            ]
            != [str(value) for value in part.get("chunkIds", [])]
        ):
            raise StagedDraftV2Error(
                "Part source segments do not match its chunk references"
            )
        (
            _selected_by_coordinate,
            selected_by_key,
            _selected_source_order,
        ) = _source_cell_maps(universe["selectedChunks"])
        result: list[dict[str, Any]] = []
        reconstructed_owned: list[str] = []
        for descriptor in source_segments:
            if (
                not isinstance(descriptor, dict)
                or descriptor.get("schemaVersion")
                != SOURCE_CHUNK_SEGMENT_SCHEMA_VERSION
            ):
                raise StagedDraftV2Error(
                    "Part contains an invalid source segment descriptor"
                )
            chunk_id = str(descriptor.get("sourceChunkId") or "")
            try:
                source_chunk = by_id[chunk_id]
            except KeyError as exc:
                raise StagedDraftV2Error(
                    f"Part references unknown source chunk {chunk_id}"
                ) from exc
            if json_sha256(source_chunk) != str(
                descriptor.get("sourceChunkSha256") or ""
            ):
                raise StagedDraftV2Error(
                    f"Source segment {chunk_id} no longer matches its chunk"
                )
            cells = list(source_chunk.get("cells", []))
            cell_keys = [
                _cell_key(source_chunk, cell) for cell in cells
            ]
            first_key = str(
                descriptor.get("firstSourceCellKey") or ""
            )
            last_key = str(descriptor.get("lastSourceCellKey") or "")
            try:
                start = cell_keys.index(first_key)
                end = cell_keys.index(last_key, start) + 1
            except ValueError as exc:
                raise StagedDraftV2Error(
                    f"Source segment {chunk_id} boundary is stale"
                ) from exc
            segment_cells = cells[start:end]
            if len(segment_cells) != int(
                descriptor.get("sourceCellCount") or 0
            ):
                raise StagedDraftV2Error(
                    f"Source segment {chunk_id} is not source-contiguous"
                )
            segment = _source_segment_chunk(
                source_chunk=source_chunk,
                cells=segment_cells,
                descriptor=descriptor,
                selected_by_key=selected_by_key,
            )
            result.append(segment)
            reconstructed_owned.extend(
                _cell_key(segment, cell)
                for cell in segment.get("cells", [])
            )
        if reconstructed_owned != list(
            part.get("ownedSourceCellKeys", [])
        ):
            raise StagedDraftV2Error(
                "Reconstructed source segments changed part ownership"
            )
        shared = set(part.get("sharedAnchorCellKeys", []))
        if any(
            str(key) not in shared
            for descriptor in source_segments
            for key in descriptor.get("sharedContextCellKeys", [])
        ):
            raise StagedDraftV2Error(
                "Source segment context exceeds the part shared allowlist"
            )
        return result
    try:
        return [by_id[str(value)] for value in part["chunkIds"]]
    except KeyError as exc:
        raise StagedDraftV2Error(
            f"Part references unknown chunk {exc.args[0]}"
        ) from exc


def locators_for_part_v2(
    universe: dict[str, Any],
    part: dict[str, Any],
) -> list[dict[str, Any]]:
    by_id = {
        str(result["chunkId"]): result
        for result in universe["selectedLocatorResults"]
    }
    try:
        return [by_id[str(value)] for value in part["chunkIds"]]
    except KeyError as exc:
        raise StagedDraftV2Error(
            f"Part references unknown locator chunk {exc.args[0]}"
        ) from exc


def registry_for_part(
    registry: dict[str, Any],
    part: dict[str, Any],
) -> dict[str, Any]:
    logical_ids = set(part["logicalStudyIds"])
    studies = [
        copy.deepcopy(study)
        for study in registry["studies"]
        if study["logicalStudyId"] in logical_ids
    ]
    all_member_ids = {
        value
        for study in studies
        for value in study["memberCandidateAnchorIds"]
    }
    focused_member_ids = {
        str(value)
        for value in part.get("candidateAnchorIds", [])
    }
    if not focused_member_ids:
        focused_member_ids = {str(value) for value in all_member_ids}
    if not focused_member_ids.issubset(
        {str(value) for value in all_member_ids}
    ):
        raise StagedDraftV2Error(
            "Part candidate anchors exceed its logical Study registry"
        )
    scoped_registry = bool(part.get("sourceSegments"))
    registry_anchor_keys = {
        str(value)
        for value in part.get("registryAnchorCellKeys", [])
    }
    if scoped_registry:
        scoped_studies: list[dict[str, Any]] = []
        for study in studies:
            full_anchor_keys = [
                str(value)
                for value in study.get("anchorEvidenceCellKeys", [])
            ]
            full_link_keys = [
                str(value)
                for value in study.get(
                    "registryLinkEvidenceCellKeys",
                    [],
                )
            ]
            study["focusedCandidateAnchorIds"] = [
                str(value)
                for value in study.get(
                    "memberCandidateAnchorIds",
                    [],
                )
                if str(value) in focused_member_ids
            ]
            study["fullAnchorEvidenceSha256"] = json_sha256(
                full_anchor_keys
            )
            study["fullRegistryLinkEvidenceSha256"] = json_sha256(
                full_link_keys
            )
            study["anchorEvidenceCellKeys"] = [
                key
                for key in full_anchor_keys
                if key in registry_anchor_keys
            ]
            study["registryLinkEvidenceCellKeys"] = [
                key
                for key in full_link_keys
                if key in registry_anchor_keys
            ]
            scoped_studies.append(study)
        studies = scoped_studies
    proposal_ids = {
        value
        for study in studies
        for value in study.get("registryLinkProposalIds", [])
    }
    result = {
        "schemaVersion": STUDY_REGISTRY_V2_SCHEMA_VERSION,
        "source": copy.deepcopy(registry["source"]),
        "registrySha256": registry["registrySha256"],
        "candidateAnchors": [
            copy.deepcopy(anchor)
            for anchor in registry["candidateAnchors"]
            if str(anchor["candidateAnchorId"]) in focused_member_ids
        ],
        "linkProposals": [
            copy.deepcopy(proposal)
            for proposal in registry.get("linkProposals", [])
            if proposal.get("proposalId") in proposal_ids
            and (
                not scoped_registry
                or focused_member_ids.intersection(
                    str(value)
                    for value in proposal.get(
                        "memberCandidateAnchorIds",
                        [],
                    )
                )
            )
        ],
        "studies": studies,
    }
    if scoped_registry:
        result["scope"] = {
            "mode": "SOURCE_SEGMENT",
            "fullRegistrySha256": registry["registrySha256"],
            "logicalStudyIds": list(part["logicalStudyIds"]),
            "candidateAnchorIds": list(
                part.get("candidateAnchorIds", [])
            ),
            "sectionAnchorIds": list(
                part.get("sectionAnchorIds", [])
            ),
        }
    return result


def build_fragment_envelope(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    plan: dict[str, Any],
    part: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
    locator_results: Sequence[dict[str, Any]],
    registry_slice: dict[str, Any],
) -> dict[str, Any]:
    _fragment_identity, fragment_identity_sha256 = (
        _require_current_fragment_identity(plan, part)
    )
    envelope = {
        "schemaVersion": "study-draft-fragment-input-v2",
        "promptVersion": FRAGMENT_PROMPT_VERSION,
        "fragmentIdentitySha256": fragment_identity_sha256,
        "source": copy.deepcopy(source),
        "workbook": copy.deepcopy(workbook),
        "planId": plan["planId"],
        "partId": part["partId"],
        "ownedSourceCellKeys": list(part["ownedSourceCellKeys"]),
        "sharedAnchorCellKeys": list(part["sharedAnchorCellKeys"]),
        "registry": copy.deepcopy(registry_slice),
        "locatorResults": copy.deepcopy(list(locator_results)),
        "focusedChunks": copy.deepcopy(list(focused_chunks)),
        "contracts": {
            "fragment": FRAGMENT_CONTRACT_VERSION,
            "validator": FRAGMENT_VALIDATOR_VERSION,
            "consolidator": CONSOLIDATOR_CONTRACT_VERSION,
        },
        "imagesAnalyzed": False,
    }
    if part.get("sourceSegments"):
        envelope["sourceSegments"] = copy.deepcopy(
            part["sourceSegments"]
        )
    envelope["inputEnvelopeSha256"] = json_sha256(envelope)
    return envelope


def build_fragment_prompt(envelope: dict[str, Any]) -> str:
    return (
        "Return only one JSON object satisfying study-draft-fragment-v2. "
        "This is an append-only source fragment, never a canonical manifest. "
        "Use only logicalStudyId values in registry. Give each record a unique "
        "nonempty placeholder recordId; the transport replaces it with the "
        "source-stable ID. Every record needs exact source evidence. In the "
        "strict output transport, encode each record's complete payload object "
        "as compact JSON in payloadJson; do not emit a nested payload field. "
        "The runner decodes payloadJson losslessly before validation. "
        "Evidence may use only ownedSourceCellKeys "
        "or sharedAnchorCellKeys. Numeric observations/series must cite at "
        "least one owned value cell; shared-only numeric claims are forbidden. "
        "Every owned numeric result cell must be bound to a canonical "
        "OBSERVATION_APPEND value or to a SERIES_SEGMENT_APPEND valueRange/"
        "rowIdentityRange. Raw backing data is semantic data, not a reason for "
        "NO_SEMANTIC_RECORD. Never encode a series as points, xAxis/yAxis, or "
        "another nested free-form curve. Never encode multiple metrics inside "
        "one observation payload. "
        "Use these exact payloadJson contracts: "
        "STUDY_PATCH={title,purpose,hypothesis,objective,designType,"
        "comparisonBasis,summary}; "
        "ENTITY_DECLARATION ARM={entityType:'ARM',key,role,label,condition,"
        "sampleSize,sampleBasis,matchingBasis,factorValues}; "
        "ENTITY_DECLARATION OUTCOME={entityType:'OUTCOME',key,originalLabel,"
        "metricType,unit,favorableDirection}; "
        "ENTITY_DECLARATION FACTOR={entityType:'FACTOR',key,originalLabel,"
        "baselineCondition,changedCondition,changeDirection,isolationStatus}; "
        "ENTITY_DECLARATION CONTEXT={entityType:'CONTEXT',key,kind,"
        "originalValue,normalizedValue}; "
        "OBSERVATION_APPEND={outcome,arm,valueNumber,valueText,numerator,"
        "denominator,ratePpm,min,max,average,sampleSize,replicateKey}; "
        "SERIES_SEGMENT_APPEND={key,seriesRole:'RAW'|'AGGREGATE',"
        "aggregationFunction,aggregateOfSeries,outcome,arm,sheet,"
        "headerRange,valueRange,"
        "rowIdentityRange,aggregateReplicateRanges:[],axisSource:"
        "'HEADER'|'ROW_IDENTITY',axisLabel,axisUnit,valueUnit,stratumKey,"
        "verificationStatus:'NEEDS_REVIEW'}. "
        "Every outcome and arm referenced by an observation or series must be "
        "declared with ENTITY_DECLARATION from an owned or shared source label. "
        "A source-stable entity identity may be emitted only once. If one "
        "source label appears to support count and rate outcomes, either use "
        "one supported outcome or cite distinct source identityCellKeys and "
        "evidence that directly establish each distinction; never duplicate "
        "an OUTCOME by changing only a model-invented key or metricType. "
        "Series ranges are exact A1 ranges on payload.sheet: headerRange is one "
        "row aligned to valueRange columns, and rowIdentityRange is one column "
        "aligned to valueRange rows. "
        "Emit exactly one coverageDisposition for every owned source cell. "
        "Do not create a comparison; emit COMPARISON_LINK_INTENT only when "
        "the source explicitly supports the relationship. All results remain "
        "NEEDS_REVIEW and images are out of scope.\n\nINPUT_ENVELOPE:\n"
        + json.dumps(
            envelope,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
    )


def finalize_fragment_envelope(
    envelope: dict[str, Any],
) -> dict[str, Any]:
    result = copy.deepcopy(envelope)
    prompt = build_fragment_prompt(result)
    result["inputHashes"] = {
        "packetSha256": json_sha256(result["focusedChunks"]),
        "locatorSha256": json_sha256(result["locatorResults"]),
        "focusedSha256": json_sha256(
            {
                "chunks": result["focusedChunks"],
                "ownedSourceCellKeys": result["ownedSourceCellKeys"],
            }
        ),
        "registrySha256": str(
            result["registry"]["registrySha256"]
        ),
        "sharedAllowlistSha256": json_sha256(
            result["sharedAnchorCellKeys"]
        ),
        "inputEnvelopeSha256": result["inputEnvelopeSha256"],
        "promptSha256": bytes_sha256(prompt.encode("utf-8")),
        "fragmentContractVersion": FRAGMENT_CONTRACT_VERSION,
        "validatorContractVersion": FRAGMENT_VALIDATOR_VERSION,
        "consolidatorContractVersion": CONSOLIDATOR_CONTRACT_VERSION,
    }
    result["promptText"] = prompt
    result["promptBytes"] = len(prompt.encode("utf-8"))
    return result


def stable_record_id(
    *,
    revision_uid: str,
    logical_study_id: str,
    record_type: str,
    identity_cell_keys: Sequence[str],
    exact_source_label: str,
    semantic_subtype: str = "",
) -> str:
    return stable_uid(
        "fragment-record-v2",
        revision_uid,
        logical_study_id,
        record_type,
        " ".join(str(semantic_subtype).split()).upper(),
        *sorted(str(value) for value in identity_cell_keys),
        " ".join(exact_source_label.split()).casefold(),
    )


def _record_semantic_subtype(record: dict[str, Any]) -> str:
    payload = record.get("payload")
    if not isinstance(payload, dict):
        return ""
    record_type = str(record.get("recordType") or "").upper()
    if record_type == "ENTITY_DECLARATION":
        return str(payload.get("entityType") or "").upper()
    if record_type in {"OBSERVATION_APPEND", "SERIES_SEGMENT_APPEND"}:
        return "|".join(
            (
                str(payload.get("outcome") or ""),
                str(payload.get("arm") or ""),
            )
        )
    if record_type == "COMPARISON_LINK_INTENT":
        outcomes = payload.get("outcomes")
        return "|".join(
            (
                str(payload.get("comparedArm") or ""),
                str(payload.get("controlArm") or ""),
                ",".join(
                    sorted(str(value) for value in outcomes)
                    if isinstance(outcomes, list)
                    else []
                ),
            )
        )
    if record_type == "LIMITATION_APPEND":
        return str(payload.get("scope") or "STUDY").upper()
    if record_type == "CONCLUSION_APPEND":
        return str(payload.get("conclusionType") or "").upper()
    return ""


def normalize_fragment_record_ids(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
) -> dict[str, Any]:
    """Replace model-proposed IDs with deterministic source-derived IDs."""

    result = copy.deepcopy(fragment)
    old_to_new: dict[str, str] = {}
    for record in result.get("records", []):
        if not isinstance(record, dict):
            continue
        old_id = str(record.get("recordId") or "")
        identity_keys = record.get("identityCellKeys")
        if not isinstance(identity_keys, list):
            continue
        new_id = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=str(record.get("logicalStudyId") or ""),
            record_type=str(record.get("recordType") or "").upper(),
            identity_cell_keys=[
                str(value) for value in identity_keys
            ],
            exact_source_label=str(
                record.get("exactSourceLabel") or ""
            ),
            semantic_subtype=_record_semantic_subtype(record),
        )
        record["recordId"] = new_id
        if str(record.get("recordType") or "").upper() == (
            "ENTITY_DECLARATION"
        ) and isinstance(record.get("payload"), dict):
            record["payload"]["entityId"] = new_id
        if old_id:
            old_to_new[old_id] = new_id
    for disposition in result.get("coverageDispositions", []):
        if not isinstance(disposition, dict):
            continue
        values = disposition.get("recordIds")
        if isinstance(values, list):
            disposition["recordIds"] = [
                old_to_new.get(str(value), str(value))
                for value in values
            ]
    return result


def normalize_fragment_evidence_dispositions(
    *,
    fragment: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Make disposition record IDs exactly follow explicit record evidence.

    Models occasionally preserve a source cell as CONTEXT_ONLY while also
    copying recordIds into that disposition.  The semantic disposition remains
    model-owned, but record linkage is deterministic: a source cell is
    RECORD_EVIDENCE if and only if an emitted record explicitly cites it.
    """

    result = copy.deepcopy(fragment)
    records = result.get("records")
    dispositions = result.get("coverageDispositions")
    if not isinstance(records, list) or not isinstance(dispositions, list):
        return result

    _by_coordinate, by_key, _source_order = _source_cell_maps(
        all_selected_chunks
    )
    records_by_id = {
        str(record.get("recordId") or ""): record
        for record in records
        if isinstance(record, dict)
        and str(record.get("recordId") or "")
    }
    for disposition in dispositions:
        if (
            not isinstance(disposition, dict)
            or str(disposition.get("disposition") or "").upper()
            != "RECORD_EVIDENCE"
        ):
            continue
        key = str(disposition.get("sourceCellKey") or "")
        source_entry = by_key.get(key)
        if source_entry is None:
            continue
        sheet, coordinate, cell = source_entry
        source_text = str(
            cell.get("displayValue")
            if cell.get("displayValue") is not None
            else cell.get("rawValue")
            or ""
        ).strip()
        for record_id in disposition.get("recordIds", []):
            record = records_by_id.get(str(record_id))
            if record is None:
                continue
            evidence = record.get("evidence")
            if not isinstance(evidence, list):
                continue
            if key in evidence_cell_keys(
                evidence,
                chunks=all_selected_chunks,
            ):
                continue
            evidence.append(
                {
                    "sheet": sheet,
                    "range": coordinate,
                    "role": "DECLARED_SOURCE",
                    "sourceText": source_text,
                    "note": str(disposition.get("reason") or "").strip(),
                }
            )

    evidence_ids_by_key: dict[str, list[str]] = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        record_id = str(record.get("recordId") or "")
        evidence = record.get("evidence")
        if not record_id or not isinstance(evidence, list):
            continue
        for key in evidence_cell_keys(
            evidence,
            chunks=all_selected_chunks,
        ):
            ids = evidence_ids_by_key.setdefault(key, [])
            if record_id not in ids:
                ids.append(record_id)

    for disposition in dispositions:
        if not isinstance(disposition, dict):
            continue
        key = str(disposition.get("sourceCellKey") or "")
        expected_ids = evidence_ids_by_key.get(key, [])
        if expected_ids:
            disposition["disposition"] = "RECORD_EVIDENCE"
            disposition["recordIds"] = expected_ids
            disposition["reason"] = ""
        elif (
            str(disposition.get("disposition") or "").upper()
            != "RECORD_EVIDENCE"
        ):
            disposition["recordIds"] = []
    return result


def normalize_fragment_complete_dispositions(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Add explicit fail-closed dispositions for every owned source cell."""

    result = copy.deepcopy(fragment)
    records = result.get("records")
    dispositions = result.get("coverageDispositions")
    if not isinstance(records, list) or not isinstance(dispositions, list):
        return result
    evidence_ids_by_key: dict[str, list[str]] = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        record_id = str(record.get("recordId") or "")
        if not record_id:
            continue
        for key in evidence_cell_keys(
            record.get("evidence", []),
            chunks=all_selected_chunks,
        ):
            ids = evidence_ids_by_key.setdefault(key, [])
            if record_id not in ids:
                ids.append(record_id)
    existing = {
        str(item.get("sourceCellKey") or ""): item
        for item in dispositions
        if isinstance(item, dict)
        and str(item.get("sourceCellKey") or "")
    }
    ordered: list[dict[str, Any]] = []
    for key_value in envelope.get("ownedSourceCellKeys", []):
        key = str(key_value)
        disposition = existing.get(key)
        if disposition is None:
            record_ids = evidence_ids_by_key.get(key, [])
            disposition = {
                "sourceCellKey": key,
                "disposition": (
                    "RECORD_EVIDENCE"
                    if record_ids
                    else "CONTEXT_ONLY"
                ),
                "recordIds": record_ids,
                "reason": (
                    ""
                    if record_ids
                    else "No emitted record cites this owned source cell."
                ),
            }
        ordered.append(disposition)
    result["coverageDispositions"] = ordered
    return result


def normalize_fragment_required_fields_and_series_headers(
    *,
    fragment: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Fill source-preserving entity fields and repair exact series headers."""

    result = copy.deepcopy(fragment)
    records = result.get("records")
    if not isinstance(records, list):
        return result
    _by_coordinate, by_key, _source_order = _source_cell_maps(
        all_selected_chunks
    )
    series_payloads_by_key = {
        str(record["payload"].get("key") or "").strip().casefold(): record[
            "payload"
        ]
        for record in records
        if isinstance(record, dict)
        and str(record.get("recordType") or "").upper()
        == "SERIES_SEGMENT_APPEND"
        and isinstance(record.get("payload"), dict)
        and str(record["payload"].get("key") or "").strip()
    }
    for record in records:
        if not isinstance(record, dict):
            continue
        record_type = str(record.get("recordType") or "").upper()
        payload = record.get("payload")
        if not isinstance(payload, dict):
            continue
        exact_label = str(
            record.get("exactSourceLabel") or ""
        ).strip()
        if record_type == "ENTITY_DECLARATION" and exact_label:
            entity_type = str(payload.get("entityType") or "").upper()
            if entity_type == "OUTCOME":
                if not str(payload.get("originalLabel") or "").strip():
                    payload["originalLabel"] = exact_label
                if not str(payload.get("metricType") or "").strip():
                    payload["metricType"] = "source_labeled_result"
            elif entity_type == "FACTOR":
                if not str(payload.get("originalLabel") or "").strip():
                    payload["originalLabel"] = exact_label
            elif entity_type == "ARM":
                if not str(payload.get("label") or "").strip():
                    payload["label"] = exact_label
                if not str(payload.get("role") or "").strip():
                    payload["role"] = "OTHER"
            elif entity_type == "CONTEXT":
                if not str(payload.get("kind") or "").strip():
                    payload["kind"] = "source context"
                if not str(payload.get("originalValue") or "").strip():
                    payload["originalValue"] = exact_label
        if record_type != "SERIES_SEGMENT_APPEND" or not exact_label:
            continue
        sheet = str(payload.get("sheet") or "").strip()
        header_range = str(payload.get("headerRange") or "").strip()
        if sheet and header_range and evidence_cell_keys(
            [{"sheet": sheet, "range": header_range}],
            chunks=all_selected_chunks,
        ):
            continue
        if sheet and header_range:
            (
                header_start_row,
                header_start_column,
                header_end_row,
                header_end_column,
            ) = range_bounds(header_range)
            merged_header_candidates: list[
                tuple[int, int, str]
            ] = []
            for candidate_sheet, _coordinate, cell in by_key.values():
                merge_range = str(cell.get("mergeRange") or "").strip()
                if (
                    candidate_sheet.casefold() != sheet.casefold()
                    or not merge_range
                ):
                    continue
                (
                    merge_start_row,
                    merge_start_column,
                    merge_end_row,
                    merge_end_column,
                ) = range_bounds(merge_range)
                if (
                    merge_start_row <= header_start_row <= header_end_row
                    <= merge_end_row
                    and merge_start_column
                    <= header_start_column
                    <= header_end_column
                    <= merge_end_column
                ):
                    merged_header_candidates.append(
                        (
                            (
                                merge_end_row
                                - merge_start_row
                                + 1
                            )
                            * (
                                merge_end_column
                                - merge_start_column
                                + 1
                            ),
                            merge_start_row * 100_000
                            + merge_start_column,
                            merge_range,
                        )
                    )
            if merged_header_candidates:
                payload["headerRange"] = min(
                    merged_header_candidates
                )[2]
                continue
        normalized_label = " ".join(exact_label.split()).casefold()
        candidates: list[str] = []
        for item in record.get("evidence", []):
            if not isinstance(item, dict):
                continue
            candidate_sheet = str(item.get("sheet") or "").strip()
            candidate_range = str(item.get("range") or "").strip()
            candidate_text = " ".join(
                str(item.get("sourceText") or "").split()
            ).casefold()
            if (
                candidate_sheet
                and candidate_range
                and candidate_text == normalized_label
                and evidence_cell_keys(
                    [
                        {
                            "sheet": candidate_sheet,
                            "range": candidate_range,
                        }
                    ],
                    chunks=all_selected_chunks,
                )
            ):
                if not sheet:
                    payload["sheet"] = candidate_sheet
                    sheet = candidate_sheet
                if candidate_sheet.casefold() == sheet.casefold():
                    candidates.append(candidate_range)
        if candidates:
            payload["headerRange"] = candidates[0]
            continue
        aggregate_of = payload.get("aggregateOfSeries")
        aggregate_keys = (
            [aggregate_of]
            if isinstance(aggregate_of, str)
            else aggregate_of
            if isinstance(aggregate_of, list)
            else []
        )
        for aggregate_key in aggregate_keys:
            aggregate_payload = series_payloads_by_key.get(
                str(aggregate_key or "").strip().casefold()
            )
            if not isinstance(aggregate_payload, dict):
                continue
            aggregate_sheet = str(
                aggregate_payload.get("sheet") or ""
            ).strip()
            aggregate_header = str(
                aggregate_payload.get("headerRange") or ""
            ).strip()
            if (
                aggregate_sheet.casefold() == sheet.casefold()
                and aggregate_header
                and evidence_cell_keys(
                    [
                        {
                            "sheet": aggregate_sheet,
                            "range": aggregate_header,
                        }
                    ],
                    chunks=all_selected_chunks,
                )
            ):
                payload["headerRange"] = aggregate_header
                break
        if (
            str(payload.get("headerRange") or "").strip()
            != header_range
        ):
            continue
        value_range = str(payload.get("valueRange") or "").strip()
        if not sheet or not value_range:
            continue
        (
            value_start_row,
            value_start_column,
            _end_row,
            value_end_column,
        ) = range_bounds(value_range)
        cited_keys = evidence_cell_keys(
            record.get("evidence", []),
            chunks=all_selected_chunks,
        )
        header_candidates: list[tuple[int, str]] = []
        for key in cited_keys:
            source_entry = by_key.get(key)
            if source_entry is None:
                continue
            candidate_sheet, coordinate, cell = source_entry
            row, column, _candidate_end_row, _candidate_end_column = (
                range_bounds(coordinate)
            )
            source_text = str(
                cell.get("displayValue")
                if cell.get("displayValue") is not None
                else cell.get("rawValue")
                or ""
            ).strip()
            if (
                candidate_sheet.casefold() == sheet.casefold()
                and column == value_start_column
                and row < value_start_row
                and source_text
                and _numeric_value(cell) is None
                and value_start_column == value_end_column
            ):
                header_candidates.append((row, coordinate))
        if header_candidates:
            payload["headerRange"] = max(header_candidates)[1]
            continue
        context_rows: dict[int, set[int]] = {}
        for candidate_sheet, _coordinate, cell in by_key.values():
            if (
                candidate_sheet.casefold() != sheet.casefold()
                or not bool(cell.get("contextOnly"))
            ):
                continue
            row, column = _position(cell)
            if (
                row < value_start_row
                and value_start_column
                <= column
                <= value_end_column
            ):
                context_rows.setdefault(row, set()).add(column)
        complete_context_rows = [
            row
            for row, columns in context_rows.items()
            if all(
                column in columns
                for column in range(
                    value_start_column,
                    value_end_column + 1,
                )
            )
        ]
        if complete_context_rows:
            context_row = max(complete_context_rows)
            payload["headerRange"] = _address(
                (
                    context_row,
                    value_start_column,
                    context_row,
                    value_end_column,
                )
            )
    return result


def normalize_fragment_multi_arm_series_rows(
    *,
    fragment: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Split one-column row series when each row names an exact source Arm."""

    result = copy.deepcopy(fragment)
    records = result.get("records")
    if not isinstance(records, list):
        return result
    by_coordinate, _by_key, _source_order = _source_cell_maps(
        all_selected_chunks
    )
    arm_keys: dict[tuple[str, str], str] = {}
    for record in records:
        if (
            not isinstance(record, dict)
            or str(record.get("recordType") or "").upper()
            != "ENTITY_DECLARATION"
            or not isinstance(record.get("payload"), dict)
            or str(record["payload"].get("entityType") or "").upper()
            != "ARM"
        ):
            continue
        logical_id = str(record.get("logicalStudyId") or "")
        arm_key = str(record["payload"].get("key") or "").strip()
        labels = {
            str(record.get("exactSourceLabel") or "").strip(),
            str(record["payload"].get("label") or "").strip(),
        }
        for label in labels:
            normalized_label = " ".join(label.split()).casefold()
            if logical_id and arm_key and normalized_label:
                arm_keys[(logical_id, normalized_label)] = arm_key

    normalized_records: list[dict[str, Any]] = []
    for record in records:
        if not isinstance(record, dict):
            normalized_records.append(record)
            continue
        payload = record.get("payload")
        if (
            str(record.get("recordType") or "").upper()
            != "SERIES_SEGMENT_APPEND"
            or not isinstance(payload, dict)
            or str(payload.get("arm") or "").strip()
            or str(payload.get("seriesRole") or "RAW").upper()
            != "RAW"
            or str(payload.get("aggregateOfSeries") or "").strip()
        ):
            normalized_records.append(record)
            continue
        sheet = str(payload.get("sheet") or "").strip()
        value_range = str(payload.get("valueRange") or "").strip()
        identity_range = str(
            payload.get("rowIdentityRange") or ""
        ).strip()
        if not sheet or not value_range or not identity_range:
            normalized_records.append(record)
            continue
        (
            value_start_row,
            value_start_column,
            value_end_row,
            value_end_column,
        ) = range_bounds(value_range)
        (
            identity_start_row,
            identity_start_column,
            identity_end_row,
            identity_end_column,
        ) = range_bounds(identity_range)
        if (
            value_start_column != value_end_column
            or identity_start_column != identity_end_column
            or value_start_row != identity_start_row
            or value_end_row != identity_end_row
            or value_end_row <= value_start_row
        ):
            normalized_records.append(record)
            continue
        row_bindings: list[
            tuple[int, str, str, str, str]
        ] = []
        logical_id = str(record.get("logicalStudyId") or "")
        for row in range(value_start_row, value_end_row + 1):
            value_address = _address(
                (row, value_start_column, row, value_start_column)
            )
            identity_address = _address(
                (
                    row,
                    identity_start_column,
                    row,
                    identity_start_column,
                )
            )
            value_entry = by_coordinate.get(
                (sheet.casefold(), value_address)
            )
            identity_entry = by_coordinate.get(
                (sheet.casefold(), identity_address)
            )
            if value_entry is None or identity_entry is None:
                row_bindings = []
                break
            identity_text = str(
                identity_entry[1].get("displayValue")
                if identity_entry[1].get("displayValue") is not None
                else identity_entry[1].get("rawValue")
                or ""
            ).strip()
            arm_key = arm_keys.get(
                (
                    logical_id,
                    " ".join(identity_text.split()).casefold(),
                )
            )
            if not identity_text or not arm_key:
                row_bindings = []
                break
            row_bindings.append(
                (
                    row,
                    arm_key,
                    identity_text,
                    value_entry[0],
                    identity_entry[0],
                )
            )
        if not row_bindings:
            normalized_records.append(record)
            continue
        base_key = str(payload.get("key") or "series").strip()
        for row, arm_key, identity_text, value_key, identity_key in (
            row_bindings
        ):
            value_address = _address(
                (row, value_start_column, row, value_start_column)
            )
            identity_address = _address(
                (
                    row,
                    identity_start_column,
                    row,
                    identity_start_column,
                )
            )
            split_record = copy.deepcopy(record)
            split_record["identityCellKeys"] = [
                identity_key,
                value_key,
            ]
            split_record["exactSourceLabel"] = identity_text
            split_payload = split_record["payload"]
            split_payload["key"] = (
                f"{base_key}_{arm_key}_{row}"
            )
            split_payload["arm"] = arm_key
            split_payload["valueRange"] = value_address
            split_payload["rowIdentityRange"] = identity_address
            split_payload["aggregateReplicateRanges"] = []
            value_cell = by_coordinate[
                (sheet.casefold(), value_address)
            ][1]
            value_text = str(
                value_cell.get("displayValue")
                if value_cell.get("displayValue") is not None
                else value_cell.get("rawValue")
                or ""
            )
            split_record["evidence"] = [
                {
                    "sheet": sheet,
                    "range": identity_address,
                    "role": "ROW_IDENTITY",
                    "sourceText": identity_text,
                    "note": "",
                },
                {
                    "sheet": sheet,
                    "range": value_address,
                    "role": "SERIES_VALUE",
                    "sourceText": value_text,
                    "note": "",
                },
            ]
            normalized_records.append(split_record)
    result["records"] = normalized_records
    return result


def _numeric_value(cell: dict[str, Any]) -> float | None:
    formula = str(cell.get("formula") or "").strip()
    value = (
        cell.get("cachedValue") if formula else cell.get("rawValue")
    )
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        return numeric if math.isfinite(numeric) else None
    return None


_DIRECT_RATE_FORMULA = re.compile(
    r"^\s*=\s*\+?\s*"
    r"(\$?[A-Z]{1,3}\$?\d+)\s*/\s*"
    r"(\$?[A-Z]{1,3}\$?\d+)"
    r"(?:\s*\*\s*(100|1000000))?\s*$",
    re.IGNORECASE,
)
_CATEGORICAL_RESULT_STATUS = re.compile(
    r"^\s*(?:pass(?:ed)?|fail(?:ed)?|ok|n/?g)\s*[.!]?\s*$",
    re.IGNORECASE,
)
_RESULT_ROW_ARM_KEY = re.compile(r"^result_row_(\d+)$")


def _source_cell_text(cell: dict[str, Any]) -> str:
    for field in ("displayValue", "cachedValue", "rawValue"):
        value = cell.get(field)
        if value is not None:
            return str(value)
    return ""


def _complete_incomplete_formula_rate_pair(
    *,
    payload: dict[str, Any],
    evidence: Sequence[dict[str, Any]],
    selected_chunks: Sequence[dict[str, Any]],
    by_coordinate: dict[
        tuple[str, str],
        tuple[str, dict[str, Any]],
    ],
    by_key: dict[str, tuple[str, str, dict[str, Any]]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Recover only a directly formula-bound absent count pair or member."""

    result = copy.deepcopy(payload)
    result_evidence = copy.deepcopy(list(evidence))
    numerator_supplied = result.get("numerator") not in (None, "")
    denominator_supplied = result.get("denominator") not in (None, "")
    if numerator_supplied and denominator_supplied:
        return result, result_evidence

    candidates: list[
        tuple[
            float,
            float,
            str,
            str,
            dict[str, Any],
            str,
            dict[str, Any],
        ]
    ] = []
    for key in evidence_cell_keys(
        result_evidence,
        chunks=selected_chunks,
    ):
        source_entry = by_key.get(key)
        if source_entry is None:
            continue
        sheet, _coordinate_value, formula_cell = source_entry
        match = _DIRECT_RATE_FORMULA.fullmatch(
            str(formula_cell.get("formula") or "")
        )
        if match is None:
            continue
        numerator_coordinate = match.group(1).replace("$", "").upper()
        denominator_coordinate = match.group(2).replace("$", "").upper()
        numerator_entry = by_coordinate.get(
            (sheet.casefold(), numerator_coordinate)
        )
        denominator_entry = by_coordinate.get(
            (sheet.casefold(), denominator_coordinate)
        )
        if numerator_entry is None or denominator_entry is None:
            continue
        numerator = _numeric_value(numerator_entry[1])
        denominator = _numeric_value(denominator_entry[1])
        formula_value = _numeric_value(formula_cell)
        if (
            numerator is None
            or denominator is None
            or formula_value is None
            or numerator < 0
            or denominator <= 0
            or numerator > denominator
        ):
            continue
        factor = float(match.group(3) or 1)
        if not math.isclose(
            formula_value,
            numerator / denominator * factor,
            rel_tol=1e-9,
            abs_tol=1e-9,
        ):
            continue
        if numerator_supplied and not math.isclose(
            float(result["numerator"]),
            numerator,
            rel_tol=1e-9,
            abs_tol=1e-9,
        ):
            continue
        if denominator_supplied and not math.isclose(
            float(result["denominator"]),
            denominator,
            rel_tol=1e-9,
            abs_tol=1e-9,
        ):
            continue
        rate_ppm = result.get("ratePpm")
        if rate_ppm not in (None, "") and not math.isclose(
            float(rate_ppm),
            numerator / denominator * 1_000_000.0,
            rel_tol=1e-9,
            abs_tol=1e-9,
        ):
            continue
        candidates.append(
            (
                numerator,
                denominator,
                sheet,
                numerator_coordinate,
                numerator_entry[1],
                denominator_coordinate,
                denominator_entry[1],
            )
        )
    unique_candidates = {
        (
            numerator,
            denominator,
            sheet.casefold(),
            numerator_coordinate,
            denominator_coordinate,
        ): candidate
        for candidate in candidates
        for (
            numerator,
            denominator,
            sheet,
            numerator_coordinate,
            _numerator_cell,
            denominator_coordinate,
            _denominator_cell,
        ) in [candidate]
    }
    if len(unique_candidates) != 1:
        if numerator_supplied != denominator_supplied:
            result["numerator"] = None
            result["denominator"] = None
        return result, result_evidence

    (
        numerator,
        denominator,
        sheet,
        numerator_coordinate,
        numerator_cell,
        denominator_coordinate,
        denominator_cell,
    ) = next(iter(unique_candidates.values()))
    result["numerator"] = numerator
    result["denominator"] = denominator
    if result.get("sampleSize") in (None, ""):
        result["sampleSize"] = denominator
    existing_coordinates = {
        (
            str(item.get("sheet") or "").casefold(),
            str(item.get("range") or "").replace("$", "").upper(),
        )
        for item in result_evidence
        if isinstance(item, dict)
    }
    for coordinate, cell, role in (
        (
            numerator_coordinate,
            numerator_cell,
            "DIRECT_FORMULA_NUMERATOR",
        ),
        (
            denominator_coordinate,
            denominator_cell,
            "DIRECT_FORMULA_DENOMINATOR",
        ),
    ):
        if (sheet.casefold(), coordinate) in existing_coordinates:
            continue
        result_evidence.append(
            {
                "sheet": sheet,
                "range": coordinate,
                "role": role,
                "sourceText": _source_cell_text(cell),
                "note": (
                    "Deterministically resolved from the directly cited "
                    "same-sheet rate formula."
                ),
            }
        )
    return result, result_evidence


def _normalize_projected_percent_observation(
    *,
    payload: dict[str, Any],
    evidence: Sequence[dict[str, Any]],
    outcome_unit: object,
    outcome_label: object,
    selected_chunks: Sequence[dict[str, Any]],
    by_key: dict[str, tuple[str, str, dict[str, Any]]],
) -> dict[str, Any]:
    """Put exact percent-formatted scalar claims on the human percent scale."""

    if str(outcome_unit or "").strip().casefold() not in {
        "%",
        "percent",
        "percentage",
        "pct",
    }:
        return payload
    evidence_keys = evidence_cell_keys(
        evidence,
        chunks=selected_chunks,
    )
    cited_cells = [
        by_key[key][2]
        for key in evidence_keys
        if key in by_key
    ]
    percent_cells = [
        cell
        for cell in cited_cells
        if "%" in str(cell.get("numberFormat") or "")
    ]
    raw_percent_values = [
        number
        for cell in percent_cells
        for number in [_numeric_value(cell)]
        if number is not None
    ]
    result = copy.deepcopy(payload)
    if not raw_percent_values:
        normalized_label = " ".join(
            str(outcome_label or "").split()
        )
        aliases = [normalized_label] if normalized_label else []
        source_style_label = re.sub(
            r"\s+(?:percentage|percent|pct|rate)\s*$",
            "",
            normalized_label,
            flags=re.IGNORECASE,
        ).strip()
        if source_style_label and source_style_label not in aliases:
            aliases.append(source_style_label)
        labeled_values: set[float] = set()
        if aliases:
            label_pattern = "(?:" + "|".join(
                r"\s+".join(
                    re.escape(part) for part in alias.split()
                )
                for alias in aliases
            ) + ")"
            labeled_percent = re.compile(
                rf"(?<![A-Za-z0-9_]){label_pattern}"
                r"(?![A-Za-z0-9_])\s*[:=]?\s*"
                r"(?P<number>[+-]?(?:\d+(?:,\d{3})*|\d*)"
                r"(?:\.\d+)?(?:[Ee][+-]?\d+)?)\s*%",
                re.IGNORECASE,
            )
            for cell in cited_cells:
                text = _source_cell_text(cell)
                for match in labeled_percent.finditer(text):
                    try:
                        number = float(
                            match.group("number").replace(",", "")
                        )
                    except ValueError:
                        continue
                    if math.isfinite(number):
                        labeled_values.add(number)
        if len(labeled_values) != 1:
            return result
        labeled_value = next(iter(labeled_values))
        current_value = result.get("valueNumber")
        if current_value in (None, ""):
            result["valueNumber"] = labeled_value
        elif (
            isinstance(current_value, (int, float))
            and not isinstance(current_value, bool)
            and math.isfinite(float(current_value))
            and not math.isclose(
                float(current_value),
                labeled_value,
                rel_tol=1e-6,
                abs_tol=1e-9,
            )
        ):
            return result
        if not str(result.get("valueText") or "").strip():
            result["valueText"] = f"{labeled_value:g}%"
        rate_ppm = result.get("ratePpm")
        if (
            isinstance(rate_ppm, (int, float))
            and not isinstance(rate_ppm, bool)
            and result.get("numerator") in (None, "")
            and result.get("denominator") in (None, "")
            and math.isclose(
                float(rate_ppm),
                labeled_value * 10_000.0,
                rel_tol=1e-6,
                abs_tol=1e-9,
            )
        ):
            result["ratePpm"] = None
        return result
    unique_raw_percent_values = {
        float(value) for value in raw_percent_values
    }
    if (
        result.get("valueNumber") in (None, "")
        and len(unique_raw_percent_values) == 1
    ):
        result["valueNumber"] = (
            next(iter(unique_raw_percent_values)) * 100.0
        )
    percent_display_values = {
        str(cell.get("displayValue") or "").strip()
        for cell in percent_cells
        if str(cell.get("displayValue") or "").strip()
    }
    if (
        not str(result.get("valueText") or "").strip()
        and len(percent_display_values) == 1
    ):
        result["valueText"] = next(iter(percent_display_values))
    for field in ("valueNumber", "min", "max", "average"):
        value = result.get(field)
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(float(value))
        ):
            continue
        raw_matches = any(
            math.isclose(
                float(value),
                raw,
                rel_tol=1e-6,
                abs_tol=1e-9,
            )
            for raw in raw_percent_values
        )
        human_matches = any(
            math.isclose(
                float(value),
                raw * 100.0,
                rel_tol=1e-6,
                abs_tol=1e-9,
            )
            for raw in raw_percent_values
        )
        if raw_matches and not human_matches:
            result[field] = float(value) * 100.0
    return result


def _augment_projected_categorical_status_observations(
    *,
    studies: Sequence[dict[str, Any]],
    selected_chunks: Sequence[dict[str, Any]],
    revision_uid: str,
) -> list[dict[str, Any]]:
    """Bind exact table status cells to existing source-row arms."""

    result = copy.deepcopy(list(studies))
    cells_by_coordinate: dict[
        tuple[str, str],
        tuple[str, dict[str, Any]],
    ] = {}
    primary_cells: dict[
        tuple[str, str],
        tuple[str, dict[str, Any]],
    ] = {}
    for chunk in selected_chunks:
        _sheet_index, sheet = _sheet(chunk)
        sheet_key = sheet.casefold()
        for collection in ("cells", "contextCells"):
            for cell in chunk.get(collection, []):
                coordinate = _coordinate(cell)
                if not coordinate:
                    continue
                cells_by_coordinate.setdefault(
                    (sheet_key, coordinate),
                    (sheet, cell),
                )
                if collection == "cells":
                    primary_cells.setdefault(
                        (sheet_key, coordinate),
                        (sheet, cell),
                    )

    arm_targets: dict[
        tuple[str, int],
        list[tuple[dict[str, Any], str]],
    ] = {}
    for study in result:
        if not isinstance(study, dict):
            continue
        for arm in study.get("arms", []):
            if not isinstance(arm, dict):
                continue
            match = _RESULT_ROW_ARM_KEY.fullmatch(
                str(arm.get("key") or "")
            )
            if match is None:
                continue
            row = int(match.group(1))
            evidence_sheets = {
                str(item.get("sheet") or "").casefold()
                for item in arm.get("evidence", [])
                if isinstance(item, dict)
                and str(item.get("sheet") or "").strip()
            }
            for sheet_key in evidence_sheets:
                arm_targets.setdefault((sheet_key, row), []).append(
                    (study, str(arm["key"]))
                )

    headers_by_sheet_column: dict[
        tuple[str, int],
        list[tuple[int, str, str, str]],
    ] = {}
    for (sheet_key, coordinate), (sheet, cell) in (
        cells_by_coordinate.items()
    ):
        row, column = _position(cell)
        text = _source_cell_text(cell).strip()
        if (
            not text
            or _numeric_value(cell) is not None
            or _CATEGORICAL_RESULT_STATUS.fullmatch(text)
            or len(text) > 80
            or re.search(r"[A-Za-z가-힣]", text) is None
        ):
            continue
        headers_by_sheet_column.setdefault(
            (sheet_key, column),
            [],
        ).append((row, coordinate, sheet, text))
    for values in headers_by_sheet_column.values():
        values.sort(key=lambda value: value[0])

    fallback_studies: dict[str, dict[str, Any]] = {}

    def fallback_study(
        sheet_key: str,
        sheet: str,
    ) -> dict[str, Any]:
        existing = fallback_studies.get(sheet_key)
        if existing is not None:
            return existing
        title_candidates = [
            (
                _position(cell)[0],
                coordinate,
                _source_cell_text(cell).strip(),
            )
            for (candidate_sheet, coordinate), (
                _sheet_title,
                cell,
            ) in cells_by_coordinate.items()
            if candidate_sheet == sheet_key
            and _source_cell_text(cell).strip()
            and len(_source_cell_text(cell).strip()) <= 200
            and re.search(
                r"(?:report|result|test)",
                _source_cell_text(cell),
                re.IGNORECASE,
            )
        ]
        title_candidates.sort(key=lambda value: (value[0], value[1]))
        if title_candidates:
            _title_row, title_coordinate, title = title_candidates[0]
        else:
            exact_text_candidates = sorted(
                (
                    _position(cell)[0],
                    coordinate,
                    _source_cell_text(cell).strip(),
                )
                for (candidate_sheet, coordinate), (
                    _sheet_title,
                    cell,
                ) in cells_by_coordinate.items()
                if candidate_sheet == sheet_key
                and _source_cell_text(cell).strip()
            )
            if not exact_text_candidates:
                raise StagedDraftV2Error(
                    f"Categorical status sheet {sheet!r} has no title evidence"
                )
            _title_row, title_coordinate, title = exact_text_candidates[0]
        created = {
            "key": stable_uid(
                "source-categorical-status-study",
                revision_uid,
                sheet.casefold(),
            ),
            "title": title,
            "purpose": "",
            "hypothesis": "",
            "objective": "",
            "designType": "DESCRIPTIVE_STATUS_TABLE",
            "comparisonBasis": "",
            "verificationStatus": "NEEDS_REVIEW",
            "comparabilityStatus": "UNASSESSED",
            "confoundingStatus": "UNASSESSED",
            "summary": (
                "Exact source categorical status cells that were not part "
                "of a numeric result grid."
            ),
            "limitations": [],
            "evidence": [
                {
                    "sheet": sheet,
                    "range": title_coordinate,
                    "role": "STATUS_TABLE_TITLE",
                    "sourceText": title,
                    "note": "",
                }
            ],
            "contexts": [],
            "factors": [],
            "arms": [],
            "outcomes": [],
            "measurementSeries": [],
            "comparisons": [],
            "conclusions": [],
        }
        fallback_studies[sheet_key] = created
        result.append(created)
        return created

    def fallback_arm(
        *,
        study: dict[str, Any],
        sheet_key: str,
        sheet: str,
        row: int,
        status_column: int,
        status_coordinate: str,
        status: str,
    ) -> str:
        key = stable_uid(
            "source-categorical-status-arm",
            revision_uid,
            sheet.casefold(),
            row,
        )
        if any(
            isinstance(arm, dict)
            and str(arm.get("key") or "") == key
            for arm in study.get("arms", [])
        ):
            return key
        row_candidates = [
            (
                0
                if re.fullmatch(r"#?\d+", text)
                else 1
                if re.search(r"[A-Za-z가-힣]", text)
                else 2,
                column,
                coordinate,
                text,
            )
            for (candidate_sheet, coordinate), (
                _sheet_title,
                cell,
            ) in cells_by_coordinate.items()
            for candidate_row, column in [_position(cell)]
            for text in [_source_cell_text(cell).strip()]
            if candidate_sheet == sheet_key
            and candidate_row == row
            and column < status_column
            and text
            and len(text) <= 240
            and _CATEGORICAL_RESULT_STATUS.fullmatch(text) is None
        ]
        row_candidates.sort(
            key=lambda value: (value[0], value[1])
        )
        if row_candidates:
            _score, _column, identity_coordinate, label = (
                row_candidates[0]
            )
        else:
            identity_coordinate = status_coordinate
            label = status
        study.setdefault("arms", []).append(
            {
                "key": key,
                "role": "OTHER",
                "label": label,
                "condition": label,
                "sampleSize": None,
                "sampleBasis": "",
                "matchingBasis": "",
                "factorValues": [],
                "evidence": [
                    {
                        "sheet": sheet,
                        "range": identity_coordinate,
                        "role": "STATUS_ROW_IDENTITY",
                        "sourceText": label,
                        "note": "",
                    }
                ],
            }
        )
        return key

    for (sheet_key, coordinate), (sheet, cell) in (
        primary_cells.items()
    ):
        status = _source_cell_text(cell).strip()
        if _CATEGORICAL_RESULT_STATUS.fullmatch(status) is None:
            continue
        row, column = _position(cell)
        targets = arm_targets.get((sheet_key, row), [])
        unique_targets = {
            (id(study), arm_key): (study, arm_key)
            for study, arm_key in targets
        }
        if len(unique_targets) == 1:
            study, arm_key = next(iter(unique_targets.values()))
        else:
            study = fallback_study(sheet_key, sheet)
            arm_key = fallback_arm(
                study=study,
                sheet_key=sheet_key,
                sheet=sheet,
                row=row,
                status_column=column,
                status_coordinate=coordinate,
                status=status,
            )
        header_candidates = [
            value
            for value in headers_by_sheet_column.get(
                (sheet_key, column),
                [],
            )
            if value[0] < row
        ][-2:]
        column_label = _column_label(column)
        outcome_key = f"source_status_c{column}"
        outcome = next(
            (
                item
                for item in study.get("outcomes", [])
                if isinstance(item, dict)
                and str(item.get("key") or "") == outcome_key
            ),
            None,
        )
        header_labels = list(
            dict.fromkeys(
                value[3] for value in header_candidates
            )
        )
        header_evidence = [
            {
                "sheet": header_sheet,
                "range": header_coordinate,
                "role": "CATEGORICAL_STATUS_HEADER",
                "sourceText": header_text,
                "note": "",
            }
            for (
                _header_row,
                header_coordinate,
                header_sheet,
                header_text,
            ) in header_candidates
        ]
        if outcome is None:
            outcome = {
                "key": outcome_key,
                "originalLabel": (
                    " | ".join(header_labels)
                    or f"Column {column_label} status"
                ),
                "metricType": "categorical_status",
                "unit": "",
                "favorableDirection": "UNKNOWN",
                "evidence": header_evidence,
                "observations": [],
            }
            study.setdefault("outcomes", []).append(outcome)
        else:
            existing_labels = [
                value.strip()
                for value in str(
                    outcome.get("originalLabel") or ""
                ).split("|")
                if value.strip()
            ]
            combined_labels = list(
                dict.fromkeys([*existing_labels, *header_labels])
            )
            outcome["originalLabel"] = " | ".join(
                combined_labels[:4]
            )
            existing_evidence = {
                (
                    str(item.get("sheet") or "").casefold(),
                    str(item.get("range") or "").upper(),
                )
                for item in outcome.get("evidence", [])
                if isinstance(item, dict)
            }
            for item in header_evidence:
                identity = (
                    str(item["sheet"]).casefold(),
                    str(item["range"]).upper(),
                )
                if identity not in existing_evidence:
                    outcome.setdefault("evidence", []).append(item)
                    existing_evidence.add(identity)
        observation_key = stable_uid(
            "source-categorical-status",
            revision_uid,
            sheet.casefold(),
            coordinate,
        )
        if any(
            isinstance(observation, dict)
            and str(observation.get("key") or "") == observation_key
            for observation in outcome.get("observations", [])
        ):
            continue
        outcome.setdefault("observations", []).append(
            {
                "key": observation_key,
                "arm": arm_key,
                "valueNumber": None,
                "valueText": status,
                "numerator": None,
                "denominator": None,
                "ratePpm": None,
                "min": None,
                "max": None,
                "average": None,
                "sampleSize": None,
                "replicateKey": f"source-{coordinate.lower()}",
                "evidence": [
                    {
                        "sheet": sheet,
                        "range": coordinate,
                        "role": "CATEGORICAL_STATUS",
                        "sourceText": status,
                        "note": "",
                    }
                ],
            }
        )
    return result


def _split_mixed_numeric_series_record(
    *,
    record: dict[str, Any],
    by_coordinate: dict[
        tuple[str, str],
        tuple[str, dict[str, Any]],
    ],
    revision_uid: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Split numeric runs and preserve nonnumeric cells as text observations."""

    payload = record.get("payload")
    if not isinstance(payload, dict):
        return [record], []
    try:
        value_bounds = range_bounds(payload.get("valueRange"))
        identity_bounds = range_bounds(payload.get("rowIdentityRange"))
    except StagedDraftV2Error:
        return [record], []
    if (
        value_bounds[1] != value_bounds[3]
        or identity_bounds[1] != identity_bounds[3]
        or value_bounds[0] != identity_bounds[0]
        or value_bounds[2] != identity_bounds[2]
    ):
        return [record], []
    sheet = str(payload.get("sheet") or "")
    sheet_key = sheet.casefold()
    value_column = value_bounds[1]
    identity_column = identity_bounds[1]
    rows: list[tuple[int, bool, str, str]] = []
    for row in range(value_bounds[0], value_bounds[2] + 1):
        value_coordinate = f"{_column_label(value_column)}{row}"
        identity_coordinate = f"{_column_label(identity_column)}{row}"
        value_entry = by_coordinate.get(
            (sheet_key, value_coordinate)
        )
        if value_entry is None:
            rows.append((row, False, value_coordinate, identity_coordinate))
            continue
        rows.append(
            (
                row,
                _numeric_value(value_entry[1]) is not None,
                value_coordinate,
                identity_coordinate,
            )
        )
    if all(is_numeric for _row, is_numeric, _value, _identity in rows):
        return [record], []

    numeric_runs: list[tuple[int, int]] = []
    start: int | None = None
    for row, is_numeric, _value_coordinate, _identity_coordinate in rows:
        if is_numeric and start is None:
            start = row
        elif not is_numeric and start is not None:
            numeric_runs.append((start, row - 1))
            start = None
    if start is not None:
        numeric_runs.append((start, value_bounds[2]))
    split_records: list[dict[str, Any]] = []
    for start_row, end_row in numeric_runs:
        split = copy.deepcopy(record)
        split_payload = split["payload"]
        split_payload["valueRange"] = _address(
            (start_row, value_column, end_row, value_column)
        )
        split_payload["rowIdentityRange"] = _address(
            (
                start_row,
                identity_column,
                end_row,
                identity_column,
            )
        )
        split_payload["key"] = (
            f"{str(payload.get('key') or record.get('recordId') or 'series')}"
            f"_numeric_{start_row}_{end_row}"
        )
        split["recordId"] = stable_uid(
            "split-mixed-numeric-series",
            revision_uid,
            str(record.get("recordId") or ""),
            start_row,
            end_row,
        )
        split_records.append(split)

    text_observations: list[dict[str, Any]] = []
    for (
        row,
        is_numeric,
        value_coordinate,
        identity_coordinate,
    ) in rows:
        if is_numeric:
            continue
        value_entry = by_coordinate.get(
            (sheet_key, value_coordinate)
        )
        if value_entry is None:
            continue
        text = _source_cell_text(value_entry[1]).strip()
        if not text:
            continue
        identity_entry = by_coordinate.get(
            (sheet_key, identity_coordinate)
        )
        identity_text = (
            _source_cell_text(identity_entry[1]).strip()
            if identity_entry is not None
            else ""
        )
        evidence = [
            {
                "sheet": sheet,
                "range": value_coordinate,
                "role": "TEXT_SERIES_VALUE",
                "sourceText": text,
                "note": "",
            }
        ]
        if identity_entry is not None:
            evidence.append(
                {
                    "sheet": sheet,
                    "range": identity_coordinate,
                    "role": "TEXT_SERIES_ROW_IDENTITY",
                    "sourceText": identity_text,
                    "note": "",
                }
            )
        text_observations.append(
            {
                "key": stable_uid(
                    "mixed-series-text-observation",
                    revision_uid,
                    sheet.casefold(),
                    value_coordinate,
                ),
                "arm": str(payload.get("arm") or ""),
                "valueNumber": None,
                "valueText": text,
                "numerator": None,
                "denominator": None,
                "ratePpm": None,
                "min": None,
                "max": None,
                "average": None,
                "sampleSize": None,
                "replicateKey": (
                    identity_text
                    or f"source-{identity_coordinate.lower()}"
                ),
                "evidence": evidence,
                "_sourceOutcome": str(payload.get("outcome") or ""),
            }
        )
    return split_records, text_observations


def _observation_claim_matches_numeric_source(
    *,
    payload: dict[str, Any],
    source_number: float,
    source_cell: dict[str, Any],
) -> bool:
    for field in (
        "valueNumber",
        "numerator",
        "denominator",
        "ratePpm",
        "min",
        "max",
        "average",
        "sampleSize",
    ):
        claim = payload.get(field)
        if (
            isinstance(claim, (int, float))
            and not isinstance(claim, bool)
            and math.isfinite(float(claim))
            and math.isclose(
                source_number,
                float(claim),
                rel_tol=1e-9,
                abs_tol=1e-9,
            )
        ):
            return True
    rate_ppm = payload.get("ratePpm")
    if (
        isinstance(rate_ppm, (int, float))
        and not isinstance(rate_ppm, bool)
        and math.isfinite(float(rate_ppm))
        and math.isclose(
            source_number * 1_000_000.0,
            float(rate_ppm),
            rel_tol=1e-9,
            abs_tol=1e-9,
        )
    ):
        return True
    value_number = payload.get("valueNumber")
    if (
        "%" in str(source_cell.get("numberFormat") or "")
        and isinstance(value_number, (int, float))
        and not isinstance(value_number, bool)
        and math.isfinite(float(value_number))
        and math.isclose(
            source_number * 100.0,
            float(value_number),
            rel_tol=1e-9,
            abs_tol=1e-9,
        )
    ):
        return True
    return False


def normalize_fragment_missing_observation_arms(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Bind otherwise arm-less isolated observations to exact source labels."""

    result = copy.deepcopy(fragment)
    records = result.get("records")
    if not isinstance(records, list):
        return result
    owned = {
        str(value)
        for value in envelope.get("ownedSourceCellKeys", [])
    }
    _by_coordinate, by_key, _source_order = _source_cell_maps(
        all_selected_chunks
    )
    arms_by_logical: dict[str, list[str]] = {}
    for record in records:
        if (
            not isinstance(record, dict)
            or str(record.get("recordType") or "").upper()
            != "ENTITY_DECLARATION"
            or not isinstance(record.get("payload"), dict)
            or str(
                record["payload"].get("entityType") or ""
            ).upper()
            != "ARM"
        ):
            continue
        key = str(record["payload"].get("key") or "").strip()
        logical_id = str(record.get("logicalStudyId") or "")
        if key and logical_id:
            arms_by_logical.setdefault(logical_id, []).append(key)
    generated_by_identity: dict[tuple[str, str], str] = {}
    generated_records: list[dict[str, Any]] = []
    for record in records:
        if (
            not isinstance(record, dict)
            or str(record.get("recordType") or "").upper()
            != "OBSERVATION_APPEND"
            or not isinstance(record.get("payload"), dict)
            or str(record["payload"].get("arm") or "").strip()
        ):
            continue
        logical_id = str(record.get("logicalStudyId") or "")
        known_arms = list(dict.fromkeys(arms_by_logical.get(logical_id, [])))
        if len(known_arms) == 1:
            record["payload"]["arm"] = known_arms[0]
            continue
        if known_arms:
            continue
        evidence_keys = [
            key
            for key in evidence_cell_keys(
                record.get("evidence", []),
                chunks=all_selected_chunks,
            )
            if key in owned and key in by_key
        ]
        label_key = next(
            (
                key
                for key in evidence_keys
                if _numeric_value(by_key[key][2]) is None
                and str(
                    by_key[key][2].get("displayValue")
                    if by_key[key][2].get("displayValue") is not None
                    else by_key[key][2].get("rawValue")
                    or ""
                ).strip()
            ),
            "",
        )
        if not label_key:
            continue
        label_cell = by_key[label_key][2]
        label = str(
            label_cell.get("displayValue")
            if label_cell.get("displayValue") is not None
            else label_cell.get("rawValue")
            or ""
        ).strip()
        identity = (logical_id, label_key)
        arm_key = generated_by_identity.get(identity)
        if arm_key is None:
            arm_key = stable_uid(
                "source-arm-v2",
                logical_id,
                label_key,
                label.casefold(),
            )
            generated_by_identity[identity] = arm_key
            sheet, coordinate, _cell = by_key[label_key]
            generated_records.append(
                {
                    "recordType": "ENTITY_DECLARATION",
                    "recordId": "",
                    "logicalStudyId": logical_id,
                    "identityCellKeys": [label_key],
                    "exactSourceLabel": label,
                    "payload": {
                        "entityType": "ARM",
                        "key": arm_key,
                        "role": "OTHER",
                        "label": label,
                        "condition": label,
                        "sampleSize": None,
                        "sampleBasis": "",
                        "matchingBasis": "",
                        "factorValues": [],
                    },
                    "evidence": [
                        {
                            "sheet": sheet,
                            "range": coordinate,
                            "role": "SOURCE",
                            "sourceText": label,
                            "note": (
                                "Exact source label used as the descriptive "
                                "Arm for an otherwise arm-less result."
                            ),
                        }
                    ],
                }
            )
        record["payload"]["arm"] = arm_key
    records.extend(generated_records)
    return result


def _replicate_key_preserves_source_identity(
    replicate_key: object,
    source_text: object,
) -> bool:
    replicate = " ".join(str(replicate_key or "").split()).casefold()
    source = " ".join(str(source_text or "").split()).casefold()
    if not replicate or not source:
        return False
    return bool(
        re.match(
            rf"^{re.escape(source)}(?:\s*[-|/:]\s*|$)",
            replicate,
        )
    )


def normalize_fragment_observation_replicate_evidence(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Add exact same-row numeric replicate identity evidence."""

    result = copy.deepcopy(fragment)
    records = result.get("records")
    if not isinstance(records, list):
        return result
    owned = {
        str(value)
        for value in envelope.get("ownedSourceCellKeys", [])
    }
    _by_coordinate, by_key, source_order = _source_cell_maps(
        all_selected_chunks
    )
    for record in records:
        if (
            not isinstance(record, dict)
            or str(record.get("recordType") or "").upper()
            != "OBSERVATION_APPEND"
            or not isinstance(record.get("payload"), dict)
        ):
            continue
        replicate_key = record["payload"].get("replicateKey")
        if not str(replicate_key or "").strip():
            continue
        evidence = record.get("evidence")
        if not isinstance(evidence, list):
            continue
        current_keys = set(
            evidence_cell_keys(
                evidence,
                chunks=all_selected_chunks,
            )
        )
        evidence_rows = {
            (
                by_key[key][0].casefold(),
                range_bounds(by_key[key][1])[0],
            )
            for key in current_keys
            if key in by_key
        }
        candidates: list[str] = []
        for key in owned - current_keys:
            if key not in by_key:
                continue
            sheet, coordinate, cell = by_key[key]
            if _numeric_value(cell) is None:
                continue
            if (
                sheet.casefold(),
                range_bounds(coordinate)[0],
            ) not in evidence_rows:
                continue
            source_text = (
                cell.get("displayValue")
                if cell.get("displayValue") is not None
                else cell.get("rawValue")
            )
            if _replicate_key_preserves_source_identity(
                replicate_key,
                source_text,
            ):
                candidates.append(key)
        for key in sorted(
            candidates,
            key=lambda value: source_order.get(value, 10**12),
        ):
            sheet, coordinate, cell = by_key[key]
            source_text = str(
                cell.get("displayValue")
                if cell.get("displayValue") is not None
                else cell.get("rawValue")
                or ""
            ).strip()
            evidence.append(
                {
                    "sheet": sheet,
                    "range": coordinate,
                    "role": "REPLICATE_IDENTITY",
                    "sourceText": source_text,
                    "note": (
                        "Exact source row identity preserved by replicateKey."
                    ),
                }
            )
    return result


def _arm_sample_size_has_owned_text_evidence(
    *,
    record_type: str,
    payload: dict[str, Any],
    owned_evidence_keys: set[str],
    by_key: dict[str, tuple[str, str, dict[str, Any]]],
) -> bool:
    """Allow exact ``10pcs``-style source text to support ARM sampleSize 10."""

    if (
        record_type != "ENTITY_DECLARATION"
        or str(payload.get("entityType") or "").upper() != "ARM"
    ):
        return False
    sample_size = payload.get("sampleSize")
    if isinstance(sample_size, bool) or not isinstance(
        sample_size,
        (int, float),
    ):
        return False
    expected = float(sample_size)
    if not math.isfinite(expected):
        return False
    for key in owned_evidence_keys:
        if key not in by_key:
            continue
        cell = by_key[key][2]
        value = (
            cell.get("displayValue")
            if cell.get("displayValue") is not None
            else cell.get("rawValue")
        )
        for token in re.findall(
            r"(?<![A-Za-z0-9])([-+]?(?:\d+(?:\.\d*)?|\.\d+))"
            r"\s*(?:pcs?|ea|samples?)(?![A-Za-z0-9])",
            str(value or ""),
            re.IGNORECASE,
        ):
            if math.isclose(
                float(token),
                expected,
                rel_tol=1e-9,
                abs_tol=1e-9,
            ):
                return True
    return False


_STRICT_COUNT_RATIO_TEXT = re.compile(
    r"^\s*(\d+)\s*/\s*(\d+)\s*(?:pcs?|ea|samples?)?\s*$",
    re.IGNORECASE,
)


def normalize_fragment_unsupported_text_numeric_claims(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Remove numeric promotions supported only by embedded narrative text."""

    result = copy.deepcopy(fragment)
    records = result.get("records")
    if not isinstance(records, list):
        return result
    owned = {
        str(value)
        for value in envelope.get("ownedSourceCellKeys", [])
    }
    _by_coordinate, by_key, _source_order = _source_cell_maps(
        all_selected_chunks
    )
    numeric_fields = (
        "valueNumber",
        "numerator",
        "denominator",
        "ratePpm",
        "min",
        "max",
        "average",
        "sampleSize",
    )
    for record in records:
        if not isinstance(record, dict):
            continue
        payload = record.get("payload")
        if not isinstance(payload, dict):
            continue
        evidence_keys = owned.intersection(
            evidence_cell_keys(
                record.get("evidence", []),
                chunks=all_selected_chunks,
            )
        )
        if any(
            key in by_key
            and _numeric_value(by_key[key][2]) is not None
            for key in evidence_keys
        ):
            continue
        record_type = str(record.get("recordType") or "").upper()
        if (
            record_type == "ENTITY_DECLARATION"
            and str(payload.get("entityType") or "").upper() == "ARM"
            and isinstance(payload.get("sampleSize"), (int, float))
            and not isinstance(payload.get("sampleSize"), bool)
            and not _arm_sample_size_has_owned_text_evidence(
                record_type=record_type,
                payload=payload,
                owned_evidence_keys=evidence_keys,
                by_key=by_key,
            )
        ):
            payload["sampleSize"] = None
            payload["sampleBasis"] = ""
        if record_type != "OBSERVATION_APPEND":
            continue
        claim_fields = [
            field
            for field in numeric_fields
            if isinstance(payload.get(field), (int, float))
            and not isinstance(payload.get(field), bool)
        ]
        if not claim_fields:
            continue
        source_texts = [
            str(
                by_key[key][2].get("displayValue")
                if by_key[key][2].get("displayValue") is not None
                else by_key[key][2].get("rawValue")
                or ""
            ).strip()
            for key in evidence_keys
            if key in by_key
        ]
        strict_ratio = next(
            (
                match
                for source_text in source_texts
                if (
                    match := _STRICT_COUNT_RATIO_TEXT.fullmatch(
                        source_text
                    )
                )
            ),
            None,
        )
        if strict_ratio is not None:
            numerator = payload.get("numerator")
            denominator = payload.get("denominator")
            if (
                isinstance(numerator, (int, float))
                and not isinstance(numerator, bool)
                and isinstance(denominator, (int, float))
                and not isinstance(denominator, bool)
                and math.isclose(
                    float(numerator),
                    float(strict_ratio.group(1)),
                )
                and math.isclose(
                    float(denominator),
                    float(strict_ratio.group(2)),
                )
            ):
                continue
        value_text = " ".join(
            str(payload.get("valueText") or "").split()
        ).casefold()
        source_preserves_value_text = bool(value_text) and any(
            value_text in " ".join(source_text.split()).casefold()
            for source_text in source_texts
        )
        if not source_preserves_value_text:
            continue
        for field in numeric_fields:
            payload[field] = None
    return result


def _append_semantic_label_records_v2(
    *,
    envelope: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
    logical_by_source_key: dict[str, str],
    records: list[dict[str, Any]],
    inventory: dict[str, Any] | None = None,
) -> None:
    """Append exact semantic entities required by strict content coverage."""

    if inventory is None:
        inventory = build_content_coverage_inventory(
            chunks=focused_chunks,
            locator_results=envelope.get("locatorResults", []),
            expected_source_cell_keys=envelope["ownedSourceCellKeys"],
        )
    _by_coordinate, by_key, _source_order = _source_cell_maps(
        focused_chunks
    )
    existing_ids = {str(record["recordId"]) for record in records}
    revision_uid = str(envelope["source"]["revisionUid"])
    for semantic_cell in inventory.get("semanticLabelCells", []):
        if not isinstance(semantic_cell, dict):
            continue
        source_key = str(semantic_cell.get("sourceCellKey") or "")
        logical_id = str(logical_by_source_key.get(source_key) or "")
        source_entry = by_key.get(source_key)
        if not logical_id or source_entry is None:
            raise StagedDraftV2Error(
                "Semantic label lacks deterministic Study ownership"
            )
        sheet_title, coordinate, _cell = source_entry
        label = str(semantic_cell.get("sourceText") or "").strip()
        roles = set(semantic_cell.get("semanticRoles", []))
        if not label or not roles:
            continue
        coordinate_key = re.sub(
            r"[^a-z0-9]+",
            "_",
            f"{sheet_title}_{coordinate}".casefold(),
        ).strip("_")
        if "OUTCOME_LABEL" in roles:
            entity_type = "OUTCOME"
            payload = {
                "entityType": entity_type,
                "key": f"semantic_outcome_{coordinate_key}",
                "originalLabel": label,
                "metricType": "source_labeled_result",
                "unit": "",
                "favorableDirection": (
                    "LOWER"
                    if any(
                        token in label.upper()
                        for token in (
                            "NG",
                            "THD",
                            "NOISE",
                            "TOUCH",
                        )
                    )
                    else "UNKNOWN"
                ),
            }
        elif "ARM_LABEL" in roles:
            entity_type = "ARM"
            upper_label = label.upper()
            payload = {
                "entityType": entity_type,
                "key": f"semantic_arm_{coordinate_key}",
                "role": (
                    "CONTROL"
                    if "CONTROL" in upper_label
                    else (
                        "REFERENCE"
                        if any(
                            token in upper_label
                            for token in (
                                "NORMAL",
                                "REFERENCE",
                                "STANDARD",
                                "SPEC",
                            )
                        )
                        else "OTHER"
                    )
                ),
                "label": label,
                "condition": label,
                "sampleSize": None,
                "sampleBasis": "",
                "matchingBasis": "",
                "factorValues": [],
            }
        elif "FACTOR_LABEL" in roles:
            entity_type = "FACTOR"
            payload = {
                "entityType": entity_type,
                "key": f"semantic_factor_{coordinate_key}",
                "originalLabel": label,
                "baselineCondition": "",
                "changedCondition": "",
                "changeDirection": "OTHER",
                "isolationStatus": "UNASSESSED",
            }
        elif roles.intersection(
            {"FACTOR_LEVEL", "UNIT_QUANTITY"}
        ):
            entity_type = "CONTEXT"
            payload = {
                "entityType": entity_type,
                "key": f"semantic_context_{coordinate_key}",
                "kind": (
                    "source unit quantity"
                    if "UNIT_QUANTITY" in roles
                    else "source factor level"
                ),
                "originalValue": label,
                "normalizedValue": label,
            }
        else:
            continue
        record = {
            "recordType": "ENTITY_DECLARATION",
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": [source_key],
            "exactSourceLabel": label,
            "payload": payload,
            "evidence": [
                {
                    "sheet": sheet_title,
                    "range": coordinate,
                    "role": "SOURCE",
                    "sourceText": label,
                    "note": "",
                }
            ],
        }
        record_id = stable_record_id(
            revision_uid=revision_uid,
            logical_study_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_cell_keys=[source_key],
            exact_source_label=label,
            semantic_subtype=entity_type,
        )
        if record_id in existing_ids:
            continue
        record["recordId"] = record_id
        record["payload"]["entityId"] = record_id
        records.append(record)
        existing_ids.add(record_id)


def _record_evidence(record: dict[str, Any]) -> list[dict[str, Any]]:
    value = record.get("evidence", [])
    if not isinstance(value, list):
        raise StagedDraftV2Error("record.evidence must be an array")
    return value


def _payload_source_ranges(
    value: object,
    *,
    path: str = "payload",
) -> list[tuple[str, str]]:
    """Return source-like A1 range fields nested anywhere in a payload."""

    result: list[tuple[str, str]] = []
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}"
            if (
                isinstance(child, str)
                and key.casefold().endswith("range")
                and _A1.fullmatch(child.strip())
            ):
                result.append((child_path, child))
            else:
                result.extend(
                    _payload_source_ranges(child, path=child_path)
                )
    elif isinstance(value, list):
        for index, child in enumerate(value):
            result.extend(
                _payload_source_ranges(
                    child,
                    path=f"{path}[{index}]",
                )
            )
    return result


def validate_fragment_v2(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Validate identity, evidence scope, references and exact disposition."""

    if not isinstance(fragment, dict):
        raise StagedDraftV2Error("Fragment must be an object")
    if (
        fragment.get("schemaVersion")
        != STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION
    ):
        raise StagedDraftV2Error(
            "Fragment schemaVersion must be study-draft-fragment-v2"
        )
    for field in ("planId", "partId", "inputEnvelopeSha256"):
        if str(fragment.get(field) or "") != str(
            envelope.get(field) or ""
        ):
            raise StagedDraftV2Error(
                f"Fragment {field} does not match its exact input envelope"
            )
    source = fragment.get("source")
    if not isinstance(source, dict) or bool(
        source.get("contentComplete")
    ):
        raise StagedDraftV2Error(
            "Fragment source.contentComplete must be false"
        )
    for field in ("revisionUid", "contentSha256"):
        expected = str(envelope["source"].get(field) or "")
        actual = str(source.get(field) or "")
        if field == "contentSha256":
            expected = expected.lower()
            actual = actual.lower()
        if actual != expected:
            raise StagedDraftV2Error(
                f"Fragment source mismatch: {field}"
            )

    allowed = set(envelope["ownedSourceCellKeys"]) | set(
        envelope["sharedAnchorCellKeys"]
    )
    owned = set(envelope["ownedSourceCellKeys"])
    registry_ids = {
        str(study["logicalStudyId"])
        for study in envelope["registry"]["studies"]
    }
    _by_coordinate, by_key, source_order = _source_cell_maps(
        all_selected_chunks
    )
    records = fragment.get("records")
    if not isinstance(records, list):
        raise StagedDraftV2Error("Fragment records must be an array")
    seen_record_ids: set[str] = set()
    record_evidence_keys: dict[str, set[str]] = {}
    represented_numeric_keys: set[str] = set()
    result_records: list[dict[str, Any]] = []
    for index, raw_record in enumerate(records):
        if not isinstance(raw_record, dict):
            raise StagedDraftV2Error(
                f"records[{index}] must be an object"
            )
        record = copy.deepcopy(raw_record)
        record_type = str(record.get("recordType") or "").upper()
        if record_type not in RECORD_TYPES:
            raise StagedDraftV2Error(
                f"records[{index}] has unsupported recordType"
            )
        logical_id = str(record.get("logicalStudyId") or "")
        if logical_id not in registry_ids:
            raise StagedDraftV2Error(
                f"records[{index}] creates a Study outside the registry"
            )
        identity_keys = record.get("identityCellKeys")
        if (
            not isinstance(identity_keys, list)
            or not identity_keys
            or any(str(value) not in allowed for value in identity_keys)
        ):
            raise StagedDraftV2Error(
                f"records[{index}].identityCellKeys exceed the allowlist"
            )
        exact_label = str(record.get("exactSourceLabel") or "").strip()
        if not exact_label:
            raise StagedDraftV2Error(
                f"records[{index}].exactSourceLabel is required"
            )
        expected_id = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=[str(value) for value in identity_keys],
            exact_source_label=exact_label,
            semantic_subtype=_record_semantic_subtype(record),
        )
        if str(record.get("recordId") or "") != expected_id:
            raise StagedDraftV2Error(
                f"records[{index}].recordId is not source-stable"
            )
        if expected_id in seen_record_ids:
            raise StagedDraftV2Error(
                f"Fragment contains duplicate recordId {expected_id}"
            )
        seen_record_ids.add(expected_id)
        evidence_keys = evidence_cell_keys(
            _record_evidence(record),
            chunks=all_selected_chunks,
        )
        if any(key not in allowed for key in evidence_keys):
            raise StagedDraftV2Error(
                f"records[{index}] evidence exceeds owned/shared scope"
            )
        if not evidence_keys and record_type != "LIMITATION_APPEND":
            raise StagedDraftV2Error(
                f"records[{index}] requires exact source evidence"
            )
        payload = record.get("payload")
        if not isinstance(payload, dict):
            raise StagedDraftV2Error(
                f"records[{index}].payload must be an object"
            )
        if record_type == "ENTITY_DECLARATION":
            entity_type = str(payload.get("entityType") or "").upper()
            entity_key = str(payload.get("key") or "").strip()
            if entity_type not in ENTITY_TYPES or not entity_key:
                raise StagedDraftV2Error(
                    f"records[{index}] ENTITY_DECLARATION requires "
                    "a supported entityType and key"
                )
            required_entity_fields = {
                "ARM": ("label",),
                "OUTCOME": ("originalLabel", "metricType"),
                "FACTOR": ("originalLabel",),
                "CONTEXT": ("kind", "originalValue"),
            }[entity_type]
            if any(
                not str(payload.get(field) or "").strip()
                for field in required_entity_fields
            ):
                raise StagedDraftV2Error(
                    f"records[{index}] {entity_type} declaration lacks "
                    f"required fields {required_entity_fields!r}"
                )
            if entity_type == "ARM":
                for key in owned.intersection(evidence_keys):
                    if key not in by_key:
                        continue
                    source_cell = by_key[key][2]
                    source_number = _numeric_value(source_cell)
                    if source_number is None:
                        continue
                    source_text = (
                        source_cell.get("displayValue")
                        if source_cell.get("displayValue") is not None
                        else source_cell.get("rawValue")
                    )
                    if _replicate_key_preserves_source_identity(
                        exact_label,
                        source_text,
                    ):
                        represented_numeric_keys.add(key)
        if record_type == "OBSERVATION_APPEND":
            if (
                not str(payload.get("outcome") or "").strip()
                or not str(payload.get("arm") or "").strip()
            ):
                raise StagedDraftV2Error(
                    f"records[{index}] OBSERVATION_APPEND requires outcome "
                    "and arm keys"
                )
            canonical_claim_fields = (
                "valueNumber",
                "numerator",
                "denominator",
                "ratePpm",
                "min",
                "max",
                "average",
                "sampleSize",
            )
            has_numeric_observation_claim = any(
                isinstance(payload.get(field), (int, float))
                and not isinstance(payload.get(field), bool)
                for field in canonical_claim_fields
            )
            if (
                not has_numeric_observation_claim
                and not str(payload.get("valueText") or "").strip()
            ):
                raise StagedDraftV2Error(
                    f"records[{index}] OBSERVATION_APPEND requires one "
                    "canonical value claim"
                )
        if record_type == "SERIES_SEGMENT_APPEND":
            required_series_fields = (
                "outcome",
                "arm",
                "sheet",
                "headerRange",
                "valueRange",
                "rowIdentityRange",
                "axisSource",
            )
            if any(
                not str(payload.get(field) or "").strip()
                for field in required_series_fields
            ):
                raise StagedDraftV2Error(
                    f"records[{index}] SERIES_SEGMENT_APPEND lacks required "
                    f"fields {required_series_fields!r}"
                )
            if str(payload.get("axisSource") or "").upper() not in {
                "HEADER",
                "ROW_IDENTITY",
            }:
                raise StagedDraftV2Error(
                    f"records[{index}] series axisSource must be HEADER or "
                    "ROW_IDENTITY"
                )
        payload_ranges = _payload_source_ranges(payload)
        payload_range_keys: dict[str, list[str]] = {}
        if payload_ranges:
            payload_sheet = str(payload.get("sheet") or "")
            if not payload_sheet:
                raise StagedDraftV2Error(
                    f"records[{index}] payload source ranges require sheet"
                )
            for payload_path, address in payload_ranges:
                keys = evidence_cell_keys(
                    [{"sheet": payload_sheet, "range": address}],
                    chunks=all_selected_chunks,
                )
                if not keys:
                    raise StagedDraftV2Error(
                        f"records[{index}] {payload_path} resolves to no "
                        "captured source cells"
                    )
                if any(key not in allowed for key in keys):
                    raise StagedDraftV2Error(
                        f"records[{index}] {payload_path} exceeds "
                        "owned/shared scope"
                    )
                payload_range_keys[payload_path] = keys
        if record_type == "SERIES_SEGMENT_APPEND":
            value_keys = payload_range_keys.get(
                "payload.valueRange",
                [],
            )
            owned_value_keys = owned.intersection(value_keys)
            if not owned_value_keys:
                raise StagedDraftV2Error(
                    f"records[{index}] series valueRange must include "
                    "an owned value cell; "
                    f"sheet={payload.get('sheet')!r}, "
                    f"valueRange={payload.get('valueRange')!r}, "
                    f"resolvedValueKeys={value_keys[:8]!r}, "
                    f"payloadKeys={sorted(str(key) for key in payload)!r}, "
                    f"payload={payload!r}"
                )
            if not any(
                _numeric_value(by_key[key][2]) is not None
                for key in owned_value_keys
                if key in by_key
            ):
                raise StagedDraftV2Error(
                    f"records[{index}] series valueRange lacks an "
                    "owned numeric value"
                )
            for payload_path in (
                "payload.valueRange",
                "payload.rowIdentityRange",
            ):
                represented_numeric_keys.update(
                    key
                    for key in owned.intersection(
                        payload_range_keys.get(payload_path, [])
                    )
                    if key in by_key
                    and _numeric_value(by_key[key][2]) is not None
                )
        numeric_claim = (
            record_type == "SERIES_SEGMENT_APPEND"
        ) or any(
            isinstance(value, (int, float)) and not isinstance(value, bool)
            for value in payload.values()
        )
        if numeric_claim:
            owned_evidence = owned.intersection(evidence_keys)
            if not owned_evidence:
                raise StagedDraftV2Error(
                    f"records[{index}] has a shared-only numeric claim"
                )
            has_owned_numeric_evidence = any(
                _numeric_value(by_key[key][2]) is not None
                for key in owned_evidence
                if key in by_key
            )
            if (
                not has_owned_numeric_evidence
                and not _arm_sample_size_has_owned_text_evidence(
                    record_type=record_type,
                    payload=payload,
                    owned_evidence_keys=owned_evidence,
                    by_key=by_key,
                )
            ):
                raise StagedDraftV2Error(
                    f"records[{index}] numeric claim lacks an owned value cell"
                )
        if record_type == "OBSERVATION_APPEND":
            for key in owned.intersection(evidence_keys):
                if key not in by_key:
                    continue
                source_cell = by_key[key][2]
                source_number = _numeric_value(source_cell)
                if source_number is None:
                    continue
                if _observation_claim_matches_numeric_source(
                    payload=payload,
                    source_number=source_number,
                    source_cell=source_cell,
                ):
                    represented_numeric_keys.add(key)
                    continue
                source_text = (
                    source_cell.get("displayValue")
                    if source_cell.get("displayValue") is not None
                    else source_cell.get("rawValue")
                )
                if _replicate_key_preserves_source_identity(
                    payload.get("replicateKey"),
                    source_text,
                ):
                    represented_numeric_keys.add(key)
        record["recordType"] = record_type
        record["identityCellKeys"] = [
            str(value) for value in identity_keys
        ]
        record_evidence_keys[expected_id] = set(evidence_keys)
        result_records.append(record)

    focused_chunks = envelope.get("focusedChunks")
    locator_results = envelope.get("locatorResults")
    if not isinstance(focused_chunks, list) or not isinstance(
        locator_results,
        list,
    ):
        raise StagedDraftV2Error(
            "Fragment envelope lacks focused source coverage inputs"
        )
    local_inventory = build_content_coverage_inventory(
        chunks=focused_chunks,
        locator_results=locator_results,
        expected_source_cell_keys=envelope["ownedSourceCellKeys"],
    )
    required_numeric_keys = {
        str(item["sourceCellKey"])
        for item in local_inventory.get("requiredCells", [])
        if isinstance(item, dict)
    }
    missing_numeric_keys = required_numeric_keys - represented_numeric_keys
    if missing_numeric_keys:
        missing_coordinates = [
            str(by_key[key][2].get("coordinate") or key)
            for key in envelope["ownedSourceCellKeys"]
            if key in missing_numeric_keys and key in by_key
        ]
        raise StagedDraftV2Error(
            "Fragment leaves required numeric source cells without a "
            "canonical observation/series binding: "
            f"{missing_coordinates[:12]!r}"
        )

    dispositions = fragment.get("coverageDispositions")
    if not isinstance(dispositions, list):
        raise StagedDraftV2Error(
            "Fragment coverageDispositions must be an array"
        )
    disposition_by_key: dict[str, dict[str, Any]] = {}
    for index, raw_disposition in enumerate(dispositions):
        if not isinstance(raw_disposition, dict):
            raise StagedDraftV2Error(
                f"coverageDispositions[{index}] must be an object"
            )
        disposition = copy.deepcopy(raw_disposition)
        key = str(disposition.get("sourceCellKey") or "")
        if key not in owned or key in disposition_by_key:
            raise StagedDraftV2Error(
                "coverageDispositions must assign each owned cell once"
            )
        disposition_value = str(
            disposition.get("disposition") or ""
        ).upper()
        if disposition_value not in DISPOSITIONS:
            raise StagedDraftV2Error(
                f"coverageDispositions[{index}] has invalid disposition"
            )
        record_ids = disposition.get("recordIds", [])
        if not isinstance(record_ids, list):
            raise StagedDraftV2Error(
                f"coverageDispositions[{index}].recordIds must be an array"
            )
        if disposition_value == "RECORD_EVIDENCE":
            normalized_record_ids = [
                str(record_id) for record_id in record_ids
            ]
            if (
                not normalized_record_ids
                or any(
                    record_id not in seen_record_ids
                    for record_id in normalized_record_ids
                )
                or any(
                    key not in record_evidence_keys[record_id]
                    for record_id in normalized_record_ids
                )
            ):
                raise StagedDraftV2Error(
                    f"Owned cell {key} is not evidence for its disposition records"
                )
        elif record_ids:
            raise StagedDraftV2Error(
                f"Owned cell {key} has recordIds without RECORD_EVIDENCE"
            )
        if (
            disposition_value
            in {"CONTEXT_ONLY", "NO_SEMANTIC_RECORD"}
            and not str(disposition.get("reason") or "").strip()
        ):
            raise StagedDraftV2Error(
                f"Owned cell {key} requires a deterministic disposition reason"
            )
        disposition["disposition"] = disposition_value
        disposition_by_key[key] = disposition
    if set(disposition_by_key) != owned:
        missing = sorted(owned - set(disposition_by_key))
        raise StagedDraftV2Error(
            f"Fragment disposition coverage is incomplete: {missing}"
        )
    for key in envelope["ownedSourceCellKeys"]:
        disposition = disposition_by_key[key]
        expected_record_ids = {
            record_id
            for record_id, evidence_keys in record_evidence_keys.items()
            if key in evidence_keys
        }
        actual_record_ids = {
            str(value)
            for value in disposition.get("recordIds", [])
        }
        if expected_record_ids:
            if (
                disposition["disposition"] != "RECORD_EVIDENCE"
                or actual_record_ids != expected_record_ids
            ):
                raise StagedDraftV2Error(
                    f"Owned cell {key} record evidence and disposition "
                    "recordIds must be exactly equal"
                )
        elif (
            disposition["disposition"] == "RECORD_EVIDENCE"
            or actual_record_ids
        ):
            raise StagedDraftV2Error(
                f"Owned cell {key} has a record disposition without "
                "record evidence"
            )
    result = copy.deepcopy(fragment)
    result["records"] = sorted(
        result_records,
        key=lambda record: (
            min(
                (
                    source_order.get(key, 10**12)
                    for key in record["identityCellKeys"]
                ),
                default=10**12,
            ),
            record["recordId"],
        ),
    )
    result["coverageDispositions"] = [
        disposition_by_key[key]
        for key in envelope["ownedSourceCellKeys"]
    ]
    return result


def build_deterministic_mask_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project the exact two-column MASK HZ/% lookup table without AI.

    This recognizer is deliberately narrow. It activates only when the
    fragment owns the complete AE8:AF18 table, the three source labels match,
    and both nine-row columns contain numeric values.
    """

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or len(focused_chunks) != 1:
        return None
    chunk = focused_chunks[0]
    if not isinstance(chunk, dict):
        return None
    if envelope.get("sharedAnchorCellKeys"):
        return None
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None

    expected_coordinates = [
        "AE8",
        "AE9",
        "AF9",
        *[
            coordinate
            for row in range(10, 19)
            for coordinate in (f"AE{row}", f"AF{row}")
        ],
    ]
    cells = chunk.get("cells")
    if not isinstance(cells, list):
        return None
    cells_by_coordinate = {
        str(cell.get("coordinate") or "").upper(): cell
        for cell in cells
        if isinstance(cell, dict)
    }
    if set(cells_by_coordinate) != set(expected_coordinates):
        return None
    source_keys = [
        _cell_key(chunk, cells_by_coordinate[coordinate])
        for coordinate in expected_coordinates
    ]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None

    def source_value(coordinate: str) -> object:
        cell = cells_by_coordinate[coordinate]
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    if (
        str(source_value("AE8") or "").strip().upper() != "MASK"
        or str(source_value("AE9") or "").strip().upper() != "HZ"
        or str(source_value("AF9") or "").strip() != "%"
    ):
        return None
    if any(
        _numeric_value(cells_by_coordinate[coordinate]) is None
        for coordinate in expected_coordinates[3:]
    ):
        return None

    _sheet_index, sheet_title = _sheet(chunk)
    if not sheet_title:
        return None
    key_by_coordinate = dict(zip(expected_coordinates, source_keys))

    def evidence(address: str, source_text: str) -> list[dict[str, Any]]:
        return [
            {
                "sheet": sheet_title,
                "range": address,
                "role": "SOURCE",
                "sourceText": source_text,
                "note": "",
            }
        ]

    def record(
        *,
        record_type: str,
        identity_coordinates: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_range: str,
        evidence_text: str,
    ) -> dict[str, Any]:
        identity_keys = [
            key_by_coordinate[coordinate]
            for coordinate in identity_coordinates
        ]
        record_value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": identity_keys,
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": evidence(evidence_range, evidence_text),
        }
        record_value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=identity_keys,
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(record_value),
        )
        if record_type == "ENTITY_DECLARATION":
            record_value["payload"]["entityId"] = record_value["recordId"]
        return record_value

    records = [
        record(
            record_type="STUDY_PATCH",
            identity_coordinates=["AE8"],
            exact_source_label="MASK",
            payload={
                "title": "MASK",
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "source lookup profile",
                "comparisonBasis": "",
                "summary": (
                    "Nine percentage values are mapped to source HZ "
                    "identities from 100 through 14000."
                ),
            },
            evidence_range="AE8:AF18",
            evidence_text="MASK; HZ; %",
        ),
        record(
            record_type="ENTITY_DECLARATION",
            identity_coordinates=["AE8"],
            exact_source_label="MASK",
            payload={
                "entityType": "ARM",
                "key": "mask_profile_arm",
                "role": "OTHER",
                "label": "MASK",
                "condition": "MASK",
                "sampleSize": None,
                "sampleBasis": "",
                "matchingBasis": "",
                "factorValues": [],
            },
            evidence_range="AE8:AF8",
            evidence_text="MASK",
        ),
        record(
            record_type="ENTITY_DECLARATION",
            identity_coordinates=["AF9"],
            exact_source_label="%",
            payload={
                "entityType": "OUTCOME",
                "key": "mask_percent",
                "originalLabel": "%",
                "metricType": "percentage_profile",
                "unit": "%",
                "favorableDirection": "UNKNOWN",
            },
            evidence_range="AF9",
            evidence_text="%",
        ),
        record(
            record_type="SERIES_SEGMENT_APPEND",
            identity_coordinates=["AE9", "AF9"],
            exact_source_label="HZ | %",
            payload={
                "key": "mask_percent_by_hz",
                "seriesRole": "RAW",
                "aggregationFunction": "",
                "aggregateOfSeries": [],
                "outcome": "mask_percent",
                "arm": "mask_profile_arm",
                "sheet": sheet_title,
                "headerRange": "AF9:AF9",
                "valueRange": "AF10:AF18",
                "rowIdentityRange": "AE10:AE18",
                "aggregateReplicateRanges": [],
                "axisSource": "ROW_IDENTITY",
                "axisLabel": "HZ",
                "axisUnit": "HZ",
                "valueUnit": "%",
                "stratumKey": "",
                "verificationStatus": "NEEDS_REVIEW",
            },
            evidence_range="AE9:AF18",
            evidence_text="HZ and percentage values",
        ),
    ]
    evidence_by_record = {
        str(item["recordId"]): set(
            evidence_cell_keys(
                item["evidence"],
                chunks=focused_chunks,
            )
        )
        for item in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(item["recordId"])
            for item in records
            if source_key in evidence_by_record[str(item["recordId"])]
        ]
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_fo_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project the exact B3:I15 RESULT CHECKING FO table without AI."""

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or len(focused_chunks) != 1:
        return None
    chunk = focused_chunks[0]
    if not isinstance(chunk, dict) or envelope.get("sharedAnchorCellKeys"):
        return None
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None

    expected_coordinates = ["B3"]
    for row in range(5, 15):
        expected_coordinates.extend(
            f"{column}{row}" for column in "BCDEFG"
        )
        if row == 5:
            expected_coordinates.extend(("H5", "I5"))
    expected_coordinates.extend(f"{column}15" for column in "BCDEFGHI")
    cells = chunk.get("cells")
    if not isinstance(cells, list):
        return None
    cells_by_coordinate = {
        str(cell.get("coordinate") or "").upper(): cell
        for cell in cells
        if isinstance(cell, dict)
    }
    if set(cells_by_coordinate) != set(expected_coordinates):
        return None
    source_keys = [
        _cell_key(chunk, cells_by_coordinate[coordinate])
        for coordinate in expected_coordinates
    ]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None

    def source_value(coordinate: str) -> object:
        cell = cells_by_coordinate[coordinate]
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    def source_text(coordinate: str) -> str:
        return str(source_value(coordinate) or "").strip()

    if source_text("B3").upper() != "RESULT CHECKING FO":
        return None
    group_specs = (
        ("test_1", "B", "C", re.compile(r"^TEST\s*1\s*#\s*(\d+)$", re.I)),
        ("test_2", "D", "E", re.compile(r"^TEST\s*2\s*#\s*(\d+)$", re.I)),
        ("normal", "F", "G", re.compile(r"^NORMAL\s*#\s*(\d+)$", re.I)),
    )
    for _group, label_column, value_column, label_pattern in group_specs:
        for row in range(5, 15):
            match = label_pattern.fullmatch(source_text(f"{label_column}{row}"))
            if (
                match is None
                or int(match.group(1)) != row - 4
                or _numeric_value(
                    cells_by_coordinate[f"{value_column}{row}"]
                )
                is None
            ):
                return None
    if (
        source_text("H5").upper() != "ST"
        or _numeric_value(cells_by_coordinate["I5"]) is None
    ):
        return None
    for label_coordinate, value_coordinate in (
        ("B15", "C15"),
        ("D15", "E15"),
        ("F15", "G15"),
        ("H15", "I15"),
    ):
        if (
            not source_text(label_coordinate).upper().startswith("AVG")
            or _numeric_value(cells_by_coordinate[value_coordinate]) is None
        ):
            return None

    _sheet_index, sheet_title = _sheet(chunk)
    if not sheet_title:
        return None
    key_by_coordinate = dict(zip(expected_coordinates, source_keys))

    def evidence(address: str, text: str) -> list[dict[str, Any]]:
        return [
            {
                "sheet": sheet_title,
                "range": address,
                "role": "SOURCE",
                "sourceText": text,
                "note": "",
            }
        ]

    def record(
        *,
        record_type: str,
        identity_coordinates: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_range: str,
        evidence_text: str,
    ) -> dict[str, Any]:
        identity_keys = [
            key_by_coordinate[coordinate]
            for coordinate in identity_coordinates
        ]
        record_value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": identity_keys,
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": evidence(evidence_range, evidence_text),
        }
        record_value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=identity_keys,
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(record_value),
        )
        if record_type == "ENTITY_DECLARATION":
            record_value["payload"]["entityId"] = record_value["recordId"]
        return record_value

    arm_specs = (
        (
            "fo_test_1",
            "TEST",
            re.sub(r"\s*#\s*1\s*$", "", source_text("B5")).strip(),
            source_text("B5"),
            "B5",
            "B5:B14",
        ),
        (
            "fo_test_2",
            "TEST",
            re.sub(r"\s*#\s*1\s*$", "", source_text("D5")).strip(),
            source_text("D5"),
            "D5",
            "D5:D14",
        ),
        (
            "fo_normal",
            "REFERENCE",
            (
                re.sub(
                    r"\s*#\s*1\s*$",
                    "",
                    source_text("F5"),
                ).strip()
                + " group"
            ),
            source_text("F5"),
            "F5",
            "F5:F14",
        ),
        (
            "fo_st",
            "OTHER",
            source_text("H5"),
            source_text("H5"),
            "H5",
            "H5",
        ),
    )
    records = [
        record(
            record_type="STUDY_PATCH",
            identity_coordinates=["B3"],
            exact_source_label=source_text("B3"),
            payload={
                "title": source_text("B3"),
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "parallel labeled result check",
                "comparisonBasis": "",
                "summary": (
                    "Ten FO values each are preserved for Test 1, Test 2, "
                    "and Normal, together with one ST value and four source "
                    "average values."
                ),
            },
            evidence_range="B3:I15",
            evidence_text="RESULT CHECKING FO",
        ),
        *[
            record(
                record_type="ENTITY_DECLARATION",
                identity_coordinates=[identity_coordinate],
                exact_source_label=exact_source_label,
                payload={
                    "entityType": "ARM",
                    "key": arm_key,
                    "role": role,
                    "label": label,
                    "condition": label,
                    "sampleSize": None,
                    "sampleBasis": "",
                    "matchingBasis": "",
                    "factorValues": [],
                },
                evidence_range=evidence_range,
                evidence_text=label,
            )
            for (
                arm_key,
                role,
                label,
                exact_source_label,
                identity_coordinate,
                evidence_range,
            ) in arm_specs
        ],
        record(
            record_type="ENTITY_DECLARATION",
            identity_coordinates=["B3"],
            exact_source_label=source_text("B3"),
            payload={
                "entityType": "OUTCOME",
                "key": "fo_numeric_result",
                "originalLabel": source_text("B3"),
                "metricType": "numeric_measurement",
                "unit": "",
                "favorableDirection": "UNKNOWN",
            },
            evidence_range="B3:I15",
            evidence_text="RESULT CHECKING FO numeric values",
        ),
    ]

    observation_specs: list[
        tuple[str, str, str, str, bool]
    ] = []
    for arm_key, label_column, value_column, _pattern in group_specs:
        canonical_arm = f"fo_{arm_key}"
        for row in range(5, 15):
            observation_specs.append(
                (
                    canonical_arm,
                    f"{label_column}{row}",
                    f"{value_column}{row}",
                    f"{label_column}{row}:{value_column}{row}",
                    False,
                )
            )
    observation_specs.extend(
        (
            ("fo_st", "H5", "I5", "H5:I5", False),
            ("fo_test_1", "B15", "C15", "B15:C15", True),
            ("fo_test_2", "D15", "E15", "D15:E15", True),
            ("fo_normal", "F15", "G15", "F15:G15", True),
            ("fo_st", "H15", "I15", "H15:I15", True),
        )
    )
    for arm_key, label_coordinate, value_coordinate, address, is_average in (
        observation_specs
    ):
        numeric_value = _numeric_value(
            cells_by_coordinate[value_coordinate]
        )
        if numeric_value is None:
            return None
        displayed_value = cells_by_coordinate[value_coordinate].get(
            "displayValue"
        )
        payload = {
            "outcome": "fo_numeric_result",
            "arm": arm_key,
            "valueNumber": numeric_value,
            "valueText": str(
                displayed_value
                if displayed_value is not None
                else source_value(value_coordinate)
            ),
            "numerator": None,
            "denominator": None,
            "ratePpm": None,
            "min": None,
            "max": None,
            "average": numeric_value if is_average else None,
            "sampleSize": None,
            "replicateKey": f"source-{value_coordinate.lower()}",
        }
        records.append(
            record(
                record_type="OBSERVATION_APPEND",
                identity_coordinates=[label_coordinate, value_coordinate],
                exact_source_label=source_text(label_coordinate),
                payload=payload,
                evidence_range=address,
                evidence_text=(
                    f"{source_text(label_coordinate)} "
                    f"{payload['valueText']}"
                ),
            )
        )

    evidence_by_record = {
        str(item["recordId"]): set(
            evidence_cell_keys(
                item["evidence"],
                chunks=focused_chunks,
            )
        )
        for item in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(item["recordId"])
            for item in records
            if source_key in evidence_by_record[str(item["recordId"])]
        ]
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_function_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project the exact B3:Q9 lot-function result grid without AI."""

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or len(focused_chunks) != 1:
        return None
    chunk = focused_chunks[0]
    if not isinstance(chunk, dict) or envelope.get("sharedAnchorCellKeys"):
        return None
    _sheet_index, sheet_title = _sheet(chunk)
    if sheet_title.strip().casefold() != "function":
        return None
    if str(chunk.get("primaryRange") or "").upper() != "B3:Q9":
        return None
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None

    expected_coordinates = [
        "B3",
        "B4",
        "C4",
        "D4",
        "E4",
        "F4",
        "J4",
        "N4",
        *[f"{column}5" for column in "FGHIJKLMNOPQ"],
        "B6",
        "C6",
        *[f"{column}6" for column in "DEFGHIJKLMNOPQ"],
        *[f"{column}7" for column in "FGHIJKNO"],
        "C8",
        *[f"{column}8" for column in "DEFGHIJKLMNOPQ"],
        *[f"{column}9" for column in "FGHIJKNO"],
    ]
    cells = chunk.get("cells")
    if not isinstance(cells, list):
        return None
    cells_by_coordinate = {
        _coordinate(cell): cell
        for cell in cells
        if isinstance(cell, dict)
    }
    if set(cells_by_coordinate) != set(expected_coordinates):
        return None
    source_keys = [
        _cell_key(chunk, cells_by_coordinate[coordinate])
        for coordinate in expected_coordinates
    ]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None

    def source_value(coordinate: str) -> object:
        cell = cells_by_coordinate[coordinate]
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    def source_text(coordinate: str) -> str:
        return str(source_value(coordinate) or "").strip()

    def normalized(coordinate: str) -> str:
        return " ".join(source_text(coordinate).upper().split())

    if not normalized("B3").startswith("RESULT CHECK FUNCTION"):
        return None
    expected_headers = {
        "B4": "DATE",
        "C4": "TEST TYPE",
        "D4": "INPUT",
        "E4": "OK",
        "F4": "SIGMA",
        "J4": "HEARING ( + 1V )",
        "N4": "HEARING ( + 0V )",
        "F5": "SPL",
        "G5": "THD",
        "H5": "SPL+THD",
        "I5": "SPL+THD+F0",
        "J5": "NOISE",
        "K5": "TOUCH",
        "L5": "TOTAL NG",
        "M5": "NG RATE",
        "N5": "NOISE",
        "O5": "TOUCH",
        "P5": "TOTAL NG",
        "Q5": "NG RATE",
    }
    if any(
        (
            normalized(coordinate) != expected
            if coordinate != "C4"
            else normalized(coordinate) not in {"TEST TYPE", "TYPE"}
        )
        for coordinate, expected in expected_headers.items()
    ):
        return None
    if not source_text("C6"):
        return None
    if not source_text("C8"):
        return None
    base_value_coordinates = [
        f"{column}{row}"
        for row in (6, 8)
        for column in "DEFGHIJKLMNOPQ"
    ]
    if any(
        _numeric_value(cells_by_coordinate[coordinate]) is None
        for coordinate in base_value_coordinates
    ):
        return None
    rate_coordinates = [
        f"{column}{row}"
        for row in (7, 9)
        for column in "FGHIJKNO"
    ]
    rate_states = [
        (
            "numeric"
            if _numeric_value(cells_by_coordinate[coordinate]) is not None
            else (
                "error"
                if str(source_value(coordinate) or "")
                .strip()
                .upper()
                .startswith("#")
                else "other"
            )
        )
        for coordinate in rate_coordinates
    ]
    if len(set(rate_states)) != 1 or rate_states[0] not in {
        "numeric",
        "error",
    }:
        return None
    has_numeric_rates = rate_states[0] == "numeric"
    key_by_coordinate = dict(zip(expected_coordinates, source_keys))

    def evidence(address: str, text: str) -> list[dict[str, Any]]:
        return [
            {
                "sheet": sheet_title,
                "range": address,
                "role": "SOURCE",
                "sourceText": text,
                "note": "",
            }
        ]

    def record(
        *,
        record_type: str,
        identity_coordinates: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_range: str,
        evidence_text: str,
    ) -> dict[str, Any]:
        identity_keys = [
            key_by_coordinate[coordinate]
            for coordinate in identity_coordinates
        ]
        value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": identity_keys,
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": evidence(evidence_range, evidence_text),
        }
        value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=identity_keys,
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(value),
        )
        if record_type == "ENTITY_DECLARATION":
            value["payload"]["entityId"] = value["recordId"]
        return value

    title = source_text("B3")
    records = [
        record(
            record_type="STUDY_PATCH",
            identity_coordinates=["B3"],
            exact_source_label=title,
            payload={
                "title": title,
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "parallel lot function result check",
                "comparisonBasis": "",
                "summary": (
                    "Two labeled test configurations are preserved with "
                    "input, OK, Sigma, and hearing result cells."
                ),
            },
            evidence_range="B3:Q9",
            evidence_text=title,
        ),
        *[
            record(
                record_type="ENTITY_DECLARATION",
                identity_coordinates=[coordinate],
                exact_source_label=source_text(coordinate),
                payload={
                    "entityType": "ARM",
                    "key": arm_key,
                    "role": role,
                    "label": source_text(coordinate),
                    "condition": source_text(coordinate),
                    "sampleSize": None,
                    "sampleBasis": "",
                    "matchingBasis": "",
                    "factorValues": [],
                },
                evidence_range=coordinate,
                evidence_text=source_text(coordinate),
            )
            for arm_key, coordinate, role in (
                ("function_test_1", "C6", "TEST"),
                (
                    "function_test_2",
                    "C8",
                    (
                        "REFERENCE"
                        if normalized("C8").startswith("NORMAL")
                        else "TEST"
                    ),
                ),
            )
        ],
    ]

    outcome_specs = {
        "D": ("input_count", "Input", "count", "", "UNKNOWN", "D4"),
        "E": ("ok_count", "OK", "count", "", "HIGH", "E4"),
        "F": (
            "sigma_spl_ng_count",
            "Sigma | SPL",
            "failure_count",
            "",
            "LOW",
            "F5",
        ),
        "G": (
            "sigma_thd_ng_count",
            "Sigma | THD",
            "failure_count",
            "",
            "LOW",
            "G5",
        ),
        "H": (
            "sigma_spl_thd_ng_count",
            "Sigma | SPL+THD",
            "failure_count",
            "",
            "LOW",
            "H5",
        ),
        "I": (
            "sigma_spl_thd_f0_ng_count",
            "Sigma | SPL+THD+F0",
            "failure_count",
            "",
            "LOW",
            "I5",
        ),
        "J": (
            "hearing_1v_noise_ng_count",
            "Hearing (+1V) | Noise",
            "failure_count",
            "",
            "LOW",
            "J5",
        ),
        "K": (
            "hearing_1v_touch_ng_count",
            "Hearing (+1V) | Touch",
            "failure_count",
            "",
            "LOW",
            "K5",
        ),
        "L": (
            "hearing_1v_total_ng_count",
            "Hearing (+1V) | Total NG",
            "failure_count",
            "",
            "LOW",
            "L5",
        ),
        "M": (
            "hearing_1v_ng_rate",
            "Hearing (+1V) | NG Rate",
            "rate",
            "fraction",
            "LOW",
            "M5",
        ),
        "N": (
            "hearing_0v_noise_ng_count",
            "Hearing (+0V) | Noise",
            "failure_count",
            "",
            "LOW",
            "N5",
        ),
        "O": (
            "hearing_0v_touch_ng_count",
            "Hearing (+0V) | Touch",
            "failure_count",
            "",
            "LOW",
            "O5",
        ),
        "P": (
            "hearing_0v_total_ng_count",
            "Hearing (+0V) | Total NG",
            "failure_count",
            "",
            "LOW",
            "P5",
        ),
        "Q": (
            "hearing_0v_ng_rate",
            "Hearing (+0V) | NG Rate",
            "rate",
            "fraction",
            "LOW",
            "Q5",
        ),
    }
    for (
        _column,
        (
            outcome_key,
            label,
            metric_type,
            unit,
            favorable_direction,
            header_coordinate,
        ),
    ) in outcome_specs.items():
        records.append(
            record(
                record_type="ENTITY_DECLARATION",
                identity_coordinates=[header_coordinate],
                exact_source_label=label,
                payload={
                    "entityType": "OUTCOME",
                    "key": outcome_key,
                    "originalLabel": label,
                    "metricType": metric_type,
                    "unit": unit,
                    "favorableDirection": {
                        "HIGH": "HIGHER",
                        "LOW": "LOWER",
                    }.get(favorable_direction, favorable_direction),
                },
                evidence_range=header_coordinate,
                evidence_text=source_text(header_coordinate),
            )
        )

    share_outcomes: dict[str, str] = {}
    if has_numeric_rates:
        for column in "FGHIJKNO":
            (
                base_key,
                base_label,
                _metric_type,
                _unit,
                _direction,
                header_coordinate,
            ) = outcome_specs[column]
            share_key = f"{base_key}_share"
            share_label = f"{base_label} share of subtotal"
            share_outcomes[column] = share_key
            records.append(
                record(
                    record_type="ENTITY_DECLARATION",
                    identity_coordinates=[
                        header_coordinate,
                        f"{column}7",
                    ],
                    exact_source_label=share_label,
                    payload={
                        "entityType": "OUTCOME",
                        "key": share_key,
                        "originalLabel": share_label,
                        "metricType": "component_share",
                        "unit": "fraction",
                        "favorableDirection": "LOWER",
                    },
                    evidence_range=f"{header_coordinate}:{column}7",
                    evidence_text=share_label,
                )
            )

    def append_observation(
        *,
        arm_key: str,
        outcome_key: str,
        value_coordinate: str,
        label: str,
    ) -> None:
        numeric_value = _numeric_value(
            cells_by_coordinate[value_coordinate]
        )
        if numeric_value is None:
            raise StagedDraftV2Error(
                f"Function projector lost numeric cell {value_coordinate}"
            )
        displayed_value = cells_by_coordinate[value_coordinate].get(
            "displayValue"
        )
        records.append(
            record(
                record_type="OBSERVATION_APPEND",
                identity_coordinates=[value_coordinate],
                exact_source_label=label,
                payload={
                    "outcome": outcome_key,
                    "arm": arm_key,
                    "valueNumber": numeric_value,
                    "valueText": str(
                        displayed_value
                        if displayed_value is not None
                        else source_value(value_coordinate)
                    ),
                    "numerator": None,
                    "denominator": None,
                    "ratePpm": None,
                    "min": None,
                    "max": None,
                    "average": None,
                    "sampleSize": None,
                    "replicateKey": (
                        f"source-{value_coordinate.lower()}"
                    ),
                },
                evidence_range=value_coordinate,
                evidence_text=str(
                    displayed_value
                    if displayed_value is not None
                    else source_value(value_coordinate)
                ),
            )
        )

    for row, arm_key in ((6, "function_test_1"), (8, "function_test_2")):
        for column, (
            outcome_key,
            label,
            _metric_type,
            _unit,
            _direction,
            _header_coordinate,
        ) in outcome_specs.items():
            append_observation(
                arm_key=arm_key,
                outcome_key=outcome_key,
                value_coordinate=f"{column}{row}",
                label=label,
            )
    if has_numeric_rates:
        for row, arm_key in (
            (7, "function_test_1"),
            (9, "function_test_2"),
        ):
            for column, share_key in share_outcomes.items():
                append_observation(
                    arm_key=arm_key,
                    outcome_key=share_key,
                    value_coordinate=f"{column}{row}",
                    label=(
                        f"{outcome_specs[column][1]} share of subtotal"
                    ),
                )
    else:
        for row, arm_label in ((7, "Test 1"), (9, "Test 2")):
            records.append(
                record(
                    record_type="LIMITATION_APPEND",
                    identity_coordinates=[f"F{row}", f"O{row}"],
                    exact_source_label=f"{arm_label} derived rate errors",
                    payload={
                        "text": (
                            f"{arm_label} derived component-rate cells "
                            "contain unresolved #DIV/0! source errors."
                        ),
                        "scope": "STUDY",
                    },
                    evidence_range=f"F{row}:O{row}",
                    evidence_text="#DIV/0! source errors",
                )
            )

    evidence_by_record = {
        str(item["recordId"]): set(
            evidence_cell_keys(
                item["evidence"],
                chunks=focused_chunks,
            )
        )
        for item in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(item["recordId"])
            for item in records
            if source_key in evidence_by_record[str(item["recordId"])]
        ]
        if not record_ids:
            return None
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def _coalesced_single_sheet_chunk(
    focused_chunks: object,
) -> dict[str, Any] | None:
    """Expose adjacent packet chunks as one deterministic table surface."""

    if not isinstance(focused_chunks, list) or not focused_chunks:
        return None
    if not all(isinstance(chunk, dict) for chunk in focused_chunks):
        return None
    if len(focused_chunks) == 1:
        return focused_chunks[0]
    sheet_identities = {_sheet(chunk) for chunk in focused_chunks}
    revision_uids = {
        str(
            (chunk.get("sourceRevision") or {}).get("revisionUid")
            or ""
        )
        for chunk in focused_chunks
    }
    if len(sheet_identities) != 1 or len(revision_uids) != 1:
        return None
    primary_cells: list[dict[str, Any]] = []
    primary_keys: set[str] = set()
    for chunk in focused_chunks:
        for cell in chunk.get("cells", []):
            if not isinstance(cell, dict):
                continue
            source_key = _cell_key(chunk, cell)
            if source_key in primary_keys:
                return None
            primary_keys.add(source_key)
            primary_cells.append(copy.deepcopy(cell))
    context_cells: list[dict[str, Any]] = []
    context_keys: set[str] = set()
    for chunk in focused_chunks:
        for cell in chunk.get("contextCells", []):
            if not isinstance(cell, dict):
                continue
            source_key = _cell_key(chunk, cell)
            if source_key in primary_keys or source_key in context_keys:
                continue
            context_keys.add(source_key)
            context_cells.append(copy.deepcopy(cell))
    result = copy.deepcopy(focused_chunks[0])
    result["cells"] = primary_cells
    result["contextCells"] = context_cells
    result["primaryRange"] = "MULTI_CHUNK"
    result["coalescedChunkIds"] = [
        str(chunk.get("chunkId") or "") for chunk in focused_chunks
    ]
    return result


def build_deterministic_function_grid_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project source-complete variants of lot Function result grids."""

    focused_chunks = envelope.get("focusedChunks")
    chunk = _coalesced_single_sheet_chunk(focused_chunks)
    if chunk is None:
        return None
    primary_cells = [
        cell
        for cell in chunk.get("cells", [])
        if isinstance(cell, dict)
    ]
    context_cells = [
        cell
        for cell in chunk.get("contextCells", [])
        if isinstance(cell, dict)
    ]
    if not primary_cells:
        return None
    source_keys = [_cell_key(chunk, cell) for cell in primary_cells]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None
    owned = set(source_keys)
    allowed = owned.union(
        str(value) for value in envelope.get("sharedAnchorCellKeys", [])
    )
    cells_by_position = {
        _position(cell): cell for cell in [*context_cells, *primary_cells]
    }
    primary_by_position = {
        _position(cell): cell for cell in primary_cells
    }

    def source_value(cell: dict[str, Any]) -> object:
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    def source_text(cell: dict[str, Any] | None) -> str:
        if cell is None:
            return ""
        value = source_value(cell)
        if isinstance(value, dict):
            return str(value.get("value") or "").strip()
        return str(value if value is not None else "").strip()

    def normalized(cell: dict[str, Any] | None) -> str:
        return " ".join(source_text(cell).upper().split())

    result_titles = sorted(
        (
            row,
            column,
            cell,
        )
        for (row, column), cell in cells_by_position.items()
        if normalized(cell).startswith("RESULT")
    )
    function_titles = [
        value
        for value in result_titles
        if "FUNCTION" in normalized(value[2])
    ]
    if not function_titles:
        return None
    max_primary_row = max(row for row, _column_value in primary_by_position)
    panels: list[dict[str, Any]] = []
    for title_row, title_column, title_cell in function_titles:
        next_result_rows = [
            row for row, _column_value, _cell in result_titles if row > title_row
        ]
        panel_end_row = (
            min(next_result_rows) - 1
            if next_result_rows
            else max_primary_row
        )
        panel_owned_cells = [
            cell
            for (row, _column_value), cell in primary_by_position.items()
            if title_row <= row <= panel_end_row
            or (
                title_row < min(row for row, _col in primary_by_position)
                and row <= panel_end_row
            )
        ]
        if not panel_owned_cells:
            continue
        panels.append(
            {
                "titleRow": title_row,
                "titleColumn": title_column,
                "titleCell": title_cell,
                "endRow": panel_end_row,
                "ownedCells": panel_owned_cells,
            }
        )
    if not panels:
        return None
    panel_owned_keys = {
        _cell_key(chunk, cell)
        for panel in panels
        for cell in panel["ownedCells"]
    }
    if panel_owned_keys != owned:
        return None

    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or not studies:
        return None
    for panel in panels:
        panel_start = int(panel["titleRow"])
        panel_end = int(panel["endRow"])
        panel_keys = {
            _cell_key(chunk, cell)
            for (row, _column_value), cell in cells_by_position.items()
            if panel_start <= row <= panel_end
        }
        scored_studies = sorted(
            (
                len(
                    panel_keys.intersection(
                        str(value)
                        for value in study.get(
                            "anchorEvidenceCellKeys",
                            [],
                        )
                    )
                ),
                str(study.get("logicalStudyId") or ""),
            )
            for study in studies
        )
        if (
            not scored_studies
            or scored_studies[-1][0] <= 0
            or (
                len(scored_studies) > 1
                and scored_studies[-2][0] == scored_studies[-1][0]
            )
        ):
            return None
        panel["logicalStudyId"] = scored_studies[-1][1]

    _sheet_index, sheet_title = _sheet(chunk)

    def coordinate(cell: dict[str, Any]) -> str:
        return _coordinate(cell)

    def key(cell: dict[str, Any]) -> str:
        return _cell_key(chunk, cell)

    def evidence(
        address: str,
        text: str,
    ) -> dict[str, Any]:
        return {
            "sheet": sheet_title,
            "range": address,
            "role": "SOURCE",
            "sourceText": text,
            "note": "",
        }

    records: list[dict[str, Any]] = []

    def append_record(
        *,
        logical_id: str,
        record_type: str,
        identity_keys: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_items: Sequence[dict[str, Any]],
    ) -> dict[str, Any]:
        if (
            not identity_keys
            or any(value not in allowed for value in identity_keys)
            or not exact_source_label.strip()
        ):
            raise StagedDraftV2Error(
                "Flexible Function projector produced invalid identity"
            )
        value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": list(identity_keys),
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": copy.deepcopy(list(evidence_items)),
        }
        value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=identity_keys,
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(value),
        )
        if record_type == "ENTITY_DECLARATION":
            value["payload"]["entityId"] = value["recordId"]
        records.append(value)
        return value

    def header_anchor(
        *,
        header_row: int,
        leaf_row: int,
        column: int,
    ) -> dict[str, Any] | None:
        leaf = cells_by_position.get((leaf_row, column))
        if leaf is not None and source_text(leaf):
            return leaf
        direct = cells_by_position.get((header_row, column))
        if direct is not None and source_text(direct):
            return direct
        for (row, _anchor_column), cell in cells_by_position.items():
            if row != header_row or not source_text(cell):
                continue
            merge_range = str(cell.get("mergeRange") or "").strip()
            if not merge_range:
                continue
            (
                _start_row,
                start_column,
                _end_row,
                end_column,
            ) = range_bounds(merge_range)
            if start_column <= column <= end_column:
                return cell
        return None

    for panel in panels:
        logical_id = str(panel["logicalStudyId"])
        title_row = int(panel["titleRow"])
        header_row = title_row + 1
        leaf_row = title_row + 2
        panel_end_row = int(panel["endRow"])
        title_cell = panel["titleCell"]
        title = source_text(title_cell)
        owned_cells = list(panel["ownedCells"])
        min_owned_row = min(_position(cell)[0] for cell in owned_cells)
        min_owned_column = min(_position(cell)[1] for cell in owned_cells)
        max_owned_row = max(_position(cell)[0] for cell in owned_cells)
        max_owned_column = max(_position(cell)[1] for cell in owned_cells)
        panel_range = (
            f"{_column_label(min_owned_column)}{min_owned_row}:"
            f"{_column_label(max_owned_column)}{max_owned_row}"
        )
        append_record(
            logical_id=logical_id,
            record_type="STUDY_PATCH",
            identity_keys=[key(title_cell)],
            exact_source_label=title,
            payload={
                "title": title,
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "parallel lot function result grid",
                "comparisonBasis": "",
                "summary": (
                    "Source rows preserve each labeled condition, input/OK "
                    "count, Sigma result, and hearing result."
                ),
            },
            evidence_items=[evidence(panel_range, title)],
        )

        input_columns = [
            column
            for (row, column), cell in cells_by_position.items()
            if row in {header_row, leaf_row}
            and normalized(cell) == "INPUT"
        ]
        if not input_columns:
            return None
        input_column = min(input_columns)
        numeric_rows = sorted(
            {
                row
                for (row, _column_value), cell in primary_by_position.items()
                if max(leaf_row + 1, min_owned_row)
                <= row
                <= panel_end_row
                and _numeric_value(cell) is not None
            }
        )
        if not numeric_rows:
            return None

        outcome_records: dict[
            tuple[int, bool], tuple[str, dict[str, Any]]
        ] = {}

        def outcome_for(
            column: int,
            *,
            is_share: bool,
            fallback_cell: dict[str, Any],
        ) -> str:
            cache_key = (column, is_share)
            existing = outcome_records.get(cache_key)
            if existing is not None:
                return existing[0]
            anchor = header_anchor(
                header_row=header_row,
                leaf_row=leaf_row,
                column=column,
            )
            parent = cells_by_position.get((header_row, column))
            if parent is None or not source_text(parent):
                parent = header_anchor(
                    header_row=header_row,
                    leaf_row=header_row,
                    column=column,
                )
            leaf = cells_by_position.get((leaf_row, column))
            label_parts = []
            for candidate in (parent, leaf):
                label = source_text(candidate)
                if label and label.casefold() not in {
                    value.casefold() for value in label_parts
                }:
                    label_parts.append(label)
            if not label_parts and anchor is not None:
                label_parts.append(source_text(anchor))
            if not label_parts:
                label_parts.append(f"Column {_column_label(column)}")
            label = " | ".join(label_parts)
            if is_share:
                label += " component share"
            normalized_label = " ".join(label.upper().split())
            if is_share:
                metric_type = "component_share"
                unit = "fraction"
                favorable_direction = "LOWER"
            elif "RATE" in normalized_label:
                metric_type = "rate"
                unit = "fraction"
                favorable_direction = "LOWER"
            elif normalized_label == "INPUT":
                metric_type = "count"
                unit = ""
                favorable_direction = "UNKNOWN"
            elif normalized_label == "OK":
                metric_type = "count"
                unit = ""
                favorable_direction = "HIGHER"
            elif any(
                token in normalized_label
                for token in ("NG", "NOISE", "TOUCH", "THD", "SIGMA")
            ):
                metric_type = "failure_count"
                unit = ""
                favorable_direction = "LOWER"
            else:
                metric_type = "numeric_measurement"
                unit = ""
                favorable_direction = "UNKNOWN"
            outcome_key = (
                f"function_r{title_row}_"
                f"c{column}_{'share' if is_share else 'value'}"
            )
            evidence_cell = anchor or fallback_cell
            record_value = append_record(
                logical_id=logical_id,
                record_type="ENTITY_DECLARATION",
                identity_keys=[key(evidence_cell)],
                exact_source_label=label,
                payload={
                    "entityType": "OUTCOME",
                    "key": outcome_key,
                    "originalLabel": label,
                    "metricType": metric_type,
                    "unit": unit,
                    "favorableDirection": favorable_direction,
                },
                evidence_items=[
                    evidence(
                        coordinate(evidence_cell),
                        source_text(evidence_cell) or label,
                    )
                ],
            )
            outcome_records[cache_key] = (outcome_key, record_value)
            return outcome_key

        current_arm_key = ""
        for row in numeric_rows:
            row_numeric_cells = sorted(
                (
                    column,
                    cell,
                )
                for (cell_row, column), cell in primary_by_position.items()
                if cell_row == row and _numeric_value(cell) is not None
            )
            condition_cells = sorted(
                (
                    column,
                    cell,
                )
                for (cell_row, column), cell in primary_by_position.items()
                if cell_row == row
                and column < input_column
                and source_text(cell)
                and not isinstance(source_value(cell), dict)
            )
            has_input = _numeric_value(
                primary_by_position.get((row, input_column), {})
            ) is not None
            is_base_row = bool(has_input or condition_cells)
            if is_base_row or not current_arm_key:
                identity_cell = (
                    condition_cells[0][1]
                    if condition_cells
                    else row_numeric_cells[0][1]
                )
                condition_label = " | ".join(
                    source_text(cell) for _column_value, cell in condition_cells
                )
                if not condition_label:
                    condition_label = f"Source row {row}"
                current_arm_key = f"function_r{title_row}_row_{row}"
                normalized_condition = condition_label.upper()
                role = (
                    "REFERENCE"
                    if len(condition_cells) == 1
                    and any(
                        token in normalized_condition
                        for token in ("NORMAL", "REFERENCE", "CONTROL")
                    )
                    else "TEST"
                )
                if condition_cells:
                    first_condition = condition_cells[0][1]
                    last_condition = condition_cells[-1][1]
                    condition_range = (
                        coordinate(first_condition)
                        if first_condition is last_condition
                        else (
                            f"{coordinate(first_condition)}:"
                            f"{coordinate(last_condition)}"
                        )
                    )
                else:
                    condition_range = coordinate(identity_cell)
                append_record(
                    logical_id=logical_id,
                    record_type="ENTITY_DECLARATION",
                    identity_keys=[key(identity_cell)],
                    exact_source_label=condition_label,
                    payload={
                        "entityType": "ARM",
                        "key": current_arm_key,
                        "role": role,
                        "label": condition_label,
                        "condition": condition_label,
                        "sampleSize": None,
                        "sampleBasis": "",
                        "matchingBasis": "",
                        "factorValues": [],
                    },
                    evidence_items=[
                        evidence(condition_range, condition_label)
                    ],
                )
            for column, value_cell in row_numeric_cells:
                numeric_value = _numeric_value(value_cell)
                if numeric_value is None:
                    continue
                outcome_key = outcome_for(
                    column,
                    is_share=not is_base_row,
                    fallback_cell=value_cell,
                )
                displayed_value = value_cell.get("displayValue")
                label = outcome_records[(column, not is_base_row)][1][
                    "exactSourceLabel"
                ]
                append_record(
                    logical_id=logical_id,
                    record_type="OBSERVATION_APPEND",
                    identity_keys=[key(value_cell)],
                    exact_source_label=label,
                    payload={
                        "outcome": outcome_key,
                        "arm": current_arm_key,
                        "valueNumber": numeric_value,
                        "valueText": str(
                            displayed_value
                            if displayed_value is not None
                            else source_value(value_cell)
                        ),
                        "numerator": None,
                        "denominator": None,
                        "ratePpm": None,
                        "min": None,
                        "max": None,
                        "average": None,
                        "sampleSize": None,
                        "replicateKey": (
                            f"source-{coordinate(value_cell).lower()}"
                        ),
                    },
                    evidence_items=[
                        evidence(
                            coordinate(value_cell),
                            str(
                                displayed_value
                                if displayed_value is not None
                                else source_value(value_cell)
                            ),
                        )
                    ],
                )

    _append_semantic_label_records_v2(
        envelope=envelope,
        focused_chunks=focused_chunks,
        logical_by_source_key={
            _cell_key(chunk, cell): str(panel["logicalStudyId"])
            for panel in panels
            for cell in panel["ownedCells"]
        },
        records=records,
    )
    evidence_by_record = {
        str(record["recordId"]): set(
            evidence_cell_keys(
                record["evidence"],
                chunks=focused_chunks,
            )
        )
        for record in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(record["recordId"])
            for record in records
            if source_key
            in evidence_by_record[str(record["recordId"])]
        ]
        if not record_ids:
            return None
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_result_table_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project mixed result/tension tables using locator-required values."""

    focused_chunks = envelope.get("focusedChunks")
    chunk = _coalesced_single_sheet_chunk(focused_chunks)
    if chunk is None:
        return None
    primary_cells = [
        cell
        for cell in chunk.get("cells", [])
        if isinstance(cell, dict)
    ]
    context_cells = [
        cell
        for cell in chunk.get("contextCells", [])
        if isinstance(cell, dict)
    ]
    if not primary_cells:
        return None
    source_keys = [_cell_key(chunk, cell) for cell in primary_cells]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None
    owned = set(source_keys)
    allowed = owned.union(
        str(value) for value in envelope.get("sharedAnchorCellKeys", [])
    )
    all_cells = [*context_cells, *primary_cells]
    by_key = {_cell_key(chunk, cell): cell for cell in all_cells}

    def source_value(cell: dict[str, Any]) -> object:
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    def source_text(cell: dict[str, Any] | None) -> str:
        if cell is None:
            return ""
        value = source_value(cell)
        if isinstance(value, dict):
            return str(value.get("value") or "").strip()
        return str(value if value is not None else "").strip()

    def normalized(cell: dict[str, Any] | None) -> str:
        return " ".join(source_text(cell).upper().split())

    # A safely split continuation part keeps the workbook/result title as
    # shared context while owning only the later table rows. Recognize that
    # exact context title so the deterministic projector covers continuation
    # numeric cells instead of sending an under-specified fragment to AI.
    labels = {normalized(cell) for cell in all_cells}
    has_result_title = any(
        re.search(r"(?:^|\s)RESULTS?(?:\s|$)", label)
        for label in labels
    )
    is_small_result_log = (
        {"NO", "DATE", "OK"}.issubset(labels)
        and any("TYPE" in label for label in labels)
    )
    if not has_result_title and not is_small_result_log:
        return None

    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or not studies:
        return None
    study_bounds: dict[str, tuple[int, int, int, int]] = {}
    for study in studies:
        logical_id = str(study.get("logicalStudyId") or "")
        anchor_cells = [
            by_key[str(value)]
            for value in study.get("anchorEvidenceCellKeys", [])
            if str(value) in by_key
        ]
        if not logical_id or not anchor_cells:
            return None
        positions = [_position(cell) for cell in anchor_cells]
        study_bounds[logical_id] = (
            min(row for row, _column_value in positions),
            min(column for _row, column in positions),
            max(row for row, _column_value in positions),
            max(column for _row, column in positions),
        )

    local_inventory = build_content_coverage_inventory(
        chunks=focused_chunks,
        locator_results=envelope.get("locatorResults", []),
        expected_source_cell_keys=envelope["ownedSourceCellKeys"],
    )
    required_numeric_keys = {
        str(item["sourceCellKey"])
        for item in local_inventory.get("requiredCells", [])
        if isinstance(item, dict)
    }
    assigned: dict[str, list[dict[str, Any]]] = {
        logical_id: [] for logical_id in study_bounds
    }
    for cell in primary_cells:
        row, column = _position(cell)
        ranked = []
        for logical_id, (
            min_row,
            min_column,
            max_row,
            max_column,
        ) in study_bounds.items():
            row_distance = (
                min_row - row
                if row < min_row
                else row - max_row if row > max_row else 0
            )
            column_distance = (
                min_column - column
                if column < min_column
                else column - max_column if column > max_column else 0
            )
            ranked.append(
                (
                    row_distance + column_distance,
                    row_distance,
                    column_distance,
                    (max_row - min_row + 1)
                    * (max_column - min_column + 1),
                    logical_id,
                )
            )
        ranked.sort()
        if (
            len(ranked) > 1
            and ranked[0][:4] == ranked[1][:4]
        ):
            source_key = _cell_key(chunk, cell)
            if source_key in required_numeric_keys:
                return None
            nearest = ranked[0][:4]
            for candidate in ranked:
                if candidate[:4] != nearest:
                    break
                assigned[candidate[4]].append(cell)
            continue
        assigned[ranked[0][4]].append(cell)
    if any(not cells for cells in assigned.values()):
        return None

    _sheet_index, sheet_title = _sheet(chunk)

    def key(cell: dict[str, Any]) -> str:
        return _cell_key(chunk, cell)

    def coordinate(cell: dict[str, Any]) -> str:
        return _coordinate(cell)

    def evidence(address: str, text: str) -> dict[str, Any]:
        return {
            "sheet": sheet_title,
            "range": address,
            "role": "SOURCE",
            "sourceText": text,
            "note": "",
        }

    records: list[dict[str, Any]] = []

    def append_record(
        *,
        logical_id: str,
        record_type: str,
        identity_cell: dict[str, Any],
        additional_identity_cells: Sequence[dict[str, Any]] = (),
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_items: Sequence[dict[str, Any]],
    ) -> dict[str, Any]:
        identity_keys = list(
            dict.fromkeys(
                [
                    key(identity_cell),
                    *(
                        key(cell)
                        for cell in additional_identity_cells
                    ),
                ]
            )
        )
        if (
            any(identity_key not in allowed for identity_key in identity_keys)
            or not exact_source_label.strip()
        ):
            raise StagedDraftV2Error(
                "Result-table projector produced invalid identity"
            )
        value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": identity_keys,
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": copy.deepcopy(list(evidence_items)),
        }
        value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=identity_keys,
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(value),
        )
        if record_type == "ENTITY_DECLARATION":
            value["payload"]["entityId"] = value["recordId"]
        records.append(value)
        return value

    for logical_id, study_cells in assigned.items():
        positions = [_position(cell) for cell in study_cells]
        min_row = min(row for row, _column_value in positions)
        min_column = min(column for _row, column in positions)
        max_row = max(row for row, _column_value in positions)
        max_column = max(column for _row, column in positions)
        study_range = (
            f"{_column_label(min_column)}{min_row}:"
            f"{_column_label(max_column)}{max_row}"
        )
        title_candidates = [
            cell
            for cell in all_cells
            if re.search(
                r"(?:^|\s)RESULTS?(?:\s|$)",
                normalized(cell),
            )
        ]
        if not title_candidates:
            title_candidates = [
                cell
                for cell in study_cells
                if source_text(cell)
                and _numeric_value(cell) is None
                and not isinstance(source_value(cell), dict)
            ]
        if not title_candidates:
            return None
        title_cell = min(title_candidates, key=_position)
        title = source_text(title_cell)
        append_record(
            logical_id=logical_id,
            record_type="STUDY_PATCH",
            identity_cell=title_cell,
            exact_source_label=title,
            payload={
                "title": title,
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "source result table",
                "comparisonBasis": "",
                "summary": (
                    "Each locator-required numeric result is preserved "
                    "against its source row and column label."
                ),
            },
            evidence_items=[evidence(study_range, title)],
        )

        required_cells = sorted(
            (
                cell
                for cell in study_cells
                if key(cell) in required_numeric_keys
                and _numeric_value(cell) is not None
            ),
            key=_position,
        )
        if not required_cells:
            continue
        first_required_row = min(_position(cell)[0] for cell in required_cells)
        header_area_cells = [
            cell
            for cell in all_cells
            if _position(cell)[0] < first_required_row
            and min_row <= _position(cell)[0] <= max_row
        ]
        outcomes: dict[
            tuple[int, bool], tuple[str, dict[str, Any]]
        ] = {}

        def header_for(
            column: int,
        ) -> dict[str, Any] | None:
            direct = [
                cell
                for cell in header_area_cells
                if _position(cell)[1] == column and source_text(cell)
            ]
            if direct:
                return max(direct, key=lambda cell: _position(cell)[0])
            merged = []
            for cell in header_area_cells:
                merge_range = str(cell.get("mergeRange") or "").strip()
                if not merge_range or not source_text(cell):
                    continue
                (
                    _range_start_row,
                    range_start_column,
                    _range_end_row,
                    range_end_column,
                ) = range_bounds(merge_range)
                if range_start_column <= column <= range_end_column:
                    merged.append(cell)
            return max(merged, key=lambda cell: _position(cell)[0]) if merged else None

        def outcome_for(
            *,
            column: int,
            is_share: bool,
            value_cell: dict[str, Any],
        ) -> str:
            cache_key = (column, is_share)
            if cache_key in outcomes:
                return outcomes[cache_key][0]
            header = header_for(column)
            header_label = source_text(header)
            if not header_label:
                label = f"{title} | Column {_column_label(column)}"
            elif _numeric_value(header or {}) is not None:
                label = f"{title} | Sample {header_label}"
            else:
                label = header_label
            if is_share:
                label += " component share"
            normalized_label = " ".join(label.upper().split())
            metric_type = (
                "component_share"
                if is_share
                else (
                    "rate"
                    if "RATE" in normalized_label
                    else (
                        "tension"
                        if "TENSION" in " ".join(title.upper().split())
                        else "numeric_measurement"
                    )
                )
            )
            unit = (
                "fraction"
                if is_share or "RATE" in normalized_label
                else (
                    "kgf"
                    if "TENSION" in " ".join(title.upper().split())
                    else ""
                )
            )
            favorable_direction = (
                "LOWER"
                if is_share
                or any(
                    token in normalized_label
                    for token in ("NG", "RATE", "NOISE", "TOUCH")
                )
                else "UNKNOWN"
            )
            outcome_key = (
                f"result_c{column}_"
                f"{'share' if is_share else 'value'}"
            )
            identity_cell = header or title_cell
            record_value = append_record(
                logical_id=logical_id,
                record_type="ENTITY_DECLARATION",
                identity_cell=identity_cell,
                additional_identity_cells=[value_cell],
                exact_source_label=label,
                payload={
                    "entityType": "OUTCOME",
                    "key": outcome_key,
                    "originalLabel": label,
                    "metricType": metric_type,
                    "unit": unit,
                    "favorableDirection": favorable_direction,
                },
                evidence_items=[
                    evidence(
                        coordinate(identity_cell),
                        source_text(identity_cell),
                    )
                ],
            )
            outcomes[cache_key] = (outcome_key, record_value)
            return outcome_key

        current_arm_key = ""
        for row in sorted({_position(cell)[0] for cell in required_cells}):
            row_values = [
                cell
                for cell in required_cells
                if _position(cell)[0] == row
            ]
            first_value_column = min(
                _position(cell)[1] for cell in row_values
            )
            row_label_cells = sorted(
                (
                    cell
                    for cell in study_cells
                    if _position(cell)[0] == row
                    and _position(cell)[1] < first_value_column
                    and source_text(cell)
                    and _numeric_value(cell) is None
                    and not isinstance(source_value(cell), dict)
                ),
                key=_position,
            )
            is_share = not row_label_cells and bool(current_arm_key)
            if not current_arm_key or not is_share:
                identity_cell = (
                    row_label_cells[0]
                    if row_label_cells
                    else row_values[0]
                )
                arm_label = " | ".join(
                    source_text(cell) for cell in row_label_cells
                )
                if not arm_label:
                    arm_label = f"Source row {row}"
                current_arm_key = f"result_row_{row}"
                append_record(
                    logical_id=logical_id,
                    record_type="ENTITY_DECLARATION",
                    identity_cell=identity_cell,
                    exact_source_label=arm_label,
                    payload={
                        "entityType": "ARM",
                        "key": current_arm_key,
                        "role": (
                            "REFERENCE"
                            if len(row_label_cells) == 1
                            and any(
                                token in arm_label.upper()
                                for token in (
                                    "NORMAL",
                                    "REFERENCE",
                                    "CONTROL",
                                )
                            )
                            else "TEST"
                        ),
                        "label": arm_label,
                        "condition": arm_label,
                        "sampleSize": None,
                        "sampleBasis": "",
                        "matchingBasis": "",
                        "factorValues": [],
                    },
                    evidence_items=[
                        evidence(
                            (
                                f"{coordinate(row_label_cells[0])}:"
                                f"{coordinate(row_label_cells[-1])}"
                                if len(row_label_cells) > 1
                                else coordinate(identity_cell)
                            ),
                            arm_label,
                        )
                    ],
                )
            for value_cell in row_values:
                column = _position(value_cell)[1]
                numeric_value = _numeric_value(value_cell)
                if numeric_value is None:
                    continue
                outcome_key = outcome_for(
                    column=column,
                    is_share=is_share,
                    value_cell=value_cell,
                )
                displayed_value = value_cell.get("displayValue")
                label = outcomes[(column, is_share)][1][
                    "exactSourceLabel"
                ]
                append_record(
                    logical_id=logical_id,
                    record_type="OBSERVATION_APPEND",
                    identity_cell=value_cell,
                    exact_source_label=label,
                    payload={
                        "outcome": outcome_key,
                        "arm": current_arm_key,
                        "valueNumber": numeric_value,
                        "valueText": str(
                            displayed_value
                            if displayed_value is not None
                            else source_value(value_cell)
                        ),
                        "numerator": None,
                        "denominator": None,
                        "ratePpm": None,
                        "min": None,
                        "max": None,
                        "average": None,
                        "sampleSize": None,
                        "replicateKey": (
                            f"source-{coordinate(value_cell).lower()}"
                        ),
                    },
                    evidence_items=[
                        evidence(
                            coordinate(value_cell),
                            str(
                                displayed_value
                                if displayed_value is not None
                                else source_value(value_cell)
                            ),
                        )
                    ],
                )

    _append_semantic_label_records_v2(
        envelope=envelope,
        focused_chunks=focused_chunks,
        logical_by_source_key={
            _cell_key(chunk, cell): logical_id
            for logical_id, study_cells in assigned.items()
            for cell in study_cells
        },
        records=records,
        inventory=local_inventory,
    )
    evidence_by_record = {
        str(record["recordId"]): set(
            evidence_cell_keys(
                record["evidence"],
                chunks=focused_chunks,
            )
        )
        for record in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(record["recordId"])
            for record in records
            if source_key
            in evidence_by_record[str(record["recordId"])]
        ]
        if not record_ids:
            return None
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_nti_horizontal_matrix_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project exact NTI row-sample matrices with header frequency axes."""

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or len(focused_chunks) != 1:
        return None
    chunk = focused_chunks[0]
    if not isinstance(chunk, dict):
        return None
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None
    _sheet_index, sheet_title = _sheet(chunk)
    sheet_key = " ".join(sheet_title.upper().split())
    metric_specs = {
        "NTI FREQUENCY": (
            "nti_frequency_response",
            "Frequency response [dBSPL]",
            "frequency_response",
            "dBSPL",
        ),
        "NTI IMPEDANCE": (
            "nti_impedance_response",
            "Impedance response",
            "impedance_response",
            "",
        ),
        "NTI THD": (
            "nti_thd_response",
            "Dist#1 Harmonics 2|3 [%]",
            "distortion_response",
            "%",
        ),
    }
    metric_spec = metric_specs.get(sheet_key)
    if metric_spec is None:
        return None
    outcome_key, outcome_label, metric_type, value_unit = metric_spec
    metric_key = outcome_key.removesuffix("_response")

    primary_cells = [
        cell
        for cell in chunk.get("cells", [])
        if isinstance(cell, dict)
    ]
    context_cells = [
        cell
        for cell in chunk.get("contextCells", [])
        if isinstance(cell, dict)
    ]
    cells_by_position = {
        (int(cell.get("row") or 0), int(cell.get("column") or 0)): cell
        for cell in [*context_cells, *primary_cells]
        if int(cell.get("row") or 0) > 0
        and int(cell.get("column") or 0) > 0
    }
    primary_positions = {
        (int(cell.get("row") or 0), int(cell.get("column") or 0))
        for cell in primary_cells
    }
    first_part = (1, 1) in primary_positions
    data_rows = sorted(
        {
            row
            for row, column in primary_positions
            if 2 <= row <= 12 and column == 2
        }
    )
    if not data_rows or data_rows != list(
        range(data_rows[0], data_rows[-1] + 1)
    ):
        return None
    expected_positions = {
        (row, column)
        for row in data_rows
        for column in range(2, 91)
    }
    if first_part:
        expected_positions.update((1, column) for column in range(1, 91))
    if primary_positions != expected_positions:
        return None
    required_context_positions = {
        (1, column) for column in range(2, 91)
    }
    if not required_context_positions.issubset(cells_by_position):
        return None

    def source_value(cell: dict[str, Any]) -> object:
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    def source_text(cell: dict[str, Any]) -> str:
        return str(source_value(cell) or "").strip()

    if source_text(cells_by_position[(1, 2)]).upper() != "SAMPLE":
        return None
    if source_text(cells_by_position[(1, 90)]).upper() != "RESULT":
        return None
    if first_part and source_text(
        cells_by_position[(1, 1)]
    ).casefold() != outcome_label.casefold():
        return None
    frequency_headers: list[dict[str, Any]] = []
    for column in range(3, 90):
        header = cells_by_position.get((1, column))
        if (
            header is None
            or re.fullmatch(
                r"\d+(?:\.\d+)?\s*HZ",
                source_text(header),
                re.IGNORECASE,
            )
            is None
        ):
            return None
        frequency_headers.append(header)
    for row in data_rows:
        identity = cells_by_position.get((row, 2))
        status = cells_by_position.get((row, 90))
        if identity is None or status is None or not source_text(identity):
            return None
        if source_text(status).upper() not in {"PASSED", "PASS"}:
            return None
        if any(
            (
                value_cell := cells_by_position.get((row, column))
            )
            is None
            or _numeric_value(value_cell) is None
            for column in range(3, 90)
        ):
            return None
    source_keys = [_cell_key(chunk, cell) for cell in primary_cells]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None
    source_order = {
        key: index for index, key in enumerate(source_keys)
    }

    def key_at(row: int, column: int) -> str:
        return _cell_key(chunk, cells_by_position[(row, column)])

    def address(row: int, start_column: int, end_column: int) -> str:
        start = f"{_column_label(start_column)}{row}"
        end = f"{_column_label(end_column)}{row}"
        return start if start == end else f"{start}:{end}"

    def evidence(address_value: str, text: str) -> dict[str, Any]:
        return {
            "sheet": sheet_title,
            "range": address_value,
            "role": "SOURCE",
            "sourceText": text,
            "note": "",
        }

    def record(
        *,
        record_type: str,
        identity_keys: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_items: Sequence[dict[str, Any]],
    ) -> dict[str, Any]:
        value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": list(identity_keys),
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": copy.deepcopy(list(evidence_items)),
        }
        value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=list(identity_keys),
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(value),
        )
        if record_type == "ENTITY_DECLARATION":
            value["payload"]["entityId"] = value["recordId"]
        return value

    records: list[dict[str, Any]] = []
    if first_part:
        title_key = key_at(1, 1)
        records.extend(
            [
                record(
                    record_type="STUDY_PATCH",
                    identity_keys=[title_key],
                    exact_source_label=outcome_label,
                    payload={
                        "title": outcome_label,
                        "purpose": "",
                        "hypothesis": "",
                        "objective": "",
                        "designType": "sample frequency-response matrix",
                        "comparisonBasis": "",
                        "summary": (
                            "The source Spec profile and ten numbered sample "
                            "profiles are preserved across 87 exact frequency "
                            "columns with their source result statuses."
                        ),
                    },
                    evidence_items=[
                        evidence("A1:CL4", outcome_label)
                    ],
                ),
                record(
                    record_type="ENTITY_DECLARATION",
                    identity_keys=[title_key],
                    exact_source_label=outcome_label,
                    payload={
                        "entityType": "OUTCOME",
                        "key": outcome_key,
                        "originalLabel": outcome_label,
                        "metricType": metric_type,
                        "unit": value_unit,
                        "favorableDirection": "UNKNOWN",
                    },
                    evidence_items=[
                        evidence("A1:CK1", outcome_label)
                    ],
                ),
                record(
                    record_type="ENTITY_DECLARATION",
                    identity_keys=[key_at(1, 90)],
                    exact_source_label="Result",
                    payload={
                        "entityType": "OUTCOME",
                        "key": f"{metric_key}_result_status",
                        "originalLabel": "Result",
                        "metricType": "categorical_status",
                        "unit": "",
                        "favorableDirection": "NONE",
                    },
                    evidence_items=[
                        evidence("CL1", "Result")
                    ],
                ),
            ]
        )

    for row in data_rows:
        identity_cell = cells_by_position[(row, 2)]
        identity_text = source_text(identity_cell)
        is_spec = identity_text.upper() == "SPEC"
        if is_spec:
            arm_key = f"{metric_key}_spec"
            arm_label = "Spec"
        else:
            numeric_identity = _numeric_value(identity_cell)
            if (
                numeric_identity is None
                or not float(numeric_identity).is_integer()
            ):
                return None
            sample_number = int(numeric_identity)
            if not 1 <= sample_number <= 10:
                return None
            arm_key = f"{metric_key}_sample_{sample_number}"
            arm_label = f"Sample {sample_number}"
        records.append(
            record(
                record_type="ENTITY_DECLARATION",
                identity_keys=[key_at(row, 2)],
                exact_source_label=identity_text,
                payload={
                    "entityType": "ARM",
                    "key": arm_key,
                    "role": "OTHER",
                    "label": arm_label,
                    "condition": arm_label,
                    "sampleSize": None,
                    "sampleBasis": "",
                    "matchingBasis": "",
                    "factorValues": [],
                },
                evidence_items=[
                    evidence(address(row, 2, 2), identity_text)
                ],
            )
        )
        if is_spec:
            for column, header in zip(
                range(3, 90),
                frequency_headers,
            ):
                value_cell = cells_by_position[(row, column)]
                numeric_value = _numeric_value(value_cell)
                if numeric_value is None:
                    return None
                displayed_value = value_cell.get("displayValue")
                records.append(
                    record(
                        record_type="OBSERVATION_APPEND",
                        identity_keys=[key_at(row, column)],
                        exact_source_label=source_text(header),
                        payload={
                            "outcome": outcome_key,
                            "arm": arm_key,
                            "valueNumber": numeric_value,
                            "valueText": str(
                                displayed_value
                                if displayed_value is not None
                                else source_value(value_cell)
                            ),
                            "numerator": None,
                            "denominator": None,
                            "ratePpm": None,
                            "min": None,
                            "max": None,
                            "average": None,
                            "sampleSize": None,
                            "replicateKey": (
                                f"frequency-"
                                f"{_column_label(column).lower()}"
                            ),
                        },
                        evidence_items=[
                            evidence(
                                address(1, column, column),
                                source_text(header),
                            ),
                            evidence(
                                address(row, column, column),
                                str(
                                    displayed_value
                                    if displayed_value is not None
                                    else source_value(value_cell)
                                ),
                            ),
                        ],
                    )
                )
        else:
            records.append(
                record(
                    record_type="SERIES_SEGMENT_APPEND",
                    identity_keys=[key_at(row, 2), key_at(row, 3)],
                    exact_source_label=arm_label,
                    payload={
                        "key": f"{arm_key}_series",
                        "seriesRole": "RAW",
                        "aggregationFunction": "",
                        "aggregateOfSeries": [],
                        "outcome": outcome_key,
                        "arm": arm_key,
                        "sheet": sheet_title,
                        "headerRange": "C1:CK1",
                        "valueRange": address(row, 3, 89),
                        "rowIdentityRange": address(row, 2, 2),
                        "aggregateReplicateRanges": [],
                        "axisSource": "HEADER",
                        "axisLabel": "Frequency",
                        "axisUnit": "Hz",
                        "valueUnit": value_unit,
                        "stratumKey": "",
                        "verificationStatus": "NEEDS_REVIEW",
                    },
                    evidence_items=[
                        evidence("C1:CK1", "Frequency headers"),
                        evidence(
                            address(row, 2, 89),
                            arm_label,
                        ),
                    ],
                )
            )
        status_cell = cells_by_position[(row, 90)]
        status_text = source_text(status_cell)
        records.append(
            record(
                record_type="OBSERVATION_APPEND",
                identity_keys=[key_at(row, 90)],
                exact_source_label="Result",
                payload={
                    "outcome": f"{metric_key}_result_status",
                    "arm": arm_key,
                    "valueNumber": None,
                    "valueText": status_text,
                    "numerator": None,
                    "denominator": None,
                    "ratePpm": None,
                    "min": None,
                    "max": None,
                    "average": None,
                    "sampleSize": None,
                    "replicateKey": f"source-cl{row}",
                },
                evidence_items=[
                    evidence(address(row, 90, 90), status_text)
                ],
            )
        )

    evidence_by_record = {
        str(item["recordId"]): set(
            evidence_cell_keys(
                item["evidence"],
                chunks=focused_chunks,
            )
        )
        for item in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(item["recordId"])
            for item in records
            if source_key in evidence_by_record[str(item["recordId"])]
        ]
        if not record_ids:
            return None
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": sorted(
            records,
            key=lambda item: (
                min(
                    (
                        source_order.get(key, 10**12)
                        for key in item["identityCellKeys"]
                    ),
                    default=10**12,
                ),
                item["recordId"],
            ),
        ),
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_nti_f0_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project the exact A1:E12 NTI resonance-frequency table."""

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or len(focused_chunks) != 1:
        return None
    chunk = focused_chunks[0]
    if not isinstance(chunk, dict) or envelope.get("sharedAnchorCellKeys"):
        return None
    _sheet_index, sheet_title = _sheet(chunk)
    if " ".join(sheet_title.upper().split()) != "NTI F0":
        return None
    if str(chunk.get("primaryRange") or "").upper() != "A1:E12":
        return None
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None
    expected_coordinates = [
        *[f"{column}1" for column in "ABCDE"],
        *[
            f"{column}{row}"
            for row in range(2, 13)
            for column in "BCDE"
        ],
    ]
    cells = chunk.get("cells")
    if not isinstance(cells, list):
        return None
    cells_by_coordinate = {
        _coordinate(cell): cell
        for cell in cells
        if isinstance(cell, dict)
    }
    if set(cells_by_coordinate) != set(expected_coordinates):
        return None
    source_keys = [
        _cell_key(chunk, cells_by_coordinate[coordinate])
        for coordinate in expected_coordinates
    ]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None

    def source_value(coordinate: str) -> object:
        cell = cells_by_coordinate[coordinate]
        return (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )

    def source_text(coordinate: str) -> str:
        return str(source_value(coordinate) or "").strip()

    expected_headers = {
        "A1": "RESONANCE FREQUENCIES",
        "B1": "SAMPLE",
        "C1": "F0[1]",
        "D1": "RESULT F0[1]",
        "E1": "OVERALL RESONANCE FREQ. RESULT",
    }
    if any(
        " ".join(source_text(coordinate).upper().split()) != value
        for coordinate, value in expected_headers.items()
    ):
        return None
    if source_text("B2").upper() != "SPEC":
        return None
    if _numeric_value(cells_by_coordinate["C2"]) is None:
        return None
    for row in range(2, 13):
        if any(
            source_text(f"{column}{row}").upper()
            not in {"PASSED", "PASS"}
            for column in "DE"
        ):
            return None
    for row in range(3, 13):
        if (
            _numeric_value(cells_by_coordinate[f"B{row}"]) is None
            or _numeric_value(cells_by_coordinate[f"C{row}"]) is None
        ):
            return None
    key_by_coordinate = dict(zip(expected_coordinates, source_keys))

    def evidence(address: str, text: str) -> list[dict[str, Any]]:
        return [
            {
                "sheet": sheet_title,
                "range": address,
                "role": "SOURCE",
                "sourceText": text,
                "note": "",
            }
        ]

    def record(
        *,
        record_type: str,
        identity_coordinates: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_range: str,
        evidence_text: str,
    ) -> dict[str, Any]:
        identity_keys = [
            key_by_coordinate[coordinate]
            for coordinate in identity_coordinates
        ]
        value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": identity_keys,
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": evidence(evidence_range, evidence_text),
        }
        value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=identity_keys,
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(value),
        )
        if record_type == "ENTITY_DECLARATION":
            value["payload"]["entityId"] = value["recordId"]
        return value

    records = [
        record(
            record_type="STUDY_PATCH",
            identity_coordinates=["A1"],
            exact_source_label=source_text("A1"),
            payload={
                "title": source_text("A1"),
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "sample resonance-frequency table",
                "comparisonBasis": "",
                "summary": (
                    "The source Spec resonance frequency, ten-sample F0 "
                    "series, and both exact result-status columns are "
                    "preserved."
                ),
            },
            evidence_range="A1:E12",
            evidence_text=source_text("A1"),
        ),
        *[
            record(
                record_type="ENTITY_DECLARATION",
                identity_coordinates=[coordinate],
                exact_source_label=label,
                payload={
                    "entityType": "OUTCOME",
                    "key": key,
                    "originalLabel": label,
                    "metricType": metric_type,
                    "unit": unit,
                    "favorableDirection": direction,
                },
                evidence_range=coordinate,
                evidence_text=source_text(coordinate),
            )
            for coordinate, key, label, metric_type, unit, direction in (
                (
                    "C1",
                    "nti_f0_frequency",
                    "f0[1]",
                    "resonance_frequency",
                    "Hz",
                    "UNKNOWN",
                ),
                (
                    "D1",
                    "nti_f0_result_status",
                    "Result f0[1]",
                    "categorical_status",
                    "",
                    "NONE",
                ),
                (
                    "E1",
                    "nti_f0_overall_status",
                    "Overall Resonance Freq. Result",
                    "categorical_status",
                    "",
                    "NONE",
                ),
            )
        ],
        *[
            record(
                record_type="ENTITY_DECLARATION",
                identity_coordinates=[coordinate],
                exact_source_label=label,
                payload={
                    "entityType": "ARM",
                    "key": key,
                    "role": "OTHER",
                    "label": label,
                    "condition": label,
                    "sampleSize": None,
                    "sampleBasis": "",
                    "matchingBasis": "",
                    "factorValues": [],
                },
                evidence_range=evidence_range,
                evidence_text=label,
            )
            for coordinate, key, label, evidence_range in (
                ("B2", "nti_f0_spec", "Spec", "B2"),
                ("B1", "nti_f0_samples", "Samples 1-10", "B1:B12"),
            )
        ],
    ]
    spec_value = _numeric_value(cells_by_coordinate["C2"])
    if spec_value is None:
        return None
    records.append(
        record(
            record_type="OBSERVATION_APPEND",
            identity_coordinates=["C2"],
            exact_source_label="Spec f0[1]",
            payload={
                "outcome": "nti_f0_frequency",
                "arm": "nti_f0_spec",
                "valueNumber": spec_value,
                "valueText": source_text("C2"),
                "numerator": None,
                "denominator": None,
                "ratePpm": None,
                "min": None,
                "max": None,
                "average": None,
                "sampleSize": None,
                "replicateKey": "source-c2",
            },
            evidence_range="C2",
            evidence_text=source_text("C2"),
        )
    )
    records.append(
        record(
            record_type="SERIES_SEGMENT_APPEND",
            identity_coordinates=["B3", "C3"],
            exact_source_label="Samples 1-10 f0[1]",
            payload={
                "key": "nti_f0_sample_series",
                "seriesRole": "RAW",
                "aggregationFunction": "",
                "aggregateOfSeries": [],
                "outcome": "nti_f0_frequency",
                "arm": "nti_f0_samples",
                "sheet": sheet_title,
                "headerRange": "C1",
                "valueRange": "C3:C12",
                "rowIdentityRange": "B3:B12",
                "aggregateReplicateRanges": [],
                "axisSource": "ROW_IDENTITY",
                "axisLabel": "Sample",
                "axisUnit": "",
                "valueUnit": "Hz",
                "stratumKey": "",
                "verificationStatus": "NEEDS_REVIEW",
            },
            evidence_range="B1:C12",
            evidence_text="Samples 1-10 f0[1]",
        )
    )
    for row in range(2, 13):
        arm_key = "nti_f0_spec" if row == 2 else "nti_f0_samples"
        for column, outcome_key in (
            ("D", "nti_f0_result_status"),
            ("E", "nti_f0_overall_status"),
        ):
            coordinate = f"{column}{row}"
            records.append(
                record(
                    record_type="OBSERVATION_APPEND",
                    identity_coordinates=[coordinate],
                    exact_source_label=source_text(f"{column}1"),
                    payload={
                        "outcome": outcome_key,
                        "arm": arm_key,
                        "valueNumber": None,
                        "valueText": source_text(coordinate),
                        "numerator": None,
                        "denominator": None,
                        "ratePpm": None,
                        "min": None,
                        "max": None,
                        "average": None,
                        "sampleSize": None,
                        "replicateKey": f"source-{coordinate.lower()}",
                    },
                    evidence_range=coordinate,
                    evidence_text=source_text(coordinate),
                )
            )

    evidence_by_record = {
        str(item["recordId"]): set(
            evidence_cell_keys(
                item["evidence"],
                chunks=focused_chunks,
            )
        )
        for item in records
    }
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(item["recordId"])
            for item in records
            if source_key in evidence_by_record[str(item["recordId"])]
        ]
        if not record_ids:
            return None
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_error_axis_tail_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Preserve an exact two-row numeric axis with unresolved result errors."""

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or len(focused_chunks) != 1:
        return None
    chunk = focused_chunks[0]
    if not isinstance(chunk, dict) or envelope.get("sharedAnchorCellKeys"):
        return None
    _sheet_index, sheet_title = _sheet(chunk)
    if not " ".join(sheet_title.upper().split()).startswith("SPL DATA_(NTI"):
        return None
    if str(chunk.get("primaryRange") or "").upper() != "A175:B176":
        return None
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None

    expected_coordinates = ["A175", "B175", "A176", "B176"]
    cells = chunk.get("cells")
    if not isinstance(cells, list):
        return None
    cells_by_coordinate = {
        _coordinate(cell): cell
        for cell in cells
        if isinstance(cell, dict)
    }
    if set(cells_by_coordinate) != set(expected_coordinates):
        return None
    source_keys = [
        _cell_key(chunk, cells_by_coordinate[coordinate])
        for coordinate in expected_coordinates
    ]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None
    axis_values = [
        _numeric_value(cells_by_coordinate[coordinate])
        for coordinate in ("A175", "A176")
    ]
    if (
        any(value is None or value < 1000 for value in axis_values)
        or axis_values[1] <= axis_values[0]
        or any(
            str(
                cells_by_coordinate[coordinate].get("rawValue") or ""
            ).strip().upper()
            != "#REF!"
            for coordinate in ("B175", "B176")
        )
    ):
        return None

    identity_keys = [
        _cell_key(chunk, cells_by_coordinate["B175"]),
        _cell_key(chunk, cells_by_coordinate["B176"]),
    ]
    exact_label = "Unresolved SPL reference-error rows"
    record = {
        "recordType": "LIMITATION_APPEND",
        "recordId": "",
        "logicalStudyId": logical_id,
        "identityCellKeys": identity_keys,
        "exactSourceLabel": exact_label,
        "payload": {
            "text": (
                "Source SPL rows 19000 and 20000 contain unresolved #REF! "
                "results; the numeric cells are retained only as row "
                "identities."
            ),
            "scope": "STUDY",
        },
        "evidence": [
            {
                "sheet": sheet_title,
                "range": "A175:B176",
                "role": "SOURCE",
                "sourceText": "19000 #REF!; 20000 #REF!",
                "note": "",
            }
        ],
    }
    record["recordId"] = stable_record_id(
        revision_uid=str(envelope["source"]["revisionUid"]),
        logical_study_id=logical_id,
        record_type="LIMITATION_APPEND",
        identity_cell_keys=identity_keys,
        exact_source_label=exact_label,
        semantic_subtype=_record_semantic_subtype(record),
    )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": [record],
        "coverageDispositions": [
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": [str(record["recordId"])],
                "reason": "",
            }
            for source_key in source_keys
        ],
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def build_deterministic_acoustic_matrix_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Project exact SPL/THD/IMP frequency matrices as source-bound series."""

    focused_chunks = envelope.get("focusedChunks")
    if not isinstance(focused_chunks, list) or not focused_chunks:
        return None
    if any(not isinstance(item, dict) for item in focused_chunks):
        return None
    chunk = focused_chunks[0]
    registry = envelope.get("registry")
    studies = registry.get("studies") if isinstance(registry, dict) else None
    if not isinstance(studies, list) or len(studies) != 1:
        return None
    logical_id = str(studies[0].get("logicalStudyId") or "")
    if not logical_id:
        return None

    _sheet_index, sheet_title = _sheet(chunk)
    if any(_sheet(item)[1] != sheet_title for item in focused_chunks):
        return None
    normalized_sheet = " ".join(sheet_title.upper().split())
    if normalized_sheet.startswith("SPL DATA_(NTI"):
        metric_key, metric_label = "spl", "SPL"
    elif normalized_sheet.startswith("THD DATA_(NTI"):
        metric_key, metric_label = "thd", "THD"
    elif normalized_sheet.startswith("IMP DATA_(NTI"):
        metric_key, metric_label = "imp", "IMP"
    else:
        return None

    all_cells = [
        cell
        for focused_chunk in focused_chunks
        for collection in ("cells", "contextCells")
        for cell in focused_chunk.get(collection, [])
        if isinstance(cell, dict)
    ]
    cells_by_position = {
        (int(cell.get("row") or 0), int(cell.get("column") or 0)): cell
        for cell in all_cells
        if int(cell.get("row") or 0) > 0
        and int(cell.get("column") or 0) > 0
    }
    primary_cells = [
        cell
        for focused_chunk in focused_chunks
        for cell in focused_chunk.get("cells", [])
        if isinstance(cell, dict)
    ]
    primary_by_position = {
        (int(cell.get("row") or 0), int(cell.get("column") or 0)): cell
        for cell in primary_cells
    }
    data_rows = sorted(
        {
            int(cell.get("row") or 0)
            for cell in primary_cells
            if int(cell.get("column") or 0) == 1
            and _numeric_value(cell) is not None
            and sum(
                1
                for (candidate_row, candidate_column), candidate_cell
                in primary_by_position.items()
                if candidate_row == int(cell.get("row") or 0)
                and candidate_column > 1
                and _numeric_value(candidate_cell) is not None
            )
            >= 2
        }
    )
    if (
        len(data_rows) < 2
        or data_rows
        != list(range(data_rows[0], data_rows[-1] + 1))
    ):
        return None
    start_row, end_row = data_rows[0], data_rows[-1]
    allowed_keys = {
        str(value)
        for field in ("ownedSourceCellKeys", "sharedAnchorCellKeys")
        for value in envelope.get(field, [])
    }
    owned_keys = {
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    }
    measurement_columns: list[tuple[int, dict[str, Any]]] = []
    for (header_row, column), header_cell in sorted(
        cells_by_position.items()
    ):
        if header_row != 2 or column <= 1:
            continue
        header_text = str(
            header_cell.get("rawValue")
            if not str(header_cell.get("formula") or "").strip()
            else header_cell.get("cachedValue")
            or ""
        ).strip()
        if not header_text:
            continue
        values = [
            primary_by_position.get((row, column)) for row in data_rows
        ]
        if any(
            value is None or _numeric_value(value) is None
            for value in values
        ):
            continue
        if _cell_key(chunk, header_cell) not in allowed_keys:
            continue
        measurement_columns.append((column, header_cell))
    if len(measurement_columns) < 20:
        return None

    expected_numeric_keys = {
        _cell_key(chunk, cell)
        for cell in primary_cells
        if _numeric_value(cell) is not None
    }
    represented_numeric_keys = {
        _cell_key(chunk, primary_by_position[(row, 1)])
        for row in data_rows
    }
    represented_numeric_keys.update(
        _cell_key(chunk, primary_by_position[(row, column)])
        for column, _header in measurement_columns
        for row in data_rows
    )
    measurement_column_numbers = {
        column for column, _header in measurement_columns
    }
    tail_numeric_cells = [
        cell
        for cell in primary_cells
        if _numeric_value(cell) is not None
        and _cell_key(chunk, cell) not in represented_numeric_keys
        and int(cell.get("column") or 0) in measurement_column_numbers
    ]
    represented_numeric_keys.update(
        _cell_key(chunk, cell) for cell in tail_numeric_cells
    )
    if expected_numeric_keys != represented_numeric_keys:
        return None

    def coordinate(cell: dict[str, Any]) -> str:
        return str(cell.get("coordinate") or "").upper()

    def header_text(cell: dict[str, Any]) -> str:
        value = (
            cell.get("cachedValue")
            if str(cell.get("formula") or "").strip()
            else cell.get("rawValue")
        )
        return str(value or "").strip()

    def evidence(address: str, text: str) -> dict[str, Any]:
        return {
            "sheet": sheet_title,
            "range": address,
            "role": "SOURCE",
            "sourceText": text,
            "note": "",
        }

    def record(
        *,
        record_type: str,
        identity_keys: Sequence[str],
        exact_source_label: str,
        payload: dict[str, Any],
        evidence_items: Sequence[dict[str, Any]],
    ) -> dict[str, Any]:
        value = {
            "recordType": record_type,
            "recordId": "",
            "logicalStudyId": logical_id,
            "identityCellKeys": list(identity_keys),
            "exactSourceLabel": exact_source_label,
            "payload": copy.deepcopy(payload),
            "evidence": copy.deepcopy(list(evidence_items)),
        }
        value["recordId"] = stable_record_id(
            revision_uid=str(envelope["source"]["revisionUid"]),
            logical_study_id=logical_id,
            record_type=record_type,
            identity_cell_keys=list(identity_keys),
            exact_source_label=exact_source_label,
            semantic_subtype=_record_semantic_subtype(value),
        )
        if record_type == "ENTITY_DECLARATION":
            value["payload"]["entityId"] = value["recordId"]
        return value

    first_column, first_header = measurement_columns[0]
    first_header_key = _cell_key(chunk, first_header)
    primary_ranges = [
        str(item.get("primaryRange") or "")
        for item in focused_chunks
    ]
    if any(not value for value in primary_ranges):
        return None
    outcome_key = f"{metric_key}_frequency_result"
    records = [
        record(
            record_type="STUDY_PATCH",
            identity_keys=[first_header_key],
            exact_source_label=header_text(first_header),
            payload={
                "title": f"{metric_label} frequency data",
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "source frequency-series matrix",
                "comparisonBasis": "",
                "summary": (
                    f"Source {metric_label} values are preserved by exact "
                    "measurement column and numeric row identity."
                ),
            },
            evidence_items=[
                evidence(
                    address,
                    f"{metric_label} source matrix segment",
                )
                for address in primary_ranges
            ],
        ),
        record(
            record_type="ENTITY_DECLARATION",
            identity_keys=[first_header_key],
            exact_source_label=header_text(first_header),
            payload={
                "entityType": "OUTCOME",
                "key": outcome_key,
                "originalLabel": metric_label,
                "metricType": "frequency_series",
                "unit": "",
                "favorableDirection": "UNKNOWN",
            },
            evidence_items=[
                evidence(
                    coordinate(first_header),
                    header_text(first_header),
                )
            ],
        ),
    ]

    axis_start = primary_by_position[(start_row, 1)]
    axis_end = primary_by_position[(end_row, 1)]
    axis_range = (
        f"{coordinate(axis_start)}:{coordinate(axis_end)}"
        if start_row != end_row
        else coordinate(axis_start)
    )
    axis_start_key = _cell_key(chunk, axis_start)
    for column, header_cell in measurement_columns:
        source_label = header_text(header_cell)
        normalized_label = " ".join(source_label.upper().split())
        aggregate_match = re.fullmatch(
            r"(.+?)\s*AVG\.?",
            normalized_label,
        )
        aggregate_source_series: list[str] = []
        if aggregate_match is not None:
            aggregate_group = re.sub(
                r"[^A-Z0-9]",
                "",
                aggregate_match.group(1),
            )
            for candidate_column, candidate_header in measurement_columns:
                candidate_label = " ".join(
                    header_text(candidate_header).upper().split()
                )
                replicate_match = re.fullmatch(
                    r"(.+?)\s*#\s*\d+",
                    candidate_label,
                )
                if (
                    replicate_match is not None
                    and re.sub(
                        r"[^A-Z0-9]",
                        "",
                        replicate_match.group(1),
                    )
                    == aggregate_group
                ):
                    aggregate_source_series.append(
                        f"{metric_key}_column_"
                        f"{_column_label(candidate_column).lower()}_series"
                    )
            if not aggregate_source_series:
                return None
        if normalized_label.startswith("NORMAL"):
            role = "REFERENCE"
        elif normalized_label.startswith("TEST"):
            role = "TEST"
        else:
            role = "OTHER"
        arm_key = f"{metric_key}_column_{_column_label(column).lower()}"
        header_key = _cell_key(chunk, header_cell)
        records.append(
            record(
                record_type="ENTITY_DECLARATION",
                identity_keys=[header_key],
                exact_source_label=source_label,
                payload={
                    "entityType": "ARM",
                    "key": arm_key,
                    "role": role,
                    "label": source_label,
                    "condition": source_label,
                    "sampleSize": None,
                    "sampleBasis": "",
                    "matchingBasis": "",
                    "factorValues": [],
                },
                evidence_items=[
                    evidence(coordinate(header_cell), source_label)
                ],
            )
        )
        first_value = primary_by_position[(start_row, column)]
        last_value = primary_by_position[(end_row, column)]
        value_range = (
            f"{coordinate(first_value)}:{coordinate(last_value)}"
            if start_row != end_row
            else coordinate(first_value)
        )
        records.append(
            record(
                record_type="SERIES_SEGMENT_APPEND",
                identity_keys=[
                    header_key,
                    axis_start_key,
                    _cell_key(chunk, first_value),
                ],
                exact_source_label=source_label,
                payload={
                    "key": f"{arm_key}_series",
                    "seriesRole": (
                        "AGGREGATE"
                        if aggregate_match is not None
                        else "RAW"
                    ),
                    "aggregationFunction": (
                        "AVERAGE"
                        if aggregate_match is not None
                        else ""
                    ),
                    "aggregateOfSeries": aggregate_source_series,
                    "outcome": outcome_key,
                    "arm": arm_key,
                    "sheet": sheet_title,
                    "headerRange": coordinate(header_cell),
                    "valueRange": value_range,
                    "rowIdentityRange": axis_range,
                    "aggregateReplicateRanges": [],
                    "axisSource": "ROW_IDENTITY",
                    "axisLabel": "source row identity",
                    "axisUnit": "",
                    "valueUnit": "",
                    "stratumKey": "",
                    "verificationStatus": "NEEDS_REVIEW",
                },
                evidence_items=[
                    evidence(coordinate(header_cell), source_label),
                    evidence(value_range, source_label),
                ],
            )
        )

    header_by_column = {
        column: header for column, header in measurement_columns
    }
    for tail_cell in tail_numeric_cells:
        column = int(tail_cell.get("column") or 0)
        header_cell = header_by_column[column]
        source_label = header_text(header_cell)
        arm_key = f"{metric_key}_column_{_column_label(column).lower()}"
        numeric_value = _numeric_value(tail_cell)
        if numeric_value is None:
            return None
        displayed_value = tail_cell.get("displayValue")
        records.append(
            record(
                record_type="OBSERVATION_APPEND",
                identity_keys=[
                    _cell_key(chunk, header_cell),
                    _cell_key(chunk, tail_cell),
                ],
                exact_source_label=source_label,
                payload={
                    "outcome": outcome_key,
                    "arm": arm_key,
                    "valueNumber": numeric_value,
                    "valueText": str(
                        displayed_value
                        if displayed_value is not None
                        else numeric_value
                    ),
                    "numerator": None,
                    "denominator": None,
                    "ratePpm": None,
                    "min": None,
                    "max": None,
                    "average": None,
                    "sampleSize": None,
                    "replicateKey": (
                        f"source-{coordinate(tail_cell).lower()}"
                    ),
                },
                evidence_items=[
                    evidence(coordinate(header_cell), source_label),
                    evidence(
                        coordinate(tail_cell),
                        str(
                            displayed_value
                            if displayed_value is not None
                            else numeric_value
                        ),
                    ),
                ],
            )
        )

    evidence_by_record = {
        str(item["recordId"]): set(
            evidence_cell_keys(
                item["evidence"],
                chunks=focused_chunks,
            )
        )
        for item in records
    }
    source_keys = [
        _cell_key(focused_chunk, cell)
        for focused_chunk in focused_chunks
        for cell in focused_chunk.get("cells", [])
    ]
    if source_keys != [
        str(value) for value in envelope.get("ownedSourceCellKeys", [])
    ]:
        return None
    dispositions = []
    for source_key in source_keys:
        record_ids = [
            str(item["recordId"])
            for item in records
            if source_key in evidence_by_record[str(item["recordId"])]
        ]
        if not record_ids:
            return None
        dispositions.append(
            {
                "sourceCellKey": source_key,
                "disposition": "RECORD_EVIDENCE",
                "recordIds": record_ids,
                "reason": "",
            }
        )
    fragment = {
        "schemaVersion": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
        "source": {
            "revisionUid": str(envelope["source"]["revisionUid"]),
            "contentSha256": str(envelope["source"]["contentSha256"]),
            "contentComplete": False,
        },
        "planId": str(envelope["planId"]),
        "partId": str(envelope["partId"]),
        "inputEnvelopeSha256": str(envelope["inputEnvelopeSha256"]),
        "records": records,
        "coverageDispositions": dispositions,
    }
    return validate_fragment_v2(
        fragment=fragment,
        envelope=envelope,
        all_selected_chunks=focused_chunks,
    )


def fragment_artifact_paths(
    artifact_directory: Path,
    part: dict[str, Any],
) -> tuple[Path, Path]:
    directory = artifact_directory / "draft-parts-v2"
    safe = re.sub(r"[^A-Za-z0-9._-]", "_", str(part["partId"]))
    return (
        directory / f"{safe}.fragment.json",
        directory / f"{safe}.provenance.json",
    )


def part_provenance_v2(
    *,
    plan: dict[str, Any],
    part: dict[str, Any],
    envelope: dict[str, Any],
    output_path: Path,
    output_sha256: str,
    generated_at: str,
) -> dict[str, Any]:
    fragment_identity, fragment_identity_sha256 = (
        _require_current_fragment_identity(plan, part)
    )
    result = {
        "schemaVersion": STAGED_PART_PROVENANCE_V2_SCHEMA_VERSION,
        "planId": plan["planId"],
        "partId": part["partId"],
        "fragmentIdentity": copy.deepcopy(fragment_identity),
        "fragmentIdentitySha256": fragment_identity_sha256,
        "source": copy.deepcopy(plan["source"]),
        "chunkIds": list(part["chunkIds"]),
        "sourceSegments": copy.deepcopy(
            part.get("sourceSegments", [])
        ),
        "ownedSourceCellKeys": list(part["ownedSourceCellKeys"]),
        "sharedAnchorCellKeys": list(part["sharedAnchorCellKeys"]),
        "inputHashes": copy.deepcopy(envelope["inputHashes"]),
        "outputPath": str(output_path),
        "outputSha256": output_sha256,
        "generatedAt": generated_at,
        "imagesAnalyzed": False,
    }
    if part.get("sourceSegments"):
        result["candidateAnchorIds"] = list(
            part.get("candidateAnchorIds", [])
        )
        result["registryAnchorCellKeys"] = list(
            part.get("registryAnchorCellKeys", [])
        )
    return result


def part_provenance_v2_matches(
    *,
    provenance: dict[str, Any],
    plan: dict[str, Any],
    part: dict[str, Any],
    envelope: dict[str, Any],
    output_sha256: str,
    output_path: Path | None = None,
) -> bool:
    fragment_identity = _fragment_identity_v2()
    fragment_identity_sha256 = json_sha256(fragment_identity)
    return bool(
        provenance.get("schemaVersion")
        == STAGED_PART_PROVENANCE_V2_SCHEMA_VERSION
        and plan.get("fragmentIdentity") == fragment_identity
        and plan.get("fragmentIdentitySha256")
        == fragment_identity_sha256
        and part.get("fragmentIdentitySha256")
        == fragment_identity_sha256
        and provenance.get("planId") == plan.get("planId")
        and provenance.get("partId") == part.get("partId")
        and provenance.get("fragmentIdentity") == fragment_identity
        and provenance.get("fragmentIdentitySha256")
        == fragment_identity_sha256
        and provenance.get("source") == plan.get("source")
        and provenance.get("chunkIds") == part.get("chunkIds")
        and provenance.get("sourceSegments", [])
        == part.get("sourceSegments", [])
        and (
            not part.get("sourceSegments")
            or (
                provenance.get("candidateAnchorIds")
                == part.get("candidateAnchorIds")
                and provenance.get("registryAnchorCellKeys")
                == part.get("registryAnchorCellKeys")
            )
        )
        and provenance.get("ownedSourceCellKeys")
        == part.get("ownedSourceCellKeys")
        and provenance.get("sharedAnchorCellKeys")
        == part.get("sharedAnchorCellKeys")
        and provenance.get("inputHashes") == envelope.get("inputHashes")
        and provenance.get("outputSha256") == output_sha256
        and (
            output_path is None
            or provenance.get("outputPath") == str(output_path)
        )
        and provenance.get("imagesAnalyzed") is False
    )


def _source_order_key(
    record: dict[str, Any],
    source_order: dict[str, int],
) -> tuple[int, str]:
    return (
        min(
            (
                source_order.get(str(key), 10**12)
                for key in record.get("identityCellKeys", [])
            ),
            default=10**12,
        ),
        str(record.get("recordId") or ""),
    )


def _merge_nonempty(
    target: dict[str, Any],
    incoming: dict[str, Any],
    *,
    path: str,
) -> dict[str, Any]:
    result = copy.deepcopy(target)
    for key, value in incoming.items():
        if key in {"evidence", "limitations"}:
            continue
        previous = result.get(key)
        if previous in (None, "", [], {}):
            result[key] = copy.deepcopy(value)
        elif value in (None, "", [], {}):
            continue
        elif previous != value:
            raise StagedDraftV2Error(
                f"Conflicting nonempty fragment values at {path}.{key}"
            )
    return result


_ENTITY_SOURCE_DESCRIPTION_FIELDS = frozenset(
    {
        "role",
        "label",
        "condition",
        "sampleSize",
        "sampleBasis",
        "matchingBasis",
        "factorValues",
        "originalLabel",
        "metricType",
        "baselineCondition",
        "changedCondition",
        "originalValue",
        "normalizedValue",
    }
)


def _merge_entity_payload(
    target: dict[str, Any],
    incoming: dict[str, Any],
    *,
    path: str,
) -> dict[str, Any]:
    """Merge one source-stable entity without trusting model wording.

    Entity record IDs are bound to the source identity, not to a model-created
    key or description. Adjacent fragments can therefore describe the same
    source cell with different wording. Keep the first source-ordered wording,
    while continuing to reject conflicts in semantic contract fields such as
    entityType, metricType, unit, direction, status, and context kind.
    """

    result = copy.deepcopy(target)
    for key, value in incoming.items():
        if key == "entityId":
            continue
        previous = result.get(key)
        if previous in (None, "", [], {}):
            result[key] = copy.deepcopy(value)
        elif value in (None, "", [], {}):
            continue
        elif previous == value:
            continue
        elif key in _ENTITY_SOURCE_DESCRIPTION_FIELDS:
            continue
        else:
            raise StagedDraftV2Error(
                f"Conflicting nonempty fragment values at {path}.{key}"
            )
    return result


def _register_entity_alias(
    *,
    aliases: dict[tuple[str, str, str], str],
    entity_record_ids: dict[tuple[str, str, str], str],
    logical_id: str,
    entity_type: str,
    old_key: str,
    canonical_key: str,
    record_id: str,
) -> None:
    alias_identity = (logical_id, entity_type, old_key)
    declared_record_id = entity_record_ids.get(alias_identity)
    if declared_record_id not in (None, record_id):
        raise StagedDraftV2Error(
            "Entity key alias collides with a distinct source-stable "
            f"entity at {logical_id}:{entity_type}:{old_key}"
        )
    previous_target = aliases.get(alias_identity)
    if previous_target not in (None, canonical_key):
        raise StagedDraftV2Error(
            "Entity key alias has conflicting canonical targets at "
            f"{logical_id}:{entity_type}:{old_key}"
        )
    aliases[alias_identity] = canonical_key


def _entity_source_alias_compatible(
    *,
    existing_key: str,
    incoming_key: str,
    existing_label: str,
    incoming_label: str,
) -> bool:
    """Allow one source identity to reconcile compatible model names."""

    if existing_key == incoming_key:
        return True

    def compact(value: str) -> str:
        return "".join(
            character.casefold()
            for character in value
            if character.isalnum()
        )

    existing_key_value = compact(existing_key)
    incoming_key_value = compact(incoming_key)
    existing_label_value = compact(existing_label)
    incoming_label_value = compact(incoming_label)
    return (
        min(len(existing_key_value), len(incoming_key_value)) >= 3
        and (
            existing_key_value in incoming_key_value
            or incoming_key_value in existing_key_value
        )
        and min(
            len(existing_label_value),
            len(incoming_label_value),
        )
        >= 3
        and (
            existing_label_value in incoming_label_value
            or incoming_label_value in existing_label_value
        )
    )


def _rewrite_fragment_entity_aliases(
    *,
    fragment: dict[str, Any],
    aliases: dict[tuple[str, str, str], str],
) -> dict[str, Any]:
    result = copy.deepcopy(fragment)

    def rewrite(
        *,
        logical_id: str,
        entity_type: str,
        value: Any,
    ) -> Any:
        key = str(value or "")
        if not key:
            return value
        return aliases.get((logical_id, entity_type, key), value)

    for record in result.get("records", []):
        if not isinstance(record, dict):
            continue
        payload = record.get("payload")
        if not isinstance(payload, dict):
            continue
        logical_id = str(record.get("logicalStudyId") or "")
        record_type = str(record.get("recordType") or "").upper()
        if record_type == "ENTITY_DECLARATION":
            entity_type = str(payload.get("entityType") or "").upper()
            payload["key"] = rewrite(
                logical_id=logical_id,
                entity_type=entity_type,
                value=payload.get("key"),
            )
            if entity_type == "ARM":
                factor_values = payload.get("factorValues")
                if isinstance(factor_values, list):
                    for factor_value in factor_values:
                        if isinstance(factor_value, dict):
                            factor_value["factor"] = rewrite(
                                logical_id=logical_id,
                                entity_type="FACTOR",
                                value=factor_value.get("factor"),
                            )
        elif record_type in {
            "OBSERVATION_APPEND",
            "SERIES_SEGMENT_APPEND",
        }:
            payload["outcome"] = rewrite(
                logical_id=logical_id,
                entity_type="OUTCOME",
                value=payload.get("outcome"),
            )
            payload["arm"] = rewrite(
                logical_id=logical_id,
                entity_type="ARM",
                value=payload.get("arm"),
            )
        elif record_type == "COMPARISON_LINK_INTENT":
            payload["comparedArm"] = rewrite(
                logical_id=logical_id,
                entity_type="ARM",
                value=payload.get("comparedArm"),
            )
            payload["controlArm"] = rewrite(
                logical_id=logical_id,
                entity_type="ARM",
                value=payload.get("controlArm"),
            )
            outcomes = payload.get("outcomes")
            if isinstance(outcomes, list):
                payload["outcomes"] = [
                    rewrite(
                        logical_id=logical_id,
                        entity_type="OUTCOME",
                        value=value,
                    )
                    for value in outcomes
                ]
    return result


def _ordered_evidence_union(
    values: Iterable[dict[str, Any]],
    *,
    chunks: Sequence[dict[str, Any]],
    source_order: dict[str, int],
) -> list[dict[str, Any]]:
    unique: dict[bytes, dict[str, Any]] = {}
    for value in values:
        encoded = compact_json_bytes(value)
        unique.setdefault(encoded, copy.deepcopy(value))
    return sorted(
        unique.values(),
        key=lambda item: (
            min(
                (
                    source_order.get(key, 10**12)
                    for key in evidence_cell_keys(
                        [item],
                        chunks=chunks,
                    )
                ),
                default=10**12,
            ),
            compact_json_bytes(item),
        ),
    )


def merge_fragment_records(
    *,
    plan: dict[str, Any],
    fragments: Sequence[tuple[dict[str, Any], dict[str, Any]]],
    selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Merge validated fragments independent of arrival order."""

    expected = {str(part["partId"]): part for part in plan["parts"]}
    by_part: dict[str, dict[str, Any]] = {}
    for part, fragment in fragments:
        part_id = str(part.get("partId") or "")
        if part_id not in expected or part_id in by_part:
            raise StagedDraftV2Error(
                "Fragments must cover each planned part exactly once"
            )
        by_part[part_id] = fragment
    if set(by_part) != set(expected):
        raise StagedDraftV2Error(
            "Cannot merge until every planned fragment is complete"
        )
    by_coordinate, by_key, source_order = _source_cell_maps(
        selected_chunks
    )
    record_by_id: dict[str, dict[str, Any]] = {}
    canonical_entity_key_by_record_id: dict[str, tuple[str, str, str]] = {}
    entity_record_ids: dict[tuple[str, str, str], str] = {}
    entity_aliases: dict[tuple[str, str, str], str] = {}
    entity_record_id_aliases: dict[str, str] = {}
    entity_source_declarations: dict[
        tuple[str, str, tuple[str, ...]],
        list[tuple[str, str, str]],
    ] = {}
    disposition_by_key: dict[str, dict[str, Any]] = {}
    for part in plan["parts"]:
        fragment = copy.deepcopy(by_part[str(part["partId"])])
        fragment = normalize_fragment_record_ids(
            fragment=fragment,
            envelope={"source": plan["source"]},
        )
        local_entity_record_ids: dict[tuple[str, str, str], str] = {}
        for record in fragment.get("records", []):
            if (
                not isinstance(record, dict)
                or str(record.get("recordType") or "").upper()
                != "ENTITY_DECLARATION"
                or not isinstance(record.get("payload"), dict)
            ):
                continue
            payload = record["payload"]
            record_id = str(record.get("recordId") or "")
            logical_id = str(record.get("logicalStudyId") or "")
            entity_type = str(payload.get("entityType") or "").upper()
            entity_key = str(payload.get("key") or "")
            exact_source_label = str(
                record.get("exactSourceLabel") or ""
            )
            source_identity = (
                logical_id,
                entity_type,
                tuple(
                    sorted(
                        {
                            str(value)
                            for value in record.get(
                                "identityCellKeys",
                                [],
                            )
                        }
                    )
                ),
            )
            compatible_source_declarations = [
                value
                for value in entity_source_declarations.get(
                    source_identity,
                    [],
                )
                if _entity_source_alias_compatible(
                    existing_key=value[1],
                    incoming_key=entity_key,
                    existing_label=value[2],
                    incoming_label=exact_source_label,
                )
            ]
            compatible_targets = {
                (value[0], value[1])
                for value in compatible_source_declarations
            }
            if len(compatible_targets) > 1:
                raise StagedDraftV2Error(
                    "Source-identical entity declarations have "
                    f"ambiguous aliases at {logical_id}:"
                    f"{entity_type}:{entity_key}"
                )
            if compatible_targets:
                canonical_record_id, canonical_key = next(
                    iter(compatible_targets)
                )
                if entity_key != canonical_key:
                    _register_entity_alias(
                        aliases=entity_aliases,
                        entity_record_ids=entity_record_ids,
                        logical_id=logical_id,
                        entity_type=entity_type,
                        old_key=entity_key,
                        canonical_key=canonical_key,
                        record_id=canonical_record_id,
                    )
                entity_record_id_aliases[record_id] = (
                    canonical_record_id
                )
                record_id = canonical_record_id
                record["recordId"] = canonical_record_id
                payload["entityId"] = canonical_record_id
                entity_key = canonical_key
                payload["key"] = canonical_key
            identity = (logical_id, entity_type, entity_key)
            same_key_record_id = (
                local_entity_record_ids.get(identity)
                or entity_record_ids.get(identity)
            )
            if (
                same_key_record_id is not None
                and same_key_record_id != record_id
            ):
                # Repeated table sections commonly redeclare one canonical
                # entity key from different source cells. Coalesce those
                # declarations by key; the later payload merge remains the
                # fail-closed semantic compatibility gate.
                entity_record_id_aliases[record_id] = (
                    same_key_record_id
                )
                record_id = same_key_record_id
                record["recordId"] = same_key_record_id
                payload["entityId"] = same_key_record_id
            local_previous = local_entity_record_ids.get(identity)
            if local_previous not in (None, record_id):
                raise StagedDraftV2Error(
                    "Fragment reuses an entity key for distinct "
                    f"source identities at {logical_id}:{entity_type}:"
                    f"{entity_key}"
                )
            local_entity_record_ids[identity] = record_id
            canonical = canonical_entity_key_by_record_id.get(record_id)
            if canonical is None:
                declarations = entity_source_declarations.setdefault(
                    source_identity,
                    [],
                )
                source_declaration = (
                    record_id,
                    entity_key,
                    exact_source_label,
                )
                if source_declaration not in declarations:
                    declarations.append(source_declaration)
                continue
            canonical_logical_id, canonical_type, canonical_key = canonical
            if (
                logical_id != canonical_logical_id
                or entity_type != canonical_type
            ):
                raise StagedDraftV2Error(
                    f"Conflicting record identity {record_id}"
                )
            if entity_key != canonical_key:
                _register_entity_alias(
                    aliases=entity_aliases,
                    entity_record_ids=entity_record_ids,
                    logical_id=logical_id,
                    entity_type=entity_type,
                    old_key=entity_key,
                    canonical_key=canonical_key,
                    record_id=record_id,
                )
        for identity, record_id in local_entity_record_ids.items():
            alias_target = entity_aliases.get(identity)
            if (
                alias_target is not None
                and canonical_entity_key_by_record_id.get(record_id)
                is None
            ):
                raise StagedDraftV2Error(
                    "Entity key alias collides with a distinct local "
                    f"declaration at {identity[0]}:{identity[1]}:"
                    f"{identity[2]}"
                )
        fragment = _rewrite_fragment_entity_aliases(
            fragment=fragment,
            aliases=entity_aliases,
        )
        fragment = normalize_fragment_record_ids(
            fragment=fragment,
            envelope={"source": plan["source"]},
        )
        for record in fragment.get("records", []):
            if (
                not isinstance(record, dict)
                or str(record.get("recordType") or "").upper()
                != "ENTITY_DECLARATION"
            ):
                continue
            record_id = str(record.get("recordId") or "")
            canonical_record_id = entity_record_id_aliases.get(
                record_id,
                record_id,
            )
            record["recordId"] = canonical_record_id
            if isinstance(record.get("payload"), dict):
                record["payload"]["entityId"] = canonical_record_id
        for disposition in fragment.get(
            "coverageDispositions",
            [],
        ):
            if not isinstance(disposition, dict):
                continue
            record_ids = disposition.get("recordIds")
            if isinstance(record_ids, list):
                disposition["recordIds"] = [
                    entity_record_id_aliases.get(
                        str(value),
                        str(value),
                    )
                    for value in record_ids
                ]
        part_owned = list(part["ownedSourceCellKeys"])
        fragment_dispositions = list(
            fragment.get("coverageDispositions", [])
        )
        if [
            str(item.get("sourceCellKey") or "")
            for item in fragment_dispositions
        ] != part_owned:
            raise StagedDraftV2Error(
                f"Part {part['partId']} disposition order/ownership changed"
            )
        for disposition in fragment_dispositions:
            key = str(disposition["sourceCellKey"])
            if key in disposition_by_key:
                raise StagedDraftV2Error(
                    f"Merged disposition duplicates source cell {key}"
                )
            disposition_by_key[key] = copy.deepcopy(disposition)
        for record in fragment["records"]:
            record_id = str(record["recordId"])
            previous = record_by_id.get(record_id)
            if previous is None:
                record_by_id[record_id] = copy.deepcopy(record)
                if (
                    str(record.get("recordType") or "").upper()
                    == "ENTITY_DECLARATION"
                    and isinstance(record.get("payload"), dict)
                ):
                    payload = record["payload"]
                    logical_id = str(record["logicalStudyId"])
                    entity_type = str(
                        payload.get("entityType") or ""
                    ).upper()
                    entity_key = str(payload.get("key") or "")
                    identity = (logical_id, entity_type, entity_key)
                    known_record_id = entity_record_ids.get(identity)
                    if known_record_id not in (None, record_id):
                        raise StagedDraftV2Error(
                            "Entity key is declared by distinct "
                            f"source identities at {logical_id}:"
                            f"{entity_type}:{entity_key}"
                        )
                    entity_record_ids[identity] = record_id
                    canonical_entity_key_by_record_id[record_id] = identity
                continue
            if (
                previous["recordType"] != record["recordType"]
                or previous["logicalStudyId"]
                != record["logicalStudyId"]
            ):
                raise StagedDraftV2Error(
                    f"Conflicting record identity {record_id}"
                )
            if (
                str(record.get("recordType") or "").upper()
                == "ENTITY_DECLARATION"
            ):
                merged_payload = _merge_entity_payload(
                    previous["payload"],
                    record["payload"],
                    path=f"record[{record_id}].payload",
                )
            else:
                merged_payload = _merge_nonempty(
                    previous["payload"],
                    record["payload"],
                    path=f"record[{record_id}].payload",
                )
            previous["payload"] = merged_payload
            previous["evidence"] = _ordered_evidence_union(
                [
                    *previous.get("evidence", []),
                    *record.get("evidence", []),
                ],
                chunks=selected_chunks,
                source_order=source_order,
            )
    records = sorted(
        record_by_id.values(),
        key=lambda record: _source_order_key(record, source_order),
    )
    expected_owned = list(plan["ownedSourceCellKeys"])
    if set(disposition_by_key) != set(expected_owned):
        raise StagedDraftV2Error(
            "Merged dispositions do not exactly cover planned ownership"
        )
    # Repeated source-stable records accumulate evidence while adjacent
    # fragments are merged. Rebuild RECORD_EVIDENCE dispositions from that
    # final evidence union so every source cell cites all and only the records
    # that actually preserve it. Keeping each fragment's pre-merge recordIds
    # would make those dispositions stale as soon as one entity or Study patch
    # spans more than one part.
    final_evidence_by_record = {
        str(record["recordId"]): set(
            evidence_cell_keys(
                record.get("evidence", []),
                chunks=selected_chunks,
            )
        )
        for record in records
    }
    merged_dispositions: list[dict[str, Any]] = []
    for key in expected_owned:
        record_ids = [
            str(record["recordId"])
            for record in records
            if key
            in final_evidence_by_record[str(record["recordId"])]
        ]
        if record_ids:
            merged_dispositions.append(
                {
                    "sourceCellKey": key,
                    "disposition": "RECORD_EVIDENCE",
                    "recordIds": record_ids,
                    "reason": "",
                }
            )
            continue
        original = copy.deepcopy(disposition_by_key[key])
        if (
            str(original.get("disposition") or "").upper()
            == "RECORD_EVIDENCE"
            or original.get("recordIds")
        ):
            raise StagedDraftV2Error(
                f"Merged record evidence disappeared for source cell {key}"
            )
        merged_dispositions.append(original)
    result = {
        "schemaVersion": "study-draft-merged-records-v2",
        "planId": plan["planId"],
        "records": records,
        "coverageDispositions": merged_dispositions,
    }
    result["recordsSha256"] = json_sha256(result)
    return result


def _entity_maps(
    records: Sequence[dict[str, Any]],
) -> tuple[
    dict[tuple[str, str, str], dict[str, Any]],
    dict[tuple[str, str], set[str]],
]:
    entities: dict[tuple[str, str, str], dict[str, Any]] = {}
    cross_index: dict[tuple[str, str], set[str]] = {}
    for record in records:
        if record["recordType"] != "ENTITY_DECLARATION":
            continue
        payload = record["payload"]
        entity_type = str(payload.get("entityType") or "").upper()
        key = str(payload.get("key") or "")
        if entity_type not in ENTITY_TYPES or not key:
            raise StagedDraftV2Error(
                "ENTITY_DECLARATION requires entityType and key"
            )
        logical_id = str(record["logicalStudyId"])
        identity = (logical_id, entity_type, key)
        previous = entities.get(identity)
        semantic_payload = copy.deepcopy(payload)
        # entityId is the source-bound declaration recordId, not part of the
        # canonical entity's semantic identity. The same entity can be
        # declared from more than one source fragment, so those declaration
        # IDs are expected to differ and must not create a semantic conflict.
        semantic_payload.pop("entityId", None)
        if previous is not None:
            previous_payload = copy.deepcopy(previous["payload"])
            previous_payload.pop("entityId", None)
            previous["payload"] = _merge_entity_payload(
                previous_payload,
                semantic_payload,
                path=f"entity[{logical_id}:{entity_type}:{key}]",
            )
            evidence_values = [
                *previous.get("evidence", []),
                *record.get("evidence", []),
            ]
            seen_evidence: set[bytes] = set()
            previous["evidence"] = [
                copy.deepcopy(item)
                for item in evidence_values
                if not (
                    compact_json_bytes(item) in seen_evidence
                    or seen_evidence.add(compact_json_bytes(item))
                )
            ]
        else:
            entity_record = copy.deepcopy(record)
            entity_record["payload"] = semantic_payload
            entities[identity] = entity_record
        cross_index.setdefault((entity_type, key), set()).add(logical_id)
    return entities, cross_index


def _require_local_entity(
    *,
    entities: dict[tuple[str, str, str], dict[str, Any]],
    cross_index: dict[tuple[str, str], set[str]],
    logical_id: str,
    entity_type: str,
    key: object,
    path: str,
) -> dict[str, Any]:
    normalized_key = str(key or "")
    entity = entities.get((logical_id, entity_type, normalized_key))
    if entity is not None:
        return entity
    if cross_index.get((entity_type, normalized_key)):
        raise StagedDraftV2Error(
            f"{path} contains a cross-Study reference {normalized_key}"
        )
    raise StagedDraftV2Error(
        f"{path} references unknown {entity_type} {normalized_key}"
    )


def _series_bounds(payload: dict[str, Any]) -> tuple[
    tuple[int, int, int, int],
    tuple[int, int, int, int],
    tuple[int, int, int, int],
]:
    return (
        range_bounds(payload.get("headerRange")),
        range_bounds(payload.get("valueRange")),
        range_bounds(payload.get("rowIdentityRange")),
    )


def _column_label(value: int) -> str:
    result = ""
    current = value
    while current:
        current, remainder = divmod(current - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _address(bounds: tuple[int, int, int, int]) -> str:
    start_row, start_column, end_row, end_column = bounds
    start = f"{_column_label(start_column)}{start_row}"
    end = f"{_column_label(end_column)}{end_row}"
    return start if start == end else f"{start}:{end}"


def merge_adjacent_series_segments(
    records: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Merge only exact adjacent rectangles with identical semantic axes."""

    groups: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for record in records:
        payload = record["payload"]
        group = (
            str(record["logicalStudyId"]),
            str(payload.get("outcome") or ""),
            str(payload.get("arm") or ""),
            str(payload.get("axisSource") or ""),
            str(payload.get("sheet") or ""),
            str(payload.get("valueUnit") or ""),
        )
        groups.setdefault(group, []).append(copy.deepcopy(record))
    result: list[dict[str, Any]] = []
    for values in groups.values():
        values.sort(
            key=lambda record: (
                _series_bounds(record["payload"])[1],
                record["recordId"],
            )
        )
        current = values[0]
        for incoming in values[1:]:
            current_header, current_value, current_identity = (
                _series_bounds(current["payload"])
            )
            next_header, next_value, next_identity = _series_bounds(
                incoming["payload"]
            )
            vertical = (
                current_value[1] == next_value[1]
                and current_value[3] == next_value[3]
                and current_value[2] + 1 == next_value[0]
                and current_header == next_header
                and current_identity[1] == next_identity[1]
                and current_identity[3] == next_identity[3]
                and current_identity[2] + 1 == next_identity[0]
            )
            horizontal = (
                current_value[0] == next_value[0]
                and current_value[2] == next_value[2]
                and current_value[3] + 1 == next_value[1]
                and current_identity == next_identity
                and current_header[0] == next_header[0]
                and current_header[2] == next_header[2]
                and current_header[3] + 1 == next_header[1]
            )
            if vertical or horizontal:
                payload = current["payload"]
                payload["valueRange"] = _address(
                    (
                        min(current_value[0], next_value[0]),
                        min(current_value[1], next_value[1]),
                        max(current_value[2], next_value[2]),
                        max(current_value[3], next_value[3]),
                    )
                )
                payload["headerRange"] = _address(
                    (
                        min(current_header[0], next_header[0]),
                        min(current_header[1], next_header[1]),
                        max(current_header[2], next_header[2]),
                        max(current_header[3], next_header[3]),
                    )
                )
                payload["rowIdentityRange"] = _address(
                    (
                        min(current_identity[0], next_identity[0]),
                        min(current_identity[1], next_identity[1]),
                        max(current_identity[2], next_identity[2]),
                        max(current_identity[3], next_identity[3]),
                    )
                )
                current["evidence"] = [
                    *current.get("evidence", []),
                    *incoming.get("evidence", []),
                ]
                current["recordId"] = stable_uid(
                    "merged-series-v2",
                    current["logicalStudyId"],
                    payload["outcome"],
                    payload["arm"],
                    payload["sheet"],
                    payload["valueRange"],
                )
            else:
                result.append(current)
                current = incoming
        result.append(current)
    return result


def project_canonical_manifest(
    *,
    merged: dict[str, Any],
    registry: dict[str, Any],
    source: dict[str, Any],
    workbook: dict[str, Any],
    selected_chunks: Sequence[dict[str, Any]],
    semantic_inventory: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Resolve append records and project one canonical manifest."""

    records = list(merged["records"])
    dispositions = list(merged.get("coverageDispositions", []))
    registry_ids = {
        str(study["logicalStudyId"]) for study in registry["studies"]
    }
    if any(
        str(record["logicalStudyId"]) not in registry_ids
        for record in records
    ):
        raise StagedDraftV2Error(
            "Merged records contain an unregistered Study"
        )
    by_coordinate, by_key, source_order = _source_cell_maps(
        selected_chunks
    )
    selected_owned_keys = [
        _cell_key(chunk, cell)
        for chunk in selected_chunks
        for cell in chunk.get("cells", [])
    ]
    if [
        str(item.get("sourceCellKey") or "")
        for item in dispositions
    ] != selected_owned_keys:
        raise StagedDraftV2Error(
            "Final merged dispositions are not the exact source-ordered "
            "selected ownership"
        )
    final_record_evidence: dict[str, set[str]] = {
        str(record["recordId"]): set(
            evidence_cell_keys(
                record.get("evidence", []),
                chunks=selected_chunks,
            )
        )
        for record in records
    }
    for disposition in dispositions:
        key = str(disposition["sourceCellKey"])
        expected_ids = {
            record_id
            for record_id, evidence_keys in final_record_evidence.items()
            if key in evidence_keys
        }
        actual_ids = {
            str(value)
            for value in disposition.get("recordIds", [])
        }
        disposition_value = str(
            disposition.get("disposition") or ""
        ).upper()
        if expected_ids:
            if (
                disposition_value != "RECORD_EVIDENCE"
                or actual_ids != expected_ids
            ):
                raise StagedDraftV2Error(
                    f"Final record coverage differs from disposition for {key}"
                )
        elif disposition_value == "RECORD_EVIDENCE" or actual_ids:
            raise StagedDraftV2Error(
                f"Final disposition for {key} cites absent record evidence"
            )
    if semantic_inventory is None:
        semantic_inventory = build_content_coverage_inventory(
            chunks=selected_chunks,
            locator_results=[],
            expected_source_cell_keys=selected_owned_keys,
        )
    logical_by_source_key: dict[str, str] = {}
    for registry_study in registry["studies"]:
        logical_id = str(registry_study["logicalStudyId"])
        for source_key in registry_study.get(
            "anchorEvidenceCellKeys",
            [],
        ):
            normalized_key = str(source_key)
            existing_logical = logical_by_source_key.get(
                normalized_key
            )
            if (
                existing_logical is not None
                and existing_logical != logical_id
            ):
                raise StagedDraftV2Error(
                    "Registry assigns one source cell to multiple Studies"
                )
            logical_by_source_key[normalized_key] = logical_id
    _append_semantic_label_records_v2(
        envelope={
            "source": source,
            "ownedSourceCellKeys": selected_owned_keys,
            "locatorResults": [],
        },
        focused_chunks=selected_chunks,
        logical_by_source_key=logical_by_source_key,
        records=records,
        inventory=semantic_inventory,
    )
    entities, cross_index = _entity_maps(records)
    studies: list[dict[str, Any]] = []
    workbook_limitations: list[str] = []
    workbook_evidence: list[dict[str, Any]] = []
    for registry_study in registry["studies"]:
        logical_id = str(registry_study["logicalStudyId"])
        local = [
            record
            for record in records
            if record["logicalStudyId"] == logical_id
        ]
        patches = [
            record for record in local if record["recordType"] == "STUDY_PATCH"
        ]
        study_value: dict[str, Any] = {}
        for patch in patches:
            study_value = _merge_nonempty(
                study_value,
                patch["payload"],
                path=f"Study[{logical_id}]",
            )
        evidence = _ordered_evidence_union(
            (
                item
                for record in local
                for item in record.get("evidence", [])
            ),
            chunks=selected_chunks,
            source_order=source_order,
        )
        workbook_evidence.extend(evidence)
        contexts: list[dict[str, Any]] = []
        factors: list[dict[str, Any]] = []
        arms: list[dict[str, Any]] = []
        outcomes: list[dict[str, Any]] = []
        outcome_by_key: dict[str, dict[str, Any]] = {}
        for (
            entity_logical,
            entity_type,
            entity_key,
        ), record in entities.items():
            if entity_logical != logical_id:
                continue
            payload = copy.deepcopy(record["payload"])
            payload.pop("entityType", None)
            payload.pop("entityId", None)
            payload["key"] = entity_key
            payload["evidence"] = copy.deepcopy(
                record.get("evidence", [])
            )
            if entity_type == "CONTEXT":
                contexts.append(payload)
            elif entity_type == "FACTOR":
                factors.append(payload)
            elif entity_type == "ARM":
                payload.setdefault("role", "OTHER")
                if (
                    re.fullmatch(
                        r"normal(?:\s*\([^()]*\))?",
                        str(payload.get("label") or "").strip(),
                        re.IGNORECASE,
                    )
                    and str(payload.get("role") or "").upper()
                    != "REFERENCE"
                ):
                    payload["role"] = "REFERENCE"
                if not isinstance(payload.get("factorValues"), list):
                    payload["factorValues"] = []
                for factor_value_index, factor_value in enumerate(
                    payload["factorValues"]
                ):
                    if not isinstance(factor_value, dict):
                        raise StagedDraftV2Error(
                            "Arm factorValues must contain objects"
                        )
                    _require_local_entity(
                        entities=entities,
                        cross_index=cross_index,
                        logical_id=logical_id,
                        entity_type="FACTOR",
                        key=factor_value.get("factor"),
                        path=(
                            f"Arm[{entity_key}].factorValues"
                            f"[{factor_value_index}]"
                        ),
                    )
                arms.append(payload)
            elif entity_type == "OUTCOME":
                if not isinstance(payload.get("observations"), list):
                    payload["observations"] = []
                outcomes.append(payload)
                outcome_by_key[entity_key] = payload
        for record in local:
            if record["recordType"] != "OBSERVATION_APPEND":
                continue
            payload = copy.deepcopy(record["payload"])
            payload, observation_evidence = (
                _complete_incomplete_formula_rate_pair(
                    payload=payload,
                    evidence=record.get("evidence", []),
                    selected_chunks=selected_chunks,
                    by_coordinate=by_coordinate,
                    by_key=by_key,
                )
            )
            outcome_key = str(payload.pop("outcome", ""))
            arm_key = str(payload.get("arm") or "")
            _require_local_entity(
                entities=entities,
                cross_index=cross_index,
                logical_id=logical_id,
                entity_type="OUTCOME",
                key=outcome_key,
                path=f"Observation[{record['recordId']}]",
            )
            _require_local_entity(
                entities=entities,
                cross_index=cross_index,
                logical_id=logical_id,
                entity_type="ARM",
                key=arm_key,
                path=f"Observation[{record['recordId']}]",
            )
            payload = _normalize_projected_percent_observation(
                payload=payload,
                evidence=observation_evidence,
                outcome_unit=outcome_by_key[outcome_key].get("unit"),
                outcome_label=outcome_by_key[outcome_key].get(
                    "originalLabel"
                ),
                selected_chunks=selected_chunks,
                by_key=by_key,
            )
            payload.setdefault("key", record["recordId"])
            payload["evidence"] = observation_evidence
            outcome_by_key[outcome_key]["observations"].append(payload)
        series_records = [
            record
            for record in local
            if record["recordType"] == "SERIES_SEGMENT_APPEND"
        ]
        for record in series_records:
            payload = record["payload"]
            _require_local_entity(
                entities=entities,
                cross_index=cross_index,
                logical_id=logical_id,
                entity_type="OUTCOME",
                key=payload.get("outcome"),
                path=f"Series[{record['recordId']}]",
            )
            _require_local_entity(
                entities=entities,
                cross_index=cross_index,
                logical_id=logical_id,
                entity_type="ARM",
                key=payload.get("arm"),
                path=f"Series[{record['recordId']}]",
            )
        normalized_series_records: list[dict[str, Any]] = []
        series_text_observations: list[dict[str, Any]] = []
        for record in series_records:
            split_records, text_observations = (
                _split_mixed_numeric_series_record(
                    record=record,
                    by_coordinate=by_coordinate,
                    revision_uid=str(source["revisionUid"]),
                )
            )
            normalized_series_records.extend(split_records)
            series_text_observations.extend(text_observations)
        text_outcomes: dict[str, dict[str, Any]] = {}
        for observation in series_text_observations:
            source_outcome_key = str(
                observation.pop("_sourceOutcome", "")
            )
            text_outcome_key = f"{source_outcome_key}_source_text"
            text_outcome = text_outcomes.get(text_outcome_key)
            if text_outcome is None:
                text_outcome = {
                    "key": text_outcome_key,
                    "originalLabel": str(
                        observation.get("valueText")
                        or "Recorded text status"
                    ),
                    "metricType": "categorical_status",
                    "unit": "",
                    "favorableDirection": "UNKNOWN",
                    "evidence": copy.deepcopy(
                        observation.get("evidence", [])[:1]
                    ),
                    "observations": [],
                }
                text_outcomes[text_outcome_key] = text_outcome
                outcomes.append(text_outcome)
                outcome_by_key[text_outcome_key] = text_outcome
            text_outcome["observations"].append(observation)
        measurement_series = []
        for record in merge_adjacent_series_segments(
            normalized_series_records
        ):
            payload = copy.deepcopy(record["payload"])
            payload.setdefault("key", record["recordId"])
            payload.setdefault("seriesRole", "RAW")
            payload.setdefault("verificationStatus", "NEEDS_REVIEW")
            for field in (
                "aggregateOfSeries",
                "aggregateReplicateRanges",
            ):
                if not isinstance(payload.get(field), list):
                    payload[field] = []
            measurement_series.append(payload)

        observation_pairs = {
            (
                str(payload.get("outcome") or ""),
                str(payload.get("arm") or ""),
            )
            for record in local
            if record["recordType"] == "OBSERVATION_APPEND"
            for payload in [record["payload"]]
        }
        comparisons: list[dict[str, Any]] = []
        for record in local:
            if record["recordType"] != "COMPARISON_LINK_INTENT":
                continue
            payload = record["payload"]
            compared = str(payload.get("comparedArm") or "")
            control = str(payload.get("controlArm") or "")
            _require_local_entity(
                entities=entities,
                cross_index=cross_index,
                logical_id=logical_id,
                entity_type="ARM",
                key=compared,
                path=f"ComparisonIntent[{record['recordId']}]",
            )
            _require_local_entity(
                entities=entities,
                cross_index=cross_index,
                logical_id=logical_id,
                entity_type="ARM",
                key=control,
                path=f"ComparisonIntent[{record['recordId']}]",
            )
            outcome_keys = payload.get("outcomes")
            if not isinstance(outcome_keys, list) or not outcome_keys:
                raise StagedDraftV2Error(
                    "Comparison intent requires explicit outcomes"
                )
            for outcome_key in outcome_keys:
                normalized_outcome = str(outcome_key)
                _require_local_entity(
                    entities=entities,
                    cross_index=cross_index,
                    logical_id=logical_id,
                    entity_type="OUTCOME",
                    key=normalized_outcome,
                    path=f"ComparisonIntent[{record['recordId']}]",
                )
                if (
                    (normalized_outcome, compared)
                    not in observation_pairs
                    or (normalized_outcome, control)
                    not in observation_pairs
                ):
                    raise StagedDraftV2Error(
                        "Comparison intent is dangling without both arm observations"
                    )
            comparisons.append(
                {
                    "key": record["recordId"],
                    "comparedArm": compared,
                    "controlArm": control,
                    "designType": str(
                        payload.get("designType") or "UNKNOWN"
                    ),
                    "matchingBasis": str(
                        payload.get("matchingBasis") or ""
                    ),
                    "validityStatus": "NEEDS_REVIEW",
                    "confoundingStatus": "UNASSESSED",
                    "verificationStatus": "NEEDS_REVIEW",
                    "aggregationEligible": False,
                    "evidence": copy.deepcopy(
                        record.get("evidence", [])
                    ),
                    "effects": [],
                }
            )
        conclusions = []
        limitations = []
        for record in local:
            if record["recordType"] == "CONCLUSION_APPEND":
                payload = copy.deepcopy(record["payload"])
                payload.setdefault("key", record["recordId"])
                payload["evidence"] = copy.deepcopy(
                    record.get("evidence", [])
                )
                conclusions.append(payload)
            elif record["recordType"] == "LIMITATION_APPEND":
                text = str(record["payload"].get("text") or "").strip()
                if not text:
                    raise StagedDraftV2Error(
                        "LIMITATION_APPEND requires text"
                    )
                if str(record["payload"].get("scope") or "STUDY").upper() == "WORKBOOK":
                    workbook_limitations.append(text)
                else:
                    limitations.append(text)
        study = {
            "key": logical_id,
            "title": str(
                study_value.get("title")
                or registry_study.get("titleHint")
                or logical_id
            ),
            "purpose": str(study_value.get("purpose") or ""),
            "hypothesis": str(study_value.get("hypothesis") or ""),
            "objective": str(study_value.get("objective") or ""),
            "designType": str(
                study_value.get("designType") or "UNKNOWN"
            ),
            "comparisonBasis": str(
                study_value.get("comparisonBasis") or ""
            ),
            "verificationStatus": "NEEDS_REVIEW",
            "comparabilityStatus": "UNASSESSED",
            "confoundingStatus": "UNASSESSED",
            "summary": str(study_value.get("summary") or ""),
            "limitations": list(dict.fromkeys(limitations)),
            "evidence": evidence,
            "contexts": contexts,
            "factors": factors,
            "arms": arms,
            "outcomes": outcomes,
            "measurementSeries": measurement_series,
            "comparisons": comparisons,
            "conclusions": conclusions,
        }
        studies.append(study)
    studies = _augment_projected_categorical_status_observations(
        studies=studies,
        selected_chunks=selected_chunks,
        revision_uid=str(source["revisionUid"]),
    )
    final_evidence = _ordered_evidence_union(
        workbook_evidence,
        chunks=selected_chunks,
        source_order=source_order,
    )
    return {
        "schemaVersion": "canonical-study-manifest-v1",
        "source": {
            **copy.deepcopy(source),
            "contentComplete": True,
        },
        "workbookAnalysis": {
            "key": stable_uid(
                "staged-workbook-analysis-v2",
                source["revisionUid"],
                merged["recordsSha256"],
            ),
            "title": str(
                workbook.get("fileName")
                or Path(str(source.get("sourcePath") or "")).name
            ),
            "summary": (
                f"{len(studies)} source-registered Study record(s) "
                "were projected from complete staged fragments."
            ),
            "status": "NEEDS_REVIEW",
            "verificationStatus": "NEEDS_REVIEW",
            "limitations": list(dict.fromkeys(workbook_limitations)),
            "evidence": final_evidence,
        },
        "studies": studies,
    }


def final_provenance_v2(
    *,
    plan: dict[str, Any],
    registry: dict[str, Any],
    ordered_part_hashes: Sequence[dict[str, str]],
    merged_path: Path,
    merged_sha256: str,
    final_path: Path,
    final_sha256: str,
    generated_at: str,
) -> dict[str, Any]:
    fragment_identity, fragment_identity_sha256 = (
        _require_current_fragment_identity(plan)
    )
    return {
        "schemaVersion": STAGED_FINAL_PROVENANCE_V2_SCHEMA_VERSION,
        "planId": plan["planId"],
        "fragmentIdentity": copy.deepcopy(fragment_identity),
        "fragmentIdentitySha256": fragment_identity_sha256,
        "source": copy.deepcopy(plan["source"]),
        "registrySha256": registry["registrySha256"],
        "orderedPartOutputHashes": copy.deepcopy(
            list(ordered_part_hashes)
        ),
        "mergedArtifactPath": str(merged_path),
        "mergedArtifactSha256": merged_sha256,
        "fragmentContractVersion": FRAGMENT_CONTRACT_VERSION,
        "validatorContractVersion": FRAGMENT_VALIDATOR_VERSION,
        "consolidatorContractVersion": CONSOLIDATOR_CONTRACT_VERSION,
        "finalManifestPath": str(final_path),
        "finalManifestSha256": final_sha256,
        "generatedAt": generated_at,
        "imagesAnalyzed": False,
    }


def final_provenance_v2_matches(
    *,
    provenance: dict[str, Any],
    plan: dict[str, Any],
    registry: dict[str, Any],
    final_sha256: str,
    ordered_part_hashes: Sequence[dict[str, str]] | None = None,
    merged_path: Path | None = None,
    merged_sha256: str | None = None,
    final_path: Path | None = None,
) -> bool:
    fragment_identity = _fragment_identity_v2()
    fragment_identity_sha256 = json_sha256(fragment_identity)
    return bool(
        provenance.get("schemaVersion")
        == STAGED_FINAL_PROVENANCE_V2_SCHEMA_VERSION
        and plan.get("fragmentIdentity") == fragment_identity
        and plan.get("fragmentIdentitySha256")
        == fragment_identity_sha256
        and provenance.get("planId") == plan.get("planId")
        and provenance.get("fragmentIdentity") == fragment_identity
        and provenance.get("fragmentIdentitySha256")
        == fragment_identity_sha256
        and provenance.get("source") == plan.get("source")
        and provenance.get("registrySha256")
        == registry.get("registrySha256")
        and provenance.get("fragmentContractVersion")
        == FRAGMENT_CONTRACT_VERSION
        and provenance.get("validatorContractVersion")
        == FRAGMENT_VALIDATOR_VERSION
        and provenance.get("consolidatorContractVersion")
        == CONSOLIDATOR_CONTRACT_VERSION
        and provenance.get("finalManifestSha256") == final_sha256
        and (
            ordered_part_hashes is None
            or provenance.get("orderedPartOutputHashes")
            == list(ordered_part_hashes)
        )
        and (
            merged_path is None
            or provenance.get("mergedArtifactPath")
            == str(merged_path)
        )
        and (
            merged_sha256 is None
            or provenance.get("mergedArtifactSha256")
            == merged_sha256
        )
        and (
            final_path is None
            or provenance.get("finalManifestPath") == str(final_path)
        )
        and provenance.get("imagesAnalyzed") is False
    )


__all__ = [
    "CONSOLIDATOR_CONTRACT_VERSION",
    "FRAGMENT_CONTRACT_VERSION",
    "FRAGMENT_PROMPT_VERSION",
    "FRAGMENT_VALIDATOR_VERSION",
    "SOURCE_CHUNK_SEGMENT_SCHEMA_VERSION",
    "STAGED_DRAFT_PLAN_V2_SCHEMA_VERSION",
    "STAGED_FINAL_PROVENANCE_V2_SCHEMA_VERSION",
    "STAGED_PART_PROVENANCE_V2_SCHEMA_VERSION",
    "STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION",
    "STUDY_REGISTRY_V2_SCHEMA_VERSION",
    "StagedDraftV2Error",
    "assess_one_call_budget",
    "audit_no_candidate_source_inventory",
    "audit_unselected_source_inventory",
    "build_deterministic_acoustic_matrix_fragment_v2",
    "build_deterministic_error_axis_tail_fragment_v2",
    "build_deterministic_fo_fragment_v2",
    "build_deterministic_function_fragment_v2",
    "build_deterministic_function_grid_fragment_v2",
    "build_deterministic_mask_fragment_v2",
    "build_deterministic_nti_f0_fragment_v2",
    "build_deterministic_nti_horizontal_matrix_fragment_v2",
    "build_deterministic_result_table_fragment_v2",
    "build_fragment_envelope",
    "build_fragment_prompt",
    "build_monolithic_request",
    "build_study_registry_v2",
    "bytes_sha256",
    "chunks_for_part_v2",
    "compact_json_bytes",
    "evidence_cell_keys",
    "final_provenance_v2",
    "final_provenance_v2_matches",
    "finalize_fragment_envelope",
    "fragment_artifact_paths",
    "json_sha256",
    "locators_for_part_v2",
    "merge_adjacent_series_segments",
    "merge_fragment_records",
    "normalize_fragment_missing_observation_arms",
    "normalize_fragment_multi_arm_series_rows",
    "normalize_fragment_observation_replicate_evidence",
    "normalize_fragment_required_fields_and_series_headers",
    "normalize_fragment_unsupported_text_numeric_claims",
    "normalize_fragment_record_ids",
    "normalize_fragment_evidence_dispositions",
    "normalize_fragment_complete_dispositions",
    "part_provenance_v2",
    "part_provenance_v2_matches",
    "plan_study_draft_v2",
    "promote_required_source_locator_sections",
    "project_canonical_manifest",
    "range_bounds",
    "registry_for_part",
    "select_draft_universe",
    "stable_record_id",
    "validate_fragment_v2",
]
