from __future__ import annotations

import json
import hashlib
import math
import os
import re
import shutil
import subprocess
import tempfile
import copy
import uuid
from collections.abc import Callable, Sequence
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from inference_data_ai_arm_identity_repair import (
    ArmIdentityRepairError,
    apply_arm_identity_repair,
    arm_identity_repair_target,
)
from inference_data_ai_b04_b08_repair import (
    B04B08RepairError,
    apply_b04_b08_repair,
    b04_b08_repair_target,
)
from inference_data_ai_b17_report_table_repair import (
    B17ReportTableRepairError,
    apply_b17_report_table_repair,
    b17_report_table_repair_applicable,
)
from inference_data_ai_composite_outcome_repair import (
    CompositeOutcomeRepairError,
    apply_deterministic_composite_outcome_repair,
    composite_outcome_repair_applicable,
)
from inference_data_ai_merged_header_repair import (
    MergedHeaderRepairError,
    apply_merged_header_series_repair,
    merged_header_series_repair_target,
)
from inference_data_ai_numeric_header_repair import (
    NumericHeaderRepairError,
    apply_numeric_header_series_repair,
    numeric_header_series_repair_target,
)
from inference_data_ai_single_outcome_repair import (
    SingleOutcomeRepairError,
    apply_deterministic_single_outcome_repair,
    single_outcome_repair_applicable,
)


LOCATOR_SCHEMA_VERSION = "semantic-locator-v1"
LOCATOR_PROMPT_VERSION = "semantic-locator-prompt-v1"
BATCH_LOCATOR_SCHEMA_VERSION = "semantic-locator-batch-v1"
BATCH_LOCATOR_PROMPT_VERSION = "semantic-locator-batch-prompt-v1"
STUDY_DRAFT_PROMPT_VERSION = "canonical-study-draft-prompt-v25"
STUDY_DRAFT_MAX_INPUT_CHARS = 1_000_000


class SemanticAiError(RuntimeError):
    """Raised when an AI semantic pass fails or returns an invalid contract."""


def locator_output_schema() -> dict[str, Any]:
    """Return the strict JSON schema used for a source-region locator pass."""

    text_array = {
        "type": "array",
        "items": {"type": "string"},
    }
    evidence_item = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "sheet": {"type": "string"},
            "range": {"type": "string"},
            "role": {"type": "string"},
        },
        "required": ["sheet", "range", "role"],
    }
    candidate = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": {"type": "string"},
            "title": {"type": "string"},
            "summary": {"type": "string"},
            "designHint": {"type": "string"},
            "contexts": text_array,
            "changedFactors": text_array,
            "outcomes": text_array,
            "comparisonHints": text_array,
            "evidence": {"type": "array", "items": evidence_item},
            "limitations": text_array,
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
        },
        "required": [
            "key",
            "title",
            "summary",
            "designHint",
            "contexts",
            "changedFactors",
            "outcomes",
            "comparisonHints",
            "evidence",
            "limitations",
            "confidence",
        ],
    }
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": LOCATOR_SCHEMA_VERSION,
            },
            "promptVersion": {
                "type": "string",
                "const": LOCATOR_PROMPT_VERSION,
            },
            "revisionUid": {"type": "string"},
            "contentSha256": {"type": "string"},
            "chunkId": {"type": "string"},
            "status": {
                "type": "string",
                "enum": ["CANDIDATES", "NO_CANDIDATE", "NEEDS_REVIEW"],
            },
            "candidates": {"type": "array", "items": candidate},
            "notes": text_array,
        },
        "required": [
            "schemaVersion",
            "promptVersion",
            "revisionUid",
            "contentSha256",
            "chunkId",
            "status",
            "candidates",
            "notes",
        ],
    }


def batch_locator_output_schema() -> dict[str, Any]:
    """Return a strict wrapper for several independently validated locator results."""

    result_schema = copy.deepcopy(locator_output_schema())
    result_schema.pop("$schema", None)
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": BATCH_LOCATOR_SCHEMA_VERSION,
            },
            "promptVersion": {
                "type": "string",
                "const": BATCH_LOCATOR_PROMPT_VERSION,
            },
            "revisionUid": {"type": "string"},
            "contentSha256": {"type": "string"},
            "results": {
                "type": "array",
                "items": result_schema,
            },
        },
        "required": [
            "schemaVersion",
            "promptVersion",
            "revisionUid",
            "contentSha256",
            "results",
        ],
    }


def study_draft_output_schema() -> dict[str, Any]:
    """Return the strict AI-draft subset of canonical-study-manifest-v1."""

    text = {"type": "string"}
    number_or_null = {"type": ["number", "null"]}
    integer_or_null = {"type": ["integer", "null"]}
    evidence = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "sheet": text,
            "range": text,
            "role": text,
            "sourceText": text,
            "note": text,
        },
        "required": ["sheet", "range", "role", "sourceText", "note"],
    }
    evidence_array = {"type": "array", "items": evidence}
    context = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "kind": text,
            "originalValue": text,
            "normalizedValue": text,
            "unit": text,
            "evidence": evidence_array,
        },
        "required": [
            "key",
            "kind",
            "originalValue",
            "normalizedValue",
            "unit",
            "evidence",
        ],
    }
    factor = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "originalLabel": text,
            "baselineCondition": text,
            "changedCondition": text,
            "changeDirection": text,
            "isolationStatus": {
                "type": "string",
                "enum": ["ISOLATED", "MULTI_FACTOR", "CONFOUNDED", "UNASSESSED"],
            },
            "evidence": evidence_array,
        },
        "required": [
            "key",
            "originalLabel",
            "baselineCondition",
            "changedCondition",
            "changeDirection",
            "isolationStatus",
            "evidence",
        ],
    }
    factor_value = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "factor": text,
            "value": text,
            "valueNumber": number_or_null,
            "unit": text,
            "isBaseline": {"type": "boolean"},
            "heldConstant": {"type": "boolean"},
        },
        "required": [
            "factor",
            "value",
            "valueNumber",
            "unit",
            "isBaseline",
            "heldConstant",
        ],
    }
    arm = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "role": {
                "type": "string",
                "enum": [
                    "CONTROL",
                    "COMPARATOR",
                    "TREATMENT",
                    "TEST",
                    "BEFORE",
                    "AFTER",
                    "REFERENCE",
                    "OTHER",
                ],
            },
            "label": text,
            "condition": text,
            "sampleSize": integer_or_null,
            "sampleBasis": text,
            "matchingBasis": text,
            "factorValues": {"type": "array", "items": factor_value},
            "evidence": evidence_array,
        },
        "required": [
            "key",
            "role",
            "label",
            "condition",
            "sampleSize",
            "sampleBasis",
            "matchingBasis",
            "factorValues",
            "evidence",
        ],
    }
    observation = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "arm": text,
            "valueNumber": number_or_null,
            "valueText": text,
            "numerator": number_or_null,
            "denominator": number_or_null,
            "ratePpm": number_or_null,
            "min": number_or_null,
            "max": number_or_null,
            "average": number_or_null,
            "sampleSize": integer_or_null,
            "evidence": evidence_array,
        },
        "required": [
            "key",
            "arm",
            "valueNumber",
            "valueText",
            "numerator",
            "denominator",
            "ratePpm",
            "min",
            "max",
            "average",
            "sampleSize",
            "evidence",
        ],
    }
    outcome = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "originalLabel": text,
            "metricType": text,
            "unit": text,
            "favorableDirection": {
                "type": "string",
                "enum": ["HIGHER", "LOWER", "TARGET", "NONE", "UNKNOWN"],
            },
            "evidence": evidence_array,
            "observations": {"type": "array", "items": observation},
        },
        "required": [
            "key",
            "originalLabel",
            "metricType",
            "unit",
            "favorableDirection",
            "evidence",
            "observations",
        ],
    }
    comparison = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "comparedArm": text,
            "controlArm": text,
            "designType": text,
            "matchingBasis": text,
            "validityStatus": {
                "type": "string",
                "enum": ["NEEDS_REVIEW", "INVALID", "EXCLUDED"],
            },
            "confoundingStatus": {
                "type": "string",
                "enum": ["POSSIBLE", "CONFOUNDED", "UNASSESSED"],
            },
            "verificationStatus": {
                "type": "string",
                "const": "NEEDS_REVIEW",
            },
            "aggregationEligible": {
                "type": "boolean",
                "const": False,
            },
            "evidence": evidence_array,
            "effects": {
                "type": "array",
                "items": {"type": "string"},
                "maxItems": 0,
            },
        },
        "required": [
            "key",
            "comparedArm",
            "controlArm",
            "designType",
            "matchingBasis",
            "validityStatus",
            "confoundingStatus",
            "verificationStatus",
            "aggregationEligible",
            "evidence",
            "effects",
        ],
    }
    conclusion = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "text": text,
            "claimType": {
                "type": "string",
                "enum": [
                    "SOURCE_CONCLUSION",
                    "AI_DERIVED_DESCRIPTIVE",
                ],
            },
            "causalStrength": {
                "type": "string",
                "enum": ["ASSOCIATION", "DESCRIPTIVE", "UNSPECIFIED"],
            },
            "evidence": evidence_array,
        },
        "required": [
            "key",
            "text",
            "claimType",
            "causalStrength",
            "evidence",
        ],
    }
    study = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "key": text,
            "title": text,
            "purpose": text,
            "hypothesis": text,
            "objective": text,
            "designType": text,
            "comparisonBasis": text,
            "verificationStatus": {
                "type": "string",
                "const": "NEEDS_REVIEW",
            },
            "comparabilityStatus": {
                "type": "string",
                "enum": ["PARTIAL", "INVALID", "UNASSESSED"],
            },
            "confoundingStatus": {
                "type": "string",
                "enum": ["POSSIBLE", "CONFOUNDED", "UNASSESSED"],
            },
            "summary": text,
            "limitations": {"type": "array", "items": text},
            "evidence": evidence_array,
            "contexts": {"type": "array", "items": context},
            "factors": {"type": "array", "items": factor},
            "arms": {"type": "array", "items": arm},
            "outcomes": {"type": "array", "items": outcome},
            "comparisons": {"type": "array", "items": comparison},
            "conclusions": {"type": "array", "items": conclusion},
        },
        "required": [
            "key",
            "title",
            "purpose",
            "hypothesis",
            "objective",
            "designType",
            "comparisonBasis",
            "verificationStatus",
            "comparabilityStatus",
            "confoundingStatus",
            "summary",
            "limitations",
            "evidence",
            "contexts",
            "factors",
            "arms",
            "outcomes",
            "comparisons",
            "conclusions",
        ],
    }
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": "canonical-study-manifest-v1",
            },
            "source": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "dataset": text,
                    "sourcePath": text,
                    "revisionUid": text,
                    "contentSha256": text,
                    "contentComplete": {"type": "boolean"},
                },
                "required": [
                    "dataset",
                    "sourcePath",
                    "revisionUid",
                    "contentSha256",
                    "contentComplete",
                ],
            },
            "workbookAnalysis": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "key": text,
                    "title": text,
                    "summary": text,
                    "status": text,
                    "verificationStatus": {
                        "type": "string",
                        "enum": ["NEEDS_REVIEW", "EXCLUDED"],
                    },
                    "limitations": {"type": "array", "items": text},
                    "evidence": evidence_array,
                },
                "required": [
                    "key",
                    "title",
                    "summary",
                    "status",
                    "verificationStatus",
                    "limitations",
                    "evidence",
                ],
            },
            "studies": {"type": "array", "items": study},
        },
        "required": ["schemaVersion", "source", "workbookAnalysis", "studies"],
    }


def _text(value: object, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise SemanticAiError(f"{field} is required")
    return text


def _a1_bounds(address: str) -> tuple[int, int, int, int]:
    from inference_data_ai_study_import import parse_a1_range

    try:
        return parse_a1_range(address)
    except ValueError as exc:
        raise SemanticAiError(str(exc)) from exc


def _range_is_within(
    inner: tuple[int, int, int, int],
    outer: tuple[int, int, int, int],
) -> bool:
    return (
        inner[0] >= outer[0]
        and inner[1] >= outer[1]
        and inner[2] <= outer[2]
        and inner[3] <= outer[3]
    )


def validate_locator_result(
    result: dict[str, Any],
    *,
    revision_uid: str,
    content_sha256: str,
    chunk: dict[str, Any],
) -> dict[str, Any]:
    """Validate identity and ensure every cited range belongs to the source chunk."""

    if not isinstance(result, dict):
        raise SemanticAiError("locator result must be a JSON object")
    if result.get("schemaVersion") != LOCATOR_SCHEMA_VERSION:
        raise SemanticAiError(f"schemaVersion must be {LOCATOR_SCHEMA_VERSION}")
    if result.get("promptVersion") != LOCATOR_PROMPT_VERSION:
        raise SemanticAiError(f"promptVersion must be {LOCATOR_PROMPT_VERSION}")
    if _text(result.get("revisionUid"), "revisionUid") != revision_uid:
        raise SemanticAiError("locator revisionUid does not match the source packet")
    if _text(result.get("contentSha256"), "contentSha256").lower() != content_sha256.lower():
        raise SemanticAiError("locator contentSha256 does not match the source packet")
    chunk_id = _text(
        chunk.get("chunkId") or chunk.get("packetId"),
        "chunk.chunkId",
    )
    if _text(result.get("chunkId"), "chunkId") != chunk_id:
        raise SemanticAiError("locator chunkId does not match the source packet")
    status = _text(result.get("status"), "status").upper()
    if status not in {"CANDIDATES", "NO_CANDIDATE", "NEEDS_REVIEW"}:
        raise SemanticAiError("invalid locator status")
    candidates = result.get("candidates")
    if not isinstance(candidates, list):
        raise SemanticAiError("candidates must be a list")
    if status == "CANDIDATES" and not candidates:
        raise SemanticAiError("CANDIDATES status requires at least one candidate")
    if status == "NO_CANDIDATE" and candidates:
        raise SemanticAiError("NO_CANDIDATE status cannot contain candidates")

    sheet_value = chunk.get("sheet")
    sheet = _text(
        sheet_value.get("title") if isinstance(sheet_value, dict) else sheet_value,
        "chunk.sheet",
    )
    chunk_bounds = _a1_bounds(
        _text(chunk.get("range") or chunk.get("primaryRange"), "chunk.range")
    )
    context_coordinates = {
        (
            int(cell.get("row") or 0),
            int(cell.get("column") or 0),
        )
        for cell in chunk.get("contextCells", [])
    }
    keys: set[str] = set()
    for candidate_index, candidate in enumerate(candidates):
        if not isinstance(candidate, dict):
            raise SemanticAiError(f"candidates[{candidate_index}] must be an object")
        key = _text(candidate.get("key"), f"candidates[{candidate_index}].key")
        if key in keys:
            raise SemanticAiError(f"duplicate candidate key: {key}")
        keys.add(key)
        for field in ("title", "summary", "designHint"):
            _text(candidate.get(field), f"candidates[{candidate_index}].{field}")
        confidence = candidate.get("confidence")
        if not isinstance(confidence, (int, float)) or isinstance(confidence, bool):
            raise SemanticAiError(f"candidates[{candidate_index}].confidence must be numeric")
        if float(confidence) < 0 or float(confidence) > 1:
            raise SemanticAiError(f"candidates[{candidate_index}].confidence must be between 0 and 1")
        evidence = candidate.get("evidence")
        if not isinstance(evidence, list) or not evidence:
            raise SemanticAiError(f"candidates[{candidate_index}].evidence requires source ranges")
        for evidence_index, item in enumerate(evidence):
            if not isinstance(item, dict):
                raise SemanticAiError(
                    f"candidates[{candidate_index}].evidence[{evidence_index}] must be an object"
                )
            cited_sheet = _text(
                item.get("sheet"),
                f"candidates[{candidate_index}].evidence[{evidence_index}].sheet",
            )
            if cited_sheet != sheet:
                raise SemanticAiError("locator evidence cannot cite a different sheet")
            cited_bounds = _a1_bounds(
                _text(
                    item.get("range"),
                    f"candidates[{candidate_index}].evidence[{evidence_index}].range",
                )
            )
            is_context_cell = (
                cited_bounds[0] == cited_bounds[2]
                and cited_bounds[1] == cited_bounds[3]
                and (cited_bounds[0], cited_bounds[1]) in context_coordinates
            )
            if not _range_is_within(cited_bounds, chunk_bounds) and not is_context_cell:
                raise SemanticAiError("locator evidence cannot cite outside its source chunk")
    return result


def build_locator_prompt(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    chunk: dict[str, Any],
) -> str:
    """Build a domain-neutral locator prompt containing one bounded source chunk."""

    payload = {
        "source": source,
        "workbook": workbook,
        "chunk": _semantic_chunk_view(chunk),
    }
    return (
        "You are the first, source-location-only pass in an evidence database pipeline.\n"
        "Read the supplied sparse Excel cell chunk and identify every region that may contain "
        "a review study, condition change, control/comparison arm, measured outcome, conclusion, "
        "or important context. The domain is open-ended: never use a whitelist of products, "
        "processes, factors, conditions, or outcomes. VP+CD and FUNCTION NG are examples only.\n"
        "Do not calculate effects, decide causality, mark anything verified, or invent missing "
        "values. Preserve unfamiliar wording in the candidate text. Cite only A1 ranges inside "
        "the supplied chunk and on its exact sheet. If the chunk is only raw measurements or "
        "reference data, locate it and state that limitation. Embedded images are out of scope "
        "and must not be requested, extracted, described, or analyzed.\n"
        f"Return exactly {LOCATOR_SCHEMA_VERSION} JSON using promptVersion "
        f"{LOCATOR_PROMPT_VERSION}. Copy revisionUid, contentSha256, and chunkId exactly.\n\n"
        "SOURCE PACKET:\n"
        + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    )


def build_batch_locator_prompt(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    chunks: Sequence[dict[str, Any]],
) -> str:
    """Build one request that locates several chunks without mixing their evidence."""

    if not chunks:
        raise SemanticAiError("batch locator requires at least one source chunk")
    payload = {
        "source": source,
        "workbook": workbook,
        "chunks": [_semantic_chunk_view(chunk) for chunk in chunks],
    }
    return (
        "You are the first, source-location-only pass in an evidence database pipeline.\n"
        "Process every supplied sparse Excel chunk independently. For each chunk, identify "
        "every region that may contain a review study, condition change, control/comparison "
        "arm, measured outcome, conclusion, or important context. The domain is open-ended: "
        "never use a whitelist. VP+CD and FUNCTION NG are examples only.\n"
        "Do not combine evidence across chunks, calculate effects, decide causality, mark "
        "anything verified, or invent missing values. Each nested result must copy its own "
        "chunkId and may cite only A1 ranges inside that chunk or its explicitly supplied "
        "single-cell context on the exact sheet. Return exactly one nested locator result for "
        "every supplied chunk, in the same order, including NO_CANDIDATE results. Embedded "
        "images are out of scope and must not be requested, extracted, described, or analyzed.\n"
        f"Return exactly {BATCH_LOCATOR_SCHEMA_VERSION} JSON using promptVersion "
        f"{BATCH_LOCATOR_PROMPT_VERSION}. Copy the top-level revisionUid and contentSha256 "
        f"exactly. Every nested result must use {LOCATOR_SCHEMA_VERSION} and "
        f"{LOCATOR_PROMPT_VERSION}.\n\n"
        "SOURCE PACKET:\n"
        + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    )


def _semantic_chunk_view(chunk: dict[str, Any]) -> dict[str, Any]:
    """Remove transport-only/repeated formatting fields from an AI prompt."""

    sheet_value = chunk.get("sheet")
    if isinstance(sheet_value, dict):
        sheet = {
            key: sheet_value.get(key)
            for key in (
                "sheetIndex",
                "title",
                "sheetState",
                "status",
                "hasTabularEvidence",
                "contentBounds",
            )
        }
    else:
        sheet = sheet_value
    def compact_cells(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
        cells = []
        for source_cell in items:
            cells.append(
                {
                    key: source_cell.get(key)
                    for key in (
                        "coordinate",
                        "rawValue",
                        "formula",
                        "cachedValue",
                        "displayValue",
                        "dataType",
                        "cachedDataType",
                        "numberFormat",
                        "mergeRange",
                        "mergeRole",
                        "hidden",
                        "valueSource",
                        "primary",
                        "contextOnly",
                    )
                }
            )
        return cells

    cells = compact_cells(chunk.get("cells", []))
    context_cells = compact_cells(chunk.get("contextCells", []))
    return {
        "chunkId": chunk.get("chunkId") or chunk.get("packetId"),
        "sheet": sheet,
        "primaryRange": chunk.get("primaryRange") or chunk.get("range"),
        "bounds": chunk.get("bounds"),
        "sectionIndex": chunk.get("sectionIndex"),
        "splitReason": chunk.get("splitReason"),
        "rowSegment": chunk.get("rowSegment"),
        "truncated": bool(chunk.get("truncated", False)),
        "mergedRanges": chunk.get("mergedRanges", []),
        "cells": cells,
        "contextCells": context_cells,
        "contextPolicy": chunk.get("contextPolicy", {}),
    }


def _study_draft_chunk_view(chunk: dict[str, Any]) -> dict[str, Any]:
    """Keep every supplied coordinate/value while minimizing prompt transport."""

    base = _semantic_chunk_view(chunk)

    def encode_cell(cell: dict[str, Any]) -> dict[str, Any]:
        encoded: dict[str, Any] = {"c": cell.get("coordinate")}
        display = cell.get("displayValue")
        cached = cell.get("cachedValue")
        raw = cell.get("rawValue")
        preferred = (
            display
            if display is not None
            else cached
            if cached is not None
            else raw
        )
        encoded["v"] = preferred
        if raw is not None and raw != preferred:
            encoded["r"] = raw
        if cached is not None and cached != preferred:
            encoded["k"] = cached
        if cell.get("formula"):
            encoded["f"] = cell["formula"]
        number_format = cell.get("numberFormat")
        if number_format and number_format != "General":
            encoded["n"] = number_format
        if cell.get("mergeRange"):
            encoded["m"] = cell["mergeRange"]
        hidden = cell.get("hidden")
        if isinstance(hidden, dict) and any(bool(value) for value in hidden.values()):
            encoded["h"] = {
                key: True for key, value in hidden.items() if bool(value)
            }
        return encoded

    base["cells"] = [encode_cell(cell) for cell in base["cells"]]
    base["contextCells"] = [
        encode_cell(cell) for cell in base["contextCells"]
    ]
    return base


def build_study_draft_prompt(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    locator_results: list[dict[str, Any]],
    focused_chunks: list[dict[str, Any]],
) -> str:
    """Build a conservative whole-workbook Study draft prompt."""

    evidence_shape = {
        "sheet": "<sheet>",
        "range": "<A1 range>",
        "role": "SOURCE",
        "sourceText": "<brief source wording>",
        "note": "",
    }
    canonical_shape = {
        "schemaVersion": "canonical-study-manifest-v1",
        "source": {
            "dataset": "<copy>",
            "sourcePath": "<copy>",
            "revisionUid": "<copy>",
            "contentSha256": "<copy>",
            "contentComplete": "<builder boolean>",
        },
        "workbookAnalysis": {
            "key": "<stable descriptive key>",
            "title": "<source-backed title>",
            "summary": "<consolidated summary>",
            "status": "NEEDS_REVIEW",
            "verificationStatus": "NEEDS_REVIEW",
            "limitations": ["<limitation>"],
            "evidence": [evidence_shape],
        },
        "studies": [
            {
                "key": "<study key>",
                "title": "<title>",
                "purpose": "",
                "hypothesis": "",
                "objective": "",
                "designType": "<open-ended design label>",
                "comparisonBasis": "<source-backed basis or empty>",
                "verificationStatus": "NEEDS_REVIEW",
                "comparabilityStatus": "UNASSESSED",
                "confoundingStatus": "UNASSESSED",
                "summary": "<study summary>",
                "limitations": ["<limitation>"],
                "evidence": [evidence_shape],
                "contexts": [
                    {
                        "key": "<key>",
                        "kind": "<open-ended context kind>",
                        "originalValue": "<source wording>",
                        "normalizedValue": "",
                        "unit": "",
                        "evidence": [evidence_shape],
                    }
                ],
                "factors": [
                    {
                        "key": "<key>",
                        "originalLabel": "<source wording>",
                        "baselineCondition": "<explicit source value or empty>",
                        "changedCondition": "<explicit source value or empty>",
                        "changeDirection": "",
                        "isolationStatus": "UNASSESSED",
                        "evidence": [evidence_shape],
                    }
                ],
                "arms": [
                    {
                        "key": "<key>",
                        "role": "OTHER",
                        "label": "<source wording>",
                        "condition": "<source condition>",
                        "sampleSize": None,
                        "sampleBasis": "",
                        "matchingBasis": "",
                        "factorValues": [
                            {
                                "factor": "<factor key>",
                                "value": "<source value>",
                                "valueNumber": None,
                                "unit": "",
                                "isBaseline": False,
                                "heldConstant": False,
                            }
                        ],
                        "evidence": [evidence_shape],
                    }
                ],
                "outcomes": [
                    {
                        "key": "<key>",
                        "originalLabel": "<source wording>",
                        "metricType": "<open-ended metric type>",
                        "unit": "",
                        "favorableDirection": "UNKNOWN",
                        "evidence": [evidence_shape],
                        "observations": [
                            {
                                "key": "<key>",
                                "arm": "<arm key>",
                                "valueNumber": None,
                                "valueText": "",
                                "numerator": None,
                                "denominator": None,
                                "ratePpm": None,
                                "min": None,
                                "max": None,
                                "average": None,
                                "sampleSize": None,
                                "evidence": [evidence_shape],
                            }
                        ],
                    }
                ],
                "measurementSeries": [
                    {
                        "key": "<series key>",
                        "seriesRole": "<RAW or AGGREGATE>",
                        "aggregationFunction": (
                            "<AVERAGE for AGGREGATE, otherwise empty>"
                        ),
                        "aggregateOfSeries": [
                            "<RAW series keys averaged by an AGGREGATE series>"
                        ],
                        "outcome": "<outcome key>",
                        "arm": "<arm key>",
                        "sheet": "<exact sheet>",
                        "headerRange": "<one-row axis/header range>",
                        "valueRange": "<numeric matrix range>",
                        "rowIdentityRange": (
                            "<one-column row/replicate identity range>"
                        ),
                        "aggregateReplicateRanges": [
                            "<exact AVG/MEAN replicate identity ranges, or empty>"
                        ],
                        "axisSource": "<HEADER or ROW_IDENTITY>",
                        "axisLabel": "<open-domain axis label>",
                        "axisUnit": "<axis unit or empty>",
                        "valueUnit": "<value unit or empty>",
                        "stratumKey": "<optional fixed stratum or empty>",
                        "verificationStatus": "NEEDS_REVIEW",
                    }
                ],
                "comparisons": [
                    {
                        "key": "<key>",
                        "comparedArm": "<arm key>",
                        "controlArm": "<reference arm key>",
                        "designType": "<design label>",
                        "matchingBasis": "<basis or empty>",
                        "validityStatus": "NEEDS_REVIEW",
                        "confoundingStatus": "UNASSESSED",
                        "verificationStatus": "NEEDS_REVIEW",
                        "aggregationEligible": False,
                        "evidence": [evidence_shape],
                        "effects": [],
                    }
                ],
                "conclusions": [
                    {
                        "key": "<key>",
                        "text": "<only a source-backed descriptive statement>",
                        "claimType": (
                            "<SOURCE_CONCLUSION only for exact source "
                            "narrative; otherwise AI_DERIVED_DESCRIPTIVE>"
                        ),
                        "causalStrength": "DESCRIPTIVE",
                        "evidence": [evidence_shape],
                    }
                ],
            }
        ],
    }
    payload = {
        "source": source,
        "workbook": workbook,
        "locatorResults": locator_results,
        "focusedCellEncoding": {
            "c": "A1 coordinate",
            "v": "preferred displayed/cached/raw value",
            "r": "raw value when different from v",
            "k": "cached formula value when different from v",
            "f": "formula",
            "n": "non-General number format",
            "m": "merged range",
            "h": "hidden sheet/row/column flags",
        },
        "focusedChunks": [
            _study_draft_chunk_view(chunk)
            for chunk in focused_chunks
        ],
    }
    return (
        "You are the semantic drafting pass for an evidence-traceable Excel review database.\n"
        "Produce one consolidated workbook analysis and zero or more distinct studies. The "
        "domain is open-ended. Never restrict extraction to known models, products, processes, "
        "factors, materials, equipment, conditions, or outcomes. VP+CD and FUNCTION NG are only "
        "examples of possible questions.\n"
        "Use only supplied cells. Preserve unfamiliar original wording. Every factor, numeric "
        "observation, comparison, and conclusion must cite exact A1 evidence. Do not infer a "
        "SOURCE_CONCLUSION from condition labels, numeric observations, calculated differences, "
        "or workbook limitations. Use claimType SOURCE_CONCLUSION only when an exact cited "
        "captured text cell contains an explicit narrative decision/conclusion supporting the "
        "claim, and copy that literal narrative into the matching evidence.sourceText. If the "
        "wording is your descriptive synthesis of source numbers, use claimType "
        "AI_DERIVED_DESCRIPTIVE with causalStrength DESCRIPTIVE or omit the conclusion; keep all "
        "numeric observations either way. "
        "Do not infer a control merely from column position. A literal whole-cell source label "
        "Normal or Normal (...) maps to arm.role REFERENCE, never CONTROL by itself. Use "
        "CONTROL only when the Arm evidence directly cites an exact captured Arm label or "
        "condition containing explicit Control wording. comparison.controlArm may reference a "
        "REFERENCE Arm without changing that Arm's role. Use REFERENCE only when exact Arm "
        "evidence and the matching label/condition contain full Normal, Reference, Standard, "
        "Spec, or equivalent reference semantics. A bare abbreviation such as ST is not "
        "Standard and must remain OTHER or a source-supported COMPARATOR. A descriptive grouped "
        "REFERENCE label such as 'Normal #1 through Normal #10' is allowed only when its exact "
        "Arm evidence contains at least two nonempty identity cells, every cell is a full "
        "Normal/Reference/Standard/Spec #N identity, and the #N values are ordered and distinct. "
        "Mixed Test/Normal evidence, merged-cell inference, or a lone Normal token cannot create "
        "that group. Preserve every replicate identity, measurement axis, and exact evidence "
        "cell; grouping never authorizes collapsing or sampling replicates. A literal label "
        "Normal still requires an exact cited Normal cell. When a row's exact identity cell is "
        "Normal and a separate governing merged cell supplies an entity such as VP+Coil, use "
        "Normal as the Arm label and condition and keep the merged entity as a Factor or Context; "
        "never synthesize a composite REFERENCE Arm identity such as 'VP+Coil Normal'. When the "
        "source explicitly labels an "
        "arm as Control, Reference, Baseline, Before, or equivalent open-domain wording, "
        "preserve the supported role and draft source-backed NEEDS_REVIEW comparisons for test arms "
        "that share the same outcomes. Do not omit an explicit comparison merely because it "
        "still requires human verification. Separate studies or mark confounding when model, "
        "lot, period, line, unit, denominator, baseline, or multiple changed conditions differ. "
        "arm.role must be exactly one of CONTROL, COMPARATOR, TREATMENT, TEST, BEFORE, AFTER, "
        "REFERENCE, or OTHER. BASELINE is not an allowed role: use CONTROL only when the source "
        "explicitly establishes a control. Use BEFORE or AFTER only when the cited captured "
        "source cell's raw/display value or custom number_format explicitly labels that Arm "
        "Before/After or pre-change/post-change. Locator titles, "
        "summaries, design hints, comparison hints, filenames, column order, repeated specimen "
        "labels, and adjacent measurement blocks are not source labels and cannot authorize "
        "BEFORE/AFTER. If no cited captured cell value or number format supplies the temporal "
        "label, do not put "
        "Before, After, pre-change, or post-change in a Study title, Arm label/condition/role, "
        "measurementSeries stratumKey, or comparison. Repeated #1..#N blocks under the same "
        "pressure or condition belong to one shared condition Arm only when the captured cells "
        "do not provide distinct source-authored stage labels. Preserve unlabeled blocks "
        "as separate measurementSeries or strata under that same Arm. Use a literal "
        "source-backed block/run label when present; otherwise use only a neutral identity such "
        "as 'Block 1', 'Block 2', or the exact source coordinates. Never rename an unlabeled "
        "block Before/After. Keep the source-stated sampleSize unchanged (for example, 10 stays "
        "10); repeated measurements never inflate it to 20. "
        "A captured Excel custom number_format is source-authored display meaning. Quoted or "
        "literal format tokens such as pressure, replicate identity, Before/After, or pre/post "
        "may be used only for the exact captured cells and evidence ranges carrying that format. "
        "For example, raw header values 1..10 with a custom format that displays an 18kPa "
        "replicate and Before/After label preserve that pressure, replicate, and temporal "
        "meaning. Do not treat a locator summary, AI sourceText, filename, or a custom format on "
        "an unrelated cell/range as evidence for another Arm, Study, series, or comparison. "
        "An evidence-linked number-format Before/After token authorizes that Arm's BEFORE/AFTER "
        "role, label, and condition, but never authorizes CONTROL or BASELINE. Keep the physical "
        "sampleSize unchanged for each stage. A NEEDS_REVIEW paired comparison is allowed only "
        "when both BEFORE and AFTER RAW series actually contain values under the same pressure, "
        "their axes align, and both formatted replicate identities are the same ordered #1..#N. "
        "Use matchingBasis that cites the pressure and formatted identities, validityStatus "
        "NEEDS_REVIEW, confoundingStatus UNASSESSED, verificationStatus NEEDS_REVIEW, "
        "aggregationEligible false, and effects empty. This matched identity does not establish "
        "the treatment protocol, causal validity, or effect. Never compare a header-only or "
        "missing stage. Preserve a 300 After header-only Arm when its formatted headers exist, "
        "but create no numeric series or comparison for it. "
        "Across independent dose or condition Arms, repeated sample labels 1..N do not prove "
        "paired observations or the same physical specimens. Treat those Arms as independent "
        "unless exact source evidence explicitly links each physical specimen across Arms. "
        "When source-authored BEFORE and AFTER Arms carry the same measurement-stage factor, "
        "preserve the temporal direction in that factor: baselineCondition must be Before, "
        "changedCondition must be After, the BEFORE Arm's matching factorValue.isBaseline must "
        "be true, and the AFTER Arm's matching factorValue.isBaseline must be false. This is "
        "factor-level temporal identity only. It never changes either Arm to CONTROL or "
        "BASELINE, never verifies a comparison, and never makes aggregation eligible. "
        "A compound source label such as '100kPa - 18kPa (2nd)' does not make the plain "
        "18kPa arm a COMPARATOR and does not prove a comparison protocol. Preserve the compound "
        "row descriptively with role OTHER. Keep current pressure, prior exposure/condition, "
        "and measurement order as separate source-backed factors or contexts; never collapse "
        "them into one generic pressure factor. Without an explicit pairing/comparison protocol, "
        "do not compare a second-measurement row with an initial-condition cohort, even as a "
        "non-aggregatable comparison. "
        "Do not calculate effects. Set every study and comparison verificationStatus to "
        "NEEDS_REVIEW, every aggregationEligible to false, and every effects list to empty. "
        "Never collapse several numeric submetrics into one valueText. Input/sample counts "
        "belong in arm.sampleSize and the corresponding observation.sampleSize. An explicit "
        "Input, sample, or cohort-size Outcome is denominator-only and must use metricType "
        "'sample_size'; preserve the source count in valueNumber and sampleSize, but set both "
        "numerator and denominator null because it is the denominator itself, not a rate pair. "
        "It must never be treated as a continuous effect. For a count Outcome "
        "whose source row supplies an explicit Input/denominator, preserve valueNumber as the "
        "count and also supply numerator=count, denominator=Input, and sampleSize=Input. Cite "
        "both the count cell and denominator cell, use a semantic count metricType such as "
        "'defect_count' or 'success_count', and never leave a calculable count as an "
        "un-denominated continuous value. Numerator and denominator must always be supplied "
        "together from their exact cited source cells or both left null. Never invent a missing "
        "denominator; when only a raw count is supported, preserve it in valueNumber/valueText "
        "and clear the unsupported numerator/denominator pair. A whole-cell explicit count "
        "ratio such as '1/8 pcs' "
        "or '1/8 EA' must cite that exact cell and preserve valueNumber and numerator as 1, "
        "denominator and sampleSize as 8, with a count metricType. Do not parse numbers out of "
        "narrative text, ranges, or unit-bearing specifications as count evidence. "
        "Defect-category counts from the same samples may overlap. If Total NG does not equal "
        "the arithmetic sum of its component category counts, do not infer mutual exclusivity, "
        "do not recompute either value, and do not treat the components as additive. Preserve "
        "Total NG and every component category as separate Outcomes and record the overlap or "
        "non-additivity limitation. If Input is not equal to OK plus Total NG, preserve all "
        "three literal source values, record the exact unreconciled residual as a limitation, "
        "and never correct, impute, or reclassify the unexplained specimens. "
        "For a source percentage or rate with an explicit "
        "count and Input, supply numerator and denominator, cite rate/count/Input cells, and "
        "express valueNumber in the human displayed unit scale (3.8 for 3.8%, never the Excel "
        "storage fraction 0.038 with unit %). For a percent-formatted numeric source cell, "
        "valueNumber must be the exact underlying captured numeric value multiplied by 100, "
        "not the rounded screen-display string. Preserve a rounded display such as 22.1% only "
        "in valueText or evidence sourceText. When numerator and denominator are also present, "
        "the numeric percentage claim must agree with their exact arithmetic; never use display "
        "rounding as a numeric claim. Numeric NG "
        "components such as a count, denominator, rate, category, or condition must remain "
        "separate Outcomes/Observations with their own exact cells. "
        "A percentage-only total-rate cell plus an Input cell and separate category-count cells "
        "does not prove a total numerator. Preserve the percentage value, display text, rate "
        "evidence, sampleSize, and category observations, but leave numerator and denominator "
        "both null unless an exact cited total-count cell supplies the numerator. Never sum "
        "category counts to manufacture it. "
        "For a wide repeated numeric "
        "curve or matrix, do not replace source values with text such as 'profile supplied'; "
        "declare the compact source range mapping supported by the manifest contract so the "
        "deterministic importer can expand every numeric point. If a table cannot be mapped "
        "without guessing its axis, arm, or replicate identities, retain the source range and "
        "state the limitation instead of inventing a summary value. A measurementSeries entry "
        "must reference an existing outcome and arm. headerRange must be exactly one row with "
        "the same exact columns as valueRange; rowIdentityRange must be exactly one column with "
        "the same exact rows as valueRange. Set axisSource to HEADER when the column headers are "
        "the measurement axis and the row identities are replicate/specimen identities. Set it "
        "to ROW_IDENTITY when row identities are the measurement axis and column headers are "
        "replicate identities. Every raw stage series must declare seriesRole RAW, an empty "
        "aggregationFunction, an empty aggregateOfSeries, and only its own numeric stage range. "
        "For a frequency profile matrix whose rows are frequencies and whose columns are "
        "replicate headers #1..#N (for example, row identities A3:A16 and raw replicate values "
        "F3:O16), the RAW series must use axisSource ROW_IDENTITY, the frequency cells as "
        "rowIdentityRange, and the #1..#N cells as headerRange. A source AVG column such as "
        "P3:P16 is a separate AGGREGATE series with aggregationFunction AVERAGE, the same "
        "frequency rowIdentityRange and axisSource ROW_IDENTITY, aggregateOfSeries naming the "
        "exact RAW matrix series, and aggregateReplicateRanges empty. Never reinterpret AVG as "
        "a replicate or use its header as the measurement axis. "
        "AVG/MEAN summary values that pool multiple raw Arms are standalone measurementSeries "
        "with seriesRole AGGREGATE, aggregationFunction AVERAGE, aggregateOfSeries listing the "
        "exact RAW series keys, and aggregateReplicateRanges empty. Put a pooled aggregate under "
        "a distinct pressure-level summary Arm with role OTHER; never attach it to either raw "
        "BEFORE or AFTER Arm. The aggregate series valueRange contains only the source AVG "
        "values, while its headerRange and rowIdentityRange preserve their exact source axes. "
        "For IMP/Fo at 18/100/200 pressure, keep BEFORE and AFTER RAW series separate and make "
        "the pressure summary AVG series reference both. For a 300 Fo AVG whose AFTER values are "
        "blank, reference only the populated 300 BEFORE RAW series. Do not create an aggregate "
        "series when the AVG source values are blank. For a single-row Fo RAW or AVG series "
        "whose formatted column headers are replicate identities, use axisSource ROW_IDENTITY: "
        "the Fo row identity is the shared measurement axis and the formatted column headers "
        "remain replicate identities. Do not use axisSource HEADER for those Fo series, because "
        "an AVG header is not a raw measurement axis. Legacy AVG identities that coexist inside "
        "one raw Arm may still use aggregateReplicateRanges, but never use that field on a "
        "standalone AGGREGATE series. In contrast, for a vertical one-column Fo series such as "
        "C5:C14 whose row labels are #1..#10 and whose governing merged header supplies the "
        "shared Fo measurement identity, use axisSource HEADER: the header is the shared "
        "measurement axis and the row labels are replicateKey values. A source AVG row such as "
        "C15 must remain on that same header axis, either as an aggregateReplicateRanges identity "
        "inside the legacy series or as a standalone AGGREGATE with axisSource HEADER and "
        "aggregateOfSeries referencing the RAW series. Never use ROW_IDENTITY for this vertical "
        "raw/AVG layout. This does not change horizontal frequency-row matrices, which remain "
        "axisSource ROW_IDENTITY. Every cell implied by a measurementSeries.valueRange must contain a usable "
        "numeric value. Never widen a dense valueRange across blank, text, malformed, or error "
        "cells. Select only the actual data-bearing columns, split non-contiguous blocks into "
        "separate series, or preserve a standalone summary as an Observation. A repeated REF "
        "header over several columns does not make blank REF columns numeric: include only the "
        "data-bearing REF column and leave blank sibling columns out. "
        "When the same specimen labels repeat in multiple measurement blocks without distinct "
        "source-authored stage labels, keep one condition Arm, keep the source sample size "
        "unchanged, and preserve each block/run as a separate series or stratum; never add the "
        "repeated labels together as independent specimens. When captured cell values or custom "
        "number formats explicitly label the stages, preserve those stage identities instead of "
        "neutral Block names. A label such as "
        "REF is descriptive REFERENCE/OTHER data unless the source "
        "explicitly defines it as the control; do not auto-create control comparisons. Do not "
        "fill a missing arm/outcome row with zero, repair malformed numeric text, or include "
        "broken/error cells in a numeric range. Split around an invalid cell or retain an "
        "explicit limitation. Do not invent conventional units that are absent from the cells. "
        "Keep a composite source label as one Outcome unless the source supplies separate "
        "component values. Never guess or silently transpose these roles. "
        "Attach workbook-level model, report date, and applicable measurement setup such as "
        "voltage, power, and distance directly to every data-bearing Study they govern. Do not "
        "strand shared setup in a context-only Study while SPL, THD, IMP, Fo, or other numeric "
        "Studies omit it. If applicability is uncertain, preserve the context with a limitation "
        "rather than inventing a value. When a numeric row axis such as 100..14000 has no literal "
        "semantic label or unit, use a neutral axisLabel such as 'source numeric axis' and an "
        "empty axisUnit. Never reinterpret that axis as the outcome name or a corrupted generic "
        "header such as sample/specimen. If a filename identifier differs from a literal report "
        "identifier, preserve both exact identities and record the mismatch as a limitation; "
        "never silently normalize them into one model. "
        "When several scalar raw values share one outcome and arm, each Observation must have a "
        "unique source-backed replicateKey (and block stratum when labels repeat), or the values "
        "must be represented by a valid measurementSeries. Never emit duplicate empty "
        "arm/outcome/stratum/replicate identities. Preserve single-row raw measurements such as "
        "Fo and Air-leak samples; keep MIN/MAX/AVG as aggregate fields or aggregate identities, "
        "never as extra raw specimens. If a MIN/MAX/AVG formula has no cached value, never "
        "invent its numeric result or parse the formula text as numeric evidence. Preserve the "
        "raw measurementSeries, cite the exact formula cell/range as formula lineage, keep the "
        "derived numeric field null or omit the aggregate series, and record the missing-cache "
        "limitation. "
        "Use one entry per "
        "arm/value matrix and cite the exact sheet/ranges; never omit or sample points. "
        "When non-reference arms repeat the same entity and context and exactly one explicit "
        "factor changes, draft only the aligned one-to-one NEEDS_REVIEW comparison (for example, "
        "the same mold at 180 C versus 190 C); never create a cross-product of mismatched arms. "
        "A source title or purpose that explicitly identifies a DOE/comparison and supplies a "
        "multi-condition result table establishes comparison intent even when no Arm is labeled "
        "Control. Draft every unambiguous exactly-one-factor aligned pair as NEEDS_REVIEW with "
        "aggregationEligible false and effects empty. Using one Arm in the structural "
        "controlArm field does not change that Arm's role to CONTROL or verify the comparison. "
        "Every Comparison must have at least one shared Outcome represented by both Arms. For "
        "each shared Outcome, never compare one Arm's RAW measurementSeries with only scalar or "
        "summary Observations on the other Arm. RAW-versus-RAW requires compatible value units "
        "and the same ordered axis identity, shape, and stratum; scalar-versus-scalar requires "
        "a compatible shared quantitative field or qualitative-only values on both Arms. Equal "
        "replicate or sample labels do not establish physical pairing; this representation gate "
        "does not prove pairing, matching, validity, or causality. NEEDS_REVIEW does not waive "
        "representation compatibility. Omit an incompatible Comparison, preserve all Arms, "
        "Outcomes, Observations, and measurementSeries, and add the mismatch as a limitation. "
        "Do not draft a pair when two or more factors differ. For an ordered dose or interval "
        "series with no explicit baseline/reference Arm, compare only adjacent source-order "
        "conditions; never create an all-pairs dose cross-product. "
        "Preserve every fixed condition stated by the source, including cure/dry temperature, "
        "duration, agent percentage, line, lot, model, and equipment, as a factor or context. "
        "When one source cell states independently queryable condition components such as "
        "drying temperature and drying duration, preserve them as separate Factors and factor "
        "values while retaining the exact combined source text and evidence; do not collapse "
        "them into one generic Drying factor. "
        "For an Arm factorValue whose exact cited whole-cell source value is one recognized "
        "quantity such as '1.56mg' or '1.56 mg', preserve that complete literal in value, set "
        "valueNumber to 1.56, and set unit to the resolvable source unit mg. Use only actual "
        "cells inside that Arm's or Factor's evidence. Never parse ranges, narrative text, "
        "dates, model IDs, ratios, cells with multiple numbers, unknown units, filenames, or "
        "evidence.sourceText as factor quantities. If a quantity-looking token is embedded in "
        "a composite whole-cell condition such as 'Test Led UC (VP+CD) 5s' or 'Frame clean by "
        "ethanol + Drying temperature 80°C time 5min', do not extract or normalize 5s, 80°C, "
        "or 5min. Preserve one compound Factor/factorValue whose value copies the exact cited "
        "whole-cell narrative with valueNumber null, unit empty, and isolationStatus UNASSESSED, "
        "or omit unsupported component factorValues. Never isolate an ordinal/run token such "
        "as 1st, 2nd, or Total from a composite row label unless that token exists as its own "
        "exact captured whole cell. "
        "A cited cell whose complete raw text is a strict signed integer, decimal, or scientific "
        "number (for example '-0.006' or '1.2e-3') is numeric evidence even when Excel stores it "
        "as text: preserve the exact source string in valueText and set valueNumber to the exact "
        "same number. This exception applies only to the entire cell and exact cell evidence; "
        "never parse narratives, ranges, ratios, IDs, dates, malformed numbers, Pass/Fail, or "
        "other categorical text as numeric. Never mix qualitative-only Pass/Fail "
        "observations into a quantitative Outcome with numeric observations or "
        "measurementSeries; split them into a distinct categorical Outcome. "
        "Repeated per-replicate PASSED/FAILED result cells are categorical Outcome observations "
        "with exact cell evidence and source-backed replicateKey values. Never promote them to "
        "SOURCE_CONCLUSION or rewrite them as AI narrative, and never infer a numeric threshold, "
        "spec limit, or acceptance rule that the source does not state. "
        "If a Normal/reference arm omits or differs on one of those conditions, mark the "
        "comparison confounded or unassessed instead of attributing the result to one factor. "
        "Before returning JSON, verify that every factorValues.factor, observation.arm, "
        "measurementSeries.outcome, measurementSeries.arm, comparison arm, and effect outcome "
        "reference exactly matches a declared key, including spelling and case. "
        "A measurement unit may come from the exact group header that directly governs a raw "
        "numeric column even when the leaf header omits the unit. Preserve that explicit unit "
        "and cite the governing group header; never borrow a unit from an unrelated block. "
        "Never claim causality in an AI draft. Use null for absent numeric values and an empty "
        "string/list for absent text/collections. Embedded images are out of scope and must not "
        "be requested, extracted, described, or analyzed. For every numeric observation, its "
        "evidence list must include the actual cells containing each supplied valueNumber, "
        "numerator, denominator, ratePpm, min, max, or average. Use multiple tight ranges when "
        "those values are non-contiguous; do not copy a denominator or other number whose cell "
        "is absent from that observation's evidence.\n"
        f"Return only canonical-study-manifest-v1 JSON with no Markdown or commentary, "
        f"drafted under {STUDY_DRAFT_PROMPT_VERSION}. "
        "Copy the source identity exactly. source.contentComplete must be true only if the "
        "workbook packet explicitly says semantic cell coverage is complete.\n\n"
        "Use exactly these property names and nesting. The following is a shape "
        "template, not source data; replace placeholders and omit list items that "
        "are not supported by the source:\n"
        + json.dumps(canonical_shape, ensure_ascii=False, separators=(",", ":"))
        + "\n\nFOCUSED SOURCE PACKET:\n"
        + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    )


def build_study_draft_repair_prompt(
    rejected: dict[str, Any],
    validation_error: str,
    *,
    source_prompt: str = "",
) -> str:
    repair_task = (
        "You are repairing a rejected canonical-study-manifest-v1 JSON draft. "
        "Return the complete corrected JSON object only, with no Markdown or "
        "commentary. Correct the validator error and every related unsupported semantic "
        "inference visible in the rejected draft. Preserve all otherwise valid source values, exact "
        "A1 ranges, source identity, study separation, review statuses, and "
        "measurementSeries mappings. Do not add facts, calculate effects, "
        "self-approve records, infer controls, or use images. Every reference "
        "must exactly match a declared key, including spelling and case. The "
        "literal focused source cells are authoritative; locator summaries and the rejected "
        "draft are not source evidence. The exact captured cell's custom number_format is also "
        "authoritative only for quoted/literal pressure, replicate, Before/After, or pre/post "
        "display tokens inside the cited evidence range. A literal whole-cell Normal or Normal "
        "(...) Arm label must be REFERENCE, never CONTROL by itself. CONTROL requires an exact "
        "captured Arm label or condition with explicit Control wording inside that Arm evidence; "
        "comparison.controlArm may still reference REFERENCE. REFERENCE requires exact cited "
        "full Normal, Reference, Standard, Spec, or equivalent reference semantics matching the "
        "Arm label/condition. Never expand a bare ST abbreviation into Standard; repair that Arm "
        "to OTHER or a source-supported COMPARATOR. Repair a descriptive grouped REFERENCE such "
        "as 'Normal #1 through Normal #10' only when exact Arm evidence contains at least two "
        "nonempty cells, all are full Normal/Reference/Standard/Spec #N identities, and their "
        "#N values are ordered and distinct. Reject mixed Test/Normal cells, merge inference, "
        "or a lone token. Preserve each replicate axis identity and exact evidence; never "
        "collapse the group. A literal Normal label still requires an exact Normal cell. When "
        "a row's exact identity cell is Normal and a separate governing merged cell supplies an "
        "entity such as VP+Coil, repair the Arm label and condition to Normal and preserve the "
        "merged entity only as its exact Factor or Context; never concatenate the cells into a "
        "synthetic REFERENCE identity such as 'VP+Coil Normal'. For an exact cited whole-cell "
        "factor quantity such as '1.56mg' or '1.56 mg', preserve the complete source text in "
        "factorValue.value, set valueNumber to 1.56, and use the resolvable source unit mg. Never "
        "derive factor quantities from ranges, narratives, dates, model IDs, ratios, multiple "
        "numbers, unknown units, filenames, or evidence.sourceText. A cited cell whose complete "
        "raw text is a strict signed integer, decimal, or scientific number is numeric evidence "
        "even when Excel stores it as text: preserve the exact source string in valueText and "
        "set valueNumber to the exact same number. Change no other observation field for this "
        "repair. Never parse narratives, ranges, ratios, IDs, dates, malformed numbers, "
        "Pass/Fail, or other categorical text as numeric. Never mix genuine qualitative-only "
        "Pass/Fail observations into a quantitative Outcome with numeric observations or "
        "measurementSeries; split them into a distinct categorical Outcome. "
        "Repair repeated per-replicate PASSED/FAILED cells as categorical Outcome observations "
        "with exact cell evidence and source-backed replicateKey values. Never use them as "
        "SOURCE_CONCLUSION or AI narrative and never infer an unstated threshold, spec limit, "
        "or acceptance rule. "
        "When a quantity-looking token is embedded in a composite whole-cell condition, never "
        "repair it into a standalone numeric factor quantity. Preserve one compound Factor and "
        "factorValue whose value is the exact cited whole-cell narrative, valueNumber is null, "
        "unit is empty, and isolationStatus is UNASSESSED, or omit unsupported component "
        "factorValues. This applies to combined time labels such as 'Test Led UC (VP+CD) 5s' "
        "and multi-condition narratives such as 'Frame clean by ethanol + Drying temperature "
        "80°C time 5min'. Likewise, never isolate an ordinal/run token such as 1st, 2nd, or "
        "Total from a composite row label unless that token exists as its own exact captured "
        "whole cell; keep the exact compound factor value and omit the unsupported component. "
        "A SOURCE_CONCLUSION requires an exact "
        "captured narrative decision/conclusion cell copied into evidence.sourceText and directly "
        "supporting claim text. Condition labels, numbers, calculated differences, and a "
        "limitation saying no source narrative exists cannot support SOURCE_CONCLUSION. Reclassify "
        "such synthesis as AI_DERIVED_DESCRIPTIVE with causalStrength DESCRIPTIVE or remove only "
        "the unsupported conclusion; never remove its numeric observations, and never transfer a "
        "format label or custom number_format label from an unrelated cell or range. An exact evidence-linked "
        "custom number_format directly "
        "authorizes that Arm's BEFORE/AFTER role, label, and condition, but it never authorizes "
        "CONTROL or BASELINE. When no cited raw/display value or custom number_format supplies "
        "a stage label, merge repeated #1..#N blocks for the same pressure or condition under "
        "one shared Arm, retain the source sampleSize (10 remains 10), and distinguish repeated "
        "blocks only with a literal source-backed run label or neutral Block "
        "1/Block 2/source-coordinate stratum. Never convert unlabeled repeated blocks into "
        "inferred phases or independent sampleSize 20. "
        "A compound second-measurement label does not establish the plain condition as a "
        "COMPARATOR or prove cohort pairing. Keep such rows descriptive with role OTHER, "
        "separate current pressure, prior exposure/condition, and order dimensions, and delete "
        "unsupported comparisons. "
        "Every repaired Comparison must retain at least one shared Outcome represented by both "
        "Arms. Reject RAW measurementSeries versus scalar/summary representation. RAW-versus-RAW "
        "requires compatible value units and aligned ordered axis identity, shape, and stratum; "
        "scalar-versus-scalar requires a compatible shared quantitative field or both sides "
        "qualitative-only. Equal labels do not prove pairing, and NEEDS_REVIEW does not waive "
        "this representation gate. Omit the incompatible Comparison only, preserve its Arms, "
        "Outcomes, Observations, and series, and add a limitation describing the mismatch. "
        "Propagate applicable report model/date and measurement setup "
        "(including voltage, power, and distance) into each data-bearing Study rather than a "
        "detached context-only Study. For an unlabeled numeric axis, use neutral 'source numeric "
        "axis' wording with empty unit; do not call it SPL, sample, or specimen. Preserve any "
        "filename-versus-report identifier mismatch as an explicit limitation. "
        "Every raw stage series must use seriesRole RAW, an empty aggregationFunction, an empty "
        "aggregateOfSeries, and only that stage's numeric values. Repair a source AVG/MEAN that "
        "pools raw Arms as a standalone seriesRole AGGREGATE measurementSeries with "
        "aggregationFunction AVERAGE, aggregateOfSeries listing the exact RAW series keys, "
        "aggregateReplicateRanges empty, and only the AVG source cells in valueRange. Put this "
        "pooled aggregate under a distinct pressure summary Arm with role OTHER, never under a "
        "BEFORE or AFTER Arm. For IMP/Fo at 18/100/200 pressure preserve separate BEFORE and "
        "AFTER RAW series and make the summary AVG reference both. For a 300 Fo AVG with blank "
        "AFTER values, reference the populated 300 BEFORE RAW series only. Do not create a "
        "series for blank AVG cells. For every single-row Fo RAW or AVG series with formatted "
        "replicate headers, use axisSource ROW_IDENTITY so the Fo row identity is the shared "
        "measurement axis and the column headers remain replicate identities; never treat the "
        "AVG header as a distinct measurement axis. "
        "For a frequency-by-replicate matrix such as frequency row identities A3:A16, #1..#10 "
        "headers F2:O2, raw values F3:O16, and AVG values P3:P16, repair the RAW matrix to "
        "axisSource ROW_IDENTITY with the frequency rowIdentityRange and replicate headerRange. "
        "Repair AVG as a separate AGGREGATE series with aggregationFunction AVERAGE, the same "
        "frequency rowIdentityRange and axisSource ROW_IDENTITY, aggregateOfSeries naming that "
        "exact RAW series, and aggregateReplicateRanges empty. Do not relax or work around the "
        "aggregate validator. "
        "For a vertical one-column Fo layout such as C5:C14 with #1..#10 row labels and a "
        "governing merged Fo header, repair RAW to axisSource HEADER so the header is the shared "
        "measurement axis and row labels are replicateKey values. Keep a source AVG row C15 on "
        "that same header axis, either via aggregateReplicateRanges in the legacy series or a "
        "standalone AGGREGATE using axisSource HEADER and aggregateOfSeries. Never repair this "
        "vertical raw/AVG layout to ROW_IDENTITY. Horizontal frequency-row matrices remain "
        "ROW_IDENTITY and must not be transposed into this rule. "
        "Across independent dose or condition Arms, repeated sample labels 1..N do not prove "
        "the same physical specimens and must not create a paired comparison without exact "
        "source pairing evidence. "
        "A whole-cell explicit count ratio such as '1/8 pcs' or '1/8 EA' must cite that exact "
        "cell and preserve numerator/valueNumber 1, denominator/sampleSize 8, and a count "
        "metricType; never extract count evidence from narrative text or numeric ranges. "
        "For an Input/sample/cohort-size Outcome with metricType sample_size, preserve the "
        "source count in valueNumber and sampleSize while setting numerator and denominator "
        "null; the sample-size value is the denominator itself, not a rate pair. "
        "When the validator reports that numerator and denominator must be supplied together, "
        "never invent the missing value. Set both only when exact cited source cells support "
        "both; otherwise clear both fields while preserving the source-backed raw count in "
        "valueNumber/valueText and retaining its evidence. "
        "When a validator reports that numerator or denominator is not present in the cited "
        "Capture v2 cells for a percentage-only total-rate Outcome, clear both numerator and "
        "denominator for every affected observation whose pair was inferred by summing category "
        "counts. Preserve valueNumber, valueText, rate evidence, sampleSize, category Outcomes, "
        "and every other field exactly. Never sum overlapping category counts to manufacture a "
        "rate numerator. "
        "For a percent-formatted numeric source, repair valueNumber to the exact underlying "
        "captured numeric value multiplied by 100. Keep the rounded screen-display string only "
        "in valueText or evidence sourceText. If numerator and denominator are present, require "
        "the exact arithmetic percentage and never retain display rounding as a numeric claim. "
        "Defect categories measured on the same samples may overlap. When Total NG differs "
        "from the arithmetic sum of component counts, preserve Total NG and every component as "
        "separate Outcomes, record the non-additivity limitation, and never infer mutual "
        "exclusivity or recompute totals. When Input differs from OK plus Total NG, preserve "
        "the literal values, record the exact residual as an unreconciled limitation, and never "
        "impute or reclassify specimens. A MIN/MAX/AVG formula without a cached value has no "
        "numeric result: preserve the raw measurementSeries and exact formula lineage, leave "
        "the derived numeric field null or omit its aggregate series, and record the missing-"
        "cache limitation. "
        "An explicit DOE/comparison title or purpose plus a multi-condition result table "
        "authorizes only unambiguous exactly-one-factor aligned NEEDS_REVIEW comparisons; it "
        "does not authorize a CONTROL Arm role, validity, aggregation, effects, or any pair "
        "where multiple factors differ. For ordered dose/interval Arms without an explicit "
        "baseline/reference, keep only adjacent source-order pairs rather than all pairs. "
        "Split literal independently queryable condition "
        "components such as drying temperature and drying duration into separate Factors while "
        "retaining the combined source text and evidence. A raw measurement unit may come from "
        "the exact governing group header when the leaf header omits it, but never from an "
        "unrelated block. "
        "A NEEDS_REVIEW paired comparison is allowed only when both BEFORE and AFTER RAW series "
        "contain values at the same pressure, their axes align, and their evidence-linked "
        "formatted replicate identities are the same ordered #1..#N. Its validityStatus and "
        "verificationStatus remain NEEDS_REVIEW, confoundingStatus remains UNASSESSED, "
        "aggregationEligible remains false, and effects remains empty. Never compare a missing "
        "or header-only stage. Preserve a 300 After header-only Arm when the formatted headers "
        "exist, but create no numeric series or comparison for it. "
        "For a source-authored measurement-stage factor shared by BEFORE and AFTER Arms, repair "
        "only its temporal direction: baselineCondition Before, changedCondition After, "
        "matching BEFORE factorValue.isBaseline true, and matching AFTER "
        "factorValue.isBaseline false. This does not create a CONTROL/BASELINE Arm and does not "
        "change any comparison gate, effect, aggregation status, role, range, series reference, "
        "count, context, or other semantic content. When this temporal-stage mapping is the "
        "validator error, perform an exact minimal repair and preserve every other field and "
        "list item unchanged. "
        "corrected result must remain a fail-closed NEEDS_REVIEW draft under "
        f"{STUDY_DRAFT_PROMPT_VERSION}.\n\n"
        f"VALIDATOR ERROR:\n{validation_error}\n\n"
        "REJECTED FULL JSON:\n"
        + json.dumps(rejected, ensure_ascii=False, separators=(",", ":"))
    )
    if not source_prompt:
        return repair_task
    return (
        source_prompt
        + "\n\nREPAIR OVERRIDE — THE PREVIOUS FULL DRAFT FAILED VALIDATION:\n"
        + repair_task
    )


def build_study_json_retry_prompt(
    source_prompt: str,
    parse_error: str,
) -> str:
    """Retry the same source-grounded task with strict JSON/schema emphasis."""

    return (
        source_prompt
        + "\n\nJSON RECOVERY — THE PREVIOUS RESPONSE WAS NOT VALID JSON:\n"
        "Repeat the same source-grounded task and return one complete JSON "
        "object only. Do not summarize, omit source-backed records, add "
        "facts, or weaken any NEEDS_REVIEW, evidence, comparison, numeric, "
        "or canonical validation rule. The output schema is now enforced. "
        "The invalid response is deliberately not supplied as evidence or "
        "repair input.\n"
        f"JSON PARSE ERROR: {parse_error}\n"
    )


_BEFORE_SOURCE_PATTERN = re.compile(
    r"(?<![A-Za-z])before(?![A-Za-z])|"
    r"(?<![A-Za-z])pre[\s_-]?change(?![A-Za-z])|"
    r"(?:변경|시험|측정)\s*전",
    re.IGNORECASE,
)
_AFTER_SOURCE_PATTERN = re.compile(
    r"(?<![A-Za-z])after(?![A-Za-z])|"
    r"(?<![A-Za-z])post[\s_-]?change(?![A-Za-z])|"
    r"(?:변경|시험|측정)\s*후",
    re.IGNORECASE,
)
_BEFORE_NUMBER_FORMAT_PATTERN = re.compile(
    r"(?<![A-Za-z])before(?![A-Za-z])|"
    r"(?<![A-Za-z])pre(?![A-Za-z])",
    re.IGNORECASE,
)
_AFTER_NUMBER_FORMAT_PATTERN = re.compile(
    r"(?<![A-Za-z])after(?![A-Za-z])|"
    r"(?<![A-Za-z])post(?![A-Za-z])",
    re.IGNORECASE,
)


def _temporal_terms(value: object) -> set[str]:
    if not isinstance(value, (str, int, float)) or isinstance(value, bool):
        return set()
    text = str(value)
    roles: set[str] = set()
    if _BEFORE_SOURCE_PATTERN.search(text):
        roles.add("BEFORE")
    if _AFTER_SOURCE_PATTERN.search(text):
        roles.add("AFTER")
    return roles


def _number_format_temporal_terms(value: object) -> frozenset[str]:
    if not isinstance(value, str):
        return frozenset()
    roles: set[str] = set()
    if _BEFORE_NUMBER_FORMAT_PATTERN.search(value):
        roles.add("BEFORE")
    if _AFTER_NUMBER_FORMAT_PATTERN.search(value):
        roles.add("AFTER")
    return frozenset(roles)


def _focused_source_cells(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, list[tuple[int, int, str, frozenset[str]]]]:
    indexed: dict[
        str,
        list[tuple[int, int, str, frozenset[str]]],
    ] = {}
    seen: set[tuple[str, int, int]] = set()
    for chunk in focused_chunks:
        sheet_value = chunk.get("sheet")
        sheet = str(
            sheet_value.get("title")
            if isinstance(sheet_value, dict)
            else sheet_value or ""
        ).strip()
        if not sheet:
            continue
        for cell in [
            *chunk.get("cells", []),
            *chunk.get("contextCells", []),
        ]:
            if not isinstance(cell, dict):
                continue
            coordinate = str(cell.get("coordinate") or "").strip()
            if not coordinate:
                continue
            try:
                start_row, start_col, end_row, end_col = _a1_bounds(coordinate)
            except SemanticAiError:
                continue
            if start_row != end_row or start_col != end_col:
                continue
            identity = (sheet, start_row, start_col)
            if identity in seen:
                continue
            seen.add(identity)
            texts = [
                str(value)
                for value in (
                    cell.get("displayValue"),
                    cell.get("rawValue"),
                    cell.get("cachedValue"),
                )
                if isinstance(value, (str, int, float))
                and not isinstance(value, bool)
            ]
            indexed.setdefault(sheet, []).append(
                (
                    start_row,
                    start_col,
                    " ".join(texts),
                    _number_format_temporal_terms(
                        cell.get("numberFormat")
                    ),
                )
            )
    return indexed


def _evidence_has_temporal_role(
    evidence: object,
    role: str,
    source_cells: dict[
        str,
        list[tuple[int, int, str, frozenset[str]]],
    ],
    *,
    allow_number_format: bool = True,
) -> bool:
    pattern = (
        _BEFORE_SOURCE_PATTERN
        if role == "BEFORE"
        else _AFTER_SOURCE_PATTERN
    )
    if not isinstance(evidence, list):
        return False
    for item in evidence:
        if not isinstance(item, dict):
            continue
        sheet = str(item.get("sheet") or "").strip()
        address = str(item.get("range") or "").strip()
        if not sheet or not address:
            continue
        try:
            start_row, start_col, end_row, end_col = _a1_bounds(address)
        except SemanticAiError:
            continue
        for row, column, text, format_roles in source_cells.get(sheet, []):
            if (
                start_row <= row <= end_row
                and start_col <= column <= end_col
                and (
                    pattern.search(text)
                    or (
                        allow_number_format
                        and role in format_roles
                    )
                )
            ):
                return True
    return False


def _temporal_stage_factor_keys(study: dict[str, Any]) -> set[str]:
    before_factors: set[str] = set()
    after_factors: set[str] = set()
    for arm in study.get("arms", []):
        if not isinstance(arm, dict):
            continue
        arm_role = str(arm.get("role") or "").strip().upper()
        if arm_role not in {"BEFORE", "AFTER"}:
            continue
        for factor_value in arm.get("factorValues", []):
            if not isinstance(factor_value, dict):
                continue
            factor_key = str(factor_value.get("factor") or "").strip()
            if not factor_key:
                continue
            temporal_roles = _temporal_terms(factor_value.get("value"))
            if arm_role not in temporal_roles:
                continue
            if arm_role == "BEFORE":
                before_factors.add(factor_key)
            else:
                after_factors.add(factor_key)
    return before_factors & after_factors


def _validate_source_explicit_temporal_semantics(
    manifest: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    """Reject temporal phases that exist only in locator/model interpretation."""

    source_cells = _focused_source_cells(focused_chunks)
    def require_evidence(
        value: object,
        path: str,
        evidence: object,
        *,
        allow_number_format: bool,
    ) -> None:
        for role in _temporal_terms(value):
            if not _evidence_has_temporal_role(
                evidence,
                role,
                source_cells,
                allow_number_format=allow_number_format,
            ):
                raise SemanticAiError(
                    f"{path} uses {role} without an evidence-linked literal "
                    "captured cell label or custom number_format"
                )

    for study_index, study in enumerate(manifest.get("studies", [])):
        if not isinstance(study, dict):
            continue
        for field in ("key", "title", "designType", "comparisonBasis"):
            require_evidence(
                study.get(field),
                f"studies[{study_index}].{field}",
                study.get("evidence"),
                allow_number_format=True,
            )
        for arm_index, arm in enumerate(study.get("arms", [])):
            if not isinstance(arm, dict):
                continue
            role = str(arm.get("role") or "").strip().upper()
            if role in {"BEFORE", "AFTER"} and not _evidence_has_temporal_role(
                arm.get("evidence"),
                role,
                source_cells,
                allow_number_format=True,
            ):
                raise SemanticAiError(
                    f"studies[{study_index}].arms[{arm_index}].role {role} "
                    "requires a literal label inside that arm's source evidence"
                )
            for field in (
                "key",
                "label",
                "condition",
                "sampleBasis",
                "matchingBasis",
            ):
                for temporal_role in _temporal_terms(arm.get(field)):
                    if not _evidence_has_temporal_role(
                        arm.get("evidence"),
                        temporal_role,
                        source_cells,
                        allow_number_format=True,
                    ):
                        raise SemanticAiError(
                            f"studies[{study_index}].arms[{arm_index}].{field} "
                            f"uses {temporal_role} without a literal label inside "
                            "that arm's source evidence"
                        )
        temporal_stage_factors = _temporal_stage_factor_keys(study)
        factors_by_key = {
            str(factor.get("key") or ""): (factor_index, factor)
            for factor_index, factor in enumerate(study.get("factors", []))
            if isinstance(factor, dict)
        }
        for factor_key in sorted(temporal_stage_factors):
            factor_entry = factors_by_key.get(factor_key)
            if factor_entry is None:
                continue
            factor_index, factor = factor_entry
            baseline_roles = _temporal_terms(
                factor.get("baselineCondition")
            )
            changed_roles = _temporal_terms(
                factor.get("changedCondition")
            )
            if (
                baseline_roles != {"BEFORE"}
                or changed_roles != {"AFTER"}
            ):
                raise SemanticAiError(
                    "source-authored temporal stage baseline mapping "
                    f"requires studies[{study_index}].factors"
                    f"[{factor_index}] baselineCondition Before and "
                    "changedCondition After"
                )
            require_evidence(
                factor.get("baselineCondition"),
                f"studies[{study_index}].factors"
                f"[{factor_index}].baselineCondition",
                factor.get("evidence"),
                allow_number_format=True,
            )
            require_evidence(
                factor.get("changedCondition"),
                f"studies[{study_index}].factors"
                f"[{factor_index}].changedCondition",
                factor.get("evidence"),
                allow_number_format=True,
            )
            for arm_index, arm in enumerate(study.get("arms", [])):
                if not isinstance(arm, dict):
                    continue
                arm_role = str(
                    arm.get("role") or ""
                ).strip().upper()
                if arm_role not in {"BEFORE", "AFTER"}:
                    continue
                for factor_value_index, factor_value in enumerate(
                    arm.get("factorValues", [])
                ):
                    if (
                        not isinstance(factor_value, dict)
                        or str(factor_value.get("factor") or "").strip()
                        != factor_key
                        or arm_role
                        not in _temporal_terms(factor_value.get("value"))
                    ):
                        continue
                    expected_baseline = arm_role == "BEFORE"
                    if factor_value.get("isBaseline") is not expected_baseline:
                        raise SemanticAiError(
                            "source-authored temporal stage baseline mapping "
                            f"requires studies[{study_index}].arms"
                            f"[{arm_index}].factorValues"
                            f"[{factor_value_index}].isBaseline "
                            f"{str(expected_baseline).lower()}"
                        )
                    require_evidence(
                        factor_value.get("value"),
                        f"studies[{study_index}].arms[{arm_index}]"
                        f".factorValues[{factor_value_index}].value",
                        arm.get("evidence"),
                        allow_number_format=True,
                    )
        for series_index, series in enumerate(
            study.get("measurementSeries", [])
        ):
            if not isinstance(series, dict):
                continue
            sheet = str(series.get("sheet") or "").strip()
            header_range = str(series.get("headerRange") or "").strip()
            if sheet and header_range:
                series_evidence: object = [
                    {"sheet": sheet, "range": header_range}
                ]
            else:
                series_evidence = series.get("evidence")
            for field in ("key", "stratumKey"):
                require_evidence(
                    series.get(field),
                    f"studies[{study_index}].measurementSeries"
                    f"[{series_index}].{field}",
                    series_evidence,
                    allow_number_format=True,
                )
        for comparison_index, comparison in enumerate(
            study.get("comparisons", [])
        ):
            if not isinstance(comparison, dict):
                continue
            for field in ("key", "designType", "matchingBasis"):
                require_evidence(
                    comparison.get(field),
                    f"studies[{study_index}].comparisons"
                    f"[{comparison_index}].{field}",
                    comparison.get("evidence"),
                    allow_number_format=True,
                )


def _reference_repair_projection(
    manifest: dict[str, Any],
) -> dict[str, Any]:
    projected = copy.deepcopy(manifest)
    for study in projected.get("studies", []):
        if not isinstance(study, dict):
            continue
        for arm in study.get("arms", []):
            if not isinstance(arm, dict):
                continue
            for factor_value in arm.get("factorValues", []):
                if isinstance(factor_value, dict) and "factor" in factor_value:
                    factor_value["factor"] = "<REFERENCE>"
        for outcome in study.get("outcomes", []):
            if not isinstance(outcome, dict):
                continue
            for observation in outcome.get("observations", []):
                if isinstance(observation, dict) and "arm" in observation:
                    observation["arm"] = "<REFERENCE>"
        for series in study.get("measurementSeries", []):
            if not isinstance(series, dict):
                continue
            for field in ("outcome", "arm"):
                if field in series:
                    series[field] = "<REFERENCE>"
        for comparison in study.get("comparisons", []):
            if not isinstance(comparison, dict):
                continue
            for field in ("comparedArm", "controlArm"):
                if field in comparison:
                    comparison[field] = "<REFERENCE>"
            for effect in comparison.get("effects", []):
                if isinstance(effect, dict) and "outcome" in effect:
                    effect["outcome"] = "<REFERENCE>"
    return projected


_SYNTHETIC_REFERENCE_ARM_ERROR_PATTERN = re.compile(
    r"studies\[(\d+)\]\.arms\[(\d+)\]\.role REFERENCE requires "
    r"directly cited captured full Normal, Reference, Standard, Spec, or "
    r"equivalent reference wording matching the Arm label or condition, "
    r"or at least two exact ordered distinct full reference #N identity "
    r"cells for a descriptive grouped Arm; a bare abbreviation such as "
    r"ST or mixed Test/Normal evidence is not reference semantics\s*$",
    re.IGNORECASE,
)


def _synthetic_reference_arm_target(
    validation_error: str,
) -> tuple[int, int] | None:
    """Parse only the exact unsupported REFERENCE evidence failure."""

    match = _SYNTHETIC_REFERENCE_ARM_ERROR_PATTERN.search(
        validation_error
    )
    if match is None:
        return None
    return int(match.group(1)), int(match.group(2))


def _normalized_arm_component(value: object) -> str:
    return " ".join(str(value or "").split()).casefold()


def _focused_exact_text_cells(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, list[tuple[int, int, str]]]:
    indexed: dict[str, list[tuple[int, int, str]]] = {}
    seen: set[tuple[str, int, int]] = set()
    for chunk in focused_chunks:
        sheet_value = chunk.get("sheet")
        sheet = str(
            sheet_value.get("title")
            if isinstance(sheet_value, dict)
            else sheet_value or ""
        ).strip()
        if not sheet:
            continue
        sheet_key = sheet.casefold()
        for cell in [
            *chunk.get("cells", []),
            *chunk.get("contextCells", []),
        ]:
            if not isinstance(cell, dict):
                continue
            coordinate = str(cell.get("coordinate") or "").strip()
            if not coordinate:
                continue
            try:
                start_row, start_col, end_row, end_col = _a1_bounds(
                    coordinate
                )
            except SemanticAiError:
                continue
            if start_row != end_row or start_col != end_col:
                continue
            identity = (sheet_key, start_row, start_col)
            if identity in seen:
                continue
            seen.add(identity)
            text = ""
            for field in ("rawValue", "displayValue", "cachedValue"):
                value = cell.get(field)
                if (
                    isinstance(value, (str, int, float))
                    and not isinstance(value, bool)
                    and str(value).strip()
                ):
                    text = str(value).strip()
                    break
            if text:
                indexed.setdefault(sheet_key, []).append(
                    (start_row, start_col, text)
                )
    return indexed


def _arm_exact_evidence_texts(
    arm: dict[str, Any],
    source_cells: dict[str, list[tuple[int, int, str]]],
) -> list[str]:
    texts: list[str] = []
    seen: set[tuple[str, int, int]] = set()
    evidence = arm.get("evidence")
    if not isinstance(evidence, list):
        return texts
    for item in evidence:
        if not isinstance(item, dict):
            continue
        sheet_key = str(item.get("sheet") or "").strip().casefold()
        address = str(item.get("range") or "").strip()
        if not sheet_key or not address:
            continue
        try:
            start_row, start_col, end_row, end_col = _a1_bounds(
                address
            )
        except SemanticAiError:
            continue
        for row, column, text in source_cells.get(sheet_key, []):
            identity = (sheet_key, row, column)
            if (
                identity not in seen
                and start_row <= row <= end_row
                and start_col <= column <= end_col
            ):
                seen.add(identity)
                texts.append(text)
    return texts


def _synthetic_arm_repair_plan(
    baseline: dict[str, Any],
    *,
    study_index: int,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[int, str]:
    """Plan exact source-authored identities for split-cell Arms."""

    try:
        study = baseline["studies"][study_index]
        arms = study["arms"]
    except (IndexError, KeyError, TypeError) as exc:
        raise SemanticAiError(
            "Synthetic REFERENCE Arm repair target structure is invalid"
        ) from exc
    if not isinstance(study, dict) or not isinstance(arms, list):
        raise SemanticAiError(
            "Synthetic REFERENCE Arm repair requires canonical Arm lists"
        )
    source_cells = _focused_exact_text_cells(focused_chunks)
    plan: dict[int, str] = {}
    for arm_index, arm in enumerate(arms):
        if not isinstance(arm, dict):
            continue
        label = str(arm.get("label") or "").strip()
        condition = str(arm.get("condition") or "").strip()
        parts = [
            part.strip()
            for part in label.split("/")
            if part.strip()
        ]
        if (
            len(parts) < 2
            or _normalized_arm_component(condition)
            != _normalized_arm_component(label)
        ):
            continue
        identity = _normalized_arm_component(parts[0])
        role = str(arm.get("role") or "OTHER").strip().upper()
        if (
            identity == "test"
            and role != "TEST"
            or identity == "normal"
            and role != "REFERENCE"
            or identity not in {"test", "normal"}
        ):
            continue
        cited_texts = _arm_exact_evidence_texts(
            arm,
            source_cells,
        )
        normalized_cited = {
            _normalized_arm_component(text)
            for text in cited_texts
            if _normalized_arm_component(text)
        }
        normalized_parts = {
            _normalized_arm_component(part)
            for part in parts
        }
        factor_values = {
            _normalized_arm_component(factor_value.get("value"))
            for factor_value in arm.get("factorValues", [])
            if isinstance(factor_value, dict)
            and _normalized_arm_component(factor_value.get("value"))
        }
        if (
            _normalized_arm_component(label) in normalized_cited
            or not normalized_parts.issubset(normalized_cited)
            or not normalized_parts.issubset(factor_values)
        ):
            continue
        exact_identity = next(
            (
                text
                for text in cited_texts
                if _normalized_arm_component(text) == identity
            ),
            parts[0],
        )
        plan[arm_index] = exact_identity
    return plan


def _validate_synthetic_reference_arm_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    *,
    target: tuple[int, int],
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    """Allow only planned label/condition changes in the target Study."""

    study_index, arm_index = target
    plan = _synthetic_arm_repair_plan(
        baseline,
        study_index=study_index,
        focused_chunks=focused_chunks,
    )
    try:
        target_arm = baseline["studies"][study_index]["arms"][
            arm_index
        ]
    except (IndexError, KeyError, TypeError) as exc:
        raise SemanticAiError(
            "Synthetic REFERENCE Arm repair target structure is invalid"
        ) from exc
    if (
        arm_index not in plan
        or not isinstance(target_arm, dict)
        or str(target_arm.get("role") or "").strip().upper()
        != "REFERENCE"
    ):
        raise SemanticAiError(
            "Synthetic REFERENCE Arm repair target is not an exact "
            "split-cell unsupported REFERENCE"
        )
    expected = copy.deepcopy(baseline)
    for planned_index, identity in plan.items():
        arm = expected["studies"][study_index]["arms"][
            planned_index
        ]
        arm["label"] = identity
        arm["condition"] = identity
    if repaired != expected:
        raise SemanticAiError(
            "Synthetic REFERENCE Arm repair changed fields outside the "
            "planned exact source label and condition paths"
        )


def _apply_deterministic_synthetic_reference_arm_repair(
    baseline: dict[str, Any],
    *,
    target: tuple[int, int],
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Restore split-cell Arms to their exact source-authored identity."""

    study_index, _arm_index = target
    plan = _synthetic_arm_repair_plan(
        baseline,
        study_index=study_index,
        focused_chunks=focused_chunks,
    )
    repaired = copy.deepcopy(baseline)
    for planned_index, identity in plan.items():
        arm = repaired["studies"][study_index]["arms"][planned_index]
        arm["label"] = identity
        arm["condition"] = identity
    _validate_synthetic_reference_arm_repair(
        baseline,
        repaired,
        target=target,
        focused_chunks=focused_chunks,
    )
    return repaired


def _temporal_stage_repair_projection(
    manifest: dict[str, Any],
) -> dict[str, Any]:
    projected = copy.deepcopy(manifest)
    for study in projected.get("studies", []):
        if not isinstance(study, dict):
            continue
        temporal_factor_keys = _temporal_stage_factor_keys(study)
        if not temporal_factor_keys:
            continue
        for factor in study.get("factors", []):
            if (
                isinstance(factor, dict)
                and str(factor.get("key") or "") in temporal_factor_keys
            ):
                for field in ("baselineCondition", "changedCondition"):
                    if field in factor:
                        factor[field] = "<TEMPORAL_STAGE_MAPPING>"
        for arm in study.get("arms", []):
            if not isinstance(arm, dict):
                continue
            arm_role = str(arm.get("role") or "").strip().upper()
            if arm_role not in {"BEFORE", "AFTER"}:
                continue
            for factor_value in arm.get("factorValues", []):
                if (
                    isinstance(factor_value, dict)
                    and str(factor_value.get("factor") or "")
                    in temporal_factor_keys
                    and arm_role
                    in _temporal_terms(factor_value.get("value"))
                    and "isBaseline" in factor_value
                ):
                    factor_value["isBaseline"] = (
                        "<TEMPORAL_STAGE_MAPPING>"
                    )
    return projected


_STRICT_NUMERIC_TEXT_PATTERN = re.compile(
    r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$"
)
_OBSERVATION_NUMERIC_FIELDS = (
    "valueNumber",
    "numerator",
    "denominator",
    "ratePpm",
    "min",
    "max",
    "average",
)


def _strict_numeric_text_decimal(value: object) -> Decimal | None:
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not _STRICT_NUMERIC_TEXT_PATTERN.fullmatch(text):
        return None
    try:
        number = Decimal(text)
    except InvalidOperation:
        return None
    return number if number.is_finite() else None


def _numeric_text_homogeneity_repair_targets(
    manifest: dict[str, Any],
) -> dict[tuple[int, int, int], Decimal]:
    """Find mixed-outcome text cells that are entirely strict numbers.

    Return no targets unless every qualitative-only observation in each
    affected quantitative outcome is a strict numeric string. This prevents
    the bounded repair from coercing genuine Pass/Fail or narrative values.
    """

    targets: dict[tuple[int, int, int], Decimal] = {}
    for study_index, study in enumerate(manifest.get("studies", [])):
        if not isinstance(study, dict):
            continue
        series_outcomes = {
            str(series.get("outcome") or "")
            for series in study.get("measurementSeries", [])
            if isinstance(series, dict)
        }
        for outcome_index, outcome in enumerate(study.get("outcomes", [])):
            if not isinstance(outcome, dict):
                continue
            observations = outcome.get("observations", [])
            if not isinstance(observations, list):
                continue
            has_quantitative = str(outcome.get("key") or "") in series_outcomes
            qualitative: list[tuple[int, dict[str, Any]]] = []
            for observation_index, observation in enumerate(observations):
                if not isinstance(observation, dict):
                    continue
                if any(
                    observation.get(field) not in (None, "")
                    for field in _OBSERVATION_NUMERIC_FIELDS
                ):
                    has_quantitative = True
                elif str(observation.get("valueText") or "").strip():
                    qualitative.append((observation_index, observation))
            if not has_quantitative or not qualitative:
                continue
            outcome_targets: dict[tuple[int, int, int], Decimal] = {}
            for observation_index, observation in qualitative:
                number = _strict_numeric_text_decimal(
                    observation.get("valueText")
                )
                if number is None:
                    outcome_targets = {}
                    break
                outcome_targets[
                    (study_index, outcome_index, observation_index)
                ] = number
            targets.update(outcome_targets)
    return targets


def _validate_numeric_text_homogeneity_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
) -> None:
    """Allow only exact valueNumber restoration for strict numeric text."""

    targets = _numeric_text_homogeneity_repair_targets(baseline)
    if not targets:
        raise SemanticAiError(
            "Numeric-text repair has no safe strict-numeric targets"
        )
    projected_baseline = copy.deepcopy(baseline)
    projected_repaired = copy.deepcopy(repaired)
    for path, expected in targets.items():
        study_index, outcome_index, observation_index = path
        try:
            baseline_observation = projected_baseline["studies"][
                study_index
            ]["outcomes"][outcome_index]["observations"][observation_index]
            repaired_observation = projected_repaired["studies"][
                study_index
            ]["outcomes"][outcome_index]["observations"][observation_index]
        except (IndexError, KeyError, TypeError) as exc:
            raise SemanticAiError(
                "Numeric-text repair changed outcome or observation structure"
            ) from exc
        repaired_number = repaired_observation.get("valueNumber")
        if isinstance(repaired_number, bool) or not isinstance(
            repaired_number,
            (int, float),
        ):
            raise SemanticAiError(
                "Numeric-text repair must set each strict numeric valueNumber"
            )
        try:
            actual = Decimal(str(repaired_number))
        except InvalidOperation as exc:
            raise SemanticAiError(
                "Numeric-text repair produced an invalid valueNumber"
            ) from exc
        if not actual.is_finite() or actual != expected:
            raise SemanticAiError(
                "Numeric-text repair valueNumber does not exactly match "
                "valueText"
            )
        baseline_observation["valueNumber"] = "<NUMERIC_TEXT_REPAIR>"
        repaired_observation["valueNumber"] = "<NUMERIC_TEXT_REPAIR>"
    if projected_repaired != projected_baseline:
        raise SemanticAiError(
            "Numeric-text repair changed fields outside the allowed "
            "valueNumber paths"
        )


def _apply_deterministic_numeric_text_homogeneity_repair(
    baseline: dict[str, Any],
) -> dict[str, Any]:
    """Fill only safe strict-numeric valueNumber paths from the baseline.

    This intentionally avoids a second model rewrite.  The exact captured
    valueText is authoritative, while every other draft field remains a
    deepcopy of the rejected baseline.
    """

    targets = _numeric_text_homogeneity_repair_targets(baseline)
    if not targets:
        raise SemanticAiError(
            "Numeric-text repair has no safe strict-numeric targets"
        )
    repaired = copy.deepcopy(baseline)
    for path, expected in targets.items():
        study_index, outcome_index, observation_index = path
        try:
            observation = repaired["studies"][study_index]["outcomes"][
                outcome_index
            ]["observations"][observation_index]
        except (IndexError, KeyError, TypeError) as exc:
            raise SemanticAiError(
                "Numeric-text repair target structure is invalid"
            ) from exc
        if expected == expected.to_integral_value():
            value_number: int | float = int(expected)
        else:
            value_number = float(expected)
            if (
                not math.isfinite(value_number)
                or Decimal(str(value_number)) != expected
            ):
                raise SemanticAiError(
                    "Strict numeric text cannot be represented exactly as "
                    "a canonical JSON number"
                )
        observation["valueNumber"] = value_number
    _validate_numeric_text_homogeneity_repair(baseline, repaired)
    return repaired


_UNSUPPORTED_RATE_PAIR_ERROR_PATTERN = re.compile(
    r"studies\[(\d+)\]\.outcomes\[(\d+)\]\.observations\[(\d+)\]"
    r"\.(?:numerator|denominator)=.*not present in its cited capture v2 cells",
    re.IGNORECASE,
)


def _unsupported_rate_pair_observation(
    validation_error: str,
) -> tuple[int, int, int] | None:
    match = _UNSUPPORTED_RATE_PAIR_ERROR_PATTERN.search(validation_error)
    if match is None:
        return None
    return (
        int(match.group(1)),
        int(match.group(2)),
        int(match.group(3)),
    )


def _validate_unsupported_rate_pair_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    observation_path: tuple[int, int, int],
) -> None:
    """Allow only clearing one rejected observation's count pair."""

    projected_baseline = copy.deepcopy(baseline)
    projected_repaired = copy.deepcopy(repaired)
    study_index, outcome_index, observation_index = observation_path
    try:
        baseline_observations = projected_baseline["studies"][
            study_index
        ]["outcomes"][outcome_index]["observations"]
        repaired_observations = projected_repaired["studies"][
            study_index
        ]["outcomes"][outcome_index]["observations"]
    except (IndexError, KeyError, TypeError) as exc:
        raise SemanticAiError(
            "Unsupported rate-pair repair changed study or outcome structure"
        ) from exc
    if (
        not isinstance(baseline_observations, list)
        or not isinstance(repaired_observations, list)
        or len(baseline_observations) != len(repaired_observations)
    ):
        raise SemanticAiError(
            "Unsupported rate-pair repair changed observation structure"
        )
    try:
        baseline_observation = baseline_observations[observation_index]
        repaired_observation = repaired_observations[observation_index]
    except IndexError as exc:
        raise SemanticAiError(
            "Unsupported rate-pair repair changed observation structure"
        ) from exc
    if not isinstance(baseline_observation, dict) or not isinstance(
        repaired_observation,
        dict,
    ):
        raise SemanticAiError(
            "Unsupported rate-pair repair requires object observations"
        )
    baseline_pair = (
        baseline_observation.get("numerator"),
        baseline_observation.get("denominator"),
    )
    repaired_pair = (
        repaired_observation.get("numerator"),
        repaired_observation.get("denominator"),
    )
    if repaired_pair != (None, None) or baseline_pair == (None, None):
        raise SemanticAiError(
            "Unsupported rate-pair repair must clear the rejected numerator "
            "and denominator pair"
        )
    for field in ("numerator", "denominator"):
        baseline_observation[field] = "<RATE_PAIR_REPAIR>"
        repaired_observation[field] = "<RATE_PAIR_REPAIR>"
    if projected_repaired != projected_baseline:
        raise SemanticAiError(
            "Unsupported rate-pair repair changed fields outside the allowed "
            "numerator and denominator paths"
        )


def _apply_deterministic_unsupported_rate_pair_repair(
    baseline: dict[str, Any],
    observation_path: tuple[int, int, int],
) -> dict[str, Any]:
    """Clear only the validator-identified unsupported count pair."""

    repaired = copy.deepcopy(baseline)
    study_index, outcome_index, observation_index = observation_path
    try:
        observation = repaired["studies"][study_index]["outcomes"][
            outcome_index
        ]["observations"][observation_index]
    except (IndexError, KeyError, TypeError) as exc:
        raise SemanticAiError(
            "Unsupported rate-pair repair target structure is invalid"
        ) from exc
    observation["numerator"] = None
    observation["denominator"] = None
    _validate_unsupported_rate_pair_repair(
        baseline,
        repaired,
        observation_path,
    )
    return repaired


_INCOMPATIBLE_RAW_COMPARISON_ERROR_PATTERN = re.compile(
    r"studies\[(\d+)\]\.comparisons\[(\d+)\]\s+shared Outcome "
    r"(['\"])([^'\"]+)\3 RAW representations require compatible value "
    r"units and aligned ordered axis identity, shape, and stratum; omit "
    r"the invalid Comparison, preserve its Arms/Outcomes/series, and add "
    r"a limitation\s*$",
    re.IGNORECASE,
)


def _incompatible_raw_comparison_target(
    validation_error: str,
) -> tuple[int, int, str] | None:
    """Parse only the exact RAW-series alignment validator failure."""

    match = _INCOMPATIBLE_RAW_COMPARISON_ERROR_PATTERN.search(
        validation_error
    )
    if match is None:
        return None
    return (
        int(match.group(1)),
        int(match.group(2)),
        match.group(4),
    )


def _incompatible_raw_comparison_limitation(
    comparison: dict[str, Any],
    outcome_key: str,
) -> str:
    comparison_key = str(comparison.get("key") or "").strip()
    identity = (
        f" {comparison_key!r}"
        if comparison_key
        else ""
    )
    return (
        f"Comparison{identity} was omitted because shared Outcome "
        f"{outcome_key!r} has incompatible RAW representation units, "
        "ordered axis identity, shape, or stratum."
    )


def _validate_incompatible_raw_comparison_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    target: tuple[int, int, str],
) -> None:
    """Require byte-equivalent JSON semantics outside one omission/note."""

    study_index, comparison_index, outcome_key = target
    expected = copy.deepcopy(baseline)
    try:
        study = expected["studies"][study_index]
        comparisons = study["comparisons"]
        comparison = comparisons[comparison_index]
        limitations = study["limitations"]
    except (IndexError, KeyError, TypeError) as exc:
        raise SemanticAiError(
            "Incompatible RAW comparison repair target structure is invalid"
        ) from exc
    if (
        not isinstance(study, dict)
        or not isinstance(comparisons, list)
        or not isinstance(comparison, dict)
        or not isinstance(limitations, list)
    ):
        raise SemanticAiError(
            "Incompatible RAW comparison repair requires canonical lists"
        )
    declared_outcomes = {
        str(outcome.get("key") or "")
        for outcome in study.get("outcomes", [])
        if isinstance(outcome, dict)
    }
    arm_keys = {
        str(arm.get("key") or "")
        for arm in study.get("arms", [])
        if isinstance(arm, dict)
    }
    compared_arm = str(comparison.get("comparedArm") or "")
    control_arm = str(comparison.get("controlArm") or "")
    if (
        outcome_key not in declared_outcomes
        or compared_arm not in arm_keys
        or control_arm not in arm_keys
        or compared_arm == control_arm
    ):
        raise SemanticAiError(
            "Incompatible RAW comparison repair target does not match "
            "declared Outcome and Arms"
        )
    raw_arms = {
        str(series.get("arm") or "")
        for series in study.get("measurementSeries", [])
        if isinstance(series, dict)
        and str(series.get("seriesRole") or "RAW").upper() == "RAW"
        and str(series.get("outcome") or "") == outcome_key
    }
    if not {compared_arm, control_arm}.issubset(raw_arms):
        raise SemanticAiError(
            "Incompatible RAW comparison repair requires RAW series for "
            "both validator-identified Arms"
        )
    limitation = _incompatible_raw_comparison_limitation(
        comparison,
        outcome_key,
    )
    del comparisons[comparison_index]
    limitations.append(limitation)
    if repaired != expected:
        raise SemanticAiError(
            "Incompatible RAW comparison repair changed fields outside "
            "the one validator-identified Comparison and limitation"
        )


def _apply_deterministic_incompatible_raw_comparison_repair(
    baseline: dict[str, Any],
    target: tuple[int, int, str],
) -> dict[str, Any]:
    """Omit one exact validator-identified incompatible Comparison."""

    repaired = copy.deepcopy(baseline)
    study_index, comparison_index, outcome_key = target
    try:
        study = repaired["studies"][study_index]
        comparisons = study["comparisons"]
        comparison = comparisons[comparison_index]
        limitations = study["limitations"]
    except (IndexError, KeyError, TypeError) as exc:
        raise SemanticAiError(
            "Incompatible RAW comparison repair target structure is invalid"
        ) from exc
    if (
        not isinstance(study, dict)
        or not isinstance(comparisons, list)
        or not isinstance(comparison, dict)
        or not isinstance(limitations, list)
    ):
        raise SemanticAiError(
            "Incompatible RAW comparison repair requires canonical targets"
        )
    limitation = _incompatible_raw_comparison_limitation(
        comparison,
        outcome_key,
    )
    del comparisons[comparison_index]
    limitations.append(limitation)
    _validate_incompatible_raw_comparison_repair(
        baseline,
        repaired,
        target,
    )
    return repaired


_A1_CELL_OR_RANGE_PATTERN = re.compile(
    r"\$?[A-Za-z]{1,4}\$?[1-9]\d*"
    r"(?::\$?[A-Za-z]{1,4}\$?[1-9]\d*)?"
)
_INVALID_A1_EVIDENCE_ERROR_PATTERN = re.compile(
    r"studies\[\d+\]"
    r"(?:\.[A-Za-z][A-Za-z0-9]*|\[\d+\])*"
    r"\.evidence\[\d+\]\.range must be an A1 cell or range",
    re.IGNORECASE,
)


def _split_a1_union_evidence_in_place(value: Any) -> int:
    """Split comma-separated A1 unions without changing evidence meaning."""

    split_count = 0
    if isinstance(value, list):
        for item in value:
            split_count += _split_a1_union_evidence_in_place(item)
        return split_count
    if not isinstance(value, dict):
        return split_count

    for key, item in list(value.items()):
        if key != "evidence" or not isinstance(item, list):
            split_count += _split_a1_union_evidence_in_place(item)
            continue
        expanded: list[Any] = []
        for evidence in item:
            if not isinstance(evidence, dict):
                expanded.append(evidence)
                continue
            address = str(evidence.get("range") or "").strip()
            if "," not in address:
                expanded.append(evidence)
                continue
            parts = [part.strip() for part in address.split(",")]
            if (
                len(parts) < 2
                or len(set(parts)) != len(parts)
                or any(
                    not _A1_CELL_OR_RANGE_PATTERN.fullmatch(part)
                    for part in parts
                )
            ):
                raise SemanticAiError(
                    "Comma-separated evidence repair requires distinct "
                    "valid A1 cells or ranges"
                )
            for part in parts:
                split_evidence = copy.deepcopy(evidence)
                split_evidence["range"] = part
                expanded.append(split_evidence)
            split_count += 1
        value[key] = expanded
    return split_count


def _validate_a1_union_evidence_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
) -> None:
    expected = copy.deepcopy(baseline)
    split_count = _split_a1_union_evidence_in_place(expected)
    if split_count < 1:
        raise SemanticAiError(
            "Comma-separated evidence repair found no A1 union"
        )
    if repaired != expected:
        raise SemanticAiError(
            "Comma-separated evidence repair changed fields outside "
            "the exact evidence range split"
        )


def _apply_deterministic_a1_union_evidence_repair(
    baseline: dict[str, Any],
) -> dict[str, Any]:
    repaired = copy.deepcopy(baseline)
    split_count = _split_a1_union_evidence_in_place(repaired)
    if split_count < 1:
        raise SemanticAiError(
            "Comma-separated evidence repair found no A1 union"
        )
    _validate_a1_union_evidence_repair(baseline, repaired)
    return repaired


def _a1_union_evidence_repair_applicable(
    validation_error: str,
    baseline: dict[str, Any],
) -> bool:
    if not _INVALID_A1_EVIDENCE_ERROR_PATTERN.search(
        validation_error
    ):
        return False
    try:
        repaired = _apply_deterministic_a1_union_evidence_repair(
            baseline
        )
    except SemanticAiError:
        return False
    return repaired != baseline


_AGGREGATE_IDENTITY_ALIGNMENT_ERROR = re.compile(
    r"studies\[(\d+)\]\.measurementSeries\[(\d+)\]"
    r"\.aggregateReplicateRanges\[(\d+)\] must be contained in and aligned"
)


def _a1_column_label(column: int) -> str:
    if column < 1:
        raise SemanticAiError("A1 column index must be positive")
    label = ""
    current = column
    while current:
        current, remainder = divmod(current - 1, 26)
        label = chr(ord("A") + remainder) + label
    return label


def _a1_address(bounds: tuple[int, int, int, int]) -> str:
    start_row, start_column, end_row, end_column = bounds
    start = f"{_a1_column_label(start_column)}{start_row}"
    end = f"{_a1_column_label(end_column)}{end_row}"
    return start if start == end else f"{start}:{end}"


def _aligned_aggregate_identity_address(
    series: dict[str, Any],
    aggregate_address: str,
) -> str | None:
    """Map an exact value-axis aggregate range onto its identity axis."""

    value_bounds = _a1_bounds(str(series.get("valueRange") or ""))
    header_bounds = _a1_bounds(str(series.get("headerRange") or ""))
    identity_bounds = _a1_bounds(
        str(series.get("rowIdentityRange") or "")
    )
    aggregate_bounds = _a1_bounds(aggregate_address)
    axis_source = str(series.get("axisSource") or "").upper()
    if axis_source == "HEADER":
        if (
            aggregate_bounds[1] != value_bounds[1]
            or aggregate_bounds[3] != value_bounds[3]
            or aggregate_bounds[0] < value_bounds[0]
            or aggregate_bounds[2] > value_bounds[2]
            or identity_bounds[1] != identity_bounds[3]
            or identity_bounds[0] != value_bounds[0]
            or identity_bounds[2] != value_bounds[2]
        ):
            return None
        return _a1_address(
            (
                aggregate_bounds[0],
                identity_bounds[1],
                aggregate_bounds[2],
                identity_bounds[1],
            )
        )
    if axis_source == "ROW_IDENTITY":
        if (
            aggregate_bounds[0] != value_bounds[0]
            or aggregate_bounds[2] != value_bounds[2]
            or aggregate_bounds[1] < value_bounds[1]
            or aggregate_bounds[3] > value_bounds[3]
            or header_bounds[0] != header_bounds[2]
            or header_bounds[1] != value_bounds[1]
            or header_bounds[3] != value_bounds[3]
        ):
            return None
        return _a1_address(
            (
                header_bounds[0],
                aggregate_bounds[1],
                header_bounds[0],
                aggregate_bounds[3],
            )
        )
    return None


def _apply_deterministic_aggregate_identity_alignment_repair(
    draft: dict[str, Any],
    validation_error: str,
) -> dict[str, Any] | None:
    """Repair only unambiguous value-range/identity-range transpositions."""

    target = _AGGREGATE_IDENTITY_ALIGNMENT_ERROR.search(validation_error)
    if target is None:
        return None
    target_path = tuple(int(value) for value in target.groups())
    result = copy.deepcopy(draft)
    target_changed = False
    for study_index, study in enumerate(result.get("studies", [])):
        if not isinstance(study, dict):
            continue
        for series_index, series in enumerate(
            study.get("measurementSeries", [])
        ):
            if not isinstance(series, dict):
                continue
            ranges = series.get("aggregateReplicateRanges")
            if not isinstance(ranges, list):
                continue
            for aggregate_index, value in enumerate(ranges):
                address = str(value or "").strip()
                if not address:
                    continue
                aligned = _aligned_aggregate_identity_address(
                    series,
                    address,
                )
                if aligned is None or aligned == address:
                    continue
                ranges[aggregate_index] = aligned
                if (
                    study_index,
                    series_index,
                    aggregate_index,
                ) == target_path:
                    target_changed = True
    return result if target_changed else None


def validate_ai_study_draft(
    result: dict[str, Any],
    *,
    source: dict[str, Any],
    content_complete: bool,
    evidence_checker: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """Apply canonical validation plus the stricter non-self-approval AI policy."""

    from inference_data_ai_study_contract import StudyContractError, validate_study_manifest

    result = _normalize_ai_draft_enums(result)
    result_source = result.get("source") if isinstance(result, dict) else None
    if not isinstance(result_source, dict):
        raise SemanticAiError("study draft source must be an object")
    for field in ("dataset", "sourcePath", "revisionUid", "contentSha256"):
        expected = _text(source.get(field), f"source.{field}")
        actual = _text(result_source.get(field), f"result.source.{field}")
        if field == "contentSha256":
            expected, actual = expected.lower(), actual.lower()
        if actual != expected:
            raise SemanticAiError(f"study draft source identity mismatch: {field}")
    if bool(result_source.get("contentComplete")) != bool(content_complete):
        raise SemanticAiError("study draft contentComplete does not match deterministic packet coverage")
    analysis = result.get("workbookAnalysis")
    if not isinstance(analysis, dict):
        raise SemanticAiError("workbookAnalysis must be an object")
    if str(analysis.get("verificationStatus") or "").upper() not in {"NEEDS_REVIEW", "EXCLUDED"}:
        raise SemanticAiError("an AI draft cannot self-verify a workbook analysis")
    for study in result.get("studies", []):
        if str(study.get("verificationStatus") or "").upper() != "NEEDS_REVIEW":
            raise SemanticAiError("an AI draft cannot self-verify a study")
        for comparison in study.get("comparisons", []):
            if str(comparison.get("verificationStatus") or "").upper() != "NEEDS_REVIEW":
                raise SemanticAiError("an AI draft cannot self-verify a comparison")
            if bool(comparison.get("aggregationEligible")):
                raise SemanticAiError("an AI draft cannot make a comparison aggregation-eligible")
            if comparison.get("effects"):
                raise SemanticAiError("AI drafts cannot calculate effects")
        for conclusion in study.get("conclusions", []):
            if str(conclusion.get("causalStrength") or "").upper() == "CAUSAL":
                raise SemanticAiError("AI drafts cannot claim causality")
    try:
        return validate_study_manifest(result, evidence_checker=evidence_checker)
    except StudyContractError as exc:
        raise SemanticAiError(f"invalid canonical study draft: {exc}") from exc


def _normalize_ai_draft_enums(result: dict[str, Any]) -> dict[str, Any]:
    """Normalize conservative AI spelling variants without changing approval state."""

    normalized = copy.deepcopy(result)
    status_aliases = {
        "NOT_ASSESSED": "UNASSESSED",
        "NOT ASSESSED": "UNASSESSED",
        "UNKNOWN": "UNASSESSED",
        "POSSIBLE_CONFOUNDING": "POSSIBLE",
        "POSSIBLY_CONFOUNDED": "POSSIBLE",
        "MULTIPLE_FACTORS": "MULTI_FACTOR",
        "MULTIFACTOR": "MULTI_FACTOR",
        "DESCRIPTIVE_ONLY": "DESCRIPTIVE",
    }

    def clean(value: object) -> str:
        status = str(value or "").strip().upper()
        return status_aliases.get(status, status)

    for study in normalized.get("studies", []):
        for field in (
            "verificationStatus",
            "comparabilityStatus",
            "confoundingStatus",
        ):
            if field in study:
                study[field] = clean(study[field])
        for factor in study.get("factors", []):
            if "isolationStatus" in factor:
                isolation_status = clean(factor["isolationStatus"])
                if isolation_status not in {
                    "CONFOUNDED",
                    "ISOLATED",
                    "MULTI_FACTOR",
                    "UNASSESSED",
                }:
                    isolation_status = "UNASSESSED"
                factor["isolationStatus"] = isolation_status
        for arm in study.get("arms", []):
            if "role" in arm:
                role = clean(arm["role"])
                if role == "NORMAL":
                    role = "REFERENCE"
                if role not in {
                    "CONTROL",
                    "COMPARATOR",
                    "TREATMENT",
                    "TEST",
                    "BEFORE",
                    "AFTER",
                    "REFERENCE",
                    "OTHER",
                }:
                    role = "OTHER"
                arm["role"] = role
        for outcome in study.get("outcomes", []):
            if "favorableDirection" in outcome:
                direction = clean(outcome["favorableDirection"])
                if direction == "UNASSESSED" or direction not in {
                    "HIGHER",
                    "LOWER",
                    "NONE",
                    "TARGET",
                    "UNKNOWN",
                }:
                    direction = "UNKNOWN"
                outcome["favorableDirection"] = direction
            observations_by_arm: dict[str, list[dict[str, Any]]] = {}
            for observation in outcome.get("observations", []):
                observations_by_arm.setdefault(
                    str(observation.get("arm") or ""),
                    [],
                ).append(observation)
            for observations in observations_by_arm.values():
                if len(observations) < 2:
                    continue
                for observation in observations:
                    if not str(observation.get("stratumKey") or "").strip() and not str(
                        observation.get("replicateKey") or ""
                    ).strip():
                        observation["replicateKey"] = str(
                            observation.get("key") or ""
                        ).strip()
        for comparison in study.get("comparisons", []):
            for field in (
                "validityStatus",
                "confoundingStatus",
                "verificationStatus",
            ):
                if field in comparison:
                    comparison[field] = clean(comparison[field])
        for conclusion in study.get("conclusions", []):
            if "causalStrength" in conclusion:
                conclusion["causalStrength"] = clean(
                    conclusion["causalStrength"]
                )
    return normalized


RunCommand = Callable[..., subprocess.CompletedProcess[str]]


def _codex_command(command: Sequence[str] | None) -> list[str]:
    if command:
        return list(command)
    executable = shutil.which("codex.cmd" if os.name == "nt" else "codex")
    if not executable:
        raise SemanticAiError("Codex CLI executable was not found on PATH")
    return [executable]


def run_codex_locator(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    chunk: dict[str, Any],
    output_path: str | Path,
    model: str | None = None,
    reasoning_effort: str | None = None,
    codex_command: Sequence[str] | None = None,
    timeout_seconds: int = 900,
    run_command: RunCommand = subprocess.run,
) -> dict[str, Any]:
    """Run one non-interactive, read-only Codex locator pass and validate its output."""

    revision_uid = _text(source.get("revisionUid"), "source.revisionUid")
    content_sha256 = _text(source.get("contentSha256"), "source.contentSha256")
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    prompt = build_locator_prompt(source=source, workbook=workbook, chunk=chunk)
    with tempfile.TemporaryDirectory(prefix="semantic-locator-") as temp_dir:
        schema_path = Path(temp_dir) / "locator.schema.json"
        last_message_path = Path(temp_dir) / "last-message.json"
        schema_path.write_text(
            json.dumps(locator_output_schema(), ensure_ascii=False),
            encoding="utf-8",
        )
        command = [
            *_codex_command(codex_command),
            "exec",
            "--ephemeral",
            "--sandbox",
            "read-only",
            "--output-schema",
            str(schema_path),
            "--output-last-message",
            str(last_message_path),
        ]
        if reasoning_effort:
            command.extend(
                ["-c", f'model_reasoning_effort="{reasoning_effort}"']
            )
        if model:
            command.extend(["--model", model])
        command.append("-")
        completed = run_command(
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
            detail = (completed.stderr or completed.stdout or "").strip()
            raise SemanticAiError(
                f"Codex locator failed with exit code {completed.returncode}: {detail[-2000:]}"
            )
        if not last_message_path.is_file():
            raise SemanticAiError("Codex locator did not produce an output message")
        try:
            result = json.loads(last_message_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SemanticAiError("Codex locator output is not valid JSON") from exc
    validated = validate_locator_result(
        result,
        revision_uid=revision_uid,
        content_sha256=content_sha256,
        chunk=chunk,
    )
    target.write_text(
        json.dumps(validated, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return validated


def validate_batch_locator_result(
    result: dict[str, Any],
    *,
    revision_uid: str,
    content_sha256: str,
    chunks: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Validate batch identity, exact chunk coverage, and every nested locator."""

    if not isinstance(result, dict):
        raise SemanticAiError("batch locator result must be a JSON object")
    if result.get("schemaVersion") != BATCH_LOCATOR_SCHEMA_VERSION:
        raise SemanticAiError("invalid batch locator schemaVersion")
    if result.get("promptVersion") != BATCH_LOCATOR_PROMPT_VERSION:
        raise SemanticAiError("invalid batch locator promptVersion")
    if _text(result.get("revisionUid"), "revisionUid") != revision_uid:
        raise SemanticAiError("batch locator revisionUid does not match the source packet")
    if (
        _text(result.get("contentSha256"), "contentSha256").lower()
        != content_sha256.lower()
    ):
        raise SemanticAiError(
            "batch locator contentSha256 does not match the source packet"
        )
    nested_results = result.get("results")
    if not isinstance(nested_results, list):
        raise SemanticAiError("batch locator results must be a list")
    expected_ids = [
        _text(chunk.get("chunkId") or chunk.get("packetId"), "chunk.chunkId")
        for chunk in chunks
    ]
    result_ids = [
        _text(item.get("chunkId"), f"results[{index}].chunkId")
        if isinstance(item, dict)
        else ""
        for index, item in enumerate(nested_results)
    ]
    if len(set(result_ids)) != len(result_ids):
        raise SemanticAiError("batch locator contains duplicate chunkId results")
    if result_ids != expected_ids:
        raise SemanticAiError(
            "batch locator results must cover every requested chunk in the same order"
        )
    return [
        validate_locator_result(
            item,
            revision_uid=revision_uid,
            content_sha256=content_sha256,
            chunk=chunk,
        )
        for item, chunk in zip(nested_results, chunks, strict=True)
    ]


def run_codex_locator_batch(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    chunks: Sequence[dict[str, Any]],
    output_paths: dict[str, str | Path],
    model: str | None = None,
    reasoning_effort: str | None = None,
    codex_command: Sequence[str] | None = None,
    timeout_seconds: int = 900,
    run_command: RunCommand = subprocess.run,
) -> list[dict[str, Any]]:
    """Run one read-only Codex call and atomically validate all nested locators."""

    if not chunks:
        raise SemanticAiError("batch locator requires at least one source chunk")
    revision_uid = _text(source.get("revisionUid"), "source.revisionUid")
    content_sha256 = _text(source.get("contentSha256"), "source.contentSha256")
    chunk_ids = [
        _text(chunk.get("chunkId") or chunk.get("packetId"), "chunk.chunkId")
        for chunk in chunks
    ]
    if len(set(chunk_ids)) != len(chunk_ids):
        raise SemanticAiError("batch locator source chunks contain duplicate chunkIds")
    missing_outputs = [chunk_id for chunk_id in chunk_ids if chunk_id not in output_paths]
    if missing_outputs:
        raise SemanticAiError(
            "batch locator output path is missing for: " + ", ".join(missing_outputs)
        )
    prompt = build_batch_locator_prompt(
        source=source,
        workbook=workbook,
        chunks=chunks,
    )
    with tempfile.TemporaryDirectory(prefix="semantic-locator-batch-") as temp_dir:
        schema_path = Path(temp_dir) / "locator-batch.schema.json"
        last_message_path = Path(temp_dir) / "last-message.json"
        schema_path.write_text(
            json.dumps(batch_locator_output_schema(), ensure_ascii=False),
            encoding="utf-8",
        )
        command = [
            *_codex_command(codex_command),
            "exec",
            "--ephemeral",
            "--sandbox",
            "read-only",
            "--output-schema",
            str(schema_path),
            "--output-last-message",
            str(last_message_path),
        ]
        if reasoning_effort:
            command.extend(["-c", f'model_reasoning_effort="{reasoning_effort}"'])
        if model:
            command.extend(["--model", model])
        command.append("-")
        completed = run_command(
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
            detail = (completed.stderr or completed.stdout or "").strip()
            raise SemanticAiError(
                "Codex batch locator failed with exit code "
                f"{completed.returncode}: {detail[-2000:]}"
            )
        if not last_message_path.is_file():
            raise SemanticAiError("Codex batch locator did not produce an output message")
        try:
            result = json.loads(last_message_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SemanticAiError("Codex batch locator output is not valid JSON") from exc
    validated = validate_batch_locator_result(
        result,
        revision_uid=revision_uid,
        content_sha256=content_sha256,
        chunks=chunks,
    )
    staged: list[tuple[Path, str]] = []
    for chunk_id, item in zip(chunk_ids, validated, strict=True):
        target = Path(output_paths[chunk_id])
        target.parent.mkdir(parents=True, exist_ok=True)
        staged.append(
            (
                target,
                json.dumps(item, ensure_ascii=False, indent=2) + "\n",
            )
        )
    for target, text in staged:
        target.write_text(text, encoding="utf-8")
    return validated


def run_codex_study_draft(
    *,
    source: dict[str, Any],
    workbook: dict[str, Any],
    locator_results: list[dict[str, Any]],
    focused_chunks: list[dict[str, Any]],
    content_complete: bool,
    output_path: str | Path,
    evidence_checker: Callable[[dict[str, Any]], None] | None = None,
    additional_validator: Callable[[dict[str, Any]], None] | None = None,
    model: str | None = None,
    reasoning_effort: str | None = None,
    codex_command: Sequence[str] | None = None,
    timeout_seconds: int = 1800,
    run_command: RunCommand = subprocess.run,
    structured_output: bool = False,
    exact_prompt_text: str | None = None,
    expected_prompt_sha256: str | None = None,
    ai_call_observer: Callable[[], None] | None = None,
    unsupported_rate_pair_paths: (
        Callable[
            [dict[str, Any]],
            Sequence[tuple[int, int, int]],
        ]
        | None
    ) = None,
) -> dict[str, Any]:
    """Run a read-only Codex draft pass that is unable to self-approve effects."""

    if (exact_prompt_text is None) != (
        expected_prompt_sha256 is None
    ):
        raise SemanticAiError(
            "Exact Study draft prompt text and hash must be supplied together"
        )
    if exact_prompt_text is not None and hashlib.sha256(
        exact_prompt_text.encode("utf-8")
    ).hexdigest() != str(expected_prompt_sha256):
        raise SemanticAiError(
            "Exact Study draft prompt does not match its expected hash"
        )
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    rejected_path = target.with_name(
        target.stem + ".rejected" + target.suffix
    )
    repair_rejected_path = target.with_name(
        target.stem + ".repair-rejected" + target.suffix
    )
    repair_rejected_unsafe_path = target.with_name(
        target.stem + ".repair-rejected.unsafe" + target.suffix
    )

    def write_unsafe_repair(
        result: dict[str, Any],
        *,
        reason: str,
    ) -> None:
        """Preserve a rejected repair while preventing unsafe resume reuse."""

        artifact_text = (
            json.dumps(result, ensure_ascii=False, indent=2) + "\n"
        )
        repair_rejected_path.write_text(
            artifact_text,
            encoding="utf-8",
        )
        artifact_sha256 = hashlib.sha256(
            repair_rejected_path.read_bytes()
        ).hexdigest()
        repair_rejected_unsafe_path.write_text(
            json.dumps(
                {
                    "schemaVersion": "study-repair-rejection-v1",
                    "artifactSha256": artifact_sha256,
                    "reason": reason,
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )

    def reusable_repair_rejected() -> bool:
        """Return true only when the latest repair was not safety-rejected."""

        if not repair_rejected_path.is_file():
            return False
        if not repair_rejected_unsafe_path.is_file():
            return True
        try:
            marker = json.loads(
                repair_rejected_unsafe_path.read_text(encoding="utf-8")
            )
            artifact_sha256 = hashlib.sha256(
                repair_rejected_path.read_bytes()
            ).hexdigest()
        except (OSError, json.JSONDecodeError):
            return False
        return str(marker.get("artifactSha256", "")) != artifact_sha256
    deterministic_rejected_is_current = (
        rejected_path.is_file()
        and (
            not target.is_file()
            or rejected_path.stat().st_mtime_ns
            > target.stat().st_mtime_ns
        )
    )
    source_bound_baseline_path = max(
        (
            path
            for path in (target, rejected_path)
            if path.is_file()
        ),
        key=lambda path: path.stat().st_mtime_ns,
        default=rejected_path,
    )
    if source_bound_baseline_path.is_file():
        try:
            source_bound_baseline = json.loads(
                source_bound_baseline_path.read_text(encoding="utf-8")
            )
            if not isinstance(source_bound_baseline, dict):
                raise ValueError("source-bound baseline is not an object")
            source_bound_candidate = copy.deepcopy(source_bound_baseline)
            if isinstance(source_bound_candidate.get("source"), dict):
                source_bound_candidate["source"]["contentComplete"] = bool(
                    content_complete
                )
            try:
                source_bound_validated = validate_ai_study_draft(
                    source_bound_candidate,
                    source=source,
                    content_complete=content_complete,
                    evidence_checker=evidence_checker,
                )
                _validate_source_explicit_temporal_semantics(
                    source_bound_validated,
                    focused_chunks,
                )
                if additional_validator is not None:
                    additional_validator(source_bound_validated)
            except Exception as exc:
                validation_error = f"{type(exc).__name__}: {exc}"
                arm_identity_target = arm_identity_repair_target(
                    validation_error,
                    source_bound_baseline,
                    focused_chunks,
                )
                b17_report_table_repair = (
                    b17_report_table_repair_applicable(
                        source_bound_baseline,
                        validation_error=validation_error,
                        focused_chunks=focused_chunks,
                    )
                )
                deterministic_result: dict[str, Any] | None = None
                if arm_identity_target is not None:
                    deterministic_result = apply_arm_identity_repair(
                        source_bound_baseline,
                        arm_identity_target,
                    )
                elif b17_report_table_repair:
                    deterministic_result = (
                        apply_b17_report_table_repair(
                            source_bound_baseline,
                            focused_chunks=focused_chunks,
                        )
                    )
                if deterministic_result is not None:
                    if isinstance(deterministic_result.get("source"), dict):
                        deterministic_result["source"][
                            "contentComplete"
                        ] = bool(content_complete)
                    deterministic_validated = validate_ai_study_draft(
                        deterministic_result,
                        source=source,
                        content_complete=content_complete,
                        evidence_checker=evidence_checker,
                    )
                    _validate_source_explicit_temporal_semantics(
                        deterministic_validated,
                        focused_chunks,
                    )
                    if additional_validator is not None:
                        additional_validator(deterministic_validated)
                    target.write_text(
                        json.dumps(
                            deterministic_validated,
                            ensure_ascii=False,
                            indent=2,
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                    return deterministic_validated
            else:
                if (
                    deterministic_rejected_is_current
                    and source_bound_baseline_path == rejected_path
                ):
                    target.write_text(
                        json.dumps(
                            source_bound_validated,
                            ensure_ascii=False,
                            indent=2,
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                    return source_bound_validated
        except (
            OSError,
            ValueError,
            json.JSONDecodeError,
            SemanticAiError,
            ArmIdentityRepairError,
            B17ReportTableRepairError,
        ):
            pass
    if deterministic_rejected_is_current:
        try:
            deterministic_baseline = json.loads(
                rejected_path.read_text(encoding="utf-8")
            )
            if (
                isinstance(deterministic_baseline, dict)
                and _numeric_text_homogeneity_repair_targets(
                    deterministic_baseline
                )
            ):
                deterministic_result = (
                    _apply_deterministic_numeric_text_homogeneity_repair(
                        deterministic_baseline
                    )
                )
                if isinstance(
                    deterministic_result.get("source"), dict
                ):
                    deterministic_result["source"]["contentComplete"] = bool(
                        content_complete
                    )
                try:
                    deterministic_validated = validate_ai_study_draft(
                        deterministic_result,
                        source=source,
                        content_complete=content_complete,
                        evidence_checker=evidence_checker,
                    )
                    _validate_source_explicit_temporal_semantics(
                        deterministic_validated,
                        focused_chunks,
                    )
                    if additional_validator is not None:
                        additional_validator(deterministic_validated)
                except Exception:
                    # Keep the numeric-only correction as the newest rejected
                    # baseline so any unrelated remaining validator error can
                    # be repaired without reviving a discarded model rewrite.
                    rejected_path.write_text(
                        json.dumps(
                            deterministic_result,
                            ensure_ascii=False,
                            indent=2,
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                else:
                    target.write_text(
                        json.dumps(
                            deterministic_validated,
                            ensure_ascii=False,
                            indent=2,
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                    return deterministic_validated
        except (
            OSError,
            ValueError,
            json.JSONDecodeError,
            SemanticAiError,
        ):
            pass
    if deterministic_rejected_is_current:
        try:
            deterministic_baseline = json.loads(
                rejected_path.read_text(encoding="utf-8")
            )
            if not isinstance(deterministic_baseline, dict):
                raise ValueError("rejected draft is not a JSON object")
            baseline_candidate = copy.deepcopy(deterministic_baseline)
            if isinstance(baseline_candidate.get("source"), dict):
                baseline_candidate["source"]["contentComplete"] = bool(
                    content_complete
                )
            try:
                baseline_validated = validate_ai_study_draft(
                    baseline_candidate,
                    source=source,
                    content_complete=content_complete,
                    evidence_checker=evidence_checker,
                )
                _validate_source_explicit_temporal_semantics(
                    baseline_validated,
                    focused_chunks,
                )
                if additional_validator is not None:
                    additional_validator(baseline_validated)
            except Exception as exc:
                validation_error = f"{type(exc).__name__}: {exc}"
                deterministic_path = (
                    _unsupported_rate_pair_observation(validation_error)
                )
                comparison_target = (
                    _incompatible_raw_comparison_target(
                        validation_error
                    )
                )
                synthetic_reference_target = (
                    _synthetic_reference_arm_target(
                        validation_error
                    )
                )
                numeric_header_target = (
                    numeric_header_series_repair_target(
                        validation_error,
                        deterministic_baseline,
                        focused_chunks,
                    )
                )
                merged_header_target = (
                    merged_header_series_repair_target(
                        validation_error
                    )
                )
                b04_b08_target = b04_b08_repair_target(
                    validation_error,
                    deterministic_baseline,
                    focused_chunks,
                )
                composite_outcome_repair = (
                    composite_outcome_repair_applicable(
                        deterministic_baseline,
                        validation_error=validation_error,
                        focused_chunks=focused_chunks,
                    )
                )
                single_outcome_repair = (
                    single_outcome_repair_applicable(
                        deterministic_baseline,
                        validation_error=validation_error,
                        focused_chunks=focused_chunks,
                    )
                )
                a1_union_evidence_repair = (
                    _a1_union_evidence_repair_applicable(
                        validation_error,
                        deterministic_baseline,
                    )
                )
                deterministic_result: dict[str, Any] | None = None
                if deterministic_path is not None:
                    deterministic_result = (
                        _apply_deterministic_unsupported_rate_pair_repair(
                            deterministic_baseline,
                            deterministic_path,
                        )
                    )
                elif comparison_target is not None:
                    deterministic_result = (
                        _apply_deterministic_incompatible_raw_comparison_repair(
                            deterministic_baseline,
                            comparison_target,
                        )
                    )
                elif synthetic_reference_target is not None:
                    deterministic_result = (
                        _apply_deterministic_synthetic_reference_arm_repair(
                            deterministic_baseline,
                            target=synthetic_reference_target,
                            focused_chunks=focused_chunks,
                        )
                    )
                elif numeric_header_target is not None:
                    deterministic_result = (
                        apply_numeric_header_series_repair(
                            deterministic_baseline,
                            numeric_header_target,
                        )
                    )
                elif merged_header_target is not None:
                    deterministic_result = (
                        apply_merged_header_series_repair(
                            deterministic_baseline,
                            merged_header_target,
                        )
                    )
                elif b04_b08_target is not None:
                    deterministic_result = apply_b04_b08_repair(
                        deterministic_baseline,
                        b04_b08_target,
                    )
                elif composite_outcome_repair:
                    deterministic_result = (
                        apply_deterministic_composite_outcome_repair(
                            deterministic_baseline,
                            validation_error=validation_error,
                            focused_chunks=focused_chunks,
                        )
                    )
                elif single_outcome_repair:
                    deterministic_result = (
                        apply_deterministic_single_outcome_repair(
                            deterministic_baseline,
                            validation_error=validation_error,
                            focused_chunks=focused_chunks,
                        )
                    )
                elif a1_union_evidence_repair:
                    deterministic_result = (
                        _apply_deterministic_a1_union_evidence_repair(
                            deterministic_baseline
                        )
                    )
                if deterministic_result is not None:
                    if isinstance(
                        deterministic_result.get("source"), dict
                    ):
                        deterministic_result["source"][
                            "contentComplete"
                        ] = bool(content_complete)
                    try:
                        deterministic_validated = validate_ai_study_draft(
                            deterministic_result,
                            source=source,
                            content_complete=content_complete,
                            evidence_checker=evidence_checker,
                        )
                        _validate_source_explicit_temporal_semantics(
                            deterministic_validated,
                            focused_chunks,
                        )
                        if additional_validator is not None:
                            additional_validator(deterministic_validated)
                    except Exception:
                        rejected_path.write_text(
                            json.dumps(
                                deterministic_result,
                                ensure_ascii=False,
                                indent=2,
                            )
                            + "\n",
                            encoding="utf-8",
                        )
                    else:
                        target.write_text(
                            json.dumps(
                                deterministic_validated,
                                ensure_ascii=False,
                                indent=2,
                            )
                            + "\n",
                            encoding="utf-8",
                        )
                        return deterministic_validated
        except (
            OSError,
            ValueError,
            json.JSONDecodeError,
            SemanticAiError,
            ArmIdentityRepairError,
            B04B08RepairError,
            B17ReportTableRepairError,
            CompositeOutcomeRepairError,
            MergedHeaderRepairError,
            NumericHeaderRepairError,
            SingleOutcomeRepairError,
        ):
            pass
    checkpoint_candidates = tuple(
        path
        for path in (target, rejected_path, repair_rejected_path)
        if path.is_file()
        and (
            path != repair_rejected_path
            or reusable_repair_rejected()
        )
    )
    ordered_checkpoint_candidates = tuple(sorted(
        checkpoint_candidates,
        key=lambda path: path.stat().st_mtime_ns,
        reverse=True,
    ))
    rejected_input_path = max(
        checkpoint_candidates,
        key=lambda path: path.stat().st_mtime_ns,
        default=rejected_path,
    )
    prompt = ""
    repair_baseline: dict[str, Any] | None = None
    reference_only_repair = False
    temporal_stage_only_repair = False
    numeric_text_homogeneity_repair = False
    unsupported_rate_pair_repair: tuple[int, int, int] | None = None
    if rejected_input_path.is_file():
        try:
            rejected = json.loads(
                rejected_input_path.read_text(encoding="utf-8")
            )
            if not isinstance(rejected, dict):
                raise ValueError("rejected draft is not a JSON object")
            try:
                rejected_validated = validate_ai_study_draft(
                    rejected,
                    source=source,
                    content_complete=content_complete,
                    evidence_checker=evidence_checker,
                )
                _validate_source_explicit_temporal_semantics(
                    rejected_validated,
                    focused_chunks,
                )
                if additional_validator is not None:
                    additional_validator(rejected_validated)
            except Exception as exc:
                validation_error = f"{type(exc).__name__}: {exc}"
                aggregate_alignment_result = (
                    _apply_deterministic_aggregate_identity_alignment_repair(
                        rejected,
                        validation_error,
                    )
                )
                if aggregate_alignment_result is not None:
                    if isinstance(
                        aggregate_alignment_result.get("source"),
                        dict,
                    ):
                        aggregate_alignment_result["source"][
                            "contentComplete"
                        ] = bool(content_complete)
                    try:
                        aggregate_alignment_validated = (
                            validate_ai_study_draft(
                                aggregate_alignment_result,
                                source=source,
                                content_complete=content_complete,
                                evidence_checker=evidence_checker,
                            )
                        )
                        _validate_source_explicit_temporal_semantics(
                            aggregate_alignment_validated,
                            focused_chunks,
                        )
                        if additional_validator is not None:
                            additional_validator(
                                aggregate_alignment_validated
                            )
                    except Exception as aggregate_error:
                        rejected_path.write_text(
                            json.dumps(
                                aggregate_alignment_result,
                                ensure_ascii=False,
                                indent=2,
                            )
                            + "\n",
                            encoding="utf-8",
                        )
                        rejected = aggregate_alignment_result
                        validation_error = (
                            f"{type(aggregate_error).__name__}: "
                            f"{aggregate_error}"
                        )
                    else:
                        target.write_text(
                            json.dumps(
                                aggregate_alignment_validated,
                                ensure_ascii=False,
                                indent=2,
                            )
                            + "\n",
                            encoding="utf-8",
                        )
                        return aggregate_alignment_validated
                # Interrupted runs may leave a newer invalid model response
                # beside an older repair checkpoint that becomes valid after
                # a local validator correction. Only after the newest
                # candidate fails, promote the first older candidate that
                # passes every current source-bound validator.
                for fallback_path in ordered_checkpoint_candidates[1:]:
                    try:
                        fallback = json.loads(
                            fallback_path.read_text(encoding="utf-8")
                        )
                        if not isinstance(fallback, dict):
                            continue
                        fallback_candidate = copy.deepcopy(fallback)
                        if isinstance(
                            fallback_candidate.get("source"),
                            dict,
                        ):
                            fallback_candidate["source"][
                                "contentComplete"
                            ] = bool(content_complete)
                        fallback_validated = validate_ai_study_draft(
                            fallback_candidate,
                            source=source,
                            content_complete=content_complete,
                            evidence_checker=evidence_checker,
                        )
                        _validate_source_explicit_temporal_semantics(
                            fallback_validated,
                            focused_chunks,
                        )
                        if additional_validator is not None:
                            additional_validator(fallback_validated)
                    except Exception:
                        continue
                    target.write_text(
                        json.dumps(
                            fallback_validated,
                            ensure_ascii=False,
                            indent=2,
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                    return fallback_validated
                normalized_error = validation_error.casefold()
                reference_only_repair = (
                    "referenc" in normalized_error
                    and "unknown" in normalized_error
                )
                temporal_stage_only_repair = (
                    "source-authored temporal stage baseline mapping"
                    in normalized_error
                )
                numeric_text_homogeneity_repair = (
                    "mixes quantitative measurements with qualitative-only"
                    in normalized_error
                    and bool(
                        _numeric_text_homogeneity_repair_targets(rejected)
                    )
                )
                unsupported_rate_pair_repair = (
                    _unsupported_rate_pair_observation(validation_error)
                )
                if unsupported_rate_pair_repair is not None:
                    deterministic_result = rejected
                    repaired_paths: set[tuple[int, int, int]] = set()
                    while unsupported_rate_pair_repair is not None:
                        repair_batch = [unsupported_rate_pair_repair]
                        if (
                            not repaired_paths
                            and unsupported_rate_pair_paths is not None
                        ):
                            try:
                                repair_batch = list(
                                    unsupported_rate_pair_paths(
                                        deterministic_result
                                    )
                                )
                            except Exception as exc:
                                raise SemanticAiError(
                                    "Unsupported rate-pair batch discovery "
                                    "failed"
                                ) from exc
                            if unsupported_rate_pair_repair not in repair_batch:
                                raise SemanticAiError(
                                    "Unsupported rate-pair batch discovery "
                                    "omitted the validator-identified path"
                                )
                        for repair_path in repair_batch:
                            if (
                                not isinstance(repair_path, tuple)
                                or len(repair_path) != 3
                                or not all(
                                    isinstance(index, int)
                                    and index >= 0
                                    for index in repair_path
                                )
                            ):
                                raise SemanticAiError(
                                    "Unsupported rate-pair batch returned "
                                    "an invalid observation path"
                                )
                            if repair_path in repaired_paths:
                                raise SemanticAiError(
                                    "Unsupported rate-pair repair did not "
                                    "resolve its validator-identified path"
                                )
                            repaired_paths.add(repair_path)
                            deterministic_result = (
                                _apply_deterministic_unsupported_rate_pair_repair(
                                    deterministic_result,
                                    repair_path,
                                )
                            )
                            if isinstance(
                                deterministic_result.get("source"), dict
                            ):
                                deterministic_result["source"][
                                    "contentComplete"
                                ] = bool(content_complete)
                            # Persist each safe two-field projection so an
                            # interrupted corpus run can resume from the
                            # latest locally repaired pair.
                            rejected_path.write_text(
                                json.dumps(
                                    deterministic_result,
                                    ensure_ascii=False,
                                    indent=2,
                                )
                                + "\n",
                                encoding="utf-8",
                            )
                        try:
                            deterministic_validated = (
                                validate_ai_study_draft(
                                    deterministic_result,
                                    source=source,
                                    content_complete=content_complete,
                                    evidence_checker=evidence_checker,
                                )
                            )
                            _validate_source_explicit_temporal_semantics(
                                deterministic_validated,
                                focused_chunks,
                            )
                            if additional_validator is not None:
                                additional_validator(
                                    deterministic_validated
                                )
                        except Exception as deterministic_error:
                            unsupported_rate_pair_repair = (
                                _unsupported_rate_pair_observation(
                                    f"{type(deterministic_error).__name__}: "
                                    f"{deterministic_error}"
                                )
                            )
                            if unsupported_rate_pair_repair is not None:
                                continue
                            # A second, distinct validation error may remain.
                            # Keep the pair-only projection as the next
                            # trusted repair baseline.
                        else:
                            target.write_text(
                                json.dumps(
                                    deterministic_validated,
                                    ensure_ascii=False,
                                    indent=2,
                                )
                                + "\n",
                                encoding="utf-8",
                            )
                            return deterministic_validated
                        return run_codex_study_draft(
                            source=source,
                            workbook=workbook,
                            locator_results=locator_results,
                            focused_chunks=focused_chunks,
                            content_complete=content_complete,
                            output_path=output_path,
                            evidence_checker=evidence_checker,
                            additional_validator=additional_validator,
                            model=model,
                            reasoning_effort=reasoning_effort,
                            codex_command=codex_command,
                            timeout_seconds=timeout_seconds,
                            run_command=run_command,
                            structured_output=structured_output,
                            exact_prompt_text=exact_prompt_text,
                            expected_prompt_sha256=expected_prompt_sha256,
                            ai_call_observer=ai_call_observer,
                            unsupported_rate_pair_paths=(
                                unsupported_rate_pair_paths
                            ),
                        )
                source_prompt = (
                    ""
                    if reference_only_repair
                    else build_study_draft_prompt(
                        source=source,
                        workbook=workbook,
                        locator_results=locator_results,
                        focused_chunks=focused_chunks,
                    )
                )
                prompt = build_study_draft_repair_prompt(
                    rejected,
                    validation_error,
                    source_prompt=source_prompt,
                )
                repair_baseline = rejected
            else:
                # A checkpoint candidate can become valid after a stricter
                # local validator is corrected. Promote that already
                # source-bound result instead of spending another model call
                # to regenerate equivalent content.
                target.write_text(
                    json.dumps(
                        rejected_validated,
                        ensure_ascii=False,
                        indent=2,
                    )
                    + "\n",
                    encoding="utf-8",
                )
                return rejected_validated
        except (OSError, ValueError, json.JSONDecodeError):
            prompt = ""
    if not prompt:
        prompt = build_study_draft_prompt(
            source=source,
            workbook=workbook,
            locator_results=locator_results,
            focused_chunks=focused_chunks,
        )
    if exact_prompt_text is not None:
        # Deterministic local repairs above may return without an AI call.
        # Any new AI call, however, must use the exact budgeted request and
        # cannot inherit a stale model-repair prompt.
        prompt = exact_prompt_text
        repair_baseline = None
        reference_only_repair = False
        temporal_stage_only_repair = False
        numeric_text_homogeneity_repair = False
        unsupported_rate_pair_repair = None
    if len(prompt) > STUDY_DRAFT_MAX_INPUT_CHARS:
        raise SemanticAiError(
            "Study draft prompt exceeds the fail-closed input budget: "
            f"{len(prompt)} > {STUDY_DRAFT_MAX_INPUT_CHARS} characters. "
            "No source cells were silently dropped; use a deterministic "
            "series mapping or a staged consolidation."
        )
    transport_token = uuid.uuid4().hex
    transport_prefix = f".{target.name}.{transport_token}"
    last_message_path = target.parent / (
        transport_prefix + ".last-message.json"
    )
    schema_path = target.parent / (
        transport_prefix + ".study-draft.schema.json"
    )
    try:

        def execute_attempt(
            attempt_prompt: str,
            *,
            use_schema: bool,
            attempt_label: str,
        ) -> str:
            if last_message_path.is_file():
                last_message_path.unlink()
            command = [
                *_codex_command(codex_command),
                "exec",
                "--ephemeral",
                "--sandbox",
                "read-only",
                "--output-last-message",
                str(last_message_path),
            ]
            if use_schema:
                if not schema_path.is_file():
                    schema_path.write_text(
                        json.dumps(
                            study_draft_output_schema(),
                            ensure_ascii=False,
                        ),
                        encoding="utf-8",
                    )
                command.extend(["--output-schema", str(schema_path)])
            if reasoning_effort:
                command.extend(
                    ["-c", f'model_reasoning_effort="{reasoning_effort}"']
                )
            if model:
                command.extend(["--model", model])
            command.append("-")
            if ai_call_observer is not None:
                ai_call_observer()
            completed = run_command(
                command,
                input=attempt_prompt,
                text=True,
                encoding="utf-8",
                errors="replace",
                capture_output=True,
                timeout=timeout_seconds,
                check=False,
            )
            if completed.returncode != 0:
                detail = (
                    completed.stderr or completed.stdout or ""
                ).strip()
                raise SemanticAiError(
                    f"Codex study draft {attempt_label} failed with exit "
                    f"code {completed.returncode}: {detail[-2000:]}"
                )
            if not last_message_path.is_file():
                detail = (
                    completed.stderr or completed.stdout or ""
                ).strip()
                raise SemanticAiError(
                    f"Codex study draft {attempt_label} did not produce an "
                    "output message"
                    + (f": {detail[-2000:]}" if detail else "")
                )
            response = last_message_path.read_text(
                encoding="utf-8"
            ).strip()
            if response.startswith("```"):
                lines = response.splitlines()
                if lines and lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].strip() == "```":
                    lines = lines[:-1]
                response = "\n".join(lines).strip()
            return response

        response_text = execute_attempt(
            prompt,
            use_schema=structured_output,
            attempt_label="initial attempt",
        )
        try:
            result = json.loads(response_text)
        except json.JSONDecodeError as first_error:
            if exact_prompt_text is not None:
                raise SemanticAiError(
                    "Exact Study draft output is invalid JSON; automatic "
                    "AI repair would change the budgeted request"
                ) from first_error
            retry_prompt = build_study_json_retry_prompt(
                prompt,
                str(first_error),
            )
            if len(retry_prompt) > STUDY_DRAFT_MAX_INPUT_CHARS:
                raise SemanticAiError(
                    "Study draft JSON retry exceeds the fail-closed input "
                    "budget; no source cells were removed."
                ) from first_error
            retry_text = execute_attempt(
                retry_prompt,
                use_schema=True,
                attempt_label="JSON retry",
            )
            try:
                result = json.loads(retry_text)
            except json.JSONDecodeError as retry_error:
                raise SemanticAiError(
                    "Codex study draft output is not valid JSON after one "
                    "bounded schema retry"
                ) from retry_error
    finally:
        for transport_path in (last_message_path, schema_path):
            transport_path.unlink(missing_ok=True)
    # Packet coverage is deterministic workflow state, not an AI judgment.
    # Restore only this boolean before the strict source-identity validator;
    # all remaining source fields must still match exactly.
    if isinstance(result, dict) and isinstance(result.get("source"), dict):
        result["source"]["contentComplete"] = bool(content_complete)
    if (
        repair_baseline is not None
        and reference_only_repair
        and _reference_repair_projection(result)
        != _reference_repair_projection(repair_baseline)
    ):
        write_unsafe_repair(
            result,
            reason=(
                "Codex reference repair changed fields outside the allowed "
                "reference paths"
            ),
        )
        raise SemanticAiError(
            "Codex reference repair changed fields outside the allowed "
            "reference paths"
        )
    if (
        repair_baseline is not None
        and temporal_stage_only_repair
        and _temporal_stage_repair_projection(result)
        != _temporal_stage_repair_projection(repair_baseline)
    ):
        write_unsafe_repair(
            result,
            reason=(
                "Codex temporal-stage repair changed fields outside the "
                "allowed baselineCondition, changedCondition, and "
                "factorValue.isBaseline paths"
            ),
        )
        raise SemanticAiError(
            "Codex temporal-stage repair changed fields outside the "
            "allowed baselineCondition, changedCondition, and "
            "factorValue.isBaseline paths"
        )
    if repair_baseline is not None and numeric_text_homogeneity_repair:
        try:
            _validate_numeric_text_homogeneity_repair(
                repair_baseline,
                result,
            )
        except SemanticAiError as exc:
            write_unsafe_repair(
                result,
                reason=str(exc),
            )
            raise
    if (
        repair_baseline is not None
        and unsupported_rate_pair_repair is not None
    ):
        try:
            _validate_unsupported_rate_pair_repair(
                repair_baseline,
                result,
                unsupported_rate_pair_repair,
            )
        except SemanticAiError as exc:
            write_unsafe_repair(
                result,
                reason=str(exc),
            )
            raise
    try:
        validated = validate_ai_study_draft(
            result,
            source=source,
            content_complete=content_complete,
            evidence_checker=evidence_checker,
        )
        _validate_source_explicit_temporal_semantics(
            validated,
            focused_chunks,
        )
        if additional_validator is not None:
            additional_validator(validated)
    except Exception:
        failure_path = (
            repair_rejected_path
            if repair_baseline is not None
            else rejected_path
        )
        failure_path.write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        if failure_path == repair_rejected_path:
            repair_rejected_unsafe_path.unlink(missing_ok=True)
        if repair_baseline is None:
            if exact_prompt_text is not None:
                raise SemanticAiError(
                    "Exact Study draft failed validation; automatic AI "
                    "repair would change the budgeted request"
                )
            return run_codex_study_draft(
                source=source,
                workbook=workbook,
                locator_results=locator_results,
                focused_chunks=focused_chunks,
                content_complete=content_complete,
                output_path=output_path,
                evidence_checker=evidence_checker,
                additional_validator=additional_validator,
                model=model,
                reasoning_effort=reasoning_effort,
                codex_command=codex_command,
                timeout_seconds=timeout_seconds,
                run_command=run_command,
                structured_output=structured_output,
                exact_prompt_text=exact_prompt_text,
                expected_prompt_sha256=expected_prompt_sha256,
                ai_call_observer=ai_call_observer,
                unsupported_rate_pair_paths=(
                    unsupported_rate_pair_paths
                ),
            )
        raise
    target.write_text(
        json.dumps(validated, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return validated


__all__ = [
    "BATCH_LOCATOR_PROMPT_VERSION",
    "BATCH_LOCATOR_SCHEMA_VERSION",
    "LOCATOR_PROMPT_VERSION",
    "LOCATOR_SCHEMA_VERSION",
    "STUDY_DRAFT_PROMPT_VERSION",
    "STUDY_DRAFT_MAX_INPUT_CHARS",
    "SemanticAiError",
    "batch_locator_output_schema",
    "build_batch_locator_prompt",
    "build_locator_prompt",
    "build_study_draft_prompt",
    "locator_output_schema",
    "run_codex_locator",
    "run_codex_locator_batch",
    "run_codex_study_draft",
    "study_draft_output_schema",
    "validate_ai_study_draft",
    "validate_batch_locator_result",
    "validate_locator_result",
]
