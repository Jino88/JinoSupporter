"""Deterministic Top-K matching of fingerprints to approved form templates."""

from __future__ import annotations

import json
from typing import Any, Iterable

from inference_data_ai_extraction_recipe import validate_form_template
from inference_data_ai_structure_fingerprint import (
    normalize_text,
    validate_structure_fingerprint,
)


MATCH_DECISION_SCHEMA_VERSION = "excel-template-match-decision-v1"
MATCHER_VERSION = "deterministic-template-matcher-v1.0"

EXACT_REUSE_THRESHOLD = 0.97
AI_MATCH_THRESHOLD = 0.90
VARIANT_THRESHOLD = 0.75


def _ratio(left: int, right: int) -> float:
    if left == right:
        return 1.0
    if left <= 0 or right <= 0:
        return 0.0
    return min(left, right) / max(left, right)


def _jaccard(left: Iterable[str], right: Iterable[str]) -> float:
    a, b = set(left), set(right)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _json_set(values: Iterable[Any]) -> set[str]:
    return {
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        for value in values
    }


def _anchor_set(fingerprint: dict[str, Any]) -> set[str]:
    anchors: set[str] = set()
    for sheet in fingerprint.get("sheets") or []:
        anchors.update(normalize_text(value) for value in sheet.get("titleTokens") or [])
        anchors.update(
            normalize_text(anchor.get("text"))
            for anchor in sheet.get("anchorSketches") or []
            if anchor.get("text")
        )
    return {anchor for anchor in anchors if anchor}


def _sheet_score(
    candidate: dict[str, Any],
    representative: dict[str, Any],
) -> tuple[float, dict[str, float]]:
    candidate_anchors = {
        normalize_text(item.get("text"))
        for item in candidate.get("anchorSketches") or []
        if item.get("text")
    } | {
        normalize_text(value) for value in candidate.get("titleTokens") or []
    }
    representative_anchors = {
        normalize_text(item.get("text"))
        for item in representative.get("anchorSketches") or []
        if item.get("text")
    } | {
        normalize_text(value) for value in representative.get("titleTokens") or []
    }
    anchors = _jaccard(candidate_anchors, representative_anchors)

    candidate_regions = list(candidate.get("tableRegionSketches") or [])
    representative_regions = list(representative.get("tableRegionSketches") or [])
    if candidate_regions and representative_regions:
        left = candidate_regions[0]
        right = representative_regions[0]
        table_geometry = (
            0.45
            * _ratio(
                int(left.get("columnCount") or 0),
                int(right.get("columnCount") or 0),
            )
            + 0.35
            * (
                1.0
                if left.get("rowCountBucket") == right.get("rowCountBucket")
                else _ratio(
                    int(candidate.get("rowCount") or 0),
                    int(representative.get("rowCount") or 0),
                )
            )
            + 0.20
            * _ratio(
                int(left.get("headerDepth") or 0),
                int(right.get("headerDepth") or 0),
            )
        )
    else:
        table_geometry = 1.0 if not candidate_regions and not representative_regions else 0.0

    header_roles = _jaccard(
        candidate.get("headerRoleSketch") or [],
        representative.get("headerRoleSketch") or [],
    )
    formulas = _jaccard(
        candidate.get("formulaPatternHashes") or [],
        representative.get("formulaPatternHashes") or [],
    )
    merges = _jaccard(
        _json_set(candidate.get("mergedGeometry") or []),
        _json_set(representative.get("mergedGeometry") or []),
    )
    shape = (
        0.55
        * _ratio(
            int(candidate.get("columnCount") or 0),
            int(representative.get("columnCount") or 0),
        )
        + 0.30
        * _ratio(
            int(candidate.get("rowCount") or 0),
            int(representative.get("rowCount") or 0),
        )
        + 0.15
        * (
            1.0
            if bool(candidate.get("tabular")) == bool(representative.get("tabular"))
            else 0.0
        )
    )
    components = {
        "anchors": anchors,
        "tableGeometry": table_geometry,
        "headerRoles": header_roles,
        "formulas": formulas,
        "merges": merges,
        "shape": shape,
    }
    score = (
        0.30 * anchors
        + 0.20 * table_geometry
        + 0.15 * header_roles
        + 0.15 * formulas
        + 0.10 * merges
        + 0.10 * shape
    )
    return score, components


def _workbook_sheet_score(
    fingerprint: dict[str, Any],
    representative: dict[str, Any],
) -> tuple[float, dict[str, float]]:
    candidate_sheets = list(fingerprint.get("sheets") or [])
    known_sheets = list(representative.get("sheets") or [])
    if not candidate_sheets and not known_sheets:
        return 1.0, {}
    if not candidate_sheets or not known_sheets:
        return 0.0, {}

    remaining = set(range(len(known_sheets)))
    scores: list[float] = []
    component_totals: dict[str, float] = {}
    for candidate in candidate_sheets:
        if not remaining:
            scores.append(0.0)
            continue
        ranked: list[tuple[float, int, dict[str, float]]] = []
        for index in remaining:
            score, components = _sheet_score(candidate, known_sheets[index])
            ranked.append((score, index, components))
        score, best_index, components = max(
            ranked,
            key=lambda item: (item[0], -item[1]),
        )
        remaining.remove(best_index)
        scores.append(score)
        for key, value in components.items():
            component_totals[key] = component_totals.get(key, 0.0) + value
    scores.extend(0.0 for _ in remaining)
    denominator = max(len(candidate_sheets), len(known_sheets))
    return (
        sum(scores) / denominator,
        {
            key: round(value / denominator, 6)
            for key, value in component_totals.items()
        },
    )


def _hard_gate(
    fingerprint: dict[str, Any],
    template: dict[str, Any],
) -> tuple[bool, list[str], float]:
    envelope = template["acceptedFingerprintEnvelope"]
    workbook = fingerprint["workbook"]
    failures: list[str] = []
    sheet_count = int(workbook.get("sheetCount") or 0)
    sheet_range = envelope["allowedSheetCount"]
    if not sheet_range[0] <= sheet_count <= sheet_range[1]:
        failures.append("SHEET_COUNT_OUT_OF_ENVELOPE")
    tabular_count = int(workbook.get("tabularSheetCount") or 0)
    tabular_range = envelope["allowedTabularSheetCount"]
    if not tabular_range[0] <= tabular_count <= tabular_range[1]:
        failures.append("TABULAR_SHEET_COUNT_OUT_OF_ENVELOPE")

    present = _anchor_set(fingerprint)
    required = {
        normalize_text(value)
        for value in envelope.get("requiredAnchors") or []
        if normalize_text(value)
    }
    coverage = 1.0 if not required else len(required & present) / len(required)
    if coverage < 1.0:
        failures.append("REQUIRED_ANCHOR_MISSING")
    return not failures, failures, coverage


def rank_templates(
    fingerprint: dict[str, Any],
    templates: Iterable[dict[str, Any]],
    *,
    top_k: int = 3,
    approved_only: bool = True,
) -> list[dict[str, Any]]:
    """Return deterministic Top-K candidates, including hard-gate failures."""

    validate_structure_fingerprint(fingerprint)
    ranked: list[dict[str, Any]] = []
    for template in templates:
        validate_form_template(template)
        if approved_only and template.get("status") != "APPROVED":
            continue
        representative = template["representativeFingerprint"]
        hard_gate_passed, failures, coverage = _hard_gate(fingerprint, template)
        exact = (
            fingerprint["fingerprintSha256"]
            == representative["fingerprintSha256"]
        )
        score, components = _workbook_sheet_score(fingerprint, representative)
        if exact and hard_gate_passed:
            score = 1.0
        ranked.append(
            {
                "templateId": template["templateId"],
                "familyId": template["familyId"],
                "templateVersion": int(template["templateVersion"]),
                "recipeRef": str(template["recipeRefs"][-1]),
                "score": round(score, 6),
                "fingerprintExact": exact,
                "hardGatePassed": hard_gate_passed,
                "hardGateFailures": failures,
                "requiredAnchorCoverage": round(coverage, 6),
                "components": components,
            }
        )
    ranked.sort(
        key=lambda item: (
            not item["hardGatePassed"],
            -float(item["score"]),
            str(item["templateId"]),
            int(item["templateVersion"]),
        )
    )
    return ranked[: max(int(top_k), 0)]


def fingerprint_similarity(
    candidate: dict[str, Any],
    representative: dict[str, Any],
) -> dict[str, Any]:
    """Compare two validated fingerprints without constructing a template."""

    validate_structure_fingerprint(candidate)
    validate_structure_fingerprint(representative)
    score, components = _workbook_sheet_score(candidate, representative)
    return {
        "score": round(score, 6),
        "components": components,
        "fingerprintExact": (
            candidate["fingerprintSha256"]
            == representative["fingerprintSha256"]
        ),
    }


def decide_template_match(
    fingerprint: dict[str, Any],
    templates: Iterable[dict[str, Any]],
    *,
    top_k: int = 3,
) -> dict[str, Any]:
    candidates = rank_templates(fingerprint, templates, top_k=top_k)
    viable = [item for item in candidates if item["hardGatePassed"]]
    selected = viable[0] if viable else None
    if selected is None:
        decision = "NEW_TEMPLATE_REQUIRED"
    elif selected["fingerprintExact"] or selected["score"] >= EXACT_REUSE_THRESHOLD:
        decision = "EXACT_REUSE"
    elif selected["score"] >= AI_MATCH_THRESHOLD:
        decision = "AI_MATCH_REQUIRED"
    elif selected["score"] >= VARIANT_THRESHOLD:
        decision = "VARIANT_PATCH_REQUIRED"
    else:
        decision = "NEW_TEMPLATE_REQUIRED"
    return {
        "schemaVersion": MATCH_DECISION_SCHEMA_VERSION,
        "matcherVersion": MATCHER_VERSION,
        "sourceSha256": str(fingerprint.get("sourceSha256") or ""),
        "fingerprintSha256": fingerprint["fingerprintSha256"],
        "decision": decision,
        "selectedTemplate": (
            f"{selected['templateId']}@{selected['templateVersion']}"
            if selected is not None
            else None
        ),
        "selectedRecipe": selected["recipeRef"] if selected is not None else None,
        "topCandidates": candidates,
        "requiredAnchorCoverage": (
            selected["requiredAnchorCoverage"] if selected is not None else 0.0
        ),
        "ai": {
            "used": False,
            "reason": (
                "AMBIGUOUS_TOP_K" if decision == "AI_MATCH_REQUIRED" else None
            ),
            "model": None,
            "promptVersion": None,
        },
        "observedDeviations": (
            selected["hardGateFailures"] if selected is not None else []
        ),
    }


__all__ = [
    "AI_MATCH_THRESHOLD",
    "EXACT_REUSE_THRESHOLD",
    "MATCH_DECISION_SCHEMA_VERSION",
    "MATCHER_VERSION",
    "VARIANT_THRESHOLD",
    "decide_template_match",
    "fingerprint_similarity",
    "rank_templates",
]
