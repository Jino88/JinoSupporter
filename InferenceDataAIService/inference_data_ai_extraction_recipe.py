"""Versioned contracts for executable Excel extraction recipes."""

from __future__ import annotations

import copy
import re
from typing import Any

from inference_data_ai_structure_fingerprint import (
    FINGERPRINT_SCHEMA_VERSION,
    validate_structure_fingerprint,
)


FORM_TEMPLATE_SCHEMA_VERSION = "excel-form-template-v1"
EXTRACTION_RECIPE_SCHEMA_VERSION = "excel-extraction-recipe-v1"
EXTRACTION_RESULT_SCHEMA_VERSION = "excel-deterministic-extraction-v1"
VALIDATION_REPORT_SCHEMA_VERSION = "excel-extraction-validation-v1"
RECIPE_ENGINE_VERSION = "deterministic-recipe-engine-v1.0"


class RecipeContractError(ValueError):
    """Raised when a template or recipe violates its executable contract."""


def _required_string(value: dict[str, Any], key: str, owner: str) -> str:
    result = value.get(key)
    if not isinstance(result, str) or not result.strip():
        raise RecipeContractError(f"{owner}.{key} must be a non-empty string.")
    return result


def create_form_template(
    *,
    template_id: str,
    family_id: str,
    fingerprint: dict[str, Any],
    recipe_ref: str,
    required_anchors: list[str] | None = None,
    status: str = "DRAFT",
) -> dict[str, Any]:
    """Create one template around a representative verified fingerprint."""

    validate_structure_fingerprint(fingerprint)
    workbook = fingerprint["workbook"]
    template = {
        "schemaVersion": FORM_TEMPLATE_SCHEMA_VERSION,
        "templateId": template_id,
        "familyId": family_id,
        "templateVersion": 1,
        "status": status,
        "representativeFingerprint": copy.deepcopy(fingerprint),
        "acceptedFingerprintEnvelope": {
            "requiredAnchors": list(required_anchors or []),
            "allowedSheetCount": [
                int(workbook["sheetCount"]),
                int(workbook["sheetCount"]),
            ],
            "allowedTabularSheetCount": [
                int(workbook["tabularSheetCount"]),
                int(workbook["tabularSheetCount"]),
            ],
        },
        "exampleSources": [
            {
                "sourceSha256": str(fingerprint.get("sourceSha256") or ""),
                "quality": "UNVERIFIED",
            }
        ],
        "recipeRefs": [recipe_ref],
        "quality": {
            "replayFileCount": 0,
            "replayPassCount": 0,
            "falsePositiveCount": 0,
        },
    }
    validate_form_template(template)
    return template


def validate_form_template(template: dict[str, Any]) -> None:
    if template.get("schemaVersion") != FORM_TEMPLATE_SCHEMA_VERSION:
        raise RecipeContractError("Unsupported form template schemaVersion.")
    _required_string(template, "templateId", "template")
    _required_string(template, "familyId", "template")
    if not isinstance(template.get("templateVersion"), int):
        raise RecipeContractError("template.templateVersion must be an integer.")
    if template.get("status") not in {"DRAFT", "APPROVED", "RETIRED"}:
        raise RecipeContractError("template.status is invalid.")
    fingerprint = template.get("representativeFingerprint")
    if not isinstance(fingerprint, dict):
        raise RecipeContractError("template requires representativeFingerprint.")
    if fingerprint.get("schemaVersion") != FINGERPRINT_SCHEMA_VERSION:
        raise RecipeContractError("template fingerprint schemaVersion is invalid.")
    validate_structure_fingerprint(fingerprint)
    envelope = template.get("acceptedFingerprintEnvelope")
    if not isinstance(envelope, dict):
        raise RecipeContractError("template requires acceptedFingerprintEnvelope.")
    for key in ("allowedSheetCount", "allowedTabularSheetCount"):
        bounds = envelope.get(key)
        if (
            not isinstance(bounds, list)
            or len(bounds) != 2
            or not all(isinstance(item, int) for item in bounds)
            or bounds[0] > bounds[1]
        ):
            raise RecipeContractError(f"template envelope {key} is invalid.")
    anchors = envelope.get("requiredAnchors")
    if not isinstance(anchors, list) or not all(
        isinstance(anchor, str) and anchor.strip() for anchor in anchors
    ):
        raise RecipeContractError("template requiredAnchors must be strings.")
    refs = template.get("recipeRefs")
    if not isinstance(refs, list) or not refs:
        raise RecipeContractError("template requires at least one recipeRef.")


def validate_extraction_recipe(recipe: dict[str, Any]) -> None:
    if recipe.get("schemaVersion") != EXTRACTION_RECIPE_SCHEMA_VERSION:
        raise RecipeContractError("Unsupported extraction recipe schemaVersion.")
    _required_string(recipe, "recipeId", "recipe")
    _required_string(recipe, "templateId", "recipe")
    if not isinstance(recipe.get("recipeVersion"), int):
        raise RecipeContractError("recipe.recipeVersion must be an integer.")

    sheet_ids: set[str] = set()
    selectors = recipe.get("sheetSelectors")
    if not isinstance(selectors, list) or not selectors:
        raise RecipeContractError("recipe requires sheetSelectors.")
    for selector in selectors:
        if not isinstance(selector, dict):
            raise RecipeContractError("sheet selector must be an object.")
        selector_id = _required_string(selector, "id", "sheetSelector")
        if selector_id in sheet_ids:
            raise RecipeContractError(f"duplicate sheet selector id: {selector_id}")
        sheet_ids.add(selector_id)
        if selector.get("cardinality", "exactly-one") != "exactly-one":
            raise RecipeContractError("only exactly-one sheet selectors are supported.")
        required_anchors = selector.get("requiredAnchors") or []
        if not isinstance(required_anchors, list) or not all(
            isinstance(anchor, str) and anchor.strip()
            for anchor in required_anchors
        ):
            raise RecipeContractError(
                "sheet selector requiredAnchors must be strings."
            )

    anchor_ids: set[str] = set()
    anchors = recipe.get("anchors")
    if not isinstance(anchors, list):
        raise RecipeContractError("recipe.anchors must be a list.")
    for anchor in anchors:
        if not isinstance(anchor, dict):
            raise RecipeContractError("anchor must be an object.")
        anchor_id = _required_string(anchor, "id", "anchor")
        if anchor_id in anchor_ids:
            raise RecipeContractError(f"duplicate anchor id: {anchor_id}")
        anchor_ids.add(anchor_id)
        if anchor.get("sheet") not in sheet_ids:
            raise RecipeContractError(f"anchor {anchor_id} references an unknown sheet.")
        pattern = _required_string(anchor, "textRegex", f"anchor {anchor_id}")
        try:
            re.compile(pattern)
        except re.error as error:
            raise RecipeContractError(
                f"anchor {anchor_id} textRegex is invalid: {error}"
            ) from error
        if anchor.get("uniqueness", "one") != "one":
            raise RecipeContractError("only unique anchors are supported.")

    region_ids: set[str] = set()
    regions = recipe.get("regions")
    if not isinstance(regions, list) or not regions:
        raise RecipeContractError("recipe requires regions.")
    for region in regions:
        if not isinstance(region, dict):
            raise RecipeContractError("region must be an object.")
        region_id = _required_string(region, "id", "region")
        if region_id in region_ids:
            raise RecipeContractError(f"duplicate region id: {region_id}")
        region_ids.add(region_id)
        if region.get("sheet") not in sheet_ids:
            raise RecipeContractError(f"region {region_id} references an unknown sheet.")
        start = region.get("start")
        if not isinstance(start, dict):
            raise RecipeContractError(f"region {region_id} requires start.")
        if "below" in start and start["below"] not in anchor_ids:
            raise RecipeContractError(f"region {region_id} references an unknown anchor.")
        if "below" not in start and "row" not in start:
            raise RecipeContractError(
                f"region {region_id} start requires below or row."
            )
        if int(region.get("headerDepth") or 0) < 1:
            raise RecipeContractError(f"region {region_id} requires headerDepth >= 1.")
        if region.get("repeatMode", "rows").lower() != "rows":
            raise RecipeContractError("only row-repeated regions are supported.")

    axes = recipe.get("axes")
    if not isinstance(axes, dict):
        raise RecipeContractError("recipe requires axes.")
    columns = axes.get("columns")
    if not isinstance(columns, list) or not columns:
        raise RecipeContractError("recipe axes require columns.")
    column_roles: set[str] = set()
    for column in columns:
        if not isinstance(column, dict):
            raise RecipeContractError("axis column must be an object.")
        role = _required_string(column, "role", "axis column")
        if role in column_roles:
            raise RecipeContractError(f"duplicate axis column role: {role}")
        column_roles.add(role)
        aliases = column.get("headerAliases")
        if not isinstance(aliases, list) or not aliases:
            raise RecipeContractError(f"axis role {role} requires headerAliases.")
        if not all(isinstance(alias, str) and alias.strip() for alias in aliases):
            raise RecipeContractError(
                f"axis role {role} headerAliases must be strings."
            )
    row_key = axes.get("rowKey")
    if not isinstance(row_key, dict):
        raise RecipeContractError("recipe axes require rowKey.")
    if row_key.get("region") not in region_ids:
        raise RecipeContractError("rowKey references an unknown region.")
    if row_key.get("columnRole") not in column_roles:
        raise RecipeContractError("rowKey references an unknown column role.")

    fields = recipe.get("fields")
    if not isinstance(fields, list) or not fields:
        raise RecipeContractError("recipe requires fields.")
    for field in fields:
        if not isinstance(field, dict):
            raise RecipeContractError("field must be an object.")
        _required_string(field, "parameter", "field")
        source = field.get("source")
        if not isinstance(source, dict):
            raise RecipeContractError("field requires source.")
        if source.get("region") not in region_ids:
            raise RecipeContractError("field references an unknown region.")
        if source.get("row", "each") != "each":
            raise RecipeContractError("only row=each fields are supported.")
        for key in ("valueColumnRole", "labelColumnRole"):
            if source.get(key) not in column_roles:
                raise RecipeContractError(f"field {key} references an unknown role.")
        unit_role = source.get("unitColumnRole")
        if unit_role is not None and unit_role not in column_roles:
            raise RecipeContractError("field unitColumnRole references an unknown role.")
        if field.get("valueType") not in {
            "text",
            "integer",
            "decimal",
            "percent",
            "date",
            "formula-result",
        }:
            raise RecipeContractError("field valueType is invalid.")
        if field.get("evidence", "exact-source-cell") != "exact-source-cell":
            raise RecipeContractError("only exact-source-cell evidence is supported.")


__all__ = [
    "EXTRACTION_RECIPE_SCHEMA_VERSION",
    "EXTRACTION_RESULT_SCHEMA_VERSION",
    "FORM_TEMPLATE_SCHEMA_VERSION",
    "RECIPE_ENGINE_VERSION",
    "RecipeContractError",
    "VALIDATION_REPORT_SCHEMA_VERSION",
    "create_form_template",
    "validate_extraction_recipe",
    "validate_form_template",
]
