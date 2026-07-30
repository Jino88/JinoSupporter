"""Execute a validated extraction recipe against one Capture v2 payload."""

from __future__ import annotations

import datetime as dt
import re
from typing import Any

from inference_data_ai_extraction_recipe import (
    EXTRACTION_RESULT_SCHEMA_VERSION,
    RECIPE_ENGINE_VERSION,
    validate_extraction_recipe,
)
from inference_data_ai_structure_fingerprint import normalize_text


class RecipeExecutionError(RuntimeError):
    """Raised when workbook structure cannot safely satisfy a recipe."""


def _cell_value(cell: dict[str, Any]) -> Any:
    value = cell.get("displayValue")
    if value is None:
        value = cell.get("rawValue")
    return value


def _nonblank(value: Any) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def _sheet_bounds(sheet: dict[str, Any]) -> dict[str, int]:
    value = sheet.get("contentBounds") or sheet.get("usedBounds") or {}
    min_row = int(value.get("minRow") or 1)
    min_column = int(value.get("minColumn") or 1)
    max_row = int(value.get("maxRow") or min_row)
    max_column = int(value.get("maxColumn") or min_column)
    return {
        "minRow": min_row,
        "minColumn": min_column,
        "maxRow": max_row,
        "maxColumn": max_column,
    }


def _sheet_cell_indexes(
    sheet: dict[str, Any],
) -> tuple[dict[tuple[int, int], dict[str, Any]], list[dict[str, Any]]]:
    cells = list(sheet.get("cells") or [])
    return (
        {
            (int(cell.get("row") or 0), int(cell.get("column") or 0)): cell
            for cell in cells
        },
        cells,
    )


def _sheet_contains_anchors(sheet: dict[str, Any], anchors: list[str]) -> bool:
    available = {
        normalize_text(_cell_value(cell))
        for cell in sheet.get("cells") or []
        if isinstance(_cell_value(cell), str)
    }
    return all(normalize_text(anchor) in available for anchor in anchors)


def _select_sheets(
    workbook: dict[str, Any],
    selectors: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    sheets = list(workbook.get("sheets") or [])
    selected: dict[str, dict[str, Any]] = {}
    for selector in selectors:
        aliases = {
            normalize_text(value) for value in selector.get("titleAliases") or []
        }
        required_anchors = list(selector.get("requiredAnchors") or [])
        fallback_role = str(selector.get("fallbackRole") or "")
        candidates: list[tuple[int, int, dict[str, Any]]] = []
        for index, sheet in enumerate(sheets):
            if required_anchors and not _sheet_contains_anchors(sheet, required_anchors):
                continue
            title_match = normalize_text(sheet.get("title")) in aliases if aliases else False
            fallback_match = (
                fallback_role == "tabular-result"
                and bool(sheet.get("hasTabularEvidence"))
            )
            if aliases and not title_match and not fallback_match:
                continue
            score = (2 if title_match else 0) + len(required_anchors)
            candidates.append((score, -index, sheet))
        if not candidates:
            raise RecipeExecutionError(
                f"Sheet selector {selector['id']} found no safe candidate."
            )
        candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
        best_score = candidates[0][0]
        tied = [item for item in candidates if item[0] == best_score]
        if len(tied) != 1:
            raise RecipeExecutionError(
                f"Sheet selector {selector['id']} is ambiguous."
            )
        selected[selector["id"]] = candidates[0][2]
    return selected


def _resolve_anchors(
    selected_sheets: dict[str, dict[str, Any]],
    anchors: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    resolved: dict[str, dict[str, Any]] = {}
    for anchor in anchors:
        pattern = re.compile(str(anchor["textRegex"]), re.IGNORECASE)
        sheet = selected_sheets[str(anchor["sheet"])]
        matches: list[dict[str, Any]] = []
        for cell in sheet.get("cells") or []:
            value = _cell_value(cell)
            if not isinstance(value, str):
                continue
            candidate = normalize_text(value) if anchor.get("normalized", True) else value
            if pattern.fullmatch(candidate):
                matches.append(cell)
        if len(matches) != 1:
            raise RecipeExecutionError(
                f"Anchor {anchor['id']} expected one cell, found {len(matches)}."
            )
        resolved[str(anchor["id"])] = matches[0]
    return resolved


def _resolve_regions(
    selected_sheets: dict[str, dict[str, Any]],
    anchors: dict[str, dict[str, Any]],
    regions: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    resolved: dict[str, dict[str, Any]] = {}
    for region in regions:
        sheet = selected_sheets[str(region["sheet"])]
        bounds = _sheet_bounds(sheet)
        start = region["start"]
        if "below" in start:
            anchor = anchors[str(start["below"])]
            start_row = int(anchor["row"]) + int(start.get("rows") or 0)
            start_column = int(start.get("column") or bounds["minColumn"])
        else:
            start_row = int(start["row"])
            start_column = int(start.get("column") or bounds["minColumn"])
        end_row = min(int(region.get("endRow") or bounds["maxRow"]), bounds["maxRow"])
        end_column = min(
            int(region.get("endColumn") or bounds["maxColumn"]),
            bounds["maxColumn"],
        )
        if (
            start_row < bounds["minRow"]
            or start_row > end_row
            or start_column < bounds["minColumn"]
            or start_column > end_column
        ):
            raise RecipeExecutionError(f"Region {region['id']} is outside sheet bounds.")
        resolved[str(region["id"])] = {
            "sheetId": str(region["sheet"]),
            "sheet": sheet,
            "startRow": start_row,
            "startColumn": start_column,
            "endRow": end_row,
            "endColumn": end_column,
            "headerDepth": int(region["headerDepth"]),
            "stop": dict(region.get("stop") or {}),
        }
    return resolved


def _resolve_columns(
    axes: dict[str, Any],
    regions: dict[str, dict[str, Any]],
) -> dict[str, dict[str, int]]:
    resolved: dict[str, dict[str, int]] = {}
    for region_id, region in regions.items():
        cell_index, _ = _sheet_cell_indexes(region["sheet"])
        header_end = min(
            region["startRow"] + region["headerDepth"] - 1,
            region["endRow"],
        )
        role_columns: dict[str, int] = {}
        for column_rule in axes["columns"]:
            role = str(column_rule["role"])
            aliases = {
                normalize_text(alias)
                for alias in column_rule.get("headerAliases") or []
            }
            candidates: set[int] = set()
            for row in range(region["startRow"], header_end + 1):
                for column in range(
                    region["startColumn"],
                    region["endColumn"] + 1,
                ):
                    cell = cell_index.get((row, column))
                    if cell is None:
                        continue
                    if normalize_text(_cell_value(cell)) in aliases:
                        candidates.add(column)
            if len(candidates) != 1:
                raise RecipeExecutionError(
                    f"Region {region_id} role {role} expected one header column, "
                    f"found {len(candidates)}."
                )
            role_columns[role] = next(iter(candidates))
        resolved[region_id] = role_columns
    return resolved


def _convert_value(value: Any, value_type: str) -> Any:
    if value_type == "text":
        if value is None:
            return None
        return str(value)
    if value_type == "integer":
        if isinstance(value, bool):
            raise ValueError("boolean is not an integer parameter")
        if isinstance(value, int):
            return value
        if isinstance(value, float) and value.is_integer():
            return int(value)
        if isinstance(value, str):
            return int(value.replace(",", "").strip())
        raise ValueError("value is not an integer")
    if value_type in {"decimal", "percent", "formula-result"}:
        if isinstance(value, bool):
            raise ValueError("boolean is not a numeric parameter")
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            text = value.replace(",", "").strip()
            if value_type == "percent" and text.endswith("%"):
                raise ValueError("percent text requires an explicit conversion rule")
            return float(text)
        raise ValueError("value is not numeric")
    if value_type == "date":
        if isinstance(value, (dt.date, dt.datetime)):
            return value.isoformat()
        if isinstance(value, str) and value.strip():
            return value.strip()
        raise ValueError("value is not a date")
    raise ValueError(f"unsupported value type: {value_type}")


def _evidence(
    sheet: dict[str, Any],
    cell: dict[str, Any],
    role: str,
) -> dict[str, Any]:
    return {
        "role": role,
        "sheet": str(sheet.get("title") or ""),
        "sheetIndex": int(sheet.get("sheetIndex") or 0),
        "cell": str(cell.get("coordinate") or ""),
        "row": int(cell.get("row") or 0),
        "column": int(cell.get("column") or 0),
        "rawValue": cell.get("rawValue"),
        "displayValue": cell.get("displayValue"),
        "formula": cell.get("formula"),
        "numberFormat": str(cell.get("numberFormat") or "General"),
        "mergeRange": cell.get("mergeRange"),
    }


def execute_recipe(
    capture_payload: dict[str, Any],
    recipe: dict[str, Any],
    *,
    match_decision: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Execute a recipe without invoking AI and return result plus diagnostics."""

    validate_extraction_recipe(recipe)
    recipe_ref = f"{recipe['recipeId']}@{recipe['recipeVersion']}"
    if isinstance(match_decision, dict):
        decision = str(match_decision.get("decision") or "")
        if decision not in {"EXACT_REUSE", "AI_CONFIRMED_REUSE"}:
            raise RecipeExecutionError(
                f"Match decision {decision or '<missing>'} does not authorize recipe execution."
            )
        if match_decision.get("selectedRecipe") != recipe_ref:
            raise RecipeExecutionError(
                "Match decision selectedRecipe does not match the executable recipe."
            )
    workbook = capture_payload.get("workbook")
    if not isinstance(workbook, dict):
        raise RecipeExecutionError("Capture payload requires workbook.")
    selected_sheets = _select_sheets(workbook, recipe["sheetSelectors"])
    anchors = _resolve_anchors(selected_sheets, recipe.get("anchors") or [])
    regions = _resolve_regions(selected_sheets, anchors, recipe["regions"])
    columns = _resolve_columns(recipe["axes"], regions)

    parameters: list[dict[str, Any]] = []
    consumed: set[tuple[int, int, int]] = set()
    field_failures: list[dict[str, Any]] = []
    region_rows: dict[str, dict[str, int]] = {}
    for field in recipe["fields"]:
        source = field["source"]
        region_id = str(source["region"])
        region = regions[region_id]
        sheet = region["sheet"]
        sheet_index = int(sheet.get("sheetIndex") or 0)
        cell_index, _ = _sheet_cell_indexes(sheet)
        role_columns = columns[region_id]
        label_column = role_columns[str(source["labelColumnRole"])]
        value_column = role_columns[str(source["valueColumnRole"])]
        unit_column = (
            role_columns[str(source["unitColumnRole"])]
            if source.get("unitColumnRole")
            else None
        )
        first_data_row = region["startRow"] + region["headerDepth"]
        blank_limit = int(region["stop"].get("firstBlankKeyColumnRun") or 1)
        blank_run = 0
        last_data_row = first_data_row - 1
        for row in range(first_data_row, region["endRow"] + 1):
            label_cell = cell_index.get((row, label_column))
            label_value = _cell_value(label_cell or {})
            if not _nonblank(label_value):
                blank_run += 1
                if blank_run >= blank_limit:
                    break
                continue
            blank_run = 0
            last_data_row = row
            value_cell = cell_index.get((row, value_column))
            value = _cell_value(value_cell or {})
            if not _nonblank(value):
                if field.get("required", False):
                    field_failures.append(
                        {
                            "code": "REQUIRED_FIELD_MISSING",
                            "parameter": field["parameter"],
                            "row": row,
                        }
                    )
                continue
            try:
                converted = _convert_value(value, str(field["valueType"]))
                if (
                    field["valueType"] == "formula-result"
                    and not (value_cell or {}).get("formula")
                ):
                    raise ValueError("formula-result source is not a formula cell")
            except (TypeError, ValueError) as error:
                field_failures.append(
                    {
                        "code": "VALUE_TYPE_MISMATCH",
                        "parameter": field["parameter"],
                        "row": row,
                        "message": str(error),
                    }
                )
                continue

            evidence = [
                _evidence(sheet, label_cell, "label"),
                _evidence(sheet, value_cell, "value"),
            ]
            consumed.add((sheet_index, row, label_column))
            consumed.add((sheet_index, row, value_column))
            unit: Any = None
            if unit_column is not None:
                unit_cell = cell_index.get((row, unit_column))
                unit = _cell_value(unit_cell or {})
                if unit_cell is not None:
                    evidence.append(_evidence(sheet, unit_cell, "unit"))
                    consumed.add((sheet_index, row, unit_column))
            parameters.append(
                {
                    "name": str(field["parameter"]),
                    "label": str(label_value),
                    "value": converted,
                    "valueType": str(field["valueType"]),
                    "unit": unit,
                    "evidence": evidence,
                }
            )
        region_rows[region_id] = {
            "firstDataRow": first_data_row,
            "lastDataRow": last_data_row,
        }

    source = capture_payload.get("source") or {}
    selected_template = (
        match_decision.get("selectedTemplate")
        if isinstance(match_decision, dict)
        else f"{recipe['templateId']}@unknown"
    )
    result = {
        "schemaVersion": EXTRACTION_RESULT_SCHEMA_VERSION,
        "sourceSha256": str(source.get("contentSha256") or ""),
        "sourcePath": str(source.get("sourcePath") or ""),
        "template": selected_template,
        "recipe": recipe_ref,
        "engineVersion": RECIPE_ENGINE_VERSION,
        "parameters": parameters,
    }
    diagnostics = {
        "selectedSheets": {
            selector_id: {
                "sheetIndex": int(sheet.get("sheetIndex") or 0),
                "title": str(sheet.get("title") or ""),
            }
            for selector_id, sheet in selected_sheets.items()
        },
        "resolvedAnchors": {
            anchor_id: {
                "sheet": str(
                    selected_sheets[
                        next(
                            item["sheet"]
                            for item in recipe["anchors"]
                            if item["id"] == anchor_id
                        )
                    ].get("title")
                    or ""
                ),
                "cell": str(cell.get("coordinate") or ""),
                "row": int(cell.get("row") or 0),
                "column": int(cell.get("column") or 0),
            }
            for anchor_id, cell in anchors.items()
        },
        "resolvedColumns": columns,
        "regions": {
            region_id: {
                "sheetIndex": int(region["sheet"].get("sheetIndex") or 0),
                "startRow": region["startRow"],
                "startColumn": region["startColumn"],
                "endRow": region["endRow"],
                "endColumn": region["endColumn"],
                **region_rows[region_id],
            }
            for region_id, region in regions.items()
        },
        "consumedCoordinates": [
            {
                "sheetIndex": sheet_index,
                "row": row,
                "column": column,
            }
            for sheet_index, row, column in sorted(consumed)
        ],
        "fieldFailures": field_failures,
    }
    return {"result": result, "diagnostics": diagnostics}


__all__ = [
    "RecipeExecutionError",
    "execute_recipe",
]
