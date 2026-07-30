"""Fail-closed validation for deterministic extraction results."""

from __future__ import annotations

from typing import Any

from inference_data_ai_extraction_recipe import (
    EXTRACTION_RESULT_SCHEMA_VERSION,
    VALIDATION_REPORT_SCHEMA_VERSION,
    validate_extraction_recipe,
)


def _cell_map(
    capture_payload: dict[str, Any],
) -> dict[tuple[int, int, int], dict[str, Any]]:
    result: dict[tuple[int, int, int], dict[str, Any]] = {}
    for sheet in (capture_payload.get("workbook") or {}).get("sheets") or []:
        sheet_index = int(sheet.get("sheetIndex") or 0)
        for cell in sheet.get("cells") or []:
            result[
                (
                    sheet_index,
                    int(cell.get("row") or 0),
                    int(cell.get("column") or 0),
                )
            ] = cell
    return result


def _failure(
    failures: list[dict[str, Any]],
    code: str,
    **details: Any,
) -> None:
    failures.append({"code": code, **details})


def validate_extraction(
    capture_payload: dict[str, Any],
    recipe: dict[str, Any],
    execution: dict[str, Any],
) -> dict[str, Any]:
    """Validate one executor result against Capture v2 and recipe requirements."""

    validate_extraction_recipe(recipe)
    result = execution.get("result")
    diagnostics = execution.get("diagnostics")
    if not isinstance(result, dict) or not isinstance(diagnostics, dict):
        raise ValueError("Execution requires result and diagnostics.")
    if result.get("schemaVersion") != EXTRACTION_RESULT_SCHEMA_VERSION:
        raise ValueError("Unsupported extraction result schemaVersion.")

    failures: list[dict[str, Any]] = list(diagnostics.get("fieldFailures") or [])
    checks: dict[str, str] = {}

    expected_anchor_ids = {str(item["id"]) for item in recipe.get("anchors") or []}
    resolved_anchor_ids = set((diagnostics.get("resolvedAnchors") or {}).keys())
    if expected_anchor_ids == resolved_anchor_ids:
        checks["anchorUniqueness"] = "PASS"
    else:
        checks["anchorUniqueness"] = "FAIL"
        _failure(
            failures,
            "REQUIRED_ANCHOR_MISSING",
            missing=sorted(expected_anchor_ids - resolved_anchor_ids),
        )

    regions = diagnostics.get("regions") or {}
    invalid_regions = [
        region_id
        for region_id, region in regions.items()
        if int(region["startRow"]) > int(region["endRow"])
        or int(region["startColumn"]) > int(region["endColumn"])
    ]
    if invalid_regions:
        checks["regionBounds"] = "FAIL"
        _failure(
            failures,
            "TABLE_SHAPE_OUT_OF_ENVELOPE",
            regions=invalid_regions,
        )
    else:
        checks["regionBounds"] = "PASS"

    parameters = list(result.get("parameters") or [])
    present_names = {
        str(parameter.get("name"))
        for parameter in parameters
        if parameter.get("name")
    }
    required_names = {
        str(field["parameter"])
        for field in recipe["fields"]
        if field.get("required", False)
    }
    missing_names = sorted(required_names - present_names)
    if missing_names:
        checks["requiredCoverage"] = "FAIL"
        _failure(
            failures,
            "REQUIRED_FIELD_MISSING",
            parameters=missing_names,
        )
    else:
        checks["requiredCoverage"] = "PASS"

    cells = _cell_map(capture_payload)
    evidence_ok = True
    formula_ok = True
    type_ok = True
    unit_ok = True
    for parameter in parameters:
        value_type = str(parameter.get("valueType") or "")
        value = parameter.get("value")
        if value_type == "integer" and (
            isinstance(value, bool) or not isinstance(value, int)
        ):
            type_ok = False
            _failure(
                failures,
                "VALUE_TYPE_MISMATCH",
                parameter=parameter.get("name"),
                label=parameter.get("label"),
            )
        if value_type in {"decimal", "percent", "formula-result"} and (
            isinstance(value, bool) or not isinstance(value, (int, float))
        ):
            type_ok = False
            _failure(
                failures,
                "VALUE_TYPE_MISMATCH",
                parameter=parameter.get("name"),
                label=parameter.get("label"),
            )
        evidence = parameter.get("evidence")
        if not isinstance(evidence, list) or not evidence:
            evidence_ok = False
            _failure(
                failures,
                "EVIDENCE_CELL_NOT_FOUND",
                parameter=parameter.get("name"),
                label=parameter.get("label"),
            )
            continue
        evidence_by_role = {
            str(item.get("role")): item
            for item in evidence
            if isinstance(item, dict)
        }
        if not {"label", "value"} <= evidence_by_role.keys():
            evidence_ok = False
            _failure(
                failures,
                "EVIDENCE_CELL_NOT_FOUND",
                parameter=parameter.get("name"),
                label=parameter.get("label"),
                roles=sorted(evidence_by_role),
            )
        for item in evidence_by_role.values():
            key = (
                int(item.get("sheetIndex") or 0),
                int(item.get("row") or 0),
                int(item.get("column") or 0),
            )
            source_cell = cells.get(key)
            if source_cell is None:
                evidence_ok = False
                _failure(
                    failures,
                    "EVIDENCE_CELL_NOT_FOUND",
                    parameter=parameter.get("name"),
                    cell=item.get("cell"),
                )
                continue
            for source_key in ("rawValue", "displayValue", "formula"):
                if source_cell.get(source_key) != item.get(source_key):
                    evidence_ok = False
                    _failure(
                        failures,
                        "EVIDENCE_VALUE_MISMATCH",
                        parameter=parameter.get("name"),
                        cell=item.get("cell"),
                        field=source_key,
                    )
        value_evidence = evidence_by_role.get("value")
        if value_type == "formula-result" and (
            value_evidence is None or not value_evidence.get("formula")
        ):
            formula_ok = False
            _failure(
                failures,
                "FORMULA_PATTERN_CHANGED",
                parameter=parameter.get("name"),
                label=parameter.get("label"),
            )
        if parameter.get("unit") is not None and "unit" not in evidence_by_role:
            unit_ok = False
            _failure(
                failures,
                "UNIT_CONFLICT",
                parameter=parameter.get("name"),
                label=parameter.get("label"),
            )

    checks["evidenceExists"] = "PASS" if evidence_ok else "FAIL"
    checks["valueType"] = "PASS" if type_ok else "FAIL"
    checks["formulaConsistency"] = "PASS" if formula_ok else "FAIL"
    checks["unitConsistency"] = "PASS" if unit_ok else "FAIL"

    consumed = {
        (
            int(item.get("sheetIndex") or 0),
            int(item.get("row") or 0),
            int(item.get("column") or 0),
        )
        for item in diagnostics.get("consumedCoordinates") or []
    }
    unexpected: list[dict[str, int]] = []
    for region in regions.values():
        sheet_index = int(region["sheetIndex"])
        first_row = int(region["firstDataRow"])
        last_row = int(region["lastDataRow"])
        if last_row < first_row:
            continue
        for (candidate_sheet, row, column), cell in cells.items():
            if candidate_sheet != sheet_index or not first_row <= row <= last_row:
                continue
            value = (
                cell.get("displayValue")
                if cell.get("displayValue") is not None
                else cell.get("rawValue")
            )
            if (
                isinstance(value, (int, float))
                and not isinstance(value, bool)
                and (candidate_sheet, row, column) not in consumed
            ):
                unexpected.append(
                    {
                        "sheetIndex": candidate_sheet,
                        "row": row,
                        "column": column,
                    }
                )
    if unexpected:
        checks["unusedQuantitativeCellReview"] = "FAIL"
        _failure(
            failures,
            "UNEXPECTED_QUANTITATIVE_REGION",
            cells=unexpected,
        )
    else:
        checks["unusedQuantitativeCellReview"] = "PASS"

    unique_failures: list[dict[str, Any]] = []
    seen: set[str] = set()
    for failure in failures:
        key = repr(sorted(failure.items(), key=lambda item: item[0]))
        if key in seen:
            continue
        seen.add(key)
        unique_failures.append(failure)
    return {
        "schemaVersion": VALIDATION_REPORT_SCHEMA_VERSION,
        "sourceSha256": str(result.get("sourceSha256") or ""),
        "recipe": str(result.get("recipe") or ""),
        "status": "VERIFIED" if not unique_failures else "FAILED",
        "checks": checks,
        "failureCodes": [item["code"] for item in unique_failures],
        "failures": unique_failures,
    }


__all__ = ["validate_extraction"]
