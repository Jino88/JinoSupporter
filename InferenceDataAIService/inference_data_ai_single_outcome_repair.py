"""Bounded deterministic repair for the single missing B22 L7 outcome.

This module intentionally has no AI or workflow dependency.  It accepts only
the exact B22 content-coverage failure, proves the source geometry from focused
Capture v2 chunks, and permits one exact Outcome insertion.
"""

from __future__ import annotations

import copy
import re
from collections.abc import Iterator, Sequence
from typing import Any


class SingleOutcomeRepairError(ValueError):
    """Raised when the exact single-Outcome repair cannot be proven safe."""


B22_REVISION_UID = "capture_revision_2c249104a0f8f746715107da"
B22_CONTENT_SHA256 = (
    "151551ccaed6505aab60de6e3e97033abcf6a88ad091147645147f3d6a8d0a2d"
)
B22_VALIDATION_ERROR = (
    "ContentCoverageError: Source content coverage is incomplete; "
    "1 quantitative cell(s): NG function!L7"
)
B22_OUTCOME_KEY = "hearing_plus1v_noise_ng_percentage"
B22_OBSERVATION_KEY = "plus1v_noise_0_percent"

_TARGET_SHEET = "NG function"
_TARGET_COORDINATE = "L7"
_STUDY_KEYS = [
    "function_lot_test",
    "nti_mask_profile",
    "spl_numeric_profile",
    "thd_numeric_profile",
    "imp_numeric_profile",
    "fo_result_check",
]
_UNREPAIRED_OUTCOME_KEYS = [
    "input",
    "ok_count",
    "sigma_spl_ng",
    "sigma_thd_ng",
    "sigma_spl_thd_ng",
    "sigma_spl_thd_f0_ng",
    "hearing_plus_1v_noise_ng",
    "hearing_plus_1v_touch_ng",
    "hearing_plus_1v_total_ng",
    "hearing_plus_1v_ng_rate",
    "hearing_plus_0v_noise_ng",
    "hearing_plus_0v_touch_ng",
    "hearing_plus_0v_total_ng",
    "hearing_plus_0v_ng_rate",
    "sigma_spl_percentage",
    "sigma_thd_percentage",
    "sigma_spl_thd_percentage",
    "sigma_spl_thd_f0_percentage",
    "hearing_plus_1v_touch_percentage",
]
_INSERT_INDEX = 18
_A1_PATTERN = re.compile(
    r"\$?([A-Za-z]{1,4})\$?([1-9]\d*)"
    r"(?::\$?([A-Za-z]{1,4})\$?([1-9]\d*))?"
)


def single_outcome_repair_applicable(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> bool:
    """Return whether the exact B22 L7 repair is safe and deterministic."""

    try:
        _expected_repair(
            baseline,
            validation_error=validation_error,
            focused_chunks=focused_chunks,
        )
    except (SingleOutcomeRepairError, TypeError, ValueError):
        return False
    return True


def apply_deterministic_single_outcome_repair(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Insert only the proven B22 L7 Outcome, or return an idempotent copy."""

    repaired = _expected_repair(
        baseline,
        validation_error=validation_error,
        focused_chunks=focused_chunks,
    )
    validate_deterministic_single_outcome_repair(
        baseline,
        repaired,
        validation_error=validation_error,
        focused_chunks=focused_chunks,
    )
    return repaired


def validate_deterministic_single_outcome_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    """Require an exact one-Outcome projection and no unrelated mutation."""

    expected = _expected_repair(
        baseline,
        validation_error=validation_error,
        focused_chunks=focused_chunks,
    )
    if repaired != expected:
        raise SingleOutcomeRepairError(
            "Single-Outcome repair changed fields outside the exact B22 "
            "NG function!L7 Outcome insertion"
        )


def _expected_repair(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    if str(validation_error).strip() != B22_VALIDATION_ERROR:
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires the exact single B22 L7 "
            "content-coverage error"
        )
    if not isinstance(baseline, dict):
        raise SingleOutcomeRepairError(
            "Single-Outcome repair baseline must be an object"
        )

    cells = _validate_source_geometry(baseline, focused_chunks)
    desired = _desired_outcome(cells)
    repaired = copy.deepcopy(baseline)
    studies = repaired.get("studies")
    if not isinstance(studies, list) or not studies:
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires the exact B22 Study list"
        )
    study = studies[0]
    if not isinstance(study, dict):
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires the exact B22 target Study"
        )
    outcomes = study.get("outcomes")
    if not isinstance(outcomes, list):
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires a canonical Outcome list"
        )

    matching_indexes = [
        index
        for index, outcome in enumerate(outcomes)
        if isinstance(outcome, dict)
        and str(outcome.get("key") or "") == B22_OUTCOME_KEY
    ]
    if matching_indexes:
        if matching_indexes != [_INSERT_INDEX]:
            raise SingleOutcomeRepairError(
                "Single-Outcome repair found a conflicting or duplicate "
                "B22 L7 Outcome key"
            )
        if outcomes[_INSERT_INDEX] != desired:
            raise SingleOutcomeRepairError(
                "Existing B22 L7 Outcome does not match the exact "
                "source-derived projection"
            )
        unrepaired = copy.deepcopy(repaired)
        del unrepaired["studies"][0]["outcomes"][_INSERT_INDEX]
        _validate_unrepaired_manifest(unrepaired, cells)
        return repaired

    _validate_unrepaired_manifest(repaired, cells)
    repaired["studies"][0]["outcomes"].insert(_INSERT_INDEX, desired)
    return repaired


def _validate_source_geometry(
    baseline: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    source = baseline.get("source")
    if not isinstance(source, dict):
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires canonical source identity"
        )
    if (
        str(source.get("revisionUid") or "") != B22_REVISION_UID
        or str(source.get("contentSha256") or "").lower()
        != B22_CONTENT_SHA256
    ):
        raise SingleOutcomeRepairError(
            "Single-Outcome repair source identity is not the proven B22 "
            "revision"
        )

    required = {
        "F6",
        "H4",
        "H5",
        "H7",
        "I5",
        "I7",
        "J5",
        "J7",
        "K5",
        "K7",
        "L4",
        "L5",
        "L6",
        "L7",
        "M5",
        "M6",
        "M7",
    }
    cells: dict[str, dict[str, Any]] = {}
    target_chunks: list[dict[str, Any]] = []
    for chunk in focused_chunks:
        if not isinstance(chunk, dict):
            continue
        sheet = chunk.get("sheet")
        title = (
            str(sheet.get("title") or "")
            if isinstance(sheet, dict)
            else str(sheet or "")
        )
        if title != _TARGET_SHEET:
            continue
        target_chunks.append(chunk)
        revision = chunk.get("sourceRevision")
        if not isinstance(revision, dict) or (
            str(revision.get("revisionUid") or "") != B22_REVISION_UID
            or str(revision.get("contentSha256") or "").lower()
            != B22_CONTENT_SHA256
        ):
            raise SingleOutcomeRepairError(
                "Focused chunk source identity does not match the B22 "
                "manifest"
            )
        sheet_index = sheet.get("sheetIndex") if isinstance(sheet, dict) else None
        for cell in chunk.get("cells", []):
            if not isinstance(cell, dict):
                continue
            coordinate = str(cell.get("coordinate") or "").upper()
            if coordinate not in required:
                continue
            if coordinate in cells:
                raise SingleOutcomeRepairError(
                    f"Focused chunks duplicate primary source cell "
                    f"{_TARGET_SHEET}!{coordinate}"
                )
            if cell.get("primary") is not True:
                raise SingleOutcomeRepairError(
                    f"Focused source cell {_TARGET_SHEET}!{coordinate} "
                    "is not primary"
                )
            expected_key = (
                f"{B22_REVISION_UID}:{sheet_index}:{coordinate}"
            )
            if str(cell.get("sourceCellKey") or "") != expected_key:
                raise SingleOutcomeRepairError(
                    f"Focused source cell key is invalid for "
                    f"{_TARGET_SHEET}!{coordinate}"
                )
            cells[coordinate] = cell

    missing = sorted(required - cells.keys())
    if missing:
        raise SingleOutcomeRepairError(
            "Focused chunks lack required B22 source geometry: "
            + ", ".join(missing)
        )
    merged_addresses = {
        str(item.get("address") or "")
        for chunk in target_chunks
        for item in chunk.get("mergedRanges", [])
        if isinstance(item, dict)
    }
    if not {"H4:K4", "L4:O4"}.issubset(merged_addresses):
        raise SingleOutcomeRepairError(
            "Focused chunks lack the exact Sigma and Hearing merged headers"
        )

    _require_text_cell(cells["H4"], "Sigma", merge_range="H4:K4")
    _require_text_cell(cells["H5"], "SPL")
    _require_text_cell(cells["I5"], "THD")
    _require_text_cell(cells["J5"], "SPL+THD")
    _require_text_cell(cells["K5"], "SPL+THD+F0")
    _require_text_cell(
        cells["L4"],
        "Hearing  ( + 1V )",
        merge_range="L4:O4",
    )
    _require_text_cell(cells["L5"], "Noise")
    _require_text_cell(cells["M5"], "Touch")
    _require_number_cell(cells["F6"], 57, merge_range="F6:F7")
    _require_number_cell(cells["L6"], 0)
    _require_number_cell(cells["M6"], 9)
    for coordinate in ("H7", "I7", "J7", "K7", "L7"):
        _require_number_cell(
            cells[coordinate],
            0,
            number_format="0.0%",
        )
    _require_number_cell(cells["M7"], 1, number_format="0.0%")
    return cells


def _validate_unrepaired_manifest(
    manifest: dict[str, Any],
    cells: dict[str, dict[str, Any]],
) -> None:
    if manifest.get("schemaVersion") != "canonical-study-manifest-v1":
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires the canonical Study schema"
        )
    studies = manifest.get("studies")
    if not isinstance(studies, list) or [
        str(study.get("key") or "") if isinstance(study, dict) else ""
        for study in studies
    ] != _STUDY_KEYS:
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires the exact B22 Study keys"
        )
    study = studies[0]
    arms = study.get("arms")
    if not isinstance(arms, list) or len(arms) != 1:
        raise SingleOutcomeRepairError(
            "B22 target Study must retain its single condition_test Arm"
        )
    arm = arms[0]
    if not isinstance(arm, dict) or (
        str(arm.get("key") or "") != "condition_test"
        or str(arm.get("role") or "") != "TEST"
        or str(arm.get("label") or "").strip() != "Condition test"
        or str(arm.get("condition") or "").strip() != "Condition test"
        or _number(arm.get("sampleSize")) != 57
    ):
        raise SingleOutcomeRepairError(
            "B22 condition_test Arm identity or sample value changed"
        )

    outcomes = study.get("outcomes")
    if not isinstance(outcomes, list) or [
        str(outcome.get("key") or "") if isinstance(outcome, dict) else ""
        for outcome in outcomes
    ] != _UNREPAIRED_OUTCOME_KEYS:
        raise SingleOutcomeRepairError(
            "B22 target Study Outcome keys or source order changed"
        )
    if len(set(_UNREPAIRED_OUTCOME_KEYS)) != len(_UNREPAIRED_OUTCOME_KEYS):
        raise SingleOutcomeRepairError("Internal B22 Outcome keys are invalid")

    count_outcome = outcomes[6]
    count_observation = _only_observation(
        count_outcome,
        "hearing_plus_1v_noise_ng_condition_test",
    )
    if (
        str(count_outcome.get("metricType") or "") != "defect_count"
        or str(count_outcome.get("unit") or "") != ""
        or str(count_observation.get("arm") or "") != "condition_test"
        or _number(count_observation.get("valueNumber")) != 0
        or str(count_observation.get("valueText") or "") != "0"
        or _number(count_observation.get("numerator")) != 0
        or _number(count_observation.get("denominator")) != 57
        or _number(count_observation.get("sampleSize")) != 57
        or not _has_evidence(count_outcome, "L4:L6")
        or not _has_evidence(count_observation, "L6")
        or not _has_evidence(count_observation, "F6")
    ):
        raise SingleOutcomeRepairError(
            "B22 L6 count Outcome no longer proves 0 of 57 for "
            "condition_test"
        )

    neighbor_specs = [
        (14, "sigma_spl_percentage", "H7", 0, "0.0%"),
        (15, "sigma_thd_percentage", "I7", 0, "0.0%"),
        (16, "sigma_spl_thd_percentage", "J7", 0, "0.0%"),
        (17, "sigma_spl_thd_f0_percentage", "K7", 0, "0.0%"),
        (
            18,
            "hearing_plus_1v_touch_percentage",
            "M7",
            100,
            "100.0%",
        ),
    ]
    for index, key, coordinate, value_number, value_text in neighbor_specs:
        outcome = outcomes[index]
        observation = _only_observation(outcome)
        if (
            str(outcome.get("key") or "") != key
            or str(outcome.get("unit") or "") != "%"
            or str(observation.get("arm") or "") != "condition_test"
            or _number(observation.get("valueNumber")) != value_number
            or str(observation.get("valueText") or "") != value_text
            or _number(observation.get("sampleSize")) != 57
            or not _has_evidence(observation, coordinate)
            or not _has_evidence(observation, "F6")
        ):
            raise SingleOutcomeRepairError(
                f"B22 neighboring percentage Outcome {key} changed"
            )

    observation_keys = [
        str(observation.get("key") or "")
        for outcome in outcomes
        if isinstance(outcome, dict)
        for observation in outcome.get("observations", [])
        if isinstance(observation, dict)
    ]
    if (
        B22_OBSERVATION_KEY in observation_keys
        or len(observation_keys) != len(set(observation_keys))
    ):
        raise SingleOutcomeRepairError(
            "B22 observation keys conflict with the exact L7 projection"
        )
    if any(
        _evidence_item_covers_l7(item)
        for item in _iter_evidence_items(outcomes)
    ):
        raise SingleOutcomeRepairError(
            "B22 L7 is already represented by another Outcome"
        )

    if (
        _number(cells["L6"].get("rawValue"))
        != _number(count_observation.get("numerator"))
        or _number(cells["F6"].get("rawValue"))
        != _number(count_observation.get("denominator"))
    ):
        raise SingleOutcomeRepairError(
            "B22 manifest count values do not match focused source cells"
        )


def _desired_outcome(
    cells: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    numerator = int(_number(cells["L6"].get("rawValue")))
    denominator = int(_number(cells["F6"].get("rawValue")))
    raw_percent = _number(cells["L7"].get("rawValue"))
    value_number = raw_percent * 100
    parent = str(cells["L4"].get("rawValue") or "").rstrip()
    leaf = str(cells["L5"].get("rawValue") or "").strip()
    value_text = f"{value_number:.1f}%"
    return {
        "key": B22_OUTCOME_KEY,
        "originalLabel": f"{parent} {leaf} percentage",
        "metricType": "defect_rate_percent",
        "unit": "%",
        "favorableDirection": "UNKNOWN",
        "evidence": [
            {
                "sheet": _TARGET_SHEET,
                "range": _TARGET_COORDINATE,
                "role": "SOURCE",
                "sourceText": value_text,
                "note": "",
            }
        ],
        "observations": [
            {
                "key": B22_OBSERVATION_KEY,
                "arm": "condition_test",
                "valueNumber": value_number,
                "valueText": value_text,
                "numerator": numerator,
                "denominator": denominator,
                "ratePpm": None,
                "min": None,
                "max": None,
                "average": None,
                "sampleSize": denominator,
                "evidence": [
                    {
                        "sheet": _TARGET_SHEET,
                        "range": "F6",
                        "role": "SOURCE",
                        "sourceText": str(denominator),
                        "note": "",
                    },
                    {
                        "sheet": _TARGET_SHEET,
                        "range": "L6:L7",
                        "role": "SOURCE",
                        "sourceText": f"{numerator}; {value_text}",
                        "note": "",
                    },
                ],
            }
        ],
    }


def _only_observation(
    outcome: Any,
    expected_key: str | None = None,
) -> dict[str, Any]:
    if not isinstance(outcome, dict):
        raise SingleOutcomeRepairError("B22 Outcome must be an object")
    observations = outcome.get("observations")
    if (
        not isinstance(observations, list)
        or len(observations) != 1
        or not isinstance(observations[0], dict)
    ):
        raise SingleOutcomeRepairError(
            "B22 Outcome must retain exactly one Observation"
        )
    observation = observations[0]
    if (
        expected_key is not None
        and str(observation.get("key") or "") != expected_key
    ):
        raise SingleOutcomeRepairError(
            "B22 Observation key changed"
        )
    return observation


def _has_evidence(value: dict[str, Any], address: str) -> bool:
    return any(
        isinstance(item, dict)
        and str(item.get("sheet") or "") == _TARGET_SHEET
        and str(item.get("range") or "").upper() == address
        for item in value.get("evidence", [])
    )


def _iter_evidence_items(value: Any) -> Iterator[dict[str, Any]]:
    if isinstance(value, list):
        for item in value:
            yield from _iter_evidence_items(item)
        return
    if not isinstance(value, dict):
        return
    for key, item in value.items():
        if key == "evidence" and isinstance(item, list):
            for evidence in item:
                if isinstance(evidence, dict):
                    yield evidence
            continue
        yield from _iter_evidence_items(item)


def _evidence_item_covers_l7(item: dict[str, Any]) -> bool:
    if str(item.get("sheet") or "") != _TARGET_SHEET:
        return False
    match = _A1_PATTERN.fullmatch(str(item.get("range") or "").strip())
    if match is None:
        return False
    start_column = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_column = _column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    return (
        start_column <= 12 <= end_column
        and start_row <= 7 <= end_row
    )


def _column_number(label: str) -> int:
    result = 0
    for character in label.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _require_text_cell(
    cell: dict[str, Any],
    expected: str,
    *,
    merge_range: str | None = None,
) -> None:
    raw = cell.get("rawValue")
    display = cell.get("displayValue")
    if (
        not isinstance(raw, str)
        or not isinstance(display, str)
        or raw.strip() != expected
        or display.strip() != expected
    ):
        raise SingleOutcomeRepairError(
            "Focused B22 text source geometry changed"
        )
    actual_merge = cell.get("mergeRange")
    actual_role = str(cell.get("mergeRole") or "none")
    if merge_range is None:
        if actual_merge is not None or actual_role != "none":
            raise SingleOutcomeRepairError(
                "Focused B22 leaf header unexpectedly became merged"
            )
    elif actual_merge != merge_range or actual_role != "anchor":
        raise SingleOutcomeRepairError(
            "Focused B22 merged parent header geometry changed"
        )


def _require_number_cell(
    cell: dict[str, Any],
    expected: int | float,
    *,
    number_format: str | None = None,
    merge_range: str | None = None,
) -> None:
    if (
        _number(cell.get("rawValue")) != expected
        or _number(cell.get("displayValue")) != expected
        or str(cell.get("dataType") or "") != "n"
        or (
            number_format is not None
            and str(cell.get("numberFormat") or "") != number_format
        )
    ):
        raise SingleOutcomeRepairError(
            "Focused B22 numeric source geometry changed"
        )
    actual_merge = cell.get("mergeRange")
    actual_role = str(cell.get("mergeRole") or "none")
    if (
        merge_range is None
        and (actual_merge is not None or actual_role != "none")
    ) or (
        merge_range is not None
        and (actual_merge != merge_range or actual_role != "anchor")
    ):
        raise SingleOutcomeRepairError(
            "Focused B22 numeric merge geometry changed"
        )


def _number(value: Any) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SingleOutcomeRepairError(
            "Single-Outcome repair requires exact numeric values"
        )
    return float(value)


__all__ = [
    "B22_CONTENT_SHA256",
    "B22_OBSERVATION_KEY",
    "B22_OUTCOME_KEY",
    "B22_REVISION_UID",
    "B22_VALIDATION_ERROR",
    "SingleOutcomeRepairError",
    "apply_deterministic_single_outcome_repair",
    "single_outcome_repair_applicable",
    "validate_deterministic_single_outcome_repair",
]
