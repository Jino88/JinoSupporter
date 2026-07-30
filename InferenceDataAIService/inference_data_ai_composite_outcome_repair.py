"""Exact deterministic expansion of the B16 composite numeric outcomes.

The repair is deliberately isolated from the AI runner.  It accepts one
proven source revision and one exact content-coverage error, expands only four
AWF matrix Outcomes, adds only three missing Sigma-rate Outcomes, and emits
separate metadata for three repeated numeric IR identifiers.
"""

from __future__ import annotations

import copy
from collections.abc import Sequence
from typing import Any


class CompositeOutcomeRepairError(ValueError):
    """Raised when the bounded B16 projection cannot be proven safe."""


B16_REVISION_UID = "capture_revision_dd42bb526a533b15b481a9e7"
B16_CONTENT_SHA256 = (
    "57fffc6ca819cff8c10017545176706b531d491acdb0e34e26a6b20db4fd41c1"
)
B16_VALIDATION_ERROR = (
    "ContentCoverageError: Source content coverage is incomplete; "
    "163 quantitative cell(s): Test!H25, Test!I25, Test!J25, Test!K25, "
    "Test!M25, Test!N25, Test!O25, Test!H26, Test!I26, Test!J26, "
    "Test!K26, Test!M26, Test!N26, Test!O26, Test!H27, Test!I27, "
    "Test!J27, Test!K27, Test!M27, Test!N27 (+143 more)"
)
INVENTORY_EXCLUSION_SCHEMA_VERSION = "b16-inventory-exclusions-v1"
IR_EXCLUSION_REASON = "DUPLICATE_IDENTIFIER"

_STUDY_KEYS = [
    "uc_vp_coil_function_by_head",
    "production_function_by_line_date_mold_vp",
    "awf_function_test_blocks",
    "separate_all_mold_vp_function",
]
_AWF_BASELINE_KEYS = [
    "awf_value",
    "awf_sigma_input",
    "awf_sigma_counts",
    "awf_sigma_rates",
    "awf_hearing_input",
    "awf_hearing_counts",
    "awf_hearing_rates",
    "awf_total_rate",
]
_SIGMA_CATEGORIES = [
    ("thd", "THD", "H"),
    ("spl", "SPL", "I"),
    ("thd_spl", "THD+SPL", "J"),
    ("spl_rb_fo", "SPL+R&B+Fo", "K"),
]
_HEARING_CATEGORIES = [
    ("touch", "Touch", "M"),
    ("noise", "Noise", "N"),
    ("no_sound", "No sound", "O"),
]
_AWF_REPAIRED_KEYS = [
    "awf_value",
    "awf_sigma_input",
    *[
        f"awf_sigma_{slug}_count"
        for slug, _label, _column in _SIGMA_CATEGORIES
    ],
    *[
        f"awf_sigma_{slug}_rate"
        for slug, _label, _column in _SIGMA_CATEGORIES
    ],
    "awf_hearing_input",
    *[
        f"awf_hearing_{slug}_count"
        for slug, _label, _column in _HEARING_CATEGORIES
    ],
    *[
        f"awf_hearing_{slug}_rate"
        for slug, _label, _column in _HEARING_CATEGORIES
    ],
    "awf_total_rate",
]
_SEPARATE_BASELINE_KEYS = [
    "sep_input",
    "sep_ok",
    "sep_sigma_spl",
    "sep_sigma_spl_thd",
    "sep_sigma_spl_thd_f0",
    "sep_noise",
    "sep_noise_rate",
    "sep_touch",
    "sep_touch_rate",
    "sep_total_ng",
    "sep_total_ng_rate",
]
_SEPARATE_REPAIRED_KEYS = [
    "sep_input",
    "sep_ok",
    "sep_sigma_spl",
    "sep_sigma_spl_rate",
    "sep_sigma_spl_thd",
    "sep_sigma_spl_thd_rate",
    "sep_sigma_spl_thd_f0",
    "sep_sigma_spl_thd_f0_rate",
    "sep_noise",
    "sep_noise_rate",
    "sep_touch",
    "sep_touch_rate",
    "sep_total_ng",
    "sep_total_ng_rate",
]
_AWF_COUNT_ROWS = [25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45]
_AWF_RATE_ROWS = [row + 1 for row in _AWF_COUNT_ROWS]
_SEPARATE_COUNT_ROWS = [15, 17, 19]
_SEPARATE_RATE_ROWS = [16, 18, 20]
_IR_EXCLUSION_COORDINATES = ["E17", "E19", "E21"]


def composite_outcome_repair_applicable(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> bool:
    try:
        _expected_repair(
            baseline,
            validation_error=validation_error,
            focused_chunks=focused_chunks,
        )
    except (CompositeOutcomeRepairError, TypeError, ValueError):
        return False
    return True


def apply_deterministic_composite_outcome_repair(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Return the exact B16 scalar projection or an idempotent copy."""

    repaired = _expected_repair(
        baseline,
        validation_error=validation_error,
        focused_chunks=focused_chunks,
    )
    validate_deterministic_composite_outcome_repair(
        baseline,
        repaired,
        validation_error=validation_error,
        focused_chunks=focused_chunks,
    )
    return repaired


def validate_deterministic_composite_outcome_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    """Reject every change outside the exact 17 replacement/addition Outcomes."""

    expected = _expected_repair(
        baseline,
        validation_error=validation_error,
        focused_chunks=focused_chunks,
    )
    if repaired != expected:
        raise CompositeOutcomeRepairError(
            "Composite Outcome repair changed fields outside the exact "
            "B16 scalar projection"
        )


def build_b16_inventory_exclusion_metadata(
    manifest: dict[str, Any],
    *,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Describe only E17/E19/E21 as repeated values under the IR header."""

    cells = _validate_source_geometry(manifest, focused_chunks)
    return {
        "schemaVersion": INVENTORY_EXCLUSION_SCHEMA_VERSION,
        "source": {
            "revisionUid": B16_REVISION_UID,
            "contentSha256": B16_CONTENT_SHA256,
        },
        "targets": [
            {
                "sourceCellKey": str(
                    cells["Separate mold VP"][coordinate]["sourceCellKey"]
                ),
                "sheet": "Separate mold VP",
                "coordinate": coordinate,
                "numericValue": 250418008.0,
                "headerCoordinate": "E13",
                "headerText": "IR",
                "classification": "EXCLUDED_NON_RESULT",
                "exclusionReason": IR_EXCLUSION_REASON,
            }
            for coordinate in _IR_EXCLUSION_COORDINATES
        ],
    }


def apply_b16_inventory_exclusions(
    inventory: dict[str, Any],
    metadata: dict[str, Any],
) -> dict[str, Any]:
    """Move exactly three proven IR identifiers out of required results."""

    _validate_exclusion_metadata(metadata)
    result = copy.deepcopy(inventory)
    if result.get("schemaVersion") != "study-content-coverage-v1":
        raise CompositeOutcomeRepairError(
            "B16 IR exclusions require the content coverage v1 inventory"
        )
    targets = {
        str(item["sourceCellKey"]): item
        for item in metadata["targets"]
    }
    required = result.get("requiredCells")
    numeric = result.get("numericCells")
    excluded = result.get("excludedCells")
    if not all(isinstance(value, list) for value in (required, numeric, excluded)):
        raise CompositeOutcomeRepairError(
            "B16 IR exclusions require canonical inventory lists"
        )

    required_keys = {
        str(item.get("sourceCellKey") or "")
        for item in required
        if isinstance(item, dict)
    }
    excluded_keys = {
        str(item.get("sourceCellKey") or "")
        for item in excluded
        if isinstance(item, dict)
    }
    target_keys = set(targets)
    if target_keys.issubset(excluded_keys) and not (
        target_keys & required_keys
    ):
        _validate_excluded_inventory(result, targets)
        return result
    if not target_keys.issubset(required_keys) or target_keys & excluded_keys:
        raise CompositeOutcomeRepairError(
            "B16 IR inventory targets are missing, partial, or conflicting"
        )

    moved: list[dict[str, Any]] = []
    retained: list[dict[str, Any]] = []
    for item in required:
        key = str(item.get("sourceCellKey") or "")
        if key not in targets:
            retained.append(item)
            continue
        _validate_inventory_target(item, targets[key])
        updated = copy.deepcopy(item)
        updated["classification"] = "EXCLUDED_NON_RESULT"
        updated["exclusionReason"] = IR_EXCLUSION_REASON
        moved.append(updated)
    if len(moved) != 3:
        raise CompositeOutcomeRepairError(
            "B16 IR inventory exclusion must move exactly three cells"
        )
    result["requiredCells"] = retained
    result["requiredCellCount"] = len(retained)

    updated_numeric: list[dict[str, Any]] = []
    numeric_updates = 0
    for item in numeric:
        key = str(item.get("sourceCellKey") or "")
        if key in targets:
            _validate_inventory_target(item, targets[key])
            replacement = copy.deepcopy(item)
            replacement["classification"] = "EXCLUDED_NON_RESULT"
            replacement["exclusionReason"] = IR_EXCLUSION_REASON
            updated_numeric.append(replacement)
            numeric_updates += 1
        else:
            updated_numeric.append(item)
    if numeric_updates != 3:
        raise CompositeOutcomeRepairError(
            "B16 numeric inventory lacks the exact three IR cells"
        )
    result["numericCells"] = updated_numeric
    result["excludedCells"] = [*excluded, *moved]
    result["excludedCellCount"] = len(result["excludedCells"])
    _validate_excluded_inventory(result, targets)
    return result


def validate_b16_inventory_exclusions(
    baseline: dict[str, Any],
    excluded: dict[str, Any],
    metadata: dict[str, Any],
) -> None:
    expected = apply_b16_inventory_exclusions(baseline, metadata)
    if excluded != expected:
        raise CompositeOutcomeRepairError(
            "B16 inventory exclusion changed fields outside E17/E19/E21"
        )


def _expected_repair(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    if str(validation_error).strip() != B16_VALIDATION_ERROR:
        raise CompositeOutcomeRepairError(
            "Composite Outcome repair requires the exact B16 content error"
        )
    if not isinstance(baseline, dict):
        raise CompositeOutcomeRepairError(
            "Composite Outcome repair baseline must be an object"
        )
    cells = _validate_source_geometry(baseline, focused_chunks)
    studies = baseline.get("studies")
    if not isinstance(studies, list) or [
        str(study.get("key") or "") if isinstance(study, dict) else ""
        for study in studies
    ] != _STUDY_KEYS:
        raise CompositeOutcomeRepairError(
            "Composite Outcome repair requires the exact B16 Study keys"
        )
    awf = studies[2]
    separate = studies[3]
    awf_keys = _outcome_keys(awf)
    separate_keys = _outcome_keys(separate)
    desired_awf = _build_awf_outcomes(awf, cells)
    desired_separate = _build_separate_rate_outcomes(separate, cells)

    if (
        awf_keys == _AWF_REPAIRED_KEYS
        and separate_keys == _SEPARATE_REPAIRED_KEYS
    ):
        _validate_repaired_targets(
            awf,
            separate,
            desired_awf,
            desired_separate,
        )
        return copy.deepcopy(baseline)
    if (
        awf_keys != _AWF_BASELINE_KEYS
        or separate_keys != _SEPARATE_BASELINE_KEYS
    ):
        raise CompositeOutcomeRepairError(
            "B16 target Outcome keys or source order changed"
        )
    _validate_baseline_targets(awf, separate, cells)

    repaired = copy.deepcopy(baseline)
    repaired_awf = repaired["studies"][2]["outcomes"]
    repaired["studies"][2]["outcomes"] = [
        repaired_awf[0],
        repaired_awf[1],
        *desired_awf["sigma_counts"],
        *desired_awf["sigma_rates"],
        repaired_awf[4],
        *desired_awf["hearing_counts"],
        *desired_awf["hearing_rates"],
        repaired_awf[7],
    ]
    repaired_separate = repaired["studies"][3]["outcomes"]
    repaired["studies"][3]["outcomes"] = [
        repaired_separate[0],
        repaired_separate[1],
        repaired_separate[2],
        desired_separate[0],
        repaired_separate[3],
        desired_separate[1],
        repaired_separate[4],
        desired_separate[2],
        *repaired_separate[5:],
    ]
    return repaired


def _validate_source_geometry(
    manifest: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, dict[str, dict[str, Any]]]:
    source = manifest.get("source")
    if not isinstance(source, dict) or (
        str(source.get("revisionUid") or "") != B16_REVISION_UID
        or str(source.get("contentSha256") or "").lower()
        != B16_CONTENT_SHA256
    ):
        raise CompositeOutcomeRepairError(
            "Composite Outcome repair source is not the proven B16 revision"
        )
    required: dict[str, set[str]] = {
        "Test": {
            "F23",
            "G23",
            "L23",
            *[f"{column}24" for column in "GHIJKLMNO"],
            *[
                f"{column}{row}"
                for row in _AWF_COUNT_ROWS
                for column in "FGHIJKLMNO"
            ],
            *[
                f"{column}{row}"
                for row in _AWF_RATE_ROWS
                for column in "HIJKMNO"
            ],
        },
        "Separate mold VP": {
            "D13",
            "E13",
            "F13",
            "H13",
            "H14",
            "I14",
            "J14",
            *[
                f"{column}{row}"
                for row in [15, 17, 19, 21]
                for column in "DEFHIJ"
            ],
            *[
                f"{column}{row}"
                for row in _SEPARATE_RATE_ROWS
                for column in "HIJ"
            ],
        },
    }
    cells: dict[str, dict[str, dict[str, Any]]] = {
        sheet: {} for sheet in required
    }
    for chunk in focused_chunks:
        if not isinstance(chunk, dict):
            continue
        sheet_value = chunk.get("sheet")
        sheet = (
            str(sheet_value.get("title") or "")
            if isinstance(sheet_value, dict)
            else str(sheet_value or "")
        )
        if sheet not in required:
            continue
        revision = chunk.get("sourceRevision")
        if not isinstance(revision, dict) or (
            str(revision.get("revisionUid") or "") != B16_REVISION_UID
            or str(revision.get("contentSha256") or "").lower()
            != B16_CONTENT_SHA256
        ):
            raise CompositeOutcomeRepairError(
                "Focused chunk identity differs from the B16 manifest"
            )
        sheet_index = (
            sheet_value.get("sheetIndex")
            if isinstance(sheet_value, dict)
            else None
        )
        for cell in chunk.get("cells", []):
            if not isinstance(cell, dict):
                continue
            coordinate = str(cell.get("coordinate") or "").upper()
            if coordinate not in required[sheet]:
                continue
            if coordinate in cells[sheet] or cell.get("primary") is not True:
                raise CompositeOutcomeRepairError(
                    f"Focused B16 cell {sheet}!{coordinate} is duplicate "
                    "or non-primary"
                )
            if str(cell.get("sourceCellKey") or "") != (
                f"{B16_REVISION_UID}:{sheet_index}:{coordinate}"
            ):
                raise CompositeOutcomeRepairError(
                    f"Focused B16 source key changed at {sheet}!{coordinate}"
                )
            cells[sheet][coordinate] = cell
    for sheet, coordinates in required.items():
        missing = sorted(coordinates - cells[sheet].keys())
        if missing:
            raise CompositeOutcomeRepairError(
                f"Focused chunks lack B16 {sheet} cells: "
                + ", ".join(missing[:20])
            )

    test = cells["Test"]
    _text(test["F23"], "AWF", "F23:F24")
    _text(test["G23"], "SIGMA", "G23:K23")
    _text(test["L23"], "HEARING", "L23:O23")
    for coordinate, expected in {
        "G24": "Input",
        "H24": "THD",
        "I24": "SPL",
        "J24": "THD+SPL",
        "K24": "SPL+R&B+Fo",
        "L24": "Input",
        "M24": "Touch",
        "N24": "Noise",
        "O24": "No sound",
    }.items():
        _text(test[coordinate], expected)
    for index, row in enumerate(_AWF_COUNT_ROWS):
        expected_arm = "awf_test1" if index < 6 else "awf_test2"
        expected_run: str | int = (
            "Total"
            if row in {35, 45}
            else (row - 23) // 2 if row < 37
            else (row - 35) // 2
        )
        _value(test[f"F{row}"], expected_run)
        _numeric(test[f"G{row}"])
        _numeric(test[f"L{row}"])
        for column in "HIJKMNO":
            _numeric(test[f"{column}{row}"])
            _numeric(
                test[f"{column}{row + 1}"],
                number_format="0.0%",
            )
        if expected_arm not in {"awf_test1", "awf_test2"}:
            raise CompositeOutcomeRepairError("Invalid B16 AWF row arm")

    separate = cells["Separate mold VP"]
    _text(separate["D13"], "Mold VP", "D13:D14")
    _text(separate["E13"], "IR", "E13:E14")
    _text(separate["F13"], "Input", "F13:F14")
    _text(separate["H13"], "Sigma", "H13:J13")
    for coordinate, expected in {
        "H14": "SPL",
        "I14": "SPL+THD",
        "J14": "SPL+THD+F0",
    }.items():
        _text(separate[coordinate], expected)
    for row, arm_label in zip(
        [15, 17, 19, 21],
        ["#5", "#9", "#12", "Total"],
        strict=True,
    ):
        _value(separate[f"D{row}"], arm_label)
        _value(separate[f"E{row}"], 250418008)
        _numeric(separate[f"F{row}"])
        for column in "HIJ":
            _numeric(separate[f"{column}{row}"])
            if row < 21:
                _numeric(
                    separate[f"{column}{row + 1}"],
                    number_format="0.0%",
                )
    return cells


def _validate_baseline_targets(
    awf: dict[str, Any],
    separate: dict[str, Any],
    cells: dict[str, dict[str, dict[str, Any]]],
) -> None:
    if _arm_keys(awf) != ["awf_test1", "awf_test2"]:
        raise CompositeOutcomeRepairError("B16 AWF Arm keys changed")
    if _arm_keys(separate) != [
        "sep_vp5",
        "sep_vp9",
        "sep_vp12",
        "sep_total",
    ]:
        raise CompositeOutcomeRepairError(
            "B16 Separate mold VP Arm keys changed"
        )
    outcomes = awf["outcomes"]
    specs = [
        (2, "G", "K", _AWF_COUNT_ROWS),
        (3, "H", "K", _AWF_RATE_ROWS),
        (5, "L", "O", _AWF_COUNT_ROWS),
        (6, "M", "O", _AWF_RATE_ROWS),
    ]
    for outcome_index, start_column, end_column, rows in specs:
        observations = outcomes[outcome_index].get("observations")
        if not isinstance(observations, list) or len(observations) != 11:
            raise CompositeOutcomeRepairError(
                "B16 composite Outcome row count changed"
            )
        for index, (observation, row) in enumerate(
            zip(observations, rows, strict=True)
        ):
            expected_arm = "awf_test1" if index < 6 else "awf_test2"
            if (
                not isinstance(observation, dict)
                or observation.get("valueNumber") is not None
                or str(observation.get("arm") or "") != expected_arm
                or not _has_evidence(
                    observation,
                    "Test",
                    f"{start_column}{row}:{end_column}{row}",
                )
            ):
                raise CompositeOutcomeRepairError(
                    "B16 composite observation identity changed"
                )
            input_column = "G" if outcome_index in {2, 3} else "L"
            input_row = row if outcome_index in {2, 5} else row - 1
            if _number(observation.get("sampleSize")) != _number(
                cells["Test"][f"{input_column}{input_row}"]["rawValue"]
            ):
                raise CompositeOutcomeRepairError(
                    "B16 composite observation sample size changed"
                )

    count_specs = [
        (2, "H", _SEPARATE_COUNT_ROWS),
        (3, "I", _SEPARATE_COUNT_ROWS),
        (4, "J", _SEPARATE_COUNT_ROWS),
    ]
    for outcome_index, column, rows in count_specs:
        observations = separate["outcomes"][outcome_index].get(
            "observations"
        )
        if not isinstance(observations, list) or len(observations) != 4:
            raise CompositeOutcomeRepairError(
                "B16 Separate count Outcome shape changed"
            )
        for observation, row in zip(observations[:3], rows, strict=True):
            if _number(observation.get("valueNumber")) != _number(
                cells["Separate mold VP"][f"{column}{row}"]["rawValue"]
            ):
                raise CompositeOutcomeRepairError(
                    "B16 Separate count value changed"
                )


def _build_awf_outcomes(
    awf: dict[str, Any],
    cells: dict[str, dict[str, dict[str, Any]]],
) -> dict[str, list[dict[str, Any]]]:
    outcomes = awf.get("outcomes")
    if not isinstance(outcomes, list) or len(outcomes) not in {8, 18}:
        raise CompositeOutcomeRepairError(
            "B16 AWF Outcome list shape changed"
        )
    if len(outcomes) == 18:
        base_observations = {
            "sigma_counts": _base_observations(
                outcomes[2]["observations"],
                "_thd",
            ),
            "sigma_rates": _base_observations(
                outcomes[6]["observations"],
                "_thd",
            ),
            "hearing_counts": _base_observations(
                outcomes[11]["observations"],
                "_touch",
            ),
            "hearing_rates": _base_observations(
                outcomes[14]["observations"],
                "_touch",
            ),
        }
    else:
        base_observations = {
            "sigma_counts": outcomes[2]["observations"],
            "sigma_rates": outcomes[3]["observations"],
            "hearing_counts": outcomes[5]["observations"],
            "hearing_rates": outcomes[6]["observations"],
        }
    return {
        "sigma_counts": _scalar_matrix_outcomes(
            cells=cells["Test"],
            parent="SIGMA",
            parent_coordinate="G23",
            categories=_SIGMA_CATEGORIES,
            rows=_AWF_COUNT_ROWS,
            observations=base_observations["sigma_counts"],
            outcome_prefix="awf_sigma",
            metric_suffix="count",
            metric_type="defect_count",
            sample_column="G",
            rate=False,
        ),
        "sigma_rates": _scalar_matrix_outcomes(
            cells=cells["Test"],
            parent="SIGMA",
            parent_coordinate="G23",
            categories=_SIGMA_CATEGORIES,
            rows=_AWF_RATE_ROWS,
            observations=base_observations["sigma_rates"],
            outcome_prefix="awf_sigma",
            metric_suffix="rate",
            metric_type="defect_rate_percent",
            sample_column="G",
            rate=True,
        ),
        "hearing_counts": _scalar_matrix_outcomes(
            cells=cells["Test"],
            parent="HEARING",
            parent_coordinate="L23",
            categories=_HEARING_CATEGORIES,
            rows=_AWF_COUNT_ROWS,
            observations=base_observations["hearing_counts"],
            outcome_prefix="awf_hearing",
            metric_suffix="count",
            metric_type="defect_count",
            sample_column="L",
            rate=False,
        ),
        "hearing_rates": _scalar_matrix_outcomes(
            cells=cells["Test"],
            parent="HEARING",
            parent_coordinate="L23",
            categories=_HEARING_CATEGORIES,
            rows=_AWF_RATE_ROWS,
            observations=base_observations["hearing_rates"],
            outcome_prefix="awf_hearing",
            metric_suffix="rate",
            metric_type="defect_rate_percent",
            sample_column="L",
            rate=True,
        ),
    }


def _scalar_matrix_outcomes(
    *,
    cells: dict[str, dict[str, Any]],
    parent: str,
    parent_coordinate: str,
    categories: Sequence[tuple[str, str, str]],
    rows: Sequence[int],
    observations: Sequence[dict[str, Any]],
    outcome_prefix: str,
    metric_suffix: str,
    metric_type: str,
    sample_column: str,
    rate: bool,
) -> list[dict[str, Any]]:
    if len(observations) != 11:
        raise CompositeOutcomeRepairError(
            "B16 scalar expansion requires eleven source rows"
        )
    result: list[dict[str, Any]] = []
    for slug, label, column in categories:
        scalar_observations: list[dict[str, Any]] = []
        for base, row in zip(observations, rows, strict=True):
            if not isinstance(base, dict):
                raise CompositeOutcomeRepairError(
                    "B16 composite Observation must be an object"
                )
            source_cell = cells[f"{column}{row}"]
            raw = _number(source_cell.get("rawValue"))
            value_number = raw * 100 if rate else raw
            count_row = row - 1 if rate else row
            sample_cell = cells[f"{sample_column}{count_row}"]
            sample_size = _number(sample_cell.get("rawValue"))
            base_key = str(base.get("key") or "")
            if not base_key:
                raise CompositeOutcomeRepairError(
                    "B16 composite Observation key is empty"
                )
            scalar_observations.append(
                {
                    "key": f"{base_key}_{slug}",
                    "arm": str(base.get("arm") or ""),
                    "valueNumber": value_number,
                    "valueText": _value_text(
                        value_number,
                        percent=rate,
                    ),
                    "numerator": None,
                    "denominator": None,
                    "ratePpm": None,
                    "min": None,
                    "max": None,
                    "average": None,
                    "sampleSize": sample_size,
                    "evidence": [
                        _evidence(
                            "Test",
                            f"{column}{row}",
                            _value_text(value_number, percent=rate),
                        ),
                        _evidence(
                            "Test",
                            f"{sample_column}{count_row}",
                            _value_text(sample_size),
                        ),
                        _evidence(
                            "Test",
                            f"F{count_row}",
                            _source_text(cells[f"F{count_row}"]),
                        ),
                    ],
                }
            )
        result.append(
            {
                "key": f"{outcome_prefix}_{slug}_{metric_suffix}",
                "originalLabel": f"{parent} {label} {metric_suffix}",
                "metricType": metric_type,
                "unit": "%" if rate else "",
                "favorableDirection": "UNKNOWN",
                "evidence": [
                    _evidence(
                        "Test",
                        parent_coordinate,
                        parent,
                    ),
                    _evidence(
                        "Test",
                        f"{column}24",
                        label,
                    ),
                ],
                "observations": scalar_observations,
            }
        )
    return result


def _base_observations(
    observations: Sequence[dict[str, Any]],
    suffix: str,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for observation in observations:
        if not isinstance(observation, dict):
            raise CompositeOutcomeRepairError(
                "Existing B16 scalar Observation must be an object"
            )
        key = str(observation.get("key") or "")
        if not key.endswith(suffix):
            raise CompositeOutcomeRepairError(
                "Existing B16 scalar Observation key changed"
            )
        base = copy.deepcopy(observation)
        base["key"] = key[: -len(suffix)]
        result.append(base)
    return result


def _build_separate_rate_outcomes(
    separate: dict[str, Any],
    cells: dict[str, dict[str, dict[str, Any]]],
) -> list[dict[str, Any]]:
    source = cells["Separate mold VP"]
    count_indexes = (
        [2, 4, 6]
        if _outcome_keys(separate) == _SEPARATE_REPAIRED_KEYS
        else [2, 3, 4]
    )
    specs = [
        ("sep_sigma_spl_rate", "Sigma SPL rate", "H"),
        ("sep_sigma_spl_thd_rate", "Sigma SPL+THD rate", "I"),
        ("sep_sigma_spl_thd_f0_rate", "Sigma SPL+THD+F0 rate", "J"),
    ]
    result: list[dict[str, Any]] = []
    for count_index, (key, label, column) in zip(
        count_indexes,
        specs,
        strict=True,
    ):
        count_outcome = separate["outcomes"][count_index]
        count_observations = count_outcome.get("observations")
        if not isinstance(count_observations, list) or len(
            count_observations
        ) < 3:
            raise CompositeOutcomeRepairError(
                "B16 Separate count observations changed"
            )
        observations: list[dict[str, Any]] = []
        for base, count_row, rate_row in zip(
            count_observations[:3],
            _SEPARATE_COUNT_ROWS,
            _SEPARATE_RATE_ROWS,
            strict=True,
        ):
            raw = _number(source[f"{column}{rate_row}"]["rawValue"])
            value_number = raw * 100
            sample_size = _number(source[f"F{count_row}"]["rawValue"])
            numerator = _number(source[f"{column}{count_row}"]["rawValue"])
            observations.append(
                {
                    "key": f"{base['key']}_rate",
                    "arm": str(base.get("arm") or ""),
                    "valueNumber": value_number,
                    "valueText": _value_text(value_number, percent=True),
                    "numerator": numerator,
                    "denominator": sample_size,
                    "ratePpm": None,
                    "min": None,
                    "max": None,
                    "average": None,
                    "sampleSize": sample_size,
                    "evidence": [
                        _evidence(
                            "Separate mold VP",
                            f"F{count_row}",
                            _value_text(sample_size),
                        ),
                        _evidence(
                            "Separate mold VP",
                            f"{column}{count_row}:{column}{rate_row}",
                            (
                                f"{_value_text(numerator)}; "
                                f"{_value_text(value_number, percent=True)}"
                            ),
                        ),
                    ],
                }
            )
        result.append(
            {
                "key": key,
                "originalLabel": label,
                "metricType": "defect_rate_percent",
                "unit": "%",
                "favorableDirection": "UNKNOWN",
                "evidence": [
                    _evidence("Separate mold VP", "H13", "Sigma"),
                    _evidence(
                        "Separate mold VP",
                        f"{column}14",
                        str(source[f"{column}14"]["rawValue"]).strip(),
                    ),
                ],
                "observations": observations,
            }
        )
    return result


def _validate_repaired_targets(
    awf: dict[str, Any],
    separate: dict[str, Any],
    desired_awf: dict[str, list[dict[str, Any]]],
    desired_separate: list[dict[str, Any]],
) -> None:
    awf_expected = [
        awf["outcomes"][0],
        awf["outcomes"][1],
        *desired_awf["sigma_counts"],
        *desired_awf["sigma_rates"],
        awf["outcomes"][10],
        *desired_awf["hearing_counts"],
        *desired_awf["hearing_rates"],
        awf["outcomes"][17],
    ]
    if awf["outcomes"] != awf_expected:
        raise CompositeOutcomeRepairError(
            "Existing B16 AWF scalar projection changed"
        )
    separate_expected = [
        separate["outcomes"][0],
        separate["outcomes"][1],
        separate["outcomes"][2],
        desired_separate[0],
        separate["outcomes"][4],
        desired_separate[1],
        separate["outcomes"][6],
        desired_separate[2],
        *separate["outcomes"][8:],
    ]
    if separate["outcomes"] != separate_expected:
        raise CompositeOutcomeRepairError(
            "Existing B16 Separate rate projection changed"
        )


def _outcome_keys(study: dict[str, Any]) -> list[str]:
    outcomes = study.get("outcomes")
    if not isinstance(outcomes, list):
        raise CompositeOutcomeRepairError("B16 Outcomes must be a list")
    return [
        str(outcome.get("key") or "")
        if isinstance(outcome, dict)
        else ""
        for outcome in outcomes
    ]


def _arm_keys(study: dict[str, Any]) -> list[str]:
    arms = study.get("arms")
    if not isinstance(arms, list):
        raise CompositeOutcomeRepairError("B16 Arms must be a list")
    return [
        str(arm.get("key") or "") if isinstance(arm, dict) else ""
        for arm in arms
    ]


def _validate_exclusion_metadata(metadata: dict[str, Any]) -> None:
    if not isinstance(metadata, dict) or (
        metadata.get("schemaVersion")
        != INVENTORY_EXCLUSION_SCHEMA_VERSION
        or metadata.get("source")
        != {
            "revisionUid": B16_REVISION_UID,
            "contentSha256": B16_CONTENT_SHA256,
        }
    ):
        raise CompositeOutcomeRepairError(
            "B16 inventory exclusion metadata identity is invalid"
        )
    targets = metadata.get("targets")
    if not isinstance(targets, list) or [
        str(item.get("coordinate") or "")
        if isinstance(item, dict)
        else ""
        for item in targets
    ] != _IR_EXCLUSION_COORDINATES:
        raise CompositeOutcomeRepairError(
            "B16 inventory metadata must target only E17/E19/E21"
        )
    for target in targets:
        if (
            target.get("sheet") != "Separate mold VP"
            or _number(target.get("numericValue")) != 250418008
            or target.get("headerCoordinate") != "E13"
            or target.get("headerText") != "IR"
            or target.get("classification") != "EXCLUDED_NON_RESULT"
            or target.get("exclusionReason") != IR_EXCLUSION_REASON
        ):
            raise CompositeOutcomeRepairError(
                "B16 inventory metadata target changed"
            )


def _validate_inventory_target(
    item: dict[str, Any],
    target: dict[str, Any],
) -> None:
    if (
        str(item.get("sourceCellKey") or "")
        != str(target.get("sourceCellKey") or "")
        or item.get("sheet") != "Separate mold VP"
        or item.get("coordinate") != target.get("coordinate")
        or _number(item.get("numericValue")) != 250418008
    ):
        raise CompositeOutcomeRepairError(
            "B16 inventory IR source identity or value changed"
        )


def _validate_excluded_inventory(
    inventory: dict[str, Any],
    targets: dict[str, dict[str, Any]],
) -> None:
    target_keys = set(targets)
    required_keys = {
        str(item.get("sourceCellKey") or "")
        for item in inventory["requiredCells"]
    }
    if target_keys & required_keys:
        raise CompositeOutcomeRepairError(
            "B16 excluded IR cells remain required"
        )
    excluded_by_key = {
        str(item.get("sourceCellKey") or ""): item
        for item in inventory["excludedCells"]
        if str(item.get("sourceCellKey") or "") in target_keys
    }
    if set(excluded_by_key) != target_keys:
        raise CompositeOutcomeRepairError(
            "B16 inventory does not contain exactly three IR exclusions"
        )
    for key, item in excluded_by_key.items():
        _validate_inventory_target(item, targets[key])
        if (
            item.get("classification") != "EXCLUDED_NON_RESULT"
            or item.get("exclusionReason") != IR_EXCLUSION_REASON
        ):
            raise CompositeOutcomeRepairError(
                "B16 IR exclusion classification changed"
            )
    if inventory.get("requiredCellCount") != len(
        inventory["requiredCells"]
    ) or inventory.get("excludedCellCount") != len(
        inventory["excludedCells"]
    ):
        raise CompositeOutcomeRepairError(
            "B16 inventory counts do not match their lists"
        )


def _has_evidence(
    value: dict[str, Any],
    sheet: str,
    address: str,
) -> bool:
    return any(
        isinstance(item, dict)
        and item.get("sheet") == sheet
        and str(item.get("range") or "").upper() == address
        for item in value.get("evidence", [])
    )


def _evidence(sheet: str, address: str, source_text: str) -> dict[str, Any]:
    return {
        "sheet": sheet,
        "range": address,
        "role": "SOURCE",
        "sourceText": source_text,
        "note": "",
    }


def _text(
    cell: dict[str, Any],
    expected: str,
    merge_range: str | None = None,
) -> None:
    if (
        str(cell.get("rawValue") or "").strip() != expected
        or str(cell.get("displayValue") or "").strip() != expected
    ):
        raise CompositeOutcomeRepairError(
            "Focused B16 source header changed"
        )
    actual_merge = cell.get("mergeRange")
    actual_role = str(cell.get("mergeRole") or "none")
    if merge_range is None:
        valid_merge = actual_merge is None and actual_role == "none"
    else:
        valid_merge = (
            actual_merge == merge_range and actual_role == "anchor"
        )
    if not valid_merge:
        raise CompositeOutcomeRepairError(
            "Focused B16 source merge geometry changed"
        )


def _value(cell: dict[str, Any], expected: str | int) -> None:
    raw = cell.get("rawValue")
    if isinstance(expected, str):
        valid = str(raw or "").strip() == expected
    else:
        valid = _number(raw) == expected
    if not valid:
        raise CompositeOutcomeRepairError(
            "Focused B16 source identity value changed"
        )


def _numeric(
    cell: dict[str, Any],
    *,
    number_format: str | None = None,
) -> None:
    _number(cell.get("rawValue"))
    _number(cell.get("displayValue"))
    if cell.get("dataType") != "n" or (
        number_format is not None
        and cell.get("numberFormat") != number_format
    ):
        raise CompositeOutcomeRepairError(
            "Focused B16 numeric geometry or format changed"
        )


def _number(value: Any) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise CompositeOutcomeRepairError(
            "B16 repair requires exact numeric source values"
        )
    return float(value)


def _value_text(value: float, *, percent: bool = False) -> str:
    if percent:
        return f"{value:.1f}%"
    if float(value).is_integer():
        return str(int(value))
    return str(value)


def _source_text(cell: dict[str, Any]) -> str:
    raw = cell.get("rawValue")
    if isinstance(raw, str):
        return raw.strip()
    return _value_text(_number(raw))


__all__ = [
    "B16_CONTENT_SHA256",
    "B16_REVISION_UID",
    "B16_VALIDATION_ERROR",
    "CompositeOutcomeRepairError",
    "INVENTORY_EXCLUSION_SCHEMA_VERSION",
    "IR_EXCLUSION_REASON",
    "apply_b16_inventory_exclusions",
    "apply_deterministic_composite_outcome_repair",
    "build_b16_inventory_exclusion_metadata",
    "composite_outcome_repair_applicable",
    "validate_b16_inventory_exclusions",
    "validate_deterministic_composite_outcome_repair",
]
