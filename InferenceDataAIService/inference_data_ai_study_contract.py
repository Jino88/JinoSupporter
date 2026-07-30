from __future__ import annotations

import copy
import math
import re
import unicodedata
from collections.abc import Callable
from typing import Any

from inference_data_ai_effects import EffectCalculationError, calculate_effect_bundle


SCHEMA_VERSION = "canonical-study-manifest-v1"


class StudyContractError(ValueError):
    pass


EvidenceChecker = Callable[[dict[str, Any]], None]


def _object(value: object, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise StudyContractError(f"{path} must be an object")
    return value


def _list(value: object, path: str) -> list[Any]:
    if not isinstance(value, list):
        raise StudyContractError(f"{path} must be a list")
    return value


def _text(value: object, path: str, *, required: bool = True) -> str:
    text = str(value or "").strip()
    if required and not text:
        raise StudyContractError(f"{path} is required")
    return text


def _status(value: object) -> str:
    return str(value or "").strip().upper()


def _number(value: object, path: str, *, positive: bool = False) -> float | None:
    if value is None or value == "":
        return None
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise StudyContractError(f"{path} must be numeric") from exc
    if not math.isfinite(result):
        raise StudyContractError(f"{path} must be finite")
    if positive and result <= 0:
        raise StudyContractError(f"{path} must be positive")
    return result


def _unique_keyed(items: object, path: str) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    result: list[dict[str, Any]] = []
    by_key: dict[str, dict[str, Any]] = {}
    by_normalized_key: dict[str, str] = {}
    for index, value in enumerate(_list(items, path)):
        item = _object(value, f"{path}[{index}]")
        key = _text(item.get("key"), f"{path}[{index}].key")
        if key in by_key:
            raise StudyContractError(f"{path} contains duplicate key: {key}")
        normalized_key = re.sub(
            r"\s+",
            " ",
            unicodedata.normalize("NFKC", key).strip().casefold(),
        )
        if normalized_key in by_normalized_key:
            raise StudyContractError(
                f"{path} contains stable-ID-colliding keys: "
                f"{by_normalized_key[normalized_key]} and {key}"
            )
        result.append(item)
        by_key[key] = item
        by_normalized_key[normalized_key] = key
    return result, by_key


def _evidence(
    value: object,
    path: str,
    checker: EvidenceChecker,
    *,
    required: bool,
) -> list[dict[str, Any]]:
    items = _list(value if value is not None else [], path)
    if required and not items:
        raise StudyContractError(f"{path} requires at least one source range")
    result: list[dict[str, Any]] = []
    for index, item_value in enumerate(items):
        item = _object(item_value, f"{path}[{index}]")
        _text(item.get("sheet"), f"{path}[{index}].sheet")
        address = _text(item.get("range"), f"{path}[{index}].range")
        if not re.fullmatch(r"\$?[A-Za-z]{1,4}\$?[1-9]\d*(?::\$?[A-Za-z]{1,4}\$?[1-9]\d*)?", address):
            raise StudyContractError(f"{path}[{index}].range must be an A1 cell or range")
        checker(item)
        result.append(item)
    return result


def _assert_status(value: object, path: str, allowed: set[str]) -> str:
    status = _status(value)
    if status not in allowed:
        raise StudyContractError(f"{path} must be one of {sorted(allowed)}")
    return status


def _has_numeric_claim(observation: dict[str, Any]) -> bool:
    return any(
        observation.get(key) not in (None, "")
        for key in (
            "valueNumber",
            "numerator",
            "denominator",
            "ratePpm",
            "min",
            "max",
            "average",
        )
    )


def _range_bounds(
    address: object,
    path: str,
) -> tuple[int, int, int, int]:
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,4})\$?([1-9]\d*)"
        r"(?::\$?([A-Za-z]{1,4})\$?([1-9]\d*))?",
        str(address or "").strip(),
    )
    if not match:
        raise StudyContractError(f"{path} must be an A1 cell or range")

    def column_number(label: str) -> int:
        value = 0
        for char in label.upper():
            value = value * 26 + ord(char) - ord("A") + 1
        return value

    start_col = column_number(match.group(1))
    start_row = int(match.group(2))
    end_col = column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_row < start_row or end_col < start_col:
        raise StudyContractError(f"{path} must not be reversed")
    return start_row, start_col, end_row, end_col


def _range_shape(address: object, path: str) -> tuple[int, int]:
    start_row, start_col, end_row, end_col = _range_bounds(
        address,
        path,
    )
    return end_row - start_row + 1, end_col - start_col + 1


def validate_study_manifest(
    data: dict[str, Any],
    *,
    evidence_checker: EvidenceChecker | None = None,
) -> dict[str, Any]:
    """Validate a domain-neutral, source-backed canonical study manifest.

    The validator deliberately has no product, model, factor, or outcome
    whitelist.  Unknown concepts are allowed when their original wording and
    source evidence are retained.
    """

    if not isinstance(data, dict):
        raise StudyContractError("manifest must be an object")
    if data.get("schemaVersion") != SCHEMA_VERSION:
        raise StudyContractError(f"schemaVersion must be {SCHEMA_VERSION}")
    result = copy.deepcopy(data)
    checker = evidence_checker or (lambda _item: None)
    source = _object(result.get("source"), "source")
    _text(source.get("dataset"), "source.dataset")
    _text(source.get("sourcePath"), "source.sourcePath")
    _text(source.get("revisionUid") or source.get("contentSha256") or source.get("fingerprint"), "source revision identity")
    source_complete = bool(source.get("contentComplete", True))

    analysis = _object(result.get("workbookAnalysis"), "workbookAnalysis")
    _text(analysis.get("key"), "workbookAnalysis.key")
    _text(analysis.get("title"), "workbookAnalysis.title")
    _text(analysis.get("summary"), "workbookAnalysis.summary", required=False)
    analysis_verification = _assert_status(
        analysis.get("verificationStatus", "NEEDS_REVIEW"),
        "workbookAnalysis.verificationStatus",
        {"VERIFIED", "NEEDS_REVIEW", "EXCLUDED", "FAILED", "STALE"},
    )
    _evidence(
        analysis.get("evidence", []),
        "workbookAnalysis.evidence",
        checker,
        required=analysis_verification == "VERIFIED",
    )

    studies, _ = _unique_keyed(result.get("studies", []), "studies")
    for study_index, study in enumerate(studies):
        path = f"studies[{study_index}]"
        _text(study.get("title"), f"{path}.title")
        _text(study.get("designType"), f"{path}.designType")
        verification = _assert_status(
            study.get("verificationStatus", "NEEDS_REVIEW"),
            f"{path}.verificationStatus",
            {"VERIFIED", "NEEDS_REVIEW", "EXCLUDED", "FAILED", "STALE"},
        )
        comparability = _assert_status(
            study.get("comparabilityStatus", "UNASSESSED"),
            f"{path}.comparabilityStatus",
            {"VALID", "PARTIAL", "INVALID", "UNASSESSED"},
        )
        confounding = _assert_status(
            study.get("confoundingStatus", "UNASSESSED"),
            f"{path}.confoundingStatus",
            {"NONE", "POSSIBLE", "CONFOUNDED", "UNASSESSED"},
        )
        if verification == "VERIFIED" and (not source_complete or analysis_verification != "VERIFIED"):
            raise StudyContractError(f"{path} cannot be VERIFIED from an incomplete or unverified workbook analysis")
        _evidence(study.get("evidence", []), f"{path}.evidence", checker, required=verification == "VERIFIED")

        contexts, _ = _unique_keyed(study.get("contexts", []), f"{path}.contexts")
        for context_index, context in enumerate(contexts):
            context_path = f"{path}.contexts[{context_index}]"
            _text(context.get("kind"), f"{context_path}.kind")
            _text(context.get("originalValue"), f"{context_path}.originalValue")
            _evidence(context.get("evidence", []), f"{context_path}.evidence", checker, required=verification == "VERIFIED")

        factors, factor_by_key = _unique_keyed(study.get("factors", []), f"{path}.factors")
        for factor_index, factor in enumerate(factors):
            factor_path = f"{path}.factors[{factor_index}]"
            _text(factor.get("originalLabel"), f"{factor_path}.originalLabel")
            _text(
                factor.get("baselineCondition"),
                f"{factor_path}.baselineCondition",
                required=False,
            )
            _text(
                factor.get("changedCondition"),
                f"{factor_path}.changedCondition",
                required=False,
            )
            _assert_status(
                factor.get("isolationStatus", "UNASSESSED"),
                f"{factor_path}.isolationStatus",
                {"ISOLATED", "MULTI_FACTOR", "CONFOUNDED", "UNASSESSED"},
            )
            _evidence(factor.get("evidence", []), f"{factor_path}.evidence", checker, required=True)

        arms, arm_by_key = _unique_keyed(study.get("arms", []), f"{path}.arms")
        for arm_index, arm in enumerate(arms):
            arm_path = f"{path}.arms[{arm_index}]"
            _assert_status(
                arm.get("role", "OTHER"),
                f"{arm_path}.role",
                {"CONTROL", "COMPARATOR", "TREATMENT", "TEST", "BEFORE", "AFTER", "REFERENCE", "OTHER"},
            )
            _text(arm.get("label"), f"{arm_path}.label")
            if arm.get("sampleSize") not in (None, ""):
                _number(arm.get("sampleSize"), f"{arm_path}.sampleSize", positive=True)
            _evidence(arm.get("evidence", []), f"{arm_path}.evidence", checker, required=verification == "VERIFIED")
            for fv_index, factor_value in enumerate(_list(arm.get("factorValues", []), f"{arm_path}.factorValues")):
                factor_value = _object(factor_value, f"{arm_path}.factorValues[{fv_index}]")
                factor_key = _text(factor_value.get("factor"), f"{arm_path}.factorValues[{fv_index}].factor")
                if factor_key not in factor_by_key:
                    raise StudyContractError(f"{arm_path}.factorValues[{fv_index}] references unknown factor {factor_key}")

        outcomes, outcome_by_key = _unique_keyed(study.get("outcomes", []), f"{path}.outcomes")
        outcome_representation_by_key: dict[str, dict[str, bool]] = {}
        observations_by_outcome_arm: dict[tuple[str, str], list[dict[str, Any]]] = {}
        for outcome_index, outcome in enumerate(outcomes):
            outcome_path = f"{path}.outcomes[{outcome_index}]"
            _text(outcome.get("originalLabel"), f"{outcome_path}.originalLabel")
            _text(outcome.get("metricType"), f"{outcome_path}.metricType")
            favorable_direction = _status(
                outcome.get("favorableDirection", "UNKNOWN")
            )
            favorable_direction = {
                "HIGH": "HIGHER",
                "LOW": "LOWER",
            }.get(favorable_direction, favorable_direction)
            outcome["favorableDirection"] = _assert_status(
                favorable_direction,
                f"{outcome_path}.favorableDirection",
                {"LOWER", "HIGHER", "TARGET", "NONE", "UNKNOWN"},
            )
            _evidence(outcome.get("evidence", []), f"{outcome_path}.evidence", checker, required=verification == "VERIFIED")
            representation = {
                "quantitative": False,
                "qualitativeOnly": False,
            }
            outcome_representation_by_key[str(outcome["key"])] = representation
            observations, _ = _unique_keyed(outcome.get("observations", []), f"{outcome_path}.observations")
            for observation_index, observation in enumerate(observations):
                observation_path = f"{outcome_path}.observations[{observation_index}]"
                arm_key = _text(observation.get("arm"), f"{observation_path}.arm")
                if arm_key not in arm_by_key:
                    raise StudyContractError(f"{observation_path} references unknown arm {arm_key}")
                numerator = _number(observation.get("numerator"), f"{observation_path}.numerator")
                denominator = _number(
                    observation.get("denominator"),
                    f"{observation_path}.denominator",
                    positive=True,
                )
                if (numerator is None) != (denominator is None):
                    raise StudyContractError(f"{observation_path} numerator and denominator must be supplied together")
                if numerator is not None and numerator < 0:
                    raise StudyContractError(f"{observation_path}.numerator cannot be negative")
                if numerator is not None and denominator is not None and numerator > denominator:
                    raise StudyContractError(f"{observation_path}.numerator cannot exceed denominator")
                _evidence(
                    observation.get("evidence", []),
                    f"{observation_path}.evidence",
                    checker,
                    required=_has_numeric_claim(observation) or verification == "VERIFIED",
                )
                if _has_numeric_claim(observation):
                    representation["quantitative"] = True
                elif str(observation.get("valueText") or "").strip():
                    representation["qualitativeOnly"] = True
                stratum_key = str(observation.get("stratumKey") or "").strip()
                replicate_key = str(observation.get("replicateKey") or "").strip()
                unique_key = (str(outcome["key"]), arm_key, stratum_key, replicate_key)
                if any(
                    str(item.get("stratumKey") or "").strip() == stratum_key
                    and str(item.get("replicateKey") or "").strip() == replicate_key
                    for item in observations_by_outcome_arm.get(
                        (str(outcome["key"]), arm_key),
                        [],
                    )
                ):
                    raise StudyContractError(
                        f"{outcome_path} has duplicate observation identity for arm {arm_key}, "
                        f"stratum {stratum_key or '<primary>'}, replicate {replicate_key or '<primary>'}"
                    )
                observations_by_outcome_arm.setdefault(
                    (str(outcome["key"]), arm_key),
                    [],
                ).append(observation)

        measurement_series, measurement_series_by_key = _unique_keyed(
            study.get("measurementSeries", []),
            f"{path}.measurementSeries",
        )
        for series_index, series in enumerate(measurement_series):
            series_path = f"{path}.measurementSeries[{series_index}]"
            series_role = _assert_status(
                series.get("seriesRole", "RAW"),
                f"{series_path}.seriesRole",
                {"RAW", "AGGREGATE"},
            )
            series["seriesRole"] = series_role
            aggregation_function = _status(
                series.get("aggregationFunction")
            )
            aggregate_of_values = _list(
                series.get("aggregateOfSeries", []),
                f"{series_path}.aggregateOfSeries",
            )
            aggregate_of_series: list[str] = []
            for aggregate_of_index, aggregate_of_value in enumerate(
                aggregate_of_values
            ):
                referenced_key = _text(
                    aggregate_of_value,
                    (
                        f"{series_path}.aggregateOfSeries"
                        f"[{aggregate_of_index}]"
                    ),
                )
                if referenced_key in aggregate_of_series:
                    raise StudyContractError(
                        f"{series_path}.aggregateOfSeries must be unique"
                    )
                aggregate_of_series.append(referenced_key)
            if series_role == "AGGREGATE":
                if aggregation_function != "AVERAGE":
                    raise StudyContractError(
                        f"{series_path}.aggregationFunction must be AVERAGE "
                        "for an AGGREGATE series"
                    )
                if not aggregate_of_series:
                    raise StudyContractError(
                        f"{series_path}.aggregateOfSeries requires at least "
                        "one RAW source series"
                    )
                series["aggregationFunction"] = aggregation_function
                series["aggregateOfSeries"] = aggregate_of_series
            elif aggregation_function or aggregate_of_series:
                raise StudyContractError(
                    f"{series_path} RAW series cannot declare "
                    "aggregationFunction or aggregateOfSeries"
                )
            outcome_key = _text(
                series.get("outcome"),
                f"{series_path}.outcome",
            )
            arm_key = _text(series.get("arm"), f"{series_path}.arm")
            if outcome_key not in outcome_by_key:
                raise StudyContractError(
                    f"{series_path} references unknown outcome {outcome_key}"
                )
            if arm_key not in arm_by_key:
                raise StudyContractError(
                    f"{series_path} references unknown arm {arm_key}"
                )
            sheet_name = str(series.get("sheet") or "")
            if not sheet_name.strip():
                raise StudyContractError(
                    f"{series_path}.sheet is required"
                )
            axis_source = _assert_status(
                series.get("axisSource"),
                f"{series_path}.axisSource",
                {"HEADER", "ROW_IDENTITY"},
            )
            series["axisSource"] = axis_source
            _text(
                series.get("axisLabel"),
                f"{series_path}.axisLabel",
                required=False,
            )
            _text(
                series.get("axisUnit"),
                f"{series_path}.axisUnit",
                required=False,
            )
            _text(
                series.get("valueUnit"),
                f"{series_path}.valueUnit",
                required=False,
            )
            series_verification = _assert_status(
                series.get("verificationStatus", "NEEDS_REVIEW"),
                f"{series_path}.verificationStatus",
                {
                    "VERIFIED",
                    "NEEDS_REVIEW",
                    "EXCLUDED",
                    "FAILED",
                    "STALE",
                },
            )
            series["verificationStatus"] = series_verification
            if series_verification == "VERIFIED" and verification != "VERIFIED":
                raise StudyContractError(
                    f"{series_path} cannot be VERIFIED under an unverified study"
                )
            header_range = _text(
                series.get("headerRange"),
                f"{series_path}.headerRange",
            )
            value_range = _text(
                series.get("valueRange"),
                f"{series_path}.valueRange",
            )
            row_identity_range = _text(
                series.get("rowIdentityRange"),
                f"{series_path}.rowIdentityRange",
            )
            header_rows, header_columns = _range_shape(
                header_range,
                f"{series_path}.headerRange",
            )
            value_rows, value_columns = _range_shape(
                value_range,
                f"{series_path}.valueRange",
            )
            identity_rows, identity_columns = _range_shape(
                row_identity_range,
                f"{series_path}.rowIdentityRange",
            )
            header_bounds = _range_bounds(
                header_range,
                f"{series_path}.headerRange",
            )
            value_bounds = _range_bounds(
                value_range,
                f"{series_path}.valueRange",
            )
            identity_bounds = _range_bounds(
                row_identity_range,
                f"{series_path}.rowIdentityRange",
            )
            if header_rows != 1 or header_columns != value_columns:
                raise StudyContractError(
                    f"{series_path}.headerRange must be one row with the "
                    "same column count as valueRange"
                )
            if header_bounds[1] != value_bounds[1] or header_bounds[3] != value_bounds[3]:
                raise StudyContractError(
                    f"{series_path}.headerRange must align to the exact "
                    "valueRange columns"
                )
            if identity_columns != 1 or identity_rows != value_rows:
                raise StudyContractError(
                    f"{series_path}.rowIdentityRange must be one column with "
                    "the same row count as valueRange"
                )
            if identity_bounds[0] != value_bounds[0] or identity_bounds[2] != value_bounds[2]:
                raise StudyContractError(
                    f"{series_path}.rowIdentityRange must align to the exact "
                    "valueRange rows"
                )
            aggregate_ranges = _list(
                series.get("aggregateReplicateRanges", []),
                f"{series_path}.aggregateReplicateRanges",
            )
            if series_role == "AGGREGATE" and aggregate_ranges:
                raise StudyContractError(
                    f"{series_path} standalone AGGREGATE series must not "
                    "declare aggregateReplicateRanges"
                )
            aggregate_coordinates: set[tuple[int, int]] = set()
            for aggregate_index, aggregate_value in enumerate(
                aggregate_ranges
            ):
                aggregate_path = (
                    f"{series_path}.aggregateReplicateRanges"
                    f"[{aggregate_index}]"
                )
                aggregate_address = _text(
                    aggregate_value,
                    aggregate_path,
                )
                aggregate_bounds = _range_bounds(
                    aggregate_address,
                    aggregate_path,
                )
                if axis_source == "HEADER":
                    aligned = (
                        aggregate_bounds[1] == identity_bounds[1]
                        and aggregate_bounds[3] == identity_bounds[3]
                        and aggregate_bounds[0] >= identity_bounds[0]
                        and aggregate_bounds[2] <= identity_bounds[2]
                    )
                else:
                    aligned = (
                        aggregate_bounds[0] == header_bounds[0]
                        and aggregate_bounds[2] == header_bounds[2]
                        and aggregate_bounds[1] >= header_bounds[1]
                        and aggregate_bounds[3] <= header_bounds[3]
                    )
                if not aligned:
                    identity_field = (
                        "rowIdentityRange"
                        if axis_source == "HEADER"
                        else "headerRange"
                    )
                    raise StudyContractError(
                        f"{aggregate_path} must be contained in and aligned "
                        f"with {series_path}.{identity_field}"
                    )
                for aggregate_row in range(
                    aggregate_bounds[0],
                    aggregate_bounds[2] + 1,
                ):
                    for aggregate_column in range(
                        aggregate_bounds[1],
                        aggregate_bounds[3] + 1,
                    ):
                        coordinate = (aggregate_row, aggregate_column)
                        if coordinate in aggregate_coordinates:
                            raise StudyContractError(
                                f"{series_path}.aggregateReplicateRanges "
                                "must not overlap"
                            )
                        aggregate_coordinates.add(coordinate)
            replicate_identity_size = (
                identity_rows
                if axis_source == "HEADER"
                else header_columns
            )
            if (
                aggregate_coordinates
                and len(aggregate_coordinates) >= replicate_identity_size
            ):
                raise StudyContractError(
                    f"{series_path}.aggregateReplicateRanges must retain "
                    "at least one RAW replicate identity; a standalone "
                    "aggregate must not be modeled as a measurement series"
                )
            for range_field, address, role in (
                ("headerRange", header_range, "MEASUREMENT_HEADER"),
                ("valueRange", value_range, "MEASUREMENT_VALUES"),
                (
                    "rowIdentityRange",
                    row_identity_range,
                    "ROW_IDENTITY",
                ),
            ):
                checker(
                    {
                        "sheet": sheet_name,
                        "range": address,
                        "role": role,
                    }
                )

        for series in measurement_series:
            outcome_representation_by_key[str(series["outcome"])][
                "quantitative"
            ] = True
        for outcome_index, outcome in enumerate(outcomes):
            representation = outcome_representation_by_key[
                str(outcome["key"])
            ]
            if (
                representation["quantitative"]
                and representation["qualitativeOnly"]
            ):
                raise StudyContractError(
                    f"{path}.outcomes[{outcome_index}] mixes quantitative "
                    "measurements with qualitative-only observations; split "
                    "the qualitative result into a distinct categorical "
                    "Outcome"
                )

        for series_index, series in enumerate(measurement_series):
            if series["seriesRole"] != "AGGREGATE":
                continue
            series_path = f"{path}.measurementSeries[{series_index}]"
            for referenced_key in series["aggregateOfSeries"]:
                if referenced_key == series["key"]:
                    raise StudyContractError(
                        f"{series_path}.aggregateOfSeries cannot reference "
                        "itself"
                    )
                referenced_series = measurement_series_by_key.get(
                    referenced_key
                )
                if referenced_series is None:
                    raise StudyContractError(
                        f"{series_path}.aggregateOfSeries references unknown "
                        f"measurementSeries {referenced_key}"
                    )
                if referenced_series["seriesRole"] != "RAW":
                    raise StudyContractError(
                        f"{series_path}.aggregateOfSeries cannot reference "
                        f"nested AGGREGATE series {referenced_key}"
                    )
                if referenced_series["outcome"] != series["outcome"]:
                    raise StudyContractError(
                        f"{series_path}.aggregateOfSeries outcome mismatch "
                        f"for measurementSeries {referenced_key}"
                    )

        comparisons, _ = _unique_keyed(study.get("comparisons", []), f"{path}.comparisons")
        for comparison_index, comparison in enumerate(comparisons):
            comparison_path = f"{path}.comparisons[{comparison_index}]"
            compared_key = _text(comparison.get("comparedArm"), f"{comparison_path}.comparedArm")
            control_key = _text(comparison.get("controlArm"), f"{comparison_path}.controlArm")
            if compared_key not in arm_by_key or control_key not in arm_by_key:
                raise StudyContractError(f"{comparison_path} references an unknown arm")
            if compared_key == control_key:
                raise StudyContractError(f"{comparison_path} cannot compare an arm with itself")
            _text(comparison.get("designType"), f"{comparison_path}.designType")
            validity = _assert_status(
                comparison.get("validityStatus", "NEEDS_REVIEW"),
                f"{comparison_path}.validityStatus",
                {"VALID", "NEEDS_REVIEW", "INVALID", "EXCLUDED"},
            )
            comparison_confounding = _assert_status(
                comparison.get("confoundingStatus", "UNASSESSED"),
                f"{comparison_path}.confoundingStatus",
                {"NONE", "POSSIBLE", "CONFOUNDED", "UNASSESSED"},
            )
            comparison_verification = _assert_status(
                comparison.get("verificationStatus", "NEEDS_REVIEW"),
                f"{comparison_path}.verificationStatus",
                {"VERIFIED", "NEEDS_REVIEW", "INVALID", "EXCLUDED", "STALE"},
            )
            _evidence(comparison.get("evidence", []), f"{comparison_path}.evidence", checker, required=True)
            aggregation_eligible = bool(comparison.get("aggregationEligible", False))
            if aggregation_eligible:
                if (
                    verification != "VERIFIED"
                    or comparability != "VALID"
                    or confounding != "NONE"
                    or validity != "VALID"
                    or comparison_confounding != "NONE"
                    or comparison_verification != "VERIFIED"
                ):
                    raise StudyContractError(
                        f"{comparison_path} cannot be aggregationEligible without a verified, valid, comparable, unconfounded study and comparison"
                    )
                _text(comparison.get("matchingBasis"), f"{comparison_path}.matchingBasis")
                for factor_index, factor in enumerate(factors):
                    if not str(factor.get("baselineCondition") or "").strip() or not str(
                        factor.get("changedCondition") or ""
                    ).strip():
                        raise StudyContractError(
                            f"{comparison_path} cannot be aggregationEligible when "
                            f"{path}.factors[{factor_index}] lacks an explicit baseline or changed condition"
                        )

            effects = _list(comparison.get("effects", []), f"{comparison_path}.effects")
            if aggregation_eligible and not effects:
                raise StudyContractError(f"{comparison_path} aggregationEligible comparison requires calculated effects")
            for effect_index, effect_value in enumerate(effects):
                effect = _object(effect_value, f"{comparison_path}.effects[{effect_index}]")
                effect_path = f"{comparison_path}.effects[{effect_index}]"
                outcome_key = _text(effect.get("outcome"), f"{effect_path}.outcome")
                if outcome_key not in outcome_by_key:
                    raise StudyContractError(f"{effect_path} references unknown outcome {outcome_key}")
                effect_type = _text(effect.get("effectType"), f"{effect_path}.effectType")
                estimate = _number(effect.get("estimate"), f"{effect_path}.estimate")
                effect_verification = _assert_status(
                    effect.get("verificationStatus", "NEEDS_REVIEW"),
                    f"{effect_path}.verificationStatus",
                    {"VERIFIED", "NEEDS_REVIEW", "INVALID", "EXCLUDED", "STALE"},
                )
                _evidence(effect.get("evidence", []), f"{effect_path}.evidence", checker, required=True)
                effect_stratum = str(effect.get("stratumKey") or "").strip()
                effect_replicate = str(effect.get("replicateKey") or "").strip()

                def select_observation(arm_key: str) -> dict[str, Any] | None:
                    candidates = observations_by_outcome_arm.get((outcome_key, arm_key), [])
                    if effect_stratum or effect_replicate:
                        matches = [
                            item
                            for item in candidates
                            if str(item.get("stratumKey") or "").strip() == effect_stratum
                            and str(item.get("replicateKey") or "").strip() == effect_replicate
                        ]
                    else:
                        primary = [
                            item
                            for item in candidates
                            if not str(item.get("stratumKey") or "").strip()
                            and not str(item.get("replicateKey") or "").strip()
                        ]
                        matches = primary if primary else candidates
                    if len(matches) > 1:
                        raise StudyContractError(
                            f"{effect_path} must identify stratumKey/replicateKey because "
                            f"{arm_key} has multiple candidate observations"
                        )
                    return matches[0] if matches else None

                compared_observation = select_observation(compared_key)
                control_observation = select_observation(control_key)
                if compared_observation is None or control_observation is None:
                    raise StudyContractError(f"{effect_path} requires observations for both compared and control arms")
                if effect_verification == "VERIFIED" or aggregation_eligible:
                    try:
                        calculated = calculate_effect_bundle(
                            compared_observation=compared_observation,
                            control_observation=control_observation,
                            comparison={
                                **comparison,
                                "validityStatus": validity,
                                "confoundingStatus": comparison_confounding,
                                "verificationStatus": comparison_verification,
                            },
                            outcome=outcome_by_key[outcome_key],
                            study={
                                **study,
                                "verificationStatus": verification,
                                "comparabilityStatus": comparability,
                                "confoundingStatus": confounding,
                            },
                        )
                    except EffectCalculationError as exc:
                        raise StudyContractError(f"{effect_path}: {exc}") from exc
                    expected = next((item for item in calculated if item["effectType"] == effect_type), None)
                    if expected is None:
                        raise StudyContractError(f"{effect_path}.effectType is not calculable from its observations")
                    tolerance = max(1e-9, abs(float(expected["estimate"])) * 1e-6)
                    if estimate is None or abs(estimate - float(expected["estimate"])) > tolerance:
                        raise StudyContractError(
                            f"{effect_path}.estimate does not match deterministic calculation: "
                            f"{estimate} != {expected['estimate']}"
                        )

        conclusions, _ = _unique_keyed(study.get("conclusions", []), f"{path}.conclusions")
        for conclusion_index, conclusion in enumerate(conclusions):
            conclusion_path = f"{path}.conclusions[{conclusion_index}]"
            _text(conclusion.get("text"), f"{conclusion_path}.text")
            claim_type = _assert_status(
                conclusion.get("claimType"),
                f"{conclusion_path}.claimType",
                {
                    "SOURCE_CONCLUSION",
                    "AI_DERIVED_DESCRIPTIVE",
                },
            )
            conclusion["claimType"] = claim_type
            causal_strength = _assert_status(
                conclusion.get("causalStrength", "UNSPECIFIED"),
                f"{conclusion_path}.causalStrength",
                {"CAUSAL", "ASSOCIATION", "DESCRIPTIVE", "UNSPECIFIED"},
            )
            _evidence(conclusion.get("evidence", []), f"{conclusion_path}.evidence", checker, required=True)
            if (
                claim_type == "AI_DERIVED_DESCRIPTIVE"
                and causal_strength not in {"DESCRIPTIVE", "UNSPECIFIED"}
            ):
                raise StudyContractError(
                    f"{conclusion_path} AI_DERIVED_DESCRIPTIVE claim "
                    "cannot assert association or causality"
                )
            if causal_strength == "CAUSAL" and (
                comparability != "VALID" or confounding != "NONE" or verification != "VERIFIED"
            ):
                raise StudyContractError(f"{conclusion_path} cannot claim causality from unverified or confounded evidence")
    return result
