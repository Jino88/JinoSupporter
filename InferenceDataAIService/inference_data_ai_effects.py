from __future__ import annotations

import hashlib
import math
import re
import unicodedata
from typing import Any


class EffectCalculationError(ValueError):
    pass


def _text(value: object) -> str:
    return str(value or "").strip()


def _status(value: object) -> str:
    return _text(value).upper()


def _number(record: dict[str, Any], *keys: str) -> float | None:
    for key in keys:
        value = record.get(key)
        if value is None or value == "":
            continue
        try:
            number = float(value)
        except (TypeError, ValueError) as exc:
            raise EffectCalculationError(f"{key} is not numeric: {value!r}") from exc
        if not math.isfinite(number):
            raise EffectCalculationError(f"{key} must be finite")
        return number
    return None


def _normalized(value: object) -> str:
    text = unicodedata.normalize("NFKC", _text(value)).casefold()
    text = re.sub(r"\s+", " ", text)
    return text


def validate_comparison_for_effects(
    comparison: dict[str, Any],
    *,
    study: dict[str, Any] | None = None,
) -> None:
    if _status(comparison.get("validityStatus", comparison.get("validity_status"))) != "VALID":
        raise EffectCalculationError("comparison validity must be VALID")
    if _status(comparison.get("confoundingStatus", comparison.get("confounding_status"))) != "NONE":
        raise EffectCalculationError("comparison must be explicitly unconfounded")
    if _status(comparison.get("verificationStatus", comparison.get("verification_status"))) != "VERIFIED":
        raise EffectCalculationError("comparison must be VERIFIED")
    compared = comparison.get("comparedArm", comparison.get("compared_arm_id"))
    control = comparison.get("controlArm", comparison.get("control_arm_id"))
    if compared in (None, "") or control in (None, "") or str(compared) == str(control):
        raise EffectCalculationError("comparison requires distinct compared and control arms")
    if study is not None:
        if _status(study.get("verificationStatus", study.get("verification_status"))) != "VERIFIED":
            raise EffectCalculationError("study must be VERIFIED")
        if _status(study.get("comparabilityStatus", study.get("comparability_status"))) != "VALID":
            raise EffectCalculationError("study comparability must be VALID")
        if _status(study.get("confoundingStatus", study.get("confounding_status"))) != "NONE":
            raise EffectCalculationError("study must be explicitly unconfounded")


def _rate_proportion(observation: dict[str, Any]) -> float | None:
    numerator = _number(observation, "numerator")
    denominator = _number(observation, "denominator")
    stored_ppm = _number(observation, "ratePpm", "rate_ppm")
    if numerator is None and denominator is None:
        if stored_ppm is None:
            return None
        return stored_ppm / 1_000_000.0
    if numerator is None or denominator is None:
        raise EffectCalculationError("numerator and denominator must be supplied together")
    if numerator < 0:
        raise EffectCalculationError("numerator cannot be negative")
    if denominator <= 0:
        raise EffectCalculationError("denominator must be positive")
    if numerator > denominator:
        raise EffectCalculationError("numerator cannot exceed denominator for a rate outcome")
    proportion = numerator / denominator
    if stored_ppm is not None:
        calculated_ppm = proportion * 1_000_000.0
        tolerance = max(1.0, abs(calculated_ppm) * 1e-6)
        if abs(calculated_ppm - stored_ppm) > tolerance:
            raise EffectCalculationError(
                f"stored ratePpm does not match numerator/denominator: {stored_ppm} != {calculated_ppm}"
            )
    return proportion


def _continuous_value(observation: dict[str, Any]) -> float | None:
    return _number(observation, "valueNumber", "value_number", "average", "average_value")


def _direction(delta: float, favorable_direction: str) -> str:
    if math.isclose(delta, 0.0, abs_tol=1e-12):
        return "NO_CHANGE"
    favorable = _status(favorable_direction)
    if favorable == "LOWER":
        return "IMPROVED" if delta < 0 else "WORSENED"
    if favorable == "HIGHER":
        return "IMPROVED" if delta > 0 else "WORSENED"
    return "HIGHER" if delta > 0 else "LOWER"


def calculate_effect_bundle(
    *,
    compared_observation: dict[str, Any],
    control_observation: dict[str, Any],
    comparison: dict[str, Any],
    outcome: dict[str, Any],
    study: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Calculate deterministic effects for one validated arm pair.

    This function is domain-neutral.  Product/process names never participate
    in arithmetic; they belong in the comparability partition supplied by the
    caller.
    """

    validate_comparison_for_effects(comparison, study=study)
    favorable = _text(outcome.get("favorableDirection", outcome.get("favorable_direction")))
    metric_type = _normalized(outcome.get("metricType", outcome.get("metric_type")))
    effects: list[dict[str, Any]] = []

    compared_rate = _rate_proportion(compared_observation)
    control_rate = _rate_proportion(control_observation)
    is_rate = compared_rate is not None or control_rate is not None or "rate" in metric_type
    if is_rate:
        if compared_rate is None or control_rate is None:
            raise EffectCalculationError("both arms require compatible rate observations")
        delta = compared_rate - control_rate
        direction = _direction(delta, favorable)
        effects.extend(
            [
                {
                    "effectType": "ABSOLUTE_RATE_DIFFERENCE",
                    "estimate": delta,
                    "unit": "proportion",
                    "direction": direction,
                    "formulaVersion": "rate-effects-v1",
                },
                {
                    "effectType": "PERCENTAGE_POINT_CHANGE",
                    "estimate": delta * 100.0,
                    "unit": "%p",
                    "direction": direction,
                    "formulaVersion": "rate-effects-v1",
                },
                {
                    "effectType": "RATE_DIFFERENCE_PPM",
                    "estimate": delta * 1_000_000.0,
                    "unit": "ppm",
                    "direction": direction,
                    "formulaVersion": "rate-effects-v1",
                },
            ]
        )
        if not math.isclose(control_rate, 0.0, abs_tol=1e-15):
            effects.extend(
                [
                    {
                        "effectType": "RELATIVE_CHANGE_PERCENT",
                        "estimate": delta * 100.0 / control_rate,
                        "unit": "%",
                        "direction": direction,
                        "formulaVersion": "rate-effects-v1",
                    },
                    {
                        "effectType": "RISK_RATIO",
                        "estimate": compared_rate / control_rate,
                        "unit": "ratio",
                        "direction": direction,
                        "formulaVersion": "rate-effects-v1",
                    },
                ]
            )
        return effects

    if "count" in metric_type:
        raise EffectCalculationError(
            "count outcomes require numerator/denominator or ratePpm; "
            "raw counts cannot be treated as continuous effects"
        )

    compared_value = _continuous_value(compared_observation)
    control_value = _continuous_value(control_observation)
    if compared_value is None or control_value is None:
        raise EffectCalculationError("both arms require compatible numeric observations")
    delta = compared_value - control_value
    unit = _text(outcome.get("unit", outcome.get("original_unit")))
    direction = _direction(delta, favorable)
    effects.append(
        {
            "effectType": "MEAN_DIFFERENCE",
            "estimate": delta,
            "unit": unit,
            "direction": direction,
            "formulaVersion": "continuous-effects-v1",
        }
    )
    if not math.isclose(control_value, 0.0, abs_tol=1e-15):
        effects.append(
            {
                "effectType": "RELATIVE_CHANGE_PERCENT",
                "estimate": delta * 100.0 / control_value,
                "unit": "%",
                "direction": direction,
                "formulaVersion": "continuous-effects-v1",
            }
        )
    return effects


def comparability_partition(
    *,
    contexts: list[dict[str, Any]],
    factors: list[dict[str, Any]],
    outcome: dict[str, Any],
    comparison: dict[str, Any],
) -> dict[str, Any]:
    """Build a conservative, generic partition for evidence aggregation.

    All supplied context and factor dimensions are retained.  This prevents a
    new domain (model, lot, supplier, voltage, material, etc.) from being
    silently discarded merely because it was not known when the code shipped.
    """

    context_parts = sorted(
        (
            _normalized(item.get("kind", item.get("contextKind"))),
            _normalized(
                item.get("conceptUid")
                or item.get("normalizedValue")
                or item.get("normalized_value")
                or item.get("originalValue")
                or item.get("original_value")
            ),
        )
        for item in contexts
    )
    factor_parts = sorted(
        (
            _normalized(item.get("conceptUid") or item.get("canonicalName") or item.get("originalLabel")),
            _normalized(item.get("baselineCondition")),
            _normalized(item.get("changedCondition")),
            _status(item.get("isolationStatus")),
        )
        for item in factors
    )
    outcome_part = {
        "concept": _normalized(
            outcome.get("conceptUid")
            or outcome.get("canonicalName")
            or outcome.get("originalLabel")
            or outcome.get("label")
        ),
        "metricType": _normalized(outcome.get("metricType", outcome.get("metric_type"))),
        "unit": _normalized(outcome.get("unit", outcome.get("original_unit"))),
        "denominatorBasis": _normalized(
            outcome.get("denominatorBasis", outcome.get("denominator_basis"))
        ),
    }
    comparison_part = {
        "designType": _normalized(comparison.get("designType", comparison.get("design_type"))),
        "matchingBasis": _normalized(
            comparison.get("matchingBasis", comparison.get("matching_basis"))
        ),
    }
    payload = repr((context_parts, factor_parts, outcome_part, comparison_part))
    key = hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]
    return {
        "partitionKey": f"partition_{key}",
        "contexts": context_parts,
        "factors": factor_parts,
        "outcome": outcome_part,
        "comparison": comparison_part,
    }
