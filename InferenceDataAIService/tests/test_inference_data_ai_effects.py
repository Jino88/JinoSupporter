from __future__ import annotations

import unittest

import inference_data_ai_effects as effects


class EffectCalculatorTests(unittest.TestCase):
    def comparison(self) -> dict:
        return {
            "comparedArm": "test",
            "controlArm": "control",
            "validityStatus": "VALID",
            "confoundingStatus": "NONE",
            "verificationStatus": "VERIFIED",
            "designType": "CONTROL_TEST",
            "matchingBasis": "same model, lot, line and period",
        }

    def study(self) -> dict:
        return {
            "verificationStatus": "VERIFIED",
            "comparabilityStatus": "VALID",
            "confoundingStatus": "NONE",
        }

    def test_rate_effects_include_percentage_points_relative_change_and_risk_ratio(self) -> None:
        result = effects.calculate_effect_bundle(
            compared_observation={"numerator": 20, "denominator": 100, "ratePpm": 200000},
            control_observation={"numerator": 10, "denominator": 100, "ratePpm": 100000},
            comparison=self.comparison(),
            outcome={"metricType": "defect_rate", "favorableDirection": "LOWER"},
            study=self.study(),
        )
        by_type = {item["effectType"]: item for item in result}
        self.assertAlmostEqual(10.0, by_type["PERCENTAGE_POINT_CHANGE"]["estimate"])
        self.assertAlmostEqual(100000.0, by_type["RATE_DIFFERENCE_PPM"]["estimate"])
        self.assertAlmostEqual(100.0, by_type["RELATIVE_CHANGE_PERCENT"]["estimate"])
        self.assertAlmostEqual(2.0, by_type["RISK_RATIO"]["estimate"])
        self.assertEqual("WORSENED", by_type["PERCENTAGE_POINT_CHANGE"]["direction"])

    def test_zero_control_rate_does_not_invent_relative_change(self) -> None:
        result = effects.calculate_effect_bundle(
            compared_observation={"numerator": 1, "denominator": 100},
            control_observation={"numerator": 0, "denominator": 100},
            comparison=self.comparison(),
            outcome={"metricType": "defect_rate", "favorableDirection": "LOWER"},
            study=self.study(),
        )
        self.assertNotIn("RELATIVE_CHANGE_PERCENT", {item["effectType"] for item in result})
        self.assertNotIn("RISK_RATIO", {item["effectType"] for item in result})

    def test_confounded_comparison_is_rejected(self) -> None:
        comparison = self.comparison()
        comparison["confoundingStatus"] = "CONFOUNDED"
        with self.assertRaisesRegex(effects.EffectCalculationError, "unconfounded"):
            effects.calculate_effect_bundle(
                compared_observation={"valueNumber": 12},
                control_observation={"valueNumber": 10},
                comparison=comparison,
                outcome={"metricType": "measurement", "unit": "N"},
                study=self.study(),
            )

    def test_continuous_effect_uses_original_unit(self) -> None:
        result = effects.calculate_effect_bundle(
            compared_observation={"average": 12},
            control_observation={"average": 10},
            comparison=self.comparison(),
            outcome={"metricType": "measurement", "unit": "N", "favorableDirection": "HIGHER"},
            study=self.study(),
        )
        self.assertEqual("MEAN_DIFFERENCE", result[0]["effectType"])
        self.assertEqual(2.0, result[0]["estimate"])
        self.assertEqual("N", result[0]["unit"])
        self.assertEqual("IMPROVED", result[0]["direction"])

    def test_rate_mismatch_is_rejected(self) -> None:
        with self.assertRaisesRegex(effects.EffectCalculationError, "does not match"):
            effects.calculate_effect_bundle(
                compared_observation={"numerator": 1, "denominator": 100, "ratePpm": 999999},
                control_observation={"numerator": 0, "denominator": 100, "ratePpm": 0},
                comparison=self.comparison(),
                outcome={"metricType": "defect_rate"},
                study=self.study(),
            )

    def test_raw_count_is_not_treated_as_continuous_effect(self) -> None:
        with self.assertRaisesRegex(
            effects.EffectCalculationError,
            "raw counts cannot be treated as continuous effects",
        ):
            effects.calculate_effect_bundle(
                compared_observation={"valueNumber": 13, "sampleSize": 160},
                control_observation={"valueNumber": 20, "sampleSize": 548},
                comparison=self.comparison(),
                outcome={"metricType": "defect_count"},
                study=self.study(),
            )

    def test_count_with_explicit_denominator_uses_rate_effects(self) -> None:
        result = effects.calculate_effect_bundle(
            compared_observation={
                "valueNumber": 13,
                "numerator": 13,
                "denominator": 160,
            },
            control_observation={
                "valueNumber": 20,
                "numerator": 20,
                "denominator": 548,
            },
            comparison=self.comparison(),
            outcome={"metricType": "defect_count"},
            study=self.study(),
        )
        by_type = {item["effectType"]: item for item in result}
        self.assertAlmostEqual(
            8.125 - (20 / 548 * 100),
            by_type["PERCENTAGE_POINT_CHANGE"]["estimate"],
        )
        self.assertNotIn("MEAN_DIFFERENCE", by_type)

    def test_partition_keeps_unseen_context_dimensions(self) -> None:
        base = effects.comparability_partition(
            contexts=[
                {"kind": "MODEL", "normalizedValue": "A"},
                {"kind": "CUSTOM_NEW_DOMAIN", "normalizedValue": "condition-x"},
            ],
            factors=[
                {
                    "canonicalName": "unseen factor",
                    "baselineCondition": "1",
                    "changedCondition": "2",
                    "isolationStatus": "ISOLATED",
                }
            ],
            outcome={"canonicalName": "unseen outcome", "metricType": "measurement", "unit": "u"},
            comparison={"designType": "CONTROL_TEST", "matchingBasis": "same lot"},
        )
        changed = effects.comparability_partition(
            contexts=[
                {"kind": "MODEL", "normalizedValue": "A"},
                {"kind": "CUSTOM_NEW_DOMAIN", "normalizedValue": "condition-y"},
            ],
            factors=[
                {
                    "canonicalName": "unseen factor",
                    "baselineCondition": "1",
                    "changedCondition": "2",
                    "isolationStatus": "ISOLATED",
                }
            ],
            outcome={"canonicalName": "unseen outcome", "metricType": "measurement", "unit": "u"},
            comparison={"designType": "CONTROL_TEST", "matchingBasis": "same lot"},
        )
        self.assertNotEqual(base["partitionKey"], changed["partitionKey"])


if __name__ == "__main__":
    unittest.main()
