from __future__ import annotations

import copy
import unittest

import inference_data_ai_study_contract as contract


class StudyContractTests(unittest.TestCase):
    def evidence(self) -> list[dict]:
        return [{"sheet": "Data", "range": "B2:D4", "role": "SOURCE"}]

    def manifest(self) -> dict:
        evidence = self.evidence()
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "dataset": "Fixture",
                "sourcePath": "fixture.xlsx",
                "revisionUid": "revision_1",
                "contentComplete": True,
            },
            "workbookAnalysis": {
                "key": "generic-cooling-review",
                "title": "Cooling condition review",
                "verificationStatus": "VERIFIED",
                "summary": "Cooling time was compared.",
                "evidence": evidence,
            },
            "studies": [
                {
                    "key": "cooling-time-vs-tension",
                    "title": "Cooling time and tension",
                    "designType": "CONTROL_TEST",
                    "verificationStatus": "VERIFIED",
                    "comparabilityStatus": "VALID",
                    "confoundingStatus": "NONE",
                    "evidence": evidence,
                    "contexts": [
                        {
                            "key": "model",
                            "kind": "MODEL",
                            "originalValue": "Unseen Model X",
                            "evidence": evidence,
                        }
                    ],
                    "factors": [
                        {
                            "key": "cooling-time",
                            "originalLabel": "Cooling hold duration",
                            "baselineCondition": "4 s",
                            "changedCondition": "8 s",
                            "isolationStatus": "ISOLATED",
                            "evidence": evidence,
                        }
                    ],
                    "arms": [
                        {
                            "key": "test",
                            "role": "TEST",
                            "label": "8 s",
                            "condition": "8 second hold",
                            "sampleSize": 10,
                            "evidence": evidence,
                            "factorValues": [{"factor": "cooling-time", "value": "8 s"}],
                        },
                        {
                            "key": "control",
                            "role": "CONTROL",
                            "label": "4 s",
                            "condition": "4 second hold",
                            "sampleSize": 10,
                            "evidence": evidence,
                            "factorValues": [{"factor": "cooling-time", "value": "4 s"}],
                        },
                    ],
                    "outcomes": [
                        {
                            "key": "tension",
                            "originalLabel": "Unseen custom tension metric",
                            "metricType": "measurement",
                            "unit": "N",
                            "favorableDirection": "HIGHER",
                            "evidence": evidence,
                            "observations": [
                                {
                                    "key": "test-value",
                                    "arm": "test",
                                    "average": 12,
                                    "evidence": evidence,
                                },
                                {
                                    "key": "control-value",
                                    "arm": "control",
                                    "average": 10,
                                    "evidence": evidence,
                                },
                            ],
                        }
                    ],
                    "comparisons": [
                        {
                            "key": "test-vs-control",
                            "comparedArm": "test",
                            "controlArm": "control",
                            "designType": "CONTROL_TEST",
                            "matchingBasis": "same model, lot, line and date",
                            "validityStatus": "VALID",
                            "confoundingStatus": "NONE",
                            "verificationStatus": "VERIFIED",
                            "aggregationEligible": True,
                            "evidence": evidence,
                            "effects": [
                                {
                                    "outcome": "tension",
                                    "effectType": "MEAN_DIFFERENCE",
                                    "estimate": 2,
                                    "unit": "N",
                                    "verificationStatus": "VERIFIED",
                                    "evidence": evidence,
                                }
                            ],
                        }
                    ],
                    "conclusions": [
                        {
                            "key": "result",
                            "text": "The recorded tension is higher under the changed condition.",
                            "claimType": "SOURCE_CONCLUSION",
                            "causalStrength": "ASSOCIATION",
                            "evidence": evidence,
                        }
                    ],
                }
            ],
        }

    def test_unseen_factor_and_outcome_names_are_allowed(self) -> None:
        checked: list[str] = []
        result = contract.validate_study_manifest(
            self.manifest(),
            evidence_checker=lambda item: checked.append(item["range"]),
        )
        self.assertEqual("cooling-time", result["studies"][0]["factors"][0]["key"])
        self.assertGreater(len(checked), 5)

    def test_favorable_direction_normalizes_legacy_aliases(self) -> None:
        data = self.manifest()
        outcome = data["studies"][0]["outcomes"][0]
        outcome["favorableDirection"] = "high"
        result = contract.validate_study_manifest(data)
        self.assertEqual(
            "HIGHER",
            result["studies"][0]["outcomes"][0][
                "favorableDirection"
            ],
        )

        outcome["favorableDirection"] = "sideways"
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "favorableDirection",
        ):
            contract.validate_study_manifest(data)

    def test_confounded_comparison_cannot_be_aggregation_eligible(self) -> None:
        data = self.manifest()
        data["studies"][0]["comparisons"][0]["confoundingStatus"] = "CONFOUNDED"
        with self.assertRaisesRegex(contract.StudyContractError, "aggregationEligible"):
            contract.validate_study_manifest(data)

    def test_incomplete_source_cannot_create_verified_study(self) -> None:
        data = self.manifest()
        data["source"]["contentComplete"] = False
        with self.assertRaisesRegex(contract.StudyContractError, "cannot be VERIFIED"):
            contract.validate_study_manifest(data)

    def test_numeric_claim_requires_evidence(self) -> None:
        data = self.manifest()
        data["studies"][0]["outcomes"][0]["observations"][0]["evidence"] = []
        with self.assertRaisesRegex(contract.StudyContractError, "requires at least one source range"):
            contract.validate_study_manifest(data)

    def test_verified_effect_must_match_deterministic_calculation(self) -> None:
        data = self.manifest()
        data["studies"][0]["comparisons"][0]["effects"][0]["estimate"] = 200
        with self.assertRaisesRegex(contract.StudyContractError, "does not match"):
            contract.validate_study_manifest(data)

    def test_causal_claim_is_rejected_when_study_is_confounded(self) -> None:
        data = self.manifest()
        study = data["studies"][0]
        study["verificationStatus"] = "NEEDS_REVIEW"
        study["comparabilityStatus"] = "PARTIAL"
        study["confoundingStatus"] = "CONFOUNDED"
        comparison = study["comparisons"][0]
        comparison["aggregationEligible"] = False
        comparison["effects"][0]["verificationStatus"] = "NEEDS_REVIEW"
        study["conclusions"][0]["causalStrength"] = "CAUSAL"
        with self.assertRaisesRegex(contract.StudyContractError, "cannot claim causality"):
            contract.validate_study_manifest(data)

    def test_conclusion_requires_explicit_provenance_type(self) -> None:
        data = self.manifest()
        del data["studies"][0]["conclusions"][0]["claimType"]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "claimType",
        ):
            contract.validate_study_manifest(data)

    def test_ai_derived_conclusion_cannot_assert_association(self) -> None:
        data = self.manifest()
        conclusion = data["studies"][0]["conclusions"][0]
        conclusion["claimType"] = "AI_DERIVED_DESCRIPTIVE"
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "cannot assert association",
        ):
            contract.validate_study_manifest(data)

    def test_quantitative_outcome_rejects_qualitative_only_observation(
        self,
    ) -> None:
        data = self.manifest()
        outcome = data["studies"][0]["outcomes"][0]
        outcome["observations"].append(
            {
                "key": "test-pass",
                "arm": "test",
                "valueText": "Pass",
                "replicateKey": "acceptance",
                "evidence": self.evidence(),
            }
        )
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "mixes quantitative measurements with qualitative-only",
        ):
            contract.validate_study_manifest(data)

    def test_split_quantitative_and_categorical_outcomes_are_allowed(
        self,
    ) -> None:
        data = self.manifest()
        study = data["studies"][0]
        study["outcomes"].append(
            {
                "key": "tension-acceptance",
                "originalLabel": "Tension acceptance",
                "metricType": "pass_fail",
                "unit": "",
                "favorableDirection": "UNKNOWN",
                "evidence": self.evidence(),
                "observations": [
                    {
                        "key": "test-pass",
                        "arm": "test",
                        "valueText": "Pass",
                        "evidence": self.evidence(),
                    }
                ],
            }
        )
        result = contract.validate_study_manifest(data)
        self.assertEqual(
            ["tension", "tension-acceptance"],
            [item["key"] for item in result["studies"][0]["outcomes"]],
        )

    def test_numeric_observation_display_text_remains_quantitative(
        self,
    ) -> None:
        data = self.manifest()
        data["studies"][0]["outcomes"][0]["observations"][0][
            "valueText"
        ] = "12 N"
        result = contract.validate_study_manifest(data)
        self.assertEqual(
            "12 N",
            result["studies"][0]["outcomes"][0]["observations"][0][
                "valueText"
            ],
        )

    def test_distinct_replicate_observations_are_allowed(self) -> None:
        data = self.manifest()
        study = data["studies"][0]
        study["verificationStatus"] = "NEEDS_REVIEW"
        study["comparabilityStatus"] = "UNASSESSED"
        study["confoundingStatus"] = "UNASSESSED"
        comparison = study["comparisons"][0]
        comparison["aggregationEligible"] = False
        comparison["verificationStatus"] = "NEEDS_REVIEW"
        comparison["validityStatus"] = "NEEDS_REVIEW"
        comparison["confoundingStatus"] = "UNASSESSED"
        comparison["effects"] = []
        observations = study["outcomes"][0]["observations"]
        observations[0]["replicateKey"] = "sample-1"
        observations.append(
            {
                **copy.deepcopy(observations[0]),
                "key": "test-value-2",
                "replicateKey": "sample-2",
                "average": 12.5,
            }
        )
        result = contract.validate_study_manifest(data)
        self.assertEqual(3, len(result["studies"][0]["outcomes"][0]["observations"]))

    def test_aggregation_requires_explicit_factor_conditions(self) -> None:
        data = self.manifest()
        data["studies"][0]["factors"][0]["baselineCondition"] = ""
        with self.assertRaisesRegex(contract.StudyContractError, "explicit baseline"):
            contract.validate_study_manifest(data)

    def test_keys_that_collapse_to_the_same_stable_uid_are_rejected(
        self,
    ) -> None:
        data = self.manifest()
        context = copy.deepcopy(data["studies"][0]["contexts"][0])
        context["key"] = " MODEL "
        data["studies"][0]["contexts"].append(context)
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "stable-ID-colliding keys",
        ):
            contract.validate_study_manifest(data)

    def test_measurement_series_accepts_compact_wide_range_mapping(self) -> None:
        data = self.manifest()
        data["studies"][0]["measurementSeries"] = [
            {
                "key": "tension-profile",
                "outcome": "tension",
                "arm": "test",
                "sheet": "Data",
                "headerRange": "B1:D1",
                "valueRange": "B2:D4",
                "rowIdentityRange": "A2:A4",
                "aggregateReplicateRanges": ["D1"],
                "axisSource": "ROW_IDENTITY",
                "axisLabel": "Position",
                "axisUnit": "mm",
                "valueUnit": "N",
                "stratumKey": "profile-a",
                "verificationStatus": "VERIFIED",
            }
        ]
        checked: list[tuple[str, str]] = []
        result = contract.validate_study_manifest(
            data,
            evidence_checker=lambda item: checked.append(
                (item["range"], item["role"])
            ),
        )
        self.assertEqual(
            "tension-profile",
            result["studies"][0]["measurementSeries"][0]["key"],
        )
        self.assertIn(("B1:D1", "MEASUREMENT_HEADER"), checked)
        self.assertIn(("B2:D4", "MEASUREMENT_VALUES"), checked)
        self.assertIn(("A2:A4", "ROW_IDENTITY"), checked)

    def test_measurement_series_preserves_exact_sheet_name_for_checker(
        self,
    ) -> None:
        data = self.manifest()
        series = {
            "key": "tension-profile",
            "outcome": "tension",
            "arm": "test",
            "sheet": "Data ",
            "headerRange": "B1:D1",
            "valueRange": "B2:D4",
            "rowIdentityRange": "A2:A4",
            "axisSource": "ROW_IDENTITY",
            "axisLabel": "Position",
            "axisUnit": "mm",
            "valueUnit": "N",
            "verificationStatus": "NEEDS_REVIEW",
        }
        data["studies"][0]["measurementSeries"] = [series]
        checked_sheets: list[str] = []

        contract.validate_study_manifest(
            data,
            evidence_checker=lambda item: (
                checked_sheets.append(item["sheet"])
                if item.get("role") in {
                    "MEASUREMENT_HEADER",
                    "MEASUREMENT_VALUES",
                    "ROW_IDENTITY",
                }
                else None
            ),
        )

        self.assertEqual(["Data ", "Data ", "Data "], checked_sheets)

    def test_measurement_series_validates_aggregate_replicate_ranges(
        self,
    ) -> None:
        data = self.manifest()
        series = {
            "key": "tension-profile",
            "outcome": "tension",
            "arm": "test",
            "sheet": "Data",
            "headerRange": "B1:D1",
            "valueRange": "B2:D4",
            "rowIdentityRange": "A2:A4",
            "axisSource": "ROW_IDENTITY",
            "axisLabel": "Position",
            "axisUnit": "mm",
            "valueUnit": "N",
            "aggregateReplicateRanges": ["D1"],
        }
        data["studies"][0]["measurementSeries"] = [series]
        contract.validate_study_manifest(data)

        series["aggregateReplicateRanges"] = ["D1", "D1"]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "must not overlap",
        ):
            contract.validate_study_manifest(data)

        series["aggregateReplicateRanges"] = ["D2"]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "contained in and aligned",
        ):
            contract.validate_study_manifest(data)

        series["axisSource"] = "HEADER"
        series["aggregateReplicateRanges"] = ["A4"]
        contract.validate_study_manifest(data)

        series["rowIdentityRange"] = "A2:A2"
        series["valueRange"] = "B2:D2"
        series["aggregateReplicateRanges"] = ["A2"]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "at least one RAW replicate identity",
        ):
            contract.validate_study_manifest(data)

    def test_standalone_average_series_references_raw_series(self) -> None:
        data = self.manifest()
        study = data["studies"][0]
        study["arms"].append(
            {
                "key": "summary",
                "role": "OTHER",
                "label": "Combined summary",
                "evidence": self.evidence(),
                "factorValues": [],
            }
        )
        study["measurementSeries"] = [
            {
                "key": "before-profile",
                "outcome": "tension",
                "arm": "control",
                "sheet": "Data",
                "headerRange": "B1",
                "valueRange": "B2:B4",
                "rowIdentityRange": "A2:A4",
                "axisSource": "ROW_IDENTITY",
            },
            {
                "key": "after-profile",
                "seriesRole": "RAW",
                "outcome": "tension",
                "arm": "test",
                "sheet": "Data",
                "headerRange": "C1",
                "valueRange": "C2:C4",
                "rowIdentityRange": "A2:A4",
                "axisSource": "ROW_IDENTITY",
            },
            {
                "key": "average-profile",
                "seriesRole": "AGGREGATE",
                "aggregationFunction": "AVERAGE",
                "aggregateOfSeries": [
                    "before-profile",
                    "after-profile",
                ],
                "outcome": "tension",
                "arm": "summary",
                "sheet": "Data",
                "headerRange": "D1",
                "valueRange": "D2:D4",
                "rowIdentityRange": "A2:A4",
                "axisSource": "ROW_IDENTITY",
            },
        ]

        result = contract.validate_study_manifest(data)
        series = result["studies"][0]["measurementSeries"]
        self.assertEqual("RAW", series[0]["seriesRole"])
        self.assertEqual("RAW", series[1]["seriesRole"])
        self.assertEqual("AGGREGATE", series[2]["seriesRole"])
        self.assertEqual(
            ["before-profile", "after-profile"],
            series[2]["aggregateOfSeries"],
        )

    def test_frequency_replicate_matrix_and_average_preserve_row_axis(
        self,
    ) -> None:
        data = self.manifest()
        study = data["studies"][0]
        study["arms"].append(
            {
                "key": "summary",
                "role": "OTHER",
                "label": "Frequency profile average",
                "evidence": self.evidence(),
                "factorValues": [],
            }
        )
        study["measurementSeries"] = [
            {
                "key": "raw-frequency-profile",
                "seriesRole": "RAW",
                "aggregationFunction": "",
                "aggregateOfSeries": [],
                "aggregateReplicateRanges": [],
                "outcome": "tension",
                "arm": "control",
                "sheet": "Data",
                "headerRange": "F2:O2",
                "valueRange": "F3:O16",
                "rowIdentityRange": "A3:A16",
                "axisSource": "ROW_IDENTITY",
            },
            {
                "key": "average-frequency-profile",
                "seriesRole": "AGGREGATE",
                "aggregationFunction": "AVERAGE",
                "aggregateOfSeries": ["raw-frequency-profile"],
                "aggregateReplicateRanges": [],
                "outcome": "tension",
                "arm": "summary",
                "sheet": "Data",
                "headerRange": "P2",
                "valueRange": "P3:P16",
                "rowIdentityRange": "A3:A16",
                "axisSource": "ROW_IDENTITY",
            },
        ]

        result = contract.validate_study_manifest(data)
        raw, average = result["studies"][0]["measurementSeries"]
        self.assertEqual("ROW_IDENTITY", raw["axisSource"])
        self.assertEqual("A3:A16", average["rowIdentityRange"])
        self.assertEqual(
            ["raw-frequency-profile"],
            average["aggregateOfSeries"],
        )

    def test_standalone_average_series_rejects_invalid_references(
        self,
    ) -> None:
        data = self.manifest()
        study = data["studies"][0]
        raw = {
            "key": "raw-profile",
            "outcome": "tension",
            "arm": "control",
            "sheet": "Data",
            "headerRange": "B1",
            "valueRange": "B2:B4",
            "rowIdentityRange": "A2:A4",
            "axisSource": "ROW_IDENTITY",
        }
        aggregate = {
            "key": "average-profile",
            "seriesRole": "AGGREGATE",
            "aggregationFunction": "AVERAGE",
            "aggregateOfSeries": ["raw-profile"],
            "outcome": "tension",
            "arm": "test",
            "sheet": "Data",
            "headerRange": "D1",
            "valueRange": "D2:D4",
            "rowIdentityRange": "A2:A4",
            "axisSource": "ROW_IDENTITY",
        }

        unknown = copy.deepcopy(data)
        unknown["studies"][0]["measurementSeries"] = [
            copy.deepcopy(raw),
            {
                **copy.deepcopy(aggregate),
                "aggregateOfSeries": ["missing-profile"],
            },
        ]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "references unknown measurementSeries",
        ):
            contract.validate_study_manifest(unknown)

        nested = copy.deepcopy(data)
        nested["studies"][0]["measurementSeries"] = [
            copy.deepcopy(raw),
            copy.deepcopy(aggregate),
            {
                **copy.deepcopy(aggregate),
                "key": "outer-average",
                "aggregateOfSeries": ["average-profile"],
            },
        ]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "nested AGGREGATE",
        ):
            contract.validate_study_manifest(nested)

        outcome_mismatch = copy.deepcopy(data)
        other_outcome = copy.deepcopy(
            outcome_mismatch["studies"][0]["outcomes"][0]
        )
        other_outcome["key"] = "other-tension"
        other_outcome["observations"] = []
        outcome_mismatch["studies"][0]["outcomes"].append(other_outcome)
        mismatched_aggregate = copy.deepcopy(aggregate)
        mismatched_aggregate["outcome"] = "other-tension"
        outcome_mismatch["studies"][0]["measurementSeries"] = [
            copy.deepcopy(raw),
            mismatched_aggregate,
        ]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "outcome mismatch",
        ):
            contract.validate_study_manifest(outcome_mismatch)

        unsupported = copy.deepcopy(data)
        unsupported_aggregate = copy.deepcopy(aggregate)
        unsupported_aggregate["aggregationFunction"] = "SUM"
        unsupported["studies"][0]["measurementSeries"] = [
            copy.deepcopy(raw),
            unsupported_aggregate,
        ]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "must be AVERAGE",
        ):
            contract.validate_study_manifest(unsupported)

    def test_measurement_series_rejects_dimension_mismatch(self) -> None:
        data = self.manifest()
        data["studies"][0]["measurementSeries"] = [
            {
                "key": "tension-profile",
                "outcome": "tension",
                "arm": "test",
                "sheet": "Data",
                "headerRange": "B1:C1",
                "valueRange": "B2:D4",
                "rowIdentityRange": "A2:A4",
                "axisSource": "ROW_IDENTITY",
                "axisLabel": "Position",
                "axisUnit": "mm",
                "valueUnit": "N",
            }
        ]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "same column count",
        ):
            contract.validate_study_manifest(data)

    def test_measurement_series_rejects_unknown_arm_or_outcome(self) -> None:
        data = self.manifest()
        data["studies"][0]["measurementSeries"] = [
            {
                "key": "tension-profile",
                "outcome": "missing-outcome",
                "arm": "test",
                "sheet": "Data",
                "headerRange": "B1:D1",
                "valueRange": "B2:D4",
                "rowIdentityRange": "A2:A4",
                "axisSource": "ROW_IDENTITY",
                "axisLabel": "",
                "axisUnit": "",
                "valueUnit": "",
            }
        ]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "unknown outcome",
        ):
            contract.validate_study_manifest(data)

    def test_measurement_series_requires_axis_source_and_exact_alignment(
        self,
    ) -> None:
        data = self.manifest()
        series = {
            "key": "tension-profile",
            "outcome": "tension",
            "arm": "test",
            "sheet": "Data",
            "headerRange": "B1:D1",
            "valueRange": "B2:D4",
            "rowIdentityRange": "A2:A4",
            "axisSource": "HEADER",
            "axisLabel": "Frequency",
            "axisUnit": "Hz",
            "valueUnit": "N",
        }
        data["studies"][0]["measurementSeries"] = [series]
        result = contract.validate_study_manifest(data)
        self.assertEqual(
            "HEADER",
            result["studies"][0]["measurementSeries"][0]["axisSource"],
        )
        self.assertEqual(
            "NEEDS_REVIEW",
            result["studies"][0]["measurementSeries"][0][
                "verificationStatus"
            ],
        )

        del series["axisSource"]
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "axisSource",
        ):
            contract.validate_study_manifest(data)

        series["axisSource"] = "ROW_IDENTITY"
        series["headerRange"] = "C1:E1"
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "exact valueRange columns",
        ):
            contract.validate_study_manifest(data)

        series["headerRange"] = "B1:D1"
        series["rowIdentityRange"] = "A3:A5"
        with self.assertRaisesRegex(
            contract.StudyContractError,
            "exact valueRange rows",
        ):
            contract.validate_study_manifest(data)


if __name__ == "__main__":
    unittest.main()
