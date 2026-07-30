from __future__ import annotations

import copy
import unittest

import inference_data_ai_answer as answer


class DeterministicEvidenceAnswerTests(unittest.TestCase):
    def source(self, suffix: str) -> dict:
        return {
            "revisionId": int(suffix),
            "revisionUid": f"revision-{suffix}",
            "dataset": "Fixture",
            "sourcePath": rf"D:\input\study-{suffix}.xlsx",
            "fileName": f"study-{suffix}.xlsx",
            "contentSha256": suffix * 64,
            "sourceFingerprint": suffix * 64,
            "captureContract": "openxml-v2",
            "captureStatus": "CAPTURED",
            "isCurrent": True,
        }

    def candidate(
        self,
        suffix: str,
        *,
        model: str = "Model-A",
    ) -> dict:
        data_id = f"DATA-{suffix}"
        return {
            "publicDataId": data_id,
            "studyUid": f"study-{suffix}",
            "source": self.source(suffix),
            "study": {
                "studyId": int(suffix),
                "title": f"Novel study {suffix}",
                "verificationStatus": "VERIFIED",
                "comparabilityStatus": "VALID",
                "confoundingStatus": "NONE",
            },
            "analysis": {
                "verificationStatus": "VERIFIED",
            },
            "contexts": [
                {
                    "kind": "MODEL",
                    "originalValue": model,
                    "normalizedValue": model.casefold(),
                    "valueNumber": None,
                    "unit": None,
                    "startValue": "",
                    "endValue": "",
                    "concept": None,
                }
            ],
            "factors": [
                {
                    "originalLabel": "Cryogenic dwell",
                    "baselineCondition": "2 s",
                    "changedCondition": "5 s",
                    "changeDirection": "INCREASE",
                    "isolationStatus": "ISOLATED",
                    "concept": {
                        "conceptId": 10,
                        "canonicalName": "Cryogenic dwell",
                    },
                }
            ],
            "arms": [],
            "outcomes": [],
            "relevance": {"score": 1},
            "rank": int(suffix),
        }

    def citation(
        self,
        suffix: str,
        *,
        entity_type: str,
        entity_uid: str,
    ) -> dict:
        source = self.source(suffix)
        return {
            "evidenceId": int(suffix),
            "evidenceUid": f"evidence-{suffix}",
            "publicEvidenceId": f"EVD-{suffix}",
            "kind": "CELL_RANGE",
            "sourcePath": source["sourcePath"],
            "fileName": source["fileName"],
            "revisionId": source["revisionId"],
            "revisionUid": source["revisionUid"],
            "contentSha256": source["contentSha256"],
            "sheet": "Data",
            "range": "B2:C3",
            "start": {"row": 2, "column": 2},
            "end": {"row": 3, "column": 3},
            "role": "RESULT",
            "linkRole": "RESULT",
            "claimScope": "",
            "sourceText": "",
            "note": "",
            "verificationStatus": "VERIFIED",
            "linkedEntities": [
                {
                    "entityType": entity_type,
                    "entityUid": entity_uid,
                }
            ],
        }

    def eligible_effect(
        self,
        suffix: str,
        estimate: float,
    ) -> dict:
        effect_uid = f"effect-{suffix}"
        citation = self.citation(
            suffix,
            entity_type="EFFECT",
            entity_uid=effect_uid,
        )
        factor_values_control = [
            {
                "factorUid": "factor-dwell",
                "factorLabel": "Cryogenic dwell",
                "originalValue": "2 s",
                "valueNumber": 2,
                "unit": "s",
                "isBaseline": True,
                "heldConstant": False,
            }
        ]
        factor_values_changed = [
            {
                **factor_values_control[0],
                "originalValue": "5 s",
                "valueNumber": 5,
                "isBaseline": False,
            }
        ]
        return {
            "publicDataId": f"DATA-{suffix}",
            "publicComparisonId": f"CMP-{suffix}",
            "publicEffectId": f"EFF-{suffix}",
            "publicEvidenceIds": [f"EVD-{suffix}"],
            "sourcePath": self.source(suffix)["sourcePath"],
            "source": self.source(suffix),
            "analysis": {"verificationStatus": "VERIFIED"},
            "study": {
                "title": f"Novel study {suffix}",
                "summary": "",
                "verificationStatus": "VERIFIED",
                "comparabilityStatus": "VALID",
                "confoundingStatus": "NONE",
            },
            "comparison": {
                "comparisonKey": "changed-vs-control",
                "designType": "CONTROL_COMPARISON",
                "matchingBasis": "same lot and sample basis",
                "validityStatus": "VALID",
                "confoundingStatus": "NONE",
                "verificationStatus": "VERIFIED",
                "aggregationEligible": True,
                "direction": "INCREASE" if estimate > 0 else "DECREASE",
                "summary": "",
                "exclusionReason": "",
                "factorDifferences": [
                    {
                        "factorUid": "factor-dwell",
                        "factorLabel": "Cryogenic dwell",
                        "controlValue": "2 s",
                        "comparedValue": "5 s",
                        "controlValueRecorded": True,
                        "comparedValueRecorded": True,
                    }
                ],
                "controlArm": {
                    "role": "CONTROL",
                    "label": "2 s",
                    "condition": "2 s",
                    "sampleBasis": "units",
                    "matchingBasis": "same lot",
                    "factorValues": factor_values_control,
                },
                "comparedArm": {
                    "role": "TREATMENT",
                    "label": "5 s",
                    "condition": "5 s",
                    "sampleBasis": "units",
                    "matchingBasis": "same lot",
                    "factorValues": factor_values_changed,
                },
            },
            "outcome": {
                "outcomeUid": "outcome-fracture",
                "originalLabel": "Nebula fracture rate",
                "metricType": "RATE",
                "unit": "%p",
                "denominatorBasis": "tested units",
                "concept": {
                    "conceptId": 20,
                    "canonicalName": "Nebula fracture rate",
                },
            },
            "observations": {
                "comparedArm": [
                    {
                        "observationUid": f"changed-{suffix}",
                        "stratumKey": "",
                        "replicateKey": "",
                    }
                ],
                "controlArm": [
                    {
                        "observationUid": f"control-{suffix}",
                        "stratumKey": "",
                        "replicateKey": "",
                    }
                ],
            },
            "effect": {
                "effectType": "ABSOLUTE_DIFFERENCE",
                "estimate": estimate,
                "unit": "%p",
                "formulaVersion": "canonical-v1",
                "direction": "INCREASE" if estimate > 0 else "DECREASE",
                "aggregationEligible": True,
                "verificationStatus": "VERIFIED",
                "details": {},
            },
            "evidence": [citation],
        }

    def pack(
        self,
        candidates: list[dict],
        eligible: list[dict],
        excluded: list[dict] | None = None,
    ) -> dict:
        excluded = excluded or []
        return {
            "schemaVersion": "canonical-evidence-pack-v1",
            "question": "Cryogenic dwell이 Nebula fracture에 관련 있나?",
            "normalizedQuestion": "",
            "queryTokens": ["cryogenic", "dwell", "nebula", "fracture"],
            "searchTokens": ["cryogenic", "dwell", "nebula", "fracture"],
            "queryRoleHints": {
                "outcomeTerms": ["nebula", "fracture"],
                "contextOrFactorTerms": ["cryogenic", "dwell"],
                "relationGateApplied": True,
            },
            "studyCandidates": candidates,
            "answerEligibleEffects": eligible,
            "excludedCandidates": excluded,
            "eligibleEffectSummary": [],
            "summary": {
                "relevantStudyCount": len(candidates),
                "answerEligibleEffectCount": len(eligible),
                "excludedCandidateCount": len(excluded),
            },
        }

    def test_single_verified_effect_renders_exact_ids_and_value(self) -> None:
        pack = self.pack(
            [self.candidate("1")],
            [self.eligible_effect("1", 5)],
        )
        result = answer.build_evidence_answer(pack)

        self.assertEqual("SUPPORTED", result["answerStatus"])
        self.assertEqual(1, len(result["quantitativeGroups"]))
        effect = result["quantitativeGroups"][0]["effects"][0]
        self.assertEqual(5, effect["estimate"])
        self.assertEqual(["EVD-1"], effect["evidenceIds"])
        self.assertEqual(
            [
                {
                    "factorUid": "factor-dwell",
                    "factorLabel": "Cryogenic dwell",
                    "controlValue": "2 s",
                    "comparedValue": "5 s",
                    "controlValueRecorded": True,
                    "comparedValueRecorded": True,
                }
            ],
            effect["factorDifferences"],
        )
        self.assertEqual("2 s", effect["controlCondition"])
        self.assertEqual("5 s", effect["comparedCondition"])
        self.assertIn(
            "Cryogenic dwell: 2 s → 5 s 조건에서",
            result["renderedAnswer"]["textKo"],
        )
        self.assertIn("DATA-1 / CMP-1 / EFF-1 / EVD-1", result["renderedAnswer"]["textKo"])
        self.assertEqual(
            ["DATA-1"],
            result["coverage"]["representedDataIds"],
        )

    def test_exact_compatibility_groups_do_not_mix_different_contexts(self) -> None:
        pack = self.pack(
            [
                self.candidate("1", model="Model-A"),
                self.candidate("2", model="Model-B"),
            ],
            [
                self.eligible_effect("1", 5),
                self.eligible_effect("2", 11),
            ],
        )
        result = answer.build_evidence_answer(pack)

        self.assertEqual(2, len(result["quantitativeGroups"]))
        self.assertEqual(
            {5, 11},
            {
                group["statistics"]["mean"]
                for group in result["quantitativeGroups"]
            },
        )
        self.assertIn(
            "INCOMPATIBLE_GROUPS",
            {item["code"] for item in result["limitations"]},
        )

    def test_same_signature_uses_decimal_statistics_and_detects_conflict(self) -> None:
        pack = self.pack(
            [self.candidate("1"), self.candidate("2")],
            [
                self.eligible_effect("1", 5),
                self.eligible_effect("2", -1),
            ],
        )
        result = answer.build_evidence_answer(pack)

        self.assertEqual("CONFLICTING", result["answerStatus"])
        group = result["quantitativeGroups"][0]
        self.assertEqual("CONFLICTING", group["directionStatus"])
        self.assertEqual(2, group["statistics"]["mean"])
        self.assertIn(
            "CONFLICTING_DIRECTION",
            {item["code"] for item in result["limitations"]},
        )
        self.assertNotIn("산술평균 2", result["renderedAnswer"]["textKo"])

    def test_descriptive_observation_is_shown_without_calculating_difference(self) -> None:
        candidate = self.candidate("1")
        observation_uid = "observation-1"
        citation = self.citation(
            "1",
            entity_type="OUTCOME",
            entity_uid="outcome-study-scope",
        )
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": None,
            "publicEffectId": None,
            "publicEvidenceIds": ["EVD-1"],
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": None,
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveOutcomes": [
                {
                    "outcome": {
                        "outcomeUid": "outcome-study-scope",
                        "originalLabel": "Nebula fracture rate",
                        "unit": "%",
                    },
                    "armObservations": [
                        {
                            "arm": {
                                "label": "5 s dwell",
                                "condition": "5 s",
                            },
                            "observations": [
                                {
                                    "observationUid": observation_uid,
                                    "valueNumber": 12.5,
                                    "valueText": "12.5%",
                                    "numerator": None,
                                    "denominator": None,
                                    "ratePpm": None,
                                    "min": None,
                                    "max": None,
                                    "average": None,
                                    "sampleSize": None,
                                    "verificationStatus": "NEEDS_REVIEW",
                                }
                            ],
                        }
                    ],
                }
            ],
            "evidence": [citation],
            "exclusionReasons": [
                {
                    "code": "NO_COMPARISON_RECORD",
                    "message": "No comparison.",
                }
            ],
        }
        result = answer.build_evidence_answer(
            self.pack([candidate], [], [excluded])
        )

        self.assertEqual("INSUFFICIENT_COMPARISON", result["answerStatus"])
        self.assertEqual(1, len(result["descriptiveStudies"]))
        self.assertIn("12.5%", result["renderedAnswer"]["textKo"])
        self.assertNotIn("대비 12.5", result["renderedAnswer"]["textKo"])
        self.assertEqual(["EVD-1"], [item["evidenceId"] for item in result["citations"]])

    def test_study_scoped_descriptive_outcomes_are_not_repeated_per_comparison(
        self,
    ) -> None:
        candidate = self.candidate("1")
        observation_uid = "observation-study-scope"
        citation = self.citation(
            "1",
            entity_type="OBSERVATION",
            entity_uid=observation_uid,
        )
        base = {
            "publicDataId": "DATA-1",
            "publicComparisonId": "CMP-1",
            "publicEffectId": None,
            "publicEvidenceIds": ["EVD-1"],
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": {},
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveScope": "STUDY",
            "descriptiveOutcomes": [
                {
                    "outcome": {
                        "originalLabel": "Nebula fracture rate",
                        "unit": "%",
                    },
                    "armObservations": [
                        {
                            "arm": {
                                "label": "Test",
                                "condition": "",
                            },
                            "observations": [
                                {
                                    "observationUid": observation_uid,
                                    "valueNumber": 4.2,
                                    "valueText": "4.2%",
                                    "numerator": None,
                                    "denominator": None,
                                    "ratePpm": None,
                                    "min": None,
                                    "max": None,
                                    "average": None,
                                    "sampleSize": None,
                                    "verificationStatus": "NEEDS_REVIEW",
                                }
                            ],
                        }
                    ],
                }
            ],
            "evidence": [citation],
            "exclusionReasons": [
                {"code": "NO_EFFECT_RECORD", "message": "No effect."}
            ],
        }
        duplicate = copy.deepcopy(base)
        duplicate["publicComparisonId"] = "CMP-2"

        result = answer.build_evidence_answer(
            self.pack([candidate], [], [base, duplicate])
        )

        self.assertEqual(1, len(result["descriptiveStudies"]))
        self.assertEqual(
            "4.2%",
            result["descriptiveStudies"][0]["outcomes"][0]["arms"][0][
                "observations"
            ][0]["valueText"],
        )

    def test_measurement_series_summary_is_descriptive_only_and_cited(
        self,
    ) -> None:
        candidate = self.candidate("1")
        series_uid = "series-1"
        citation = self.citation(
            "1",
            entity_type="MEASUREMENT_SERIES",
            entity_uid=series_uid,
        )
        citation["role"] = "MEASUREMENT_VALUES"
        citation["linkRole"] = "MEASUREMENT_VALUES"
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": "CMP-1",
            "publicEffectId": None,
            "publicEvidenceIds": ["EVD-1"],
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": {
                "validityStatus": "NEEDS_REVIEW",
                "confoundingStatus": "UNASSESSED",
                "verificationStatus": "NEEDS_REVIEW",
                "aggregationEligible": False,
            },
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveMeasurementSeries": [
                {
                    "seriesUid": series_uid,
                    "publicSeriesId": "SER-1",
                    "seriesKey": "spectral-sweep",
                    "outcome": {
                        "outcomeKey": "spectrum",
                        "originalLabel": "Nebula fracture spectrum",
                        "domain": "Acoustic",
                        "metricType": "PROFILE",
                    },
                    "arm": {
                        "label": "1800 V",
                        "condition": "Drive 1800 V",
                    },
                    "axisLabel": "Frequency",
                    "axisSource": "ROW_IDENTITY",
                    "axisUnit": "Hz",
                    "valueUnit": "dBSPL",
                    "stratumKey": "left channel",
                    "replicateKeys": ["sample-a", "sample-b"],
                    "aggregateReplicateKeys": ["AVG"],
                    "verificationStatus": "NEEDS_REVIEW",
                    "interpretationStatus": "DESCRIPTIVE_ONLY",
                    "pointSummary": {
                        "pointCount": 5,
                        "rawPointCount": 4,
                        "aggregatePointCount": 1,
                        "minimum": 10,
                        "maximum": 40,
                        "average": 25,
                        "distinctAxisCount": 2,
                        "distinctReplicateCount": 2,
                        "aggregateReplicateCount": 1,
                    },
                }
            ],
            "evidence": [citation],
            "exclusionReasons": [
                {
                    "code": "NO_EFFECT_RECORD",
                    "message": "No approved comparison effect.",
                }
            ],
        }
        result = answer.build_evidence_answer(
            self.pack([candidate], [], [excluded])
        )

        self.assertEqual(
            "INSUFFICIENT_COMPARISON",
            result["answerStatus"],
        )
        self.assertEqual([], result["quantitativeGroups"])
        series = result["descriptiveStudies"][0][
            "measurementSeries"
        ][0]
        self.assertEqual("DESCRIPTIVE_ONLY", series["interpretationStatus"])
        self.assertEqual(25, series["pointSummary"]["average"])
        self.assertEqual(["EVD-1"], series["evidenceIds"])
        self.assertEqual(
            ["DATA-1"],
            result["coverage"]["representedDataIds"],
        )
        self.assertEqual(
            0,
            result["coverage"][
                "uncitedDescriptiveMeasurementSeriesCount"
            ],
        )
        rendered = result["renderedAnswer"]["textKo"]
        self.assertIn("SER-1", rendered)
        self.assertIn("range 10~40 dBSPL", rendered)
        self.assertIn("descriptive average 25 dBSPL", rendered)
        self.assertIn("5 points (4 raw + 1 aggregate)", rendered)
        self.assertIn("raw replicates 2, aggregate columns/rows 1", rendered)
        self.assertNotIn("ABSOLUTE_DIFFERENCE", rendered)
        self.assertEqual(
            ["EVD-1"],
            [item["evidenceId"] for item in result["citations"]],
        )

    def test_unmatched_measurement_series_is_not_rendered_for_relation_query(
        self,
    ) -> None:
        candidate = self.candidate("1")
        citation = self.citation(
            "1",
            entity_type="MEASUREMENT_SERIES",
            entity_uid="series-acoustic",
        )
        citation["role"] = "MEASUREMENT_VALUES"
        citation["linkRole"] = "MEASUREMENT_VALUES"
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": "CMP-1",
            "publicEffectId": None,
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": {
                "validityStatus": "NEEDS_REVIEW",
                "confoundingStatus": "UNASSESSED",
                "verificationStatus": "NEEDS_REVIEW",
                "aggregationEligible": False,
            },
            "descriptiveMeasurementSeries": [
                {
                    "seriesUid": "series-acoustic",
                    "outcome": {
                        "outcomeKey": "spl",
                        "originalLabel": "SPL acoustic profile",
                        "domain": "Acoustic",
                        "metricType": "PROFILE",
                    },
                }
            ],
            "evidence": [citation],
            "exclusionReasons": [
                {
                    "code": "NO_EFFECT_RECORD",
                    "message": "No approved comparison effect.",
                }
            ],
        }

        result = answer.build_evidence_answer(
            self.pack([candidate], [], [excluded])
        )

        self.assertEqual([], result["descriptiveStudies"])
        self.assertEqual([], result["citations"])
        self.assertNotIn(
            "SPL acoustic profile",
            result["renderedAnswer"]["textKo"],
        )

    def test_confounded_comparison_explains_multi_factor_exclusion(self) -> None:
        candidate = self.candidate("1")
        comparison_uid = "comparison-1"
        citation = self.citation(
            "1",
            entity_type="COMPARISON",
            entity_uid=comparison_uid,
        )
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": "CMP-1",
            "publicEffectId": None,
            "publicEvidenceIds": ["EVD-1"],
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": {
                "comparisonUid": comparison_uid,
                "designType": "TEST_VS_NORMAL",
                "matchingBasis": (
                    "Mold and Normal processing conditions differ or "
                    "are unstated"
                ),
                "validityStatus": "NEEDS_REVIEW",
                "confoundingStatus": "CONFOUNDED",
                "verificationStatus": "NEEDS_REVIEW",
                "aggregationEligible": False,
                "summary": "",
                "exclusionReason": "",
                "factorDifferences": [
                    {
                        "factorLabel": "VP mold temperature",
                        "comparedValue": "190 C",
                        "controlValue": "",
                        "comparedValueRecorded": True,
                        "controlValueRecorded": False,
                    },
                    {
                        "factorLabel": "VP mold number",
                        "comparedValue": "#8",
                        "controlValue": "#10",
                        "comparedValueRecorded": True,
                        "controlValueRecorded": True,
                    },
                    {
                        "factorLabel": "Vulcanizing agent",
                        "comparedValue": "10%",
                        "controlValue": "",
                        "comparedValueRecorded": True,
                        "controlValueRecorded": False,
                    },
                ],
                "comparedArm": {
                    "label": "Test mold #8",
                    "condition": "190 C, Vulcanizing agent 10%",
                },
                "controlArm": {
                    "label": "Normal mold #10",
                    "condition": "Normal condition not stated",
                },
            },
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveMeasurementSeries": [],
            "evidence": [citation],
            "exclusionReasons": [
                {
                    "code": "COMPARISON_CONFOUNDED",
                    "message": "Comparison is confounded.",
                },
                {
                    "code": "NO_EFFECT_RECORD",
                    "message": "No effect was calculated.",
                },
            ],
        }

        result = answer.build_evidence_answer(
            self.pack([candidate], [], [excluded])
        )

        record = result["excludedRecords"][0]
        self.assertIn("CONFOUNDED_MULTI_FACTOR", record["reasonCodes"])
        self.assertEqual(
            "CONFOUNDED_MULTI_FACTOR",
            record["comparisonAssessment"]["code"],
        )
        self.assertEqual(["EVD-1"], record["evidenceIds"])
        self.assertIn("교란 비교 1건", result["directAnswer"]["textKo"])
        self.assertIn(
            "Vulcanizing agent",
            result["renderedAnswer"]["textKo"],
        )
        self.assertIn(
            "Mold and Normal processing conditions differ or are unstated",
            result["renderedAnswer"]["textKo"],
        )
        limitation = next(
            item
            for item in result["limitations"]
            if item["code"] == "CONFOUNDED_MULTI_FACTOR"
        )
        self.assertEqual(["CMP-1"], limitation["relatedIds"])
        self.assertEqual(
            ["EVD-1"],
            [item["evidenceId"] for item in result["citations"]],
        )

    def test_uncited_legacy_descriptive_value_is_withheld_without_failing_answer(self) -> None:
        candidate = self.candidate("1")
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": None,
            "publicEffectId": None,
            "publicEvidenceIds": [],
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": None,
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveOutcomes": [
                {
                    "outcome": {
                        "originalLabel": "Nebula fracture legacy rate",
                        "unit": "%",
                    },
                    "armObservations": [
                        {
                            "arm": {
                                "label": "Legacy condition",
                                "condition": "",
                            },
                            "observations": [
                                {
                                    "observationUid": "legacy-observation-without-evd",
                                    "valueNumber": 9.9,
                                    "valueText": "9.9%",
                                    "numerator": None,
                                    "denominator": None,
                                    "ratePpm": None,
                                    "min": None,
                                    "max": None,
                                    "average": None,
                                    "sampleSize": None,
                                    "verificationStatus": "NEEDS_REVIEW",
                                }
                            ],
                        }
                    ],
                }
            ],
            "evidence": [],
            "exclusionReasons": [
                {
                    "code": "NO_DIRECT_OBSERVATION_EVIDENCE",
                    "message": "Legacy observation has no exact cell link.",
                }
            ],
        }
        pack = self.pack([candidate], [], [excluded])

        result = answer.build_evidence_answer(pack)

        self.assertEqual("INSUFFICIENT_COMPARISON", result["answerStatus"])
        self.assertEqual([], result["descriptiveStudies"])
        self.assertEqual([], result["citations"])
        self.assertEqual(
            1,
            result["coverage"]["uncitedDescriptiveObservationCount"],
        )
        self.assertNotIn("9.9%", result["renderedAnswer"]["textKo"])
        self.assertIn(
            "DESCRIPTIVE_EVIDENCE_MISSING",
            {item["code"] for item in result["limitations"]},
        )
        self.assertEqual(
            result,
            answer.validate_evidence_answer(result, pack),
        )

    def test_no_relevant_data_and_tampering_are_detected(self) -> None:
        pack = self.pack([], [])
        result = answer.build_evidence_answer(pack)
        self.assertEqual("NO_RELEVANT_DATA", result["answerStatus"])
        self.assertEqual(result, answer.validate_evidence_answer(result, pack))

        tampered = copy.deepcopy(result)
        tampered["directAnswer"]["textKo"] = "근거 없이 영향이 확정됨"
        with self.assertRaises(answer.EvidenceAnswerError):
            answer.validate_evidence_answer(tampered, pack)

    def test_terminal_source_exclusion_is_an_answered_record_without_values(self) -> None:
        pack = self.pack([], [])
        pack["sourceExclusions"] = [
            {
                "publicAnalysisId": "ANALYSIS-TERMINAL",
                "analysisUid": "analysis-terminal",
                "revisionUid": "revision-terminal",
                "sourcePath": "/fixture/Zephyr XRAY image.xlsx",
                "fileName": "Zephyr XRAY image.xlsx",
                "contentSha256": "f" * 64,
                "sourceFingerprint": "fingerprint-terminal",
                "captureContract": "openxml-v2",
                "sourceContentStatus": "NO_TABULAR_EVIDENCE",
                "analysisStatus": "NO_TABULAR_EVIDENCE",
                "verificationStatus": "EXCLUDED",
                "summary": "No reviewable table.",
                "limitations": ["Images are outside scope."],
                "exclusionReasons": [
                    {
                        "code": "NO_TABULAR_EVIDENCE",
                        "message": "No canonical Study was created.",
                    },
                    {
                        "code": "IMAGES_NOT_ANALYZED",
                        "message": "Images are outside scope.",
                    },
                ],
                "imagesAnalyzed": False,
            }
        ]

        result = answer.build_evidence_answer(pack)

        self.assertEqual("INSUFFICIENT_COMPARISON", result["answerStatus"])
        self.assertEqual(0, result["coverage"]["relevantStudyCount"])
        self.assertEqual(1, result["coverage"]["relevantRecordCount"])
        self.assertEqual(1, result["coverage"]["relevantSourceExclusionCount"])
        self.assertEqual([], result["citations"])
        self.assertEqual(
            "NO_TABULAR_EVIDENCE",
            result["excludedSources"][0]["sourceContentStatus"],
        )
        self.assertIn(
            "ANALYSIS-TERMINAL",
            result["renderedAnswer"]["textKo"],
        )

    def test_descriptive_answer_filters_to_matching_question_outcomes(self) -> None:
        candidate = self.candidate("1")
        observation_uid = "observation-nebula"
        citation = self.citation(
            "1",
            entity_type="OBSERVATION",
            entity_uid=observation_uid,
        )
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": None,
            "publicEffectId": None,
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": None,
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveOutcomes": [
                {
                    "outcome": {
                        "outcomeKey": "input",
                        "originalLabel": "Input",
                        "domain": "",
                        "metricType": "count",
                        "definition": "",
                        "unit": "",
                    },
                    "armObservations": [],
                },
                {
                    "outcome": {
                        "outcomeKey": "nebula-fracture",
                        "originalLabel": "Nebula fracture",
                        "domain": "",
                        "metricType": "rate",
                        "definition": "",
                        "unit": "%",
                    },
                    "armObservations": [
                        {
                            "arm": {"label": "Test", "condition": ""},
                            "observations": [
                                {
                                    "observationUid": observation_uid,
                                    "valueNumber": 12.5,
                                    "valueText": "",
                                    "numerator": None,
                                    "denominator": None,
                                    "ratePpm": None,
                                    "min": None,
                                    "max": None,
                                    "average": None,
                                    "sampleSize": None,
                                    "verificationStatus": "NEEDS_REVIEW",
                                }
                            ],
                        }
                    ],
                },
            ],
            "evidence": [citation],
            "exclusionReasons": [
                {"code": "NO_COMPARISON_RECORD", "message": "No comparison."}
            ],
        }
        pack = self.pack([candidate], [], [excluded])
        result = answer.build_evidence_answer(pack)
        labels = [
            outcome["label"]
            for study in result["descriptiveStudies"]
            for outcome in study["outcomes"]
        ]
        self.assertEqual(["Nebula fracture"], labels)

    def test_title_outcome_proxy_keeps_source_backed_detailed_submetrics(
        self,
    ) -> None:
        candidate = self.candidate("1")
        candidate["relevance"]["matchedFields"] = [
            {
                "field": "study.title",
                "terms": ["function"],
                "score": 4,
            }
        ]
        observation_uid = "observation-total-ng"
        citation = self.citation(
            "1",
            entity_type="OBSERVATION",
            entity_uid=observation_uid,
        )
        excluded = {
            "publicDataId": "DATA-1",
            "publicComparisonId": "CMP-1",
            "publicEffectId": None,
            "sourcePath": self.source("1")["sourcePath"],
            "comparison": {},
            "outcome": None,
            "observations": {"comparedArm": [], "controlArm": []},
            "effect": None,
            "descriptiveScope": "STUDY",
            "descriptiveOutcomes": [
                {
                    "outcome": {
                        "outcomeUid": "outcome-total-ng",
                        "outcomeKey": "total-ng",
                        "originalLabel": "Total NG rate",
                        "domain": "",
                        "metricType": "rate",
                        "definition": "",
                        "unit": "%",
                    },
                    "armObservations": [
                        {
                            "arm": {"label": "Test", "condition": ""},
                            "observations": [
                                {
                                    "observationUid": observation_uid,
                                    "valueNumber": 3.1,
                                    "valueText": "3.1%",
                                    "numerator": None,
                                    "denominator": None,
                                    "ratePpm": None,
                                    "min": None,
                                    "max": None,
                                    "average": None,
                                    "sampleSize": None,
                                    "verificationStatus": "NEEDS_REVIEW",
                                }
                            ],
                        }
                    ],
                }
            ],
            "evidence": [citation],
            "exclusionReasons": [
                {"code": "NO_EFFECT_RECORD", "message": "No effect."}
            ],
        }
        pack = self.pack([candidate], [], [excluded])
        pack["question"] = "Function 결과를 조건별로 보여줘"
        pack["queryRoleHints"] = {
            "outcomeTerms": ["function"],
            "contextOrFactorTerms": ["조건"],
            "relationGateApplied": True,
        }

        result = answer.build_evidence_answer(pack)

        self.assertEqual(
            ["Total NG rate"],
            [
                outcome["label"]
                for study in result["descriptiveStudies"]
                for outcome in study["outcomes"]
            ],
        )
        self.assertIn("3.1%", result["renderedAnswer"]["textKo"])

    def test_eligible_effect_requires_direct_current_revision_evidence(self) -> None:
        effect = self.eligible_effect("1", 5)
        effect["evidence"][0]["linkedEntities"][0]["entityType"] = "STUDY"
        with self.assertRaisesRegex(
            answer.EvidenceAnswerError,
            "direct verified",
        ):
            answer.build_evidence_answer(
                self.pack([self.candidate("1")], [effect])
            )


if __name__ == "__main__":
    unittest.main()
