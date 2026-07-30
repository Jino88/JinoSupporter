from __future__ import annotations

import hashlib
import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import inference_data_ai_semantic_ai as semantic_ai


class SemanticAiTests(unittest.TestCase):
    def source(self) -> dict:
        return {
            "revisionUid": "capture_revision_generic",
            "contentSha256": "a" * 64,
            "sourcePath": r"D:\input\generic.xlsx",
        }

    def chunk(self) -> dict:
        return {
            "chunkId": "sheet-1-section-1-chunk-1",
            "sheet": "Unfamiliar data",
            "range": "B3:H20",
            "cells": [
                {"coordinate": "B3", "displayValue": "Cryogenic dwell"},
                {"coordinate": "C3", "displayValue": "9.5 s"},
                {"coordinate": "B8", "displayValue": "Unseen response"},
            ],
        }

    def test_exact_study_prompt_hash_is_enforced_before_execution(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(
                semantic_ai.SemanticAiError,
                "expected hash",
            ):
                semantic_ai.run_codex_study_draft(
                    source=self.source(),
                    workbook={"semanticCellCoverageComplete": True},
                    locator_results=[self.result()],
                    focused_chunks=[self.chunk()],
                    content_complete=True,
                    output_path=Path(temp_dir) / "study.json",
                    exact_prompt_text="exact prompt",
                    expected_prompt_sha256="wrong",
                    run_command=mock.Mock(
                        side_effect=AssertionError(
                            "hash mismatch must fail before execution"
                        )
                    ),
                )

    def test_aggregate_replicate_value_ranges_map_to_identity_axis(
        self,
    ) -> None:
        draft = {
            "studies": [
                {
                    "measurementSeries": [
                        {
                            "axisSource": "HEADER",
                            "headerRange": "I14:P14",
                            "valueRange": "I17:P25",
                            "rowIdentityRange": "E17:E25",
                            "aggregateReplicateRanges": [
                                "I18:P18",
                                "I20:P20",
                            ],
                        }
                    ]
                }
            ]
        }
        repaired = (
            semantic_ai
            ._apply_deterministic_aggregate_identity_alignment_repair(
                draft,
                "StudyContractError: studies[0].measurementSeries[0]"
                ".aggregateReplicateRanges[0] must be contained in and "
                "aligned with studies[0].measurementSeries[0]"
                ".rowIdentityRange",
            )
        )
        self.assertIsNotNone(repaired)
        self.assertEqual(
            ["E18", "E20"],
            repaired["studies"][0]["measurementSeries"][0][
                "aggregateReplicateRanges"
            ],
        )
        self.assertEqual(
            ["I18:P18", "I20:P20"],
            draft["studies"][0]["measurementSeries"][0][
                "aggregateReplicateRanges"
            ],
        )

    def result(self) -> dict:
        return {
            "schemaVersion": "semantic-locator-v1",
            "promptVersion": "semantic-locator-prompt-v1",
            "revisionUid": "capture_revision_generic",
            "contentSha256": "a" * 64,
            "chunkId": "sheet-1-section-1-chunk-1",
            "status": "CANDIDATES",
            "candidates": [
                {
                    "key": "cryogenic-dwell",
                    "title": "Unfamiliar dwell review",
                    "summary": "A possible condition and response are present.",
                    "designHint": "UNKNOWN",
                    "contexts": ["Model Z"],
                    "changedFactors": ["Cryogenic dwell"],
                    "outcomes": ["Unseen response"],
                    "comparisonHints": [],
                    "evidence": [
                        {
                            "sheet": "Unfamiliar data",
                            "range": "B3:H10",
                            "role": "CANDIDATE_REGION",
                        }
                    ],
                    "limitations": ["Control arm is not visible in this chunk."],
                    "confidence": 0.65,
                }
            ],
            "notes": [],
        }

    def second_chunk(self) -> dict:
        return {
            "chunkId": "sheet-1-section-1-chunk-2",
            "sheet": "Unfamiliar data",
            "range": "B21:H30",
            "cells": [
                {"coordinate": "B21", "displayValue": "Alternate process"},
                {"coordinate": "C21", "displayValue": "Result"},
            ],
        }

    def second_result(self) -> dict:
        result = json.loads(json.dumps(self.result()))
        result["chunkId"] = "sheet-1-section-1-chunk-2"
        result["candidates"][0]["key"] = "alternate-process"
        result["candidates"][0]["title"] = "Alternate process review"
        result["candidates"][0]["evidence"][0]["range"] = "B21:H25"
        return result

    def batch_result(self) -> dict:
        return {
            "schemaVersion": "semantic-locator-batch-v1",
            "promptVersion": "semantic-locator-batch-prompt-v1",
            "revisionUid": self.source()["revisionUid"],
            "contentSha256": self.source()["contentSha256"],
            "results": [self.result(), self.second_result()],
        }

    def study_draft(self) -> dict:
        evidence = [
            {
                "sheet": "Unfamiliar data",
                "range": "B3:H10",
                "role": "SOURCE",
                "sourceText": "Cryogenic dwell",
                "note": "",
            }
        ]
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                **self.source(),
                "dataset": "Fixture",
                "contentComplete": True,
            },
            "workbookAnalysis": {
                "key": "generic-review",
                "title": "Generic review",
                "summary": "A condition and an unfamiliar response were recorded.",
                "status": "NEEDS_REVIEW",
                "verificationStatus": "NEEDS_REVIEW",
                "limitations": ["Draft semantics require review."],
                "evidence": evidence,
            },
            "studies": [
                {
                    "key": "generic-study",
                    "title": "Cryogenic dwell review",
                    "purpose": "",
                    "hypothesis": "",
                    "objective": "",
                    "designType": "CONTROL_TEST",
                    "comparisonBasis": "Possible adjacent conditions",
                    "verificationStatus": "NEEDS_REVIEW",
                    "comparabilityStatus": "UNASSESSED",
                    "confoundingStatus": "UNASSESSED",
                    "summary": "Unfamiliar response values appear under two conditions.",
                    "limitations": ["Control designation requires review."],
                    "evidence": evidence,
                    "contexts": [],
                    "factors": [
                        {
                            "key": "cryogenic-dwell",
                            "originalLabel": "Cryogenic dwell",
                            "baselineCondition": "4 s",
                            "changedCondition": "9.5 s",
                            "changeDirection": "INCREASE",
                            "isolationStatus": "UNASSESSED",
                            "evidence": evidence,
                        }
                    ],
                    "arms": [
                        {
                            "key": "candidate-a",
                            "role": "OTHER",
                            "label": "9.5 s",
                            "condition": "9.5 s dwell",
                            "sampleSize": None,
                            "sampleBasis": "",
                            "matchingBasis": "",
                            "factorValues": [
                                {
                                    "factor": "cryogenic-dwell",
                                    "value": "9.5 s",
                                    "valueNumber": 9.5,
                                    "unit": "s",
                                    "isBaseline": False,
                                    "heldConstant": False,
                                }
                            ],
                            "evidence": evidence,
                        }
                    ],
                    "outcomes": [
                        {
                            "key": "unseen-response",
                            "originalLabel": "Unseen response",
                            "metricType": "measurement",
                            "unit": "",
                            "favorableDirection": "UNKNOWN",
                            "evidence": evidence,
                            "observations": [
                                {
                                    "key": "candidate-a-value",
                                    "arm": "candidate-a",
                                    "valueNumber": 12.5,
                                    "valueText": "12.5",
                                    "numerator": None,
                                    "denominator": None,
                                    "ratePpm": None,
                                    "min": None,
                                    "max": None,
                                    "average": None,
                                    "sampleSize": None,
                                    "evidence": evidence,
                                }
                            ],
                        }
                    ],
                    "comparisons": [],
                    "conclusions": [
                        {
                            "key": "descriptive-result",
                            "text": "The sheet records an unfamiliar response.",
                            "claimType": "AI_DERIVED_DESCRIPTIVE",
                            "causalStrength": "DESCRIPTIVE",
                            "evidence": evidence,
                        }
                    ],
                }
            ],
        }

    def incompatible_raw_comparison_draft(self) -> dict:
        draft = self.study_draft()
        study = draft["studies"][0]
        support_studies = [
            json.loads(json.dumps(study)),
            json.loads(json.dumps(study)),
        ]
        support_studies[0]["key"] = "generic-support-0"
        support_studies[1]["key"] = "generic-support-1"
        outcome = study["outcomes"][0]
        outcome["key"] = "gauss-measurement"
        outcome["originalLabel"] = "Gauss measurement"
        outcome["unit"] = "G"
        second_arm = json.loads(json.dumps(study["arms"][0]))
        second_arm.update(
            {
                "key": "candidate-b",
                "label": "4 s",
                "condition": "4 s dwell",
            }
        )
        second_arm["factorValues"][0].update(
            {
                "value": "4 s",
                "valueNumber": 4,
            }
        )
        study["arms"].append(second_arm)
        series = {
            "key": "gauss-raw-a",
            "seriesRole": "RAW",
            "aggregationFunction": "",
            "aggregateOfSeries": [],
            "outcome": "gauss-measurement",
            "arm": "candidate-a",
            "sheet": "Unfamiliar data",
            "headerRange": "C3",
            "valueRange": "C3",
            "rowIdentityRange": "B3",
            "aggregateReplicateRanges": [],
            "axisSource": "ROW_IDENTITY",
            "axisLabel": "source numeric axis",
            "axisUnit": "",
            "valueUnit": "G",
            "stratumKey": "profile-a",
            "verificationStatus": "NEEDS_REVIEW",
        }
        second_series = json.loads(json.dumps(series))
        second_series.update(
            {
                "key": "gauss-raw-b",
                "arm": "candidate-b",
                "valueUnit": "mT",
                "stratumKey": "profile-b",
            }
        )
        study["measurementSeries"] = [series, second_series]
        study["comparisons"] = [
            {
                "key": "candidate-a-vs-b",
                "comparedArm": "candidate-a",
                "controlArm": "candidate-b",
                "designType": "CONTROL_TEST",
                "matchingBasis": "Source-adjacent conditions",
                "validityStatus": "NEEDS_REVIEW",
                "confoundingStatus": "UNASSESSED",
                "verificationStatus": "NEEDS_REVIEW",
                "aggregationEligible": False,
                "evidence": json.loads(
                    json.dumps(study["evidence"])
                ),
                "effects": [],
            }
        ]
        draft["studies"] = [*support_studies, study]
        return draft

    def synthetic_split_arm_draft(
        self,
    ) -> tuple[dict, list[dict]]:
        draft = self.study_draft()
        study = draft["studies"][0]
        support_studies = [
            json.loads(json.dumps(study)),
            json.loads(json.dumps(study)),
        ]
        support_studies[0]["key"] = "generic-support-0"
        support_studies[1]["key"] = "generic-support-1"
        evidence = [
            {
                "sheet": "Test",
                "range": "E30:H34",
                "role": "SOURCE",
                "sourceText": "",
                "note": "",
            }
        ]
        study["factors"].extend(
            [
                {
                    "key": "drop_sample_type",
                    "originalLabel": "Sample type",
                    "baselineCondition": "Normal",
                    "changedCondition": "Test",
                    "changeDirection": "OTHER",
                    "isolationStatus": "UNASSESSED",
                    "evidence": evidence,
                },
                {
                    "key": "drop_mode",
                    "originalLabel": "Mode",
                    "baselineCondition": "Auto",
                    "changedCondition": "Manual",
                    "changeDirection": "OTHER",
                    "isolationStatus": "UNASSESSED",
                    "evidence": evidence,
                },
            ]
        )

        def arm(
            key: str,
            role: str,
            identity: str,
            sample_type: str,
            mode: str,
            arm_evidence: list[dict],
        ) -> dict:
            return {
                "key": key,
                "role": role,
                "label": identity,
                "condition": identity,
                "sampleSize": None,
                "sampleBasis": "",
                "matchingBasis": "",
                "factorValues": [
                    {
                        "factor": "drop_sample_type",
                        "value": sample_type,
                        "valueNumber": None,
                        "unit": "",
                        "isBaseline": sample_type == "Normal",
                        "heldConstant": False,
                    },
                    {
                        "factor": "drop_mode",
                        "value": mode,
                        "valueNumber": None,
                        "unit": "",
                        "isBaseline": False,
                        "heldConstant": True,
                    },
                ],
                "evidence": arm_evidence,
            }

        study["arms"] = [
            arm(
                "drop_test_auto",
                "TEST",
                "Test / Auto",
                "Test",
                "Auto",
                [{"sheet": "Test", "range": "E30:F30"}],
            ),
            arm(
                "drop_test_manual",
                "TEST",
                "Test / Manual",
                "Test",
                "Manual",
                [
                    {"sheet": "Test", "range": "E30:E31"},
                    {"sheet": "Test", "range": "F31"},
                ],
            ),
            arm(
                "drop_normal_auto",
                "REFERENCE",
                "Normal / Auto",
                "Normal",
                "Auto",
                [{"sheet": "Test", "range": "E32:F32"}],
            ),
            arm(
                "drop_normal_manual",
                "REFERENCE",
                "Normal / Manual",
                "Normal",
                "Manual",
                [
                    {"sheet": "Test", "range": "E32:E33"},
                    {"sheet": "Test", "range": "F33"},
                ],
            ),
            arm(
                "safe_literal_normal",
                "REFERENCE",
                "Normal",
                "Normal",
                "Auto",
                [{"sheet": "Test", "range": "E34"}],
            ),
            arm(
                "safe_whole_cell_composite",
                "REFERENCE",
                "Normal / Auto",
                "Normal",
                "Auto",
                [{"sheet": "Test", "range": "H34"}],
            ),
        ]
        study["outcomes"][0]["observations"][0]["arm"] = (
            "drop_test_auto"
        )
        draft["studies"] = [*support_studies, study]
        chunk = {
            "chunkId": "test-split-arms",
            "sheet": {"title": "Test"},
            "cells": [
                {
                    "coordinate": coordinate,
                    "rawValue": value,
                    "displayValue": value,
                }
                for coordinate, value in (
                    ("E30", "Test"),
                    ("F30", "Auto"),
                    ("F31", "Manual"),
                    ("E32", "Normal"),
                    ("F32", "Auto"),
                    ("F33", "Manual"),
                    ("E34", "Normal"),
                    ("H34", "Normal / Auto"),
                )
            ],
            "contextCells": [],
        }
        return draft, [chunk]

    def temporal_stage_draft(self) -> tuple[dict, dict]:
        draft = self.study_draft()
        study = draft["studies"][0]
        temporal_evidence = [
            {
                "sheet": "Unfamiliar data",
                "range": "D3:E3",
                "role": "SOURCE",
                "sourceText": "",
                "note": "",
            }
        ]
        study["factors"].append(
            {
                "key": "measurement-stage",
                "originalLabel": "Measurement stage",
                "baselineCondition": "",
                "changedCondition": "Before, After",
                "changeDirection": "OTHER",
                "isolationStatus": "UNASSESSED",
                "evidence": temporal_evidence,
            }
        )
        before_arm = json.loads(json.dumps(study["arms"][0]))
        before_arm.update(
            {
                "key": "candidate-before",
                "role": "BEFORE",
                "label": "Before",
                "condition": "Before",
                "evidence": [
                    {
                        **temporal_evidence[0],
                        "range": "D3",
                    }
                ],
            }
        )
        before_arm["factorValues"].append(
            {
                "factor": "measurement-stage",
                "value": "Before",
                "valueNumber": None,
                "unit": "",
                "isBaseline": False,
                "heldConstant": False,
            }
        )
        after_arm = json.loads(json.dumps(study["arms"][0]))
        after_arm.update(
            {
                "key": "candidate-after",
                "role": "AFTER",
                "label": "After",
                "condition": "After",
                "evidence": [
                    {
                        **temporal_evidence[0],
                        "range": "E3",
                    }
                ],
            }
        )
        after_arm["factorValues"].append(
            {
                "factor": "measurement-stage",
                "value": "After",
                "valueNumber": None,
                "unit": "",
                "isBaseline": False,
                "heldConstant": False,
            }
        )
        study["arms"] = [before_arm, after_arm]
        study["outcomes"][0]["observations"][0][
            "arm"
        ] = "candidate-before"
        chunk = self.chunk()
        chunk["cells"].extend(
            [
                {
                    "coordinate": "D3",
                    "rawValue": 1,
                    "displayValue": 1,
                    "numberFormat": '"stage "_Before',
                },
                {
                    "coordinate": "E3",
                    "rawValue": 1,
                    "displayValue": 1,
                    "numberFormat": '"stage "_After',
                },
            ]
        )
        return draft, chunk

    def test_prompt_is_generic_and_explicitly_excludes_images(self) -> None:
        prompt = semantic_ai.build_locator_prompt(
            source=self.source(),
            workbook={"status": "CAPTURED"},
            chunk=self.chunk(),
        )
        self.assertIn("open-ended", prompt)
        self.assertIn("examples only", prompt)
        self.assertIn("images are out of scope", prompt)
        self.assertIn("Cryogenic dwell", prompt)

    def test_study_prompt_preserves_explicit_baseline_comparisons_for_review(self) -> None:
        prompt = semantic_ai.build_study_draft_prompt(
            source=self.source(),
            workbook={"status": "CAPTURED"},
            locator_results=[self.result()],
            focused_chunks=[self.chunk()],
        )
        self.assertIn("canonical-study-draft-prompt-v25", prompt)
        self.assertIn("BASELINE is not an allowed role", prompt)
        self.assertIn(
            "custom number_format explicitly labels that Arm",
            prompt,
        )
        self.assertIn("one shared condition Arm", prompt)
        self.assertIn("'Block 1', 'Block 2'", prompt)
        self.assertIn("10 stays 10", prompt)
        self.assertIn(
            "do not put Before, After, pre-change, or post-change in a Study title",
            prompt,
        )
        self.assertIn("does not make the plain 18kPa arm a COMPARATOR", prompt)
        self.assertIn("current pressure, prior exposure/condition", prompt)
        self.assertIn("directly to every data-bearing Study", prompt)
        self.assertIn("'source numeric axis'", prompt)
        self.assertIn("filename identifier differs", prompt)
        self.assertIn("custom number_format is source-authored display meaning", prompt)
        self.assertIn("raw header values 1..10", prompt)
        self.assertIn("pressure, replicate, and temporal meaning", prompt)
        self.assertIn("custom format on an unrelated cell/range", prompt)
        self.assertIn("authorizes that Arm's BEFORE/AFTER", prompt)
        self.assertIn("never authorizes CONTROL or BASELINE", prompt)
        self.assertIn("same ordered #1..#N", prompt)
        self.assertIn("'VP+Coil Normal'", prompt)
        self.assertIn("300 After header-only Arm", prompt)
        self.assertIn("seriesRole RAW", prompt)
        self.assertIn("seriesRole AGGREGATE", prompt)
        self.assertIn("aggregationFunction AVERAGE", prompt)
        self.assertIn("aggregateOfSeries listing the exact RAW series keys", prompt)
        self.assertIn("frequency profile matrix", prompt)
        self.assertIn("F3:O16", prompt)
        self.assertIn("P3:P16", prompt)
        self.assertIn("distinct pressure-level summary Arm with role OTHER", prompt)
        self.assertIn("300 Fo AVG", prompt)
        self.assertIn("single-row Fo RAW or AVG series", prompt)
        self.assertIn("use axisSource ROW_IDENTITY", prompt)
        self.assertIn("AVG header is not a raw measurement axis", prompt)
        self.assertIn("vertical one-column Fo series", prompt)
        self.assertIn("C5:C14", prompt)
        self.assertIn("source AVG row such as C15", prompt)
        self.assertIn("Never use ROW_IDENTITY for this vertical", prompt)
        self.assertIn("horizontal frequency-row matrices", prompt)
        self.assertIn("baselineCondition must be Before", prompt)
        self.assertIn("changedCondition must be After", prompt)
        self.assertIn("factor-level temporal identity only", prompt)
        self.assertIn(
            "Every cell implied by a measurementSeries.valueRange",
            prompt,
        )
        self.assertIn(
            "include only the data-bearing REF column",
            prompt,
        )
        self.assertIn(
            "unique source-backed replicateKey",
            prompt,
        )
        self.assertIn(
            "Normal or Normal (...) maps to arm.role REFERENCE",
            prompt,
        )
        self.assertIn("bare abbreviation such as ST is not Standard", prompt)
        self.assertIn("Normal #1 through Normal #10", prompt)
        self.assertIn("#N values are ordered and distinct", prompt)
        self.assertIn("Mixed Test/Normal evidence", prompt)
        self.assertIn("literal label Normal still requires", prompt)
        self.assertIn("draft source-backed NEEDS_REVIEW comparisons", prompt)
        self.assertIn("Do not omit an explicit comparison", prompt)
        self.assertIn("effects list to empty", prompt)
        self.assertIn("Never collapse several numeric submetrics", prompt)
        self.assertIn("deterministic importer can expand every numeric point", prompt)
        self.assertIn("axisSource", prompt)
        self.assertIn("HEADER when the column headers are", prompt)
        self.assertIn("ROW_IDENTITY when row identities are", prompt)
        self.assertIn("aggregateReplicateRanges", prompt)
        self.assertIn("sample_size", prompt)
        self.assertIn("set both numerator and denominator null", prompt)
        self.assertIn("Input is not equal to OK plus Total NG", prompt)
        self.assertIn("explicitly identifies a DOE/comparison", prompt)
        self.assertIn("only adjacent source-order", prompt)
        self.assertIn("drying temperature and drying duration", prompt)
        self.assertIn("exact group header that directly governs", prompt)
        self.assertIn("3.8 for 3.8%", prompt)
        self.assertIn(
            "exact underlying captured numeric value multiplied by 100",
            prompt,
        )
        self.assertIn("rounded screen-display string", prompt)
        self.assertIn(
            "agree with their exact arithmetic",
            prompt,
        )
        self.assertIn(
            "never use display rounding as a numeric claim",
            prompt,
        )
        self.assertIn("percentage-only total-rate cell", prompt)
        self.assertIn("Never sum category counts", prompt)
        self.assertIn("aligned one-to-one NEEDS_REVIEW comparison", prompt)
        self.assertIn("Preserve every fixed condition", prompt)
        self.assertIn("Preserve unlabeled blocks", prompt)
        self.assertIn("never add the repeated labels together", prompt)
        self.assertIn("do not auto-create control comparisons", prompt)
        self.assertIn("Do not fill a missing arm/outcome row with zero", prompt)
        self.assertIn("reference exactly matches a declared key", prompt)
        self.assertIn(
            "repeated sample labels 1..N do not prove paired observations",
            prompt,
        )
        self.assertIn("whole-cell explicit count ratio", prompt)
        self.assertIn(
            "Numerator and denominator must always be supplied together",
            prompt,
        )
        self.assertIn(
            "Never invent a missing denominator",
            prompt,
        )
        self.assertIn(
            "Do not infer a SOURCE_CONCLUSION from condition labels",
            prompt,
        )
        self.assertIn(
            "copy that literal narrative into the matching evidence.sourceText",
            prompt,
        )
        self.assertIn("AI_DERIVED_DESCRIPTIVE", prompt)
        self.assertIn(
            "keep all numeric observations either way",
            prompt,
        )
        self.assertIn(
            "Normal or Normal (...) maps to arm.role REFERENCE",
            prompt,
        )
        self.assertIn(
            "comparison.controlArm may reference a REFERENCE Arm",
            prompt,
        )
        self.assertIn(
            "one recognized quantity such as '1.56mg' or '1.56 mg'",
            prompt,
        )
        self.assertIn(
            "Never parse ranges, narrative text, dates, model IDs, ratios",
            prompt,
        )
        self.assertIn("quantity-looking token is embedded", prompt)
        self.assertIn("Test Led UC (VP+CD) 5s", prompt)
        self.assertIn("Drying temperature 80°C time 5min", prompt)
        self.assertIn(
            "Never mix qualitative-only Pass/Fail observations",
            prompt,
        )
        self.assertIn("Repeated per-replicate PASSED/FAILED", prompt)
        self.assertIn("never infer a numeric threshold", prompt)
        self.assertIn(
            "RAW measurementSeries with only scalar or summary Observations",
            prompt,
        )
        self.assertIn(
            "same ordered axis identity, shape, and stratum",
            prompt,
        )
        self.assertIn(
            "Equal replicate or sample labels do not establish physical pairing",
            prompt,
        )
        self.assertIn(
            "Total NG does not equal the arithmetic sum",
            prompt,
        )
        self.assertIn(
            "MIN/MAX/AVG formula has no cached value",
            prompt,
        )

    def test_study_repair_prompt_rejects_inferred_temporal_phases(self) -> None:
        prompt = semantic_ai.build_study_draft_repair_prompt(
            self.study_draft(),
            "SemanticAiError: inferred phase",
            source_prompt="FOCUSED SOURCE PACKET",
        )

        self.assertIn("canonical-study-draft-prompt-v25", prompt)
        self.assertIn("locator summaries", prompt)
        self.assertIn("authorizes that Arm's BEFORE/AFTER", prompt)
        self.assertIn("never authorizes CONTROL or BASELINE", prompt)
        self.assertIn("one shared Arm", prompt)
        self.assertIn("Block 1/Block 2/source-coordinate stratum", prompt)
        self.assertIn("10 remains 10", prompt)
        self.assertIn("does not establish the plain condition as a COMPARATOR", prompt)
        self.assertIn("delete unsupported comparisons", prompt)
        self.assertIn("detached context-only Study", prompt)
        self.assertIn("filename-versus-report identifier mismatch", prompt)
        self.assertIn("exact captured cell's custom number_format", prompt)
        self.assertIn("never transfer a format label", prompt)
        self.assertIn("seriesRole AGGREGATE", prompt)
        self.assertIn("aggregateOfSeries listing the exact RAW series keys", prompt)
        self.assertIn("frequency-by-replicate matrix", prompt)
        self.assertIn("F3:O16", prompt)
        self.assertIn("P3:P16", prompt)
        self.assertIn("same ordered #1..#N", prompt)
        self.assertIn("300 After header-only Arm", prompt)
        self.assertIn("single-row Fo RAW or AVG series", prompt)
        self.assertIn("never treat the AVG header as a distinct measurement axis", prompt)
        self.assertIn("vertical one-column Fo layout", prompt)
        self.assertIn("C5:C14", prompt)
        self.assertIn("source AVG row C15", prompt)
        self.assertIn("Never repair this vertical", prompt)
        self.assertIn("Horizontal frequency-row matrices remain", prompt)
        self.assertIn("perform an exact minimal repair", prompt)
        self.assertIn("matching BEFORE factorValue.isBaseline true", prompt)
        self.assertIn(
            "repeated sample labels 1..N do not prove the same physical specimens",
            prompt,
        )
        self.assertIn("whole-cell explicit count ratio", prompt)
        self.assertIn(
            "numerator and denominator must be supplied together",
            prompt,
        )
        self.assertIn(
            "otherwise clear both fields while preserving the source-backed raw count",
            prompt,
        )
        self.assertIn("not present in the cited Capture v2 cells", prompt)
        self.assertIn("Never sum overlapping category counts", prompt)
        self.assertIn(
            "SOURCE_CONCLUSION requires an exact captured narrative",
            prompt,
        )
        self.assertIn(
            "limitation saying no source narrative exists cannot support",
            prompt,
        )
        self.assertIn(
            "never remove its numeric observations",
            prompt,
        )
        self.assertIn(
            "Normal or Normal (...) Arm label must be REFERENCE",
            prompt,
        )
        self.assertIn("Never expand a bare ST abbreviation", prompt)
        self.assertIn("descriptive grouped REFERENCE", prompt)
        self.assertIn("Reject mixed Test/Normal cells", prompt)
        self.assertIn("Preserve each replicate axis identity", prompt)
        self.assertIn("'VP+Coil Normal'", prompt)
        self.assertIn(
            "factor quantity such as '1.56mg' or '1.56 mg'",
            prompt,
        )
        self.assertIn("quantity-looking token is embedded", prompt)
        self.assertIn("valueNumber is null", prompt)
        self.assertIn(
            "split them into a distinct categorical Outcome",
            prompt,
        )
        self.assertIn("Repair repeated per-replicate PASSED/FAILED", prompt)
        self.assertIn("never infer an unstated threshold", prompt)
        self.assertIn(
            "Reject RAW measurementSeries versus scalar/summary representation",
            prompt,
        )
        self.assertIn(
            "Omit the incompatible Comparison only",
            prompt,
        )
        self.assertIn("setting numerator and denominator null", prompt)
        self.assertIn("Input differs from OK plus Total NG", prompt)
        self.assertIn("explicit DOE/comparison title or purpose", prompt)
        self.assertIn("only adjacent source-order pairs", prompt)
        self.assertIn("governing group header", prompt)
        self.assertIn("When Total NG differs", prompt)
        self.assertIn(
            "MIN/MAX/AVG formula without a cached value",
            prompt,
        )
        self.assertIn(
            "repair valueNumber to the exact underlying captured numeric value multiplied by 100",
            prompt,
        )
        self.assertIn(
            "rounded screen-display string only",
            prompt,
        )
        self.assertIn(
            "never retain display rounding as a numeric claim",
            prompt,
        )

    def test_study_prompt_compacts_transport_without_dropping_cells(
        self,
    ) -> None:
        chunk = self.chunk()
        chunk["cells"].append(
            {
                "coordinate": "D8",
                "rawValue": 0.25,
                "displayValue": "25%",
                "numberFormat": "0%",
                "formula": "=1/4",
                "cachedValue": 0.25,
                "hidden": {
                    "sheet": False,
                    "row": True,
                    "column": False,
                },
            }
        )
        view = semantic_ai._study_draft_chunk_view(chunk)

        self.assertEqual(
            [cell["coordinate"] for cell in chunk["cells"]],
            [cell["c"] for cell in view["cells"]],
        )
        encoded = view["cells"][-1]
        self.assertEqual("25%", encoded["v"])
        self.assertEqual(0.25, encoded["r"])
        self.assertEqual("=1/4", encoded["f"])
        self.assertEqual("0%", encoded["n"])
        self.assertEqual({"row": True}, encoded["h"])
        self.assertNotIn("dataType", encoded)

    def test_study_runner_rejects_oversized_input_before_codex_call(
        self,
    ) -> None:
        called = False

        def unexpected_run(*_: object, **__: object):
            nonlocal called
            called = True
            raise AssertionError("Codex must not receive oversized input.")

        with tempfile.TemporaryDirectory() as directory, mock.patch.object(
            semantic_ai,
            "build_study_draft_prompt",
            return_value="x"
            * (semantic_ai.STUDY_DRAFT_MAX_INPUT_CHARS + 1),
        ):
            with self.assertRaisesRegex(
                semantic_ai.SemanticAiError,
                "fail-closed input budget",
            ):
                semantic_ai.run_codex_study_draft(
                    source=self.source(),
                    workbook={"status": "CAPTURED"},
                    locator_results=[self.result()],
                    focused_chunks=[self.chunk()],
                    content_complete=True,
                    output_path=Path(directory) / "draft.json",
                    run_command=unexpected_run,
                )
        self.assertFalse(called)

    def test_locator_accepts_unknown_factor_and_outcome(self) -> None:
        validated = semantic_ai.validate_locator_result(
            self.result(),
            revision_uid=self.source()["revisionUid"],
            content_sha256=self.source()["contentSha256"],
            chunk=self.chunk(),
        )
        self.assertEqual("Cryogenic dwell", validated["candidates"][0]["changedFactors"][0])

    def test_locator_rejects_range_outside_chunk(self) -> None:
        result = self.result()
        result["candidates"][0]["evidence"][0]["range"] = "A1:H10"
        with self.assertRaisesRegex(semantic_ai.SemanticAiError, "outside"):
            semantic_ai.validate_locator_result(
                result,
                revision_uid=self.source()["revisionUid"],
                content_sha256=self.source()["contentSha256"],
                chunk=self.chunk(),
            )

    def test_locator_rejects_source_identity_mismatch(self) -> None:
        result = self.result()
        result["revisionUid"] = "other"
        with self.assertRaisesRegex(semantic_ai.SemanticAiError, "revisionUid"):
            semantic_ai.validate_locator_result(
                result,
                revision_uid=self.source()["revisionUid"],
                content_sha256=self.source()["contentSha256"],
                chunk=self.chunk(),
            )

    def test_codex_runner_uses_read_only_mode_and_writes_validated_output(self) -> None:
        calls: list[tuple[list[str], str]] = []

        def fake_run(command: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            calls.append((command, str(kwargs["input"])))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(self.result(), ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "locator.json"
            result = semantic_ai.run_codex_locator(
                source=self.source(),
                workbook={"status": "CAPTURED"},
                chunk=self.chunk(),
                output_path=output,
                codex_command=("codex-test",),
                run_command=fake_run,
            )
            self.assertTrue(output.is_file())
            self.assertEqual("CANDIDATES", result["status"])
        command, prompt = calls[0]
        self.assertEqual("codex-test", command[0])
        self.assertEqual("read-only", command[command.index("--sandbox") + 1])
        self.assertIn("Cryogenic dwell", prompt)

    def test_batch_locator_requires_exact_order_and_validates_each_chunk(self) -> None:
        validated = semantic_ai.validate_batch_locator_result(
            self.batch_result(),
            revision_uid=self.source()["revisionUid"],
            content_sha256=self.source()["contentSha256"],
            chunks=[self.chunk(), self.second_chunk()],
        )
        self.assertEqual(
            [
                "sheet-1-section-1-chunk-1",
                "sheet-1-section-1-chunk-2",
            ],
            [item["chunkId"] for item in validated],
        )
        invalid = self.batch_result()
        invalid["results"].reverse()
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "same order",
        ):
            semantic_ai.validate_batch_locator_result(
                invalid,
                revision_uid=self.source()["revisionUid"],
                content_sha256=self.source()["contentSha256"],
                chunks=[self.chunk(), self.second_chunk()],
            )

    def test_codex_batch_locator_uses_one_read_only_call_and_writes_all_outputs(self) -> None:
        calls: list[list[str]] = []

        def fake_run(command: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            calls.append(command)
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(self.batch_result(), ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            first_output = Path(temp_dir) / "first.json"
            second_output = Path(temp_dir) / "second.json"
            validated = semantic_ai.run_codex_locator_batch(
                source=self.source(),
                workbook={"status": "CAPTURED"},
                chunks=[self.chunk(), self.second_chunk()],
                output_paths={
                    self.chunk()["chunkId"]: first_output,
                    self.second_chunk()["chunkId"]: second_output,
                },
                codex_command=("codex-test",),
                run_command=fake_run,
            )
            self.assertEqual(2, len(validated))
            self.assertTrue(first_output.is_file())
            self.assertTrue(second_output.is_file())
        self.assertEqual(1, len(calls))
        self.assertEqual(
            "read-only",
            calls[0][calls[0].index("--sandbox") + 1],
        )

    def test_ai_study_draft_cannot_self_verify_or_calculate(self) -> None:
        draft = self.study_draft()
        draft["studies"][0]["verificationStatus"] = "VERIFIED"
        with self.assertRaisesRegex(semantic_ai.SemanticAiError, "self-verify"):
            semantic_ai.validate_ai_study_draft(
                draft,
                source=draft["source"],
                content_complete=True,
            )

    def test_ai_study_draft_accepts_unknown_domain_with_exact_source_identity(self) -> None:
        draft = self.study_draft()
        checked: list[str] = []
        validated = semantic_ai.validate_ai_study_draft(
            draft,
            source=draft["source"],
            content_complete=True,
            evidence_checker=lambda item: checked.append(item["range"]),
        )
        self.assertEqual(
            "Cryogenic dwell",
            validated["studies"][0]["factors"][0]["originalLabel"],
        )
        self.assertGreater(len(checked), 3)

    def test_ai_study_draft_rejects_packet_coverage_mismatch(self) -> None:
        draft = self.study_draft()
        draft["source"]["contentComplete"] = False

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "contentComplete does not match",
        ):
            semantic_ai.validate_ai_study_draft(
                draft,
                source=draft["source"],
                content_complete=True,
            )

    def test_ai_study_draft_normalizes_only_conservative_enum_aliases(self) -> None:
        draft = self.study_draft()
        draft["studies"][0]["confoundingStatus"] = "POSSIBLE_CONFOUNDING"
        draft["studies"][0]["comparabilityStatus"] = "NOT_ASSESSED"
        draft["studies"][0]["factors"][0]["isolationStatus"] = "NOT_DETERMINED"
        draft["studies"][0]["arms"][0]["role"] = "PRODUCTION_LOT"
        draft["studies"][0]["outcomes"][0]["favorableDirection"] = "MIXED"
        validated = semantic_ai.validate_ai_study_draft(
            draft,
            source=draft["source"],
            content_complete=True,
        )
        self.assertEqual("POSSIBLE", validated["studies"][0]["confoundingStatus"])
        self.assertEqual(
            "UNASSESSED",
            validated["studies"][0]["comparabilityStatus"],
        )
        self.assertEqual(
            "OTHER",
            validated["studies"][0]["arms"][0]["role"],
        )
        self.assertEqual(
            "UNASSESSED",
            validated["studies"][0]["factors"][0]["isolationStatus"],
        )
        self.assertEqual(
            "UNKNOWN",
            validated["studies"][0]["outcomes"][0]["favorableDirection"],
        )
        self.assertEqual("NEEDS_REVIEW", validated["studies"][0]["verificationStatus"])

    def test_ai_study_draft_assigns_identity_to_repeated_unverified_observations(self) -> None:
        draft = self.study_draft()
        observations = draft["studies"][0]["outcomes"][0]["observations"]
        observations.append(
            {
                **json.loads(json.dumps(observations[0])),
                "key": "candidate-a-value-percent",
                "valueNumber": 0.125,
            }
        )
        validated = semantic_ai.validate_ai_study_draft(
            draft,
            source=draft["source"],
            content_complete=True,
        )
        replicate_keys = [
            item["replicateKey"]
            for item in validated["studies"][0]["outcomes"][0]["observations"]
        ]
        self.assertEqual(
            ["candidate-a-value", "candidate-a-value-percent"],
            replicate_keys,
        )

    def test_temporal_role_requires_literal_label_in_arm_evidence(self) -> None:
        draft = self.study_draft()
        arm = draft["studies"][0]["arms"][0]
        arm["role"] = "BEFORE"
        arm["label"] = "Before"
        arm["condition"] = "Before"

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "requires a literal label",
        ):
            semantic_ai._validate_source_explicit_temporal_semantics(
                draft,
                [self.chunk()],
            )

        explicit_chunk = self.chunk()
        explicit_chunk["cells"].append(
            {"coordinate": "D3", "displayValue": "Before"}
        )
        semantic_ai._validate_source_explicit_temporal_semantics(
            draft,
            [explicit_chunk],
        )

    def test_temporal_series_accepts_evidence_linked_custom_number_format(
        self,
    ) -> None:
        draft = self.study_draft()
        draft["studies"][0]["measurementSeries"] = [
            {
                "key": "18kpa-before-series",
                "stratumKey": "Before",
                "sheet": "Unfamiliar data",
                "headerRange": "D3:D3",
                "valueRange": "D4:D8",
                "rowIdentityRange": "B4:B8",
            },
            {
                "key": "18kpa-after-series",
                "stratumKey": "After",
                "sheet": "Unfamiliar data",
                "headerRange": "E3:E3",
                "valueRange": "E4:E8",
                "rowIdentityRange": "B4:B8",
            },
        ]
        chunk = self.chunk()
        chunk["cells"].append(
            {
                "coordinate": "D3",
                "rawValue": 1,
                "displayValue": 1,
                "numberFormat": '"18kPa #"?"_Before',
            }
        )
        chunk["cells"].append(
            {
                "coordinate": "E3",
                "rawValue": 1,
                "displayValue": 1,
                "numberFormat": '"18kPa #"?"_After',
            }
        )

        semantic_ai._validate_source_explicit_temporal_semantics(
            draft,
            [chunk],
        )
        self.assertEqual(
            frozenset({"BEFORE"}),
            semantic_ai._number_format_temporal_terms('"phase "_Pre'),
        )
        self.assertEqual(
            frozenset({"AFTER"}),
            semantic_ai._number_format_temporal_terms('"phase "_Post'),
        )

    def test_temporal_series_uses_declared_header_range_as_evidence(
        self,
    ) -> None:
        draft = self.study_draft()
        draft["studies"][0]["measurementSeries"] = [
            {
                "key": "18kpa-before-series",
                "stratumKey": "18kPa Before",
                "sheet": "Unfamiliar data",
                "headerRange": "D3:D3",
                "valueRange": "D4:D8",
                "rowIdentityRange": "B4:B8",
            }
        ]
        chunk = self.chunk()
        chunk["cells"].append(
            {
                "coordinate": "D3",
                "rawValue": 1,
                "displayValue": 1,
                "numberFormat": '"18kPa #"?"_Before',
            }
        )

        semantic_ai._validate_source_explicit_temporal_semantics(
            draft,
            [chunk],
        )

    def test_unrelated_custom_number_format_cannot_authorize_series(
        self,
    ) -> None:
        draft = self.study_draft()
        draft["studies"][0]["measurementSeries"] = [
            {
                "key": "before-series",
                "stratumKey": "Before",
                "sheet": "Unfamiliar data",
                "headerRange": "B8:B8",
                "valueRange": "B9:B10",
                "rowIdentityRange": "C9:C10",
            }
        ]
        chunk = self.chunk()
        chunk["cells"].append(
            {
                "coordinate": "D3",
                "rawValue": 1,
                "displayValue": 1,
                "numberFormat": '"18kPa #"?"_Before',
            }
        )

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "without an evidence-linked literal captured cell label",
        ):
            semantic_ai._validate_source_explicit_temporal_semantics(
                draft,
                [chunk],
            )

    def test_custom_number_format_authorizes_temporal_arm(
        self,
    ) -> None:
        draft = self.study_draft()
        arm = draft["studies"][0]["arms"][0]
        arm["role"] = "BEFORE"
        arm["label"] = "18kPa Before"
        arm["condition"] = "18kPa Before"
        arm["evidence"] = [
            {
                "sheet": "Unfamiliar data",
                "range": "D3",
                "role": "SOURCE",
                "sourceText": "1",
                "note": "",
            }
        ]
        chunk = self.chunk()
        chunk["cells"].append(
            {
                "coordinate": "D3",
                "rawValue": 1,
                "displayValue": 1,
                "numberFormat": '"18kPa #"?"_Before',
            }
        )

        semantic_ai._validate_source_explicit_temporal_semantics(
            draft,
            [chunk],
        )

    def test_temporal_stage_factor_preserves_before_as_factor_baseline(
        self,
    ) -> None:
        draft, chunk = self.temporal_stage_draft()

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "source-authored temporal stage baseline mapping",
        ):
            semantic_ai._validate_source_explicit_temporal_semantics(
                draft,
                [chunk],
            )

        stage_factor = draft["studies"][0]["factors"][-1]
        stage_factor["baselineCondition"] = "Before"
        stage_factor["changedCondition"] = "After"
        before_value = draft["studies"][0]["arms"][0][
            "factorValues"
        ][-1]
        before_value["isBaseline"] = True

        semantic_ai._validate_source_explicit_temporal_semantics(
            draft,
            [chunk],
        )

    def test_temporal_stage_repair_projection_allows_only_mapping_fields(
        self,
    ) -> None:
        rejected, _chunk = self.temporal_stage_draft()
        corrected = json.loads(json.dumps(rejected))
        stage_factor = corrected["studies"][0]["factors"][-1]
        stage_factor["baselineCondition"] = "Before"
        stage_factor["changedCondition"] = "After"
        corrected["studies"][0]["arms"][0]["factorValues"][-1][
            "isBaseline"
        ] = True

        self.assertEqual(
            semantic_ai._temporal_stage_repair_projection(rejected),
            semantic_ai._temporal_stage_repair_projection(corrected),
        )

        corrected["studies"][0]["summary"] = "Unsafe unrelated change"
        self.assertNotEqual(
            semantic_ai._temporal_stage_repair_projection(rejected),
            semantic_ai._temporal_stage_repair_projection(corrected),
        )

    def test_unrelated_custom_number_format_cannot_authorize_temporal_arm(
        self,
    ) -> None:
        draft = self.study_draft()
        arm = draft["studies"][0]["arms"][0]
        arm["role"] = "BEFORE"
        arm["label"] = "18kPa Before"
        arm["condition"] = "18kPa Before"
        arm["evidence"] = [
            {
                "sheet": "Unfamiliar data",
                "range": "B8",
                "role": "SOURCE",
                "sourceText": "Unseen response",
                "note": "",
            }
        ]
        chunk = self.chunk()
        chunk["cells"].append(
            {
                "coordinate": "D3",
                "rawValue": 1,
                "displayValue": 1,
                "numberFormat": '"18kPa #"?"_Before',
            }
        )

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "requires a literal label",
        ):
            semantic_ai._validate_source_explicit_temporal_semantics(
                draft,
                [chunk],
            )

    def test_locator_source_text_alone_cannot_authorize_temporal_role(
        self,
    ) -> None:
        draft = self.study_draft()
        arm = draft["studies"][0]["arms"][0]
        arm["role"] = "AFTER"
        arm["label"] = "After"
        arm["condition"] = "After"
        arm["evidence"] = [
            {
                "sheet": "Unfamiliar data",
                "range": "C3",
                "role": "SOURCE",
                "sourceText": "After",
                "note": "AI/locator text only",
            }
        ]

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "requires a literal label",
        ):
            semantic_ai._validate_source_explicit_temporal_semantics(
                draft,
                [self.chunk()],
            )

    def test_temporal_words_in_title_or_stratum_require_source_label(self) -> None:
        draft = self.study_draft()
        draft["studies"][0]["title"] = "Before/After matrix"
        draft["studies"][0]["measurementSeries"] = [
            {"key": "block-before", "stratumKey": "Before"}
        ]

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "without an evidence-linked literal captured cell label",
        ):
            semantic_ai._validate_source_explicit_temporal_semantics(
                draft,
                [self.chunk()],
            )

    def test_codex_study_runner_is_read_only_and_validates_output(self) -> None:
        draft = self.study_draft()
        calls: list[list[str]] = []

        def fake_run(command: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            calls.append(command)
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(draft, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            result = semantic_ai.run_codex_study_draft(
                source=draft["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )
            self.assertEqual("NEEDS_REVIEW", result["studies"][0]["verificationStatus"])
            self.assertTrue(output.is_file())
        self.assertEqual("read-only", calls[0][calls[0].index("--sandbox") + 1])

    def test_codex_study_runner_retries_non_json_with_output_schema(
        self,
    ) -> None:
        draft = self.study_draft()
        calls: list[tuple[list[str], str]] = []

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            calls.append((command, str(kwargs["input"])))
            output_index = command.index("--output-last-message") + 1
            response = (
                "This is not JSON."
                if len(calls) == 1
                else json.dumps(draft, ensure_ascii=False)
            )
            Path(command[output_index]).write_text(
                response,
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            result = semantic_ai.run_codex_study_draft(
                source=draft["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )

        self.assertEqual(2, len(calls))
        self.assertNotIn("--output-schema", calls[0][0])
        self.assertIn("--output-schema", calls[1][0])
        self.assertIn("JSON RECOVERY", calls[1][1])
        self.assertIn("FOCUSED SOURCE PACKET", calls[1][1])
        self.assertIn(
            "invalid response is deliberately not supplied",
            calls[1][1],
        )
        self.assertNotIn("This is not JSON.", calls[1][1])
        self.assertEqual(
            "NEEDS_REVIEW",
            result["workbookAnalysis"]["verificationStatus"],
        )

    def test_study_transport_paths_are_parent_stable_unique_and_cleaned(
        self,
    ) -> None:
        draft = self.study_draft()
        paths: list[tuple[Path, Path | None]] = []
        calls = 0

        def fake_run(
            command: list[str],
            **_: object,
        ) -> subprocess.CompletedProcess[str]:
            nonlocal calls
            calls += 1
            last_message = Path(
                command[command.index("--output-last-message") + 1]
            )
            schema = (
                Path(command[command.index("--output-schema") + 1])
                if "--output-schema" in command
                else None
            )
            paths.append((last_message, schema))
            if schema is not None:
                self.assertTrue(schema.is_file())
            response = (
                "not JSON on the first attempt"
                if calls == 1
                else json.dumps(draft, ensure_ascii=False)
            )
            last_message.write_text(response, encoding="utf-8")
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            parent = Path(temp_dir)
            output = parent / "study.json"
            sentinel = parent / ".caller-owned.keep"
            sentinel.write_text("keep", encoding="utf-8")
            semantic_ai.run_codex_study_draft(
                source=draft["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )
            # This test exercises a second transport invocation. Remove the
            # valid checkpoint so the production resume path does not
            # intentionally promote it without another Codex call.
            output.unlink()
            semantic_ai.run_codex_study_draft(
                source=draft["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
                structured_output=True,
            )

            self.assertEqual(3, len(paths))
            self.assertEqual(paths[0][0], paths[1][0])
            self.assertNotEqual(paths[0][0], paths[2][0])
            self.assertIsNotNone(paths[1][1])
            self.assertIsNotNone(paths[2][1])
            self.assertNotEqual(paths[1][1], paths[2][1])
            for last_message, schema in paths:
                self.assertEqual(parent, last_message.parent)
                self.assertFalse(last_message.exists())
                if schema is not None:
                    self.assertEqual(parent, schema.parent)
                    self.assertFalse(schema.exists())
            self.assertTrue(sentinel.is_file())

    def test_study_transport_paths_are_cleaned_after_command_failure(
        self,
    ) -> None:
        draft = self.study_draft()
        captured: list[tuple[Path, Path]] = []

        def fake_run(
            command: list[str],
            **_: object,
        ) -> subprocess.CompletedProcess[str]:
            last_message = Path(
                command[command.index("--output-last-message") + 1]
            )
            schema = Path(command[command.index("--output-schema") + 1])
            captured.append((last_message, schema))
            self.assertTrue(schema.is_file())
            last_message.write_text('{"partial":', encoding="utf-8")
            return subprocess.CompletedProcess(
                command,
                9,
                stdout='noise followed by {"not":"transport output"}',
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            parent = Path(temp_dir)
            output = parent / "study.json"
            sentinel = parent / ".caller-owned.keep"
            sentinel.write_text("keep", encoding="utf-8")
            with self.assertRaisesRegex(
                semantic_ai.SemanticAiError,
                "failed with exit code 9",
            ):
                semantic_ai.run_codex_study_draft(
                    source=draft["source"],
                    workbook={"semanticCellCoverageComplete": True},
                    locator_results=[self.result()],
                    focused_chunks=[self.chunk()],
                    content_complete=True,
                    output_path=output,
                    run_command=fake_run,
                    structured_output=True,
                )

            self.assertEqual(1, len(captured))
            last_message, schema = captured[0]
            self.assertEqual(parent, last_message.parent)
            self.assertEqual(parent, schema.parent)
            self.assertFalse(last_message.exists())
            self.assertFalse(schema.exists())
            self.assertTrue(sentinel.is_file())

    def test_codex_study_runner_bounds_non_json_retry_to_one(
        self,
    ) -> None:
        calls = 0

        def fake_run(
            command: list[str],
            **_: object,
        ) -> subprocess.CompletedProcess[str]:
            nonlocal calls
            calls += 1
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                "{not-json",
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        draft = self.study_draft()
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(
                semantic_ai.SemanticAiError,
                "after one bounded schema retry",
            ):
                semantic_ai.run_codex_study_draft(
                    source=draft["source"],
                    workbook={"semanticCellCoverageComplete": True},
                    locator_results=[self.result()],
                    focused_chunks=[self.chunk()],
                    content_complete=True,
                    output_path=Path(temp_dir) / "study.json",
                    run_command=fake_run,
                )

        self.assertEqual(2, calls)

    def test_codex_study_runner_restores_deterministic_packet_coverage(
        self,
    ) -> None:
        draft = self.study_draft()
        draft["source"]["contentComplete"] = False

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(draft, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            result = semantic_ai.run_codex_study_draft(
                source=draft["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=Path(temp_dir) / "study.json",
                run_command=fake_run,
            )

        self.assertIs(True, result["source"]["contentComplete"])

    def test_codex_temporal_stage_repair_uses_current_manifest_as_baseline(
        self,
    ) -> None:
        rejected, chunk = self.temporal_stage_draft()
        corrected = json.loads(json.dumps(rejected))
        stage_factor = corrected["studies"][0]["factors"][-1]
        stage_factor["baselineCondition"] = "Before"
        stage_factor["changedCondition"] = "After"
        corrected["studies"][0]["arms"][0]["factorValues"][-1][
            "isBaseline"
        ] = True
        prompts: list[str] = []

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(corrected, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[chunk],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )

        self.assertEqual(1, len(prompts))
        self.assertIn(
            "source-authored temporal stage baseline mapping",
            prompts[0],
        )
        self.assertIn("perform an exact minimal repair", prompts[0])
        self.assertEqual(
            "Before",
            result["studies"][0]["factors"][-1][
                "baselineCondition"
            ],
        )
        self.assertTrue(
            result["studies"][0]["arms"][0]["factorValues"][-1][
                "isBaseline"
            ]
        )

    def test_codex_study_runner_promotes_current_valid_rejected_draft(
        self,
    ) -> None:
        corrected = self.study_draft()

        def unexpected_run(
            _command: list[str],
            **_kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            raise AssertionError("A validated rejected draft must be reused")

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(corrected, ensure_ascii=False),
                encoding="utf-8",
            )

            result = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=unexpected_run,
            )

            self.assertTrue(output.is_file())
            self.assertEqual(corrected["studies"], result["studies"])

    def test_codex_study_runner_repairs_rejected_reference_without_resending_cells(
        self,
    ) -> None:
        corrected = self.study_draft()
        rejected = json.loads(json.dumps(corrected))
        rejected["studies"][0]["outcomes"][0]["observations"][0][
            "arm"
        ] = "candidate-a-typo"
        prompts: list[str] = []

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(corrected, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )

        self.assertEqual(1, len(prompts))
        self.assertIn("VALIDATOR ERROR", prompts[0])
        self.assertIn("references unknown arm", prompts[0])
        self.assertIn("REJECTED FULL JSON", prompts[0])
        self.assertNotIn("FOCUSED SOURCE PACKET", prompts[0])
        self.assertEqual(
            "candidate-a",
            result["studies"][0]["outcomes"][0]["observations"][0][
                "arm"
            ],
        )

    def test_codex_reference_repair_cannot_change_other_source_content(
        self,
    ) -> None:
        corrected = self.study_draft()
        rejected = json.loads(json.dumps(corrected))
        rejected["studies"][0]["outcomes"][0]["observations"][0][
            "arm"
        ] = "candidate-a-typo"
        unsafe = json.loads(json.dumps(corrected))
        unsafe["studies"][0]["summary"] = "Changed by repair"

        def fake_run(
            command: list[str],
            **_: object,
        ) -> subprocess.CompletedProcess[str]:
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(unsafe, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            rejected_path = output.with_name("study.rejected.json")
            rejected_path.write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(
                semantic_ai.SemanticAiError,
                "outside the allowed reference paths",
            ):
                semantic_ai.run_codex_study_draft(
                    source=corrected["source"],
                    workbook={"semanticCellCoverageComplete": True},
                    locator_results=[self.result()],
                    focused_chunks=[self.chunk()],
                    content_complete=True,
                    output_path=output,
                    run_command=fake_run,
                )
            preserved = json.loads(
                rejected_path.read_text(encoding="utf-8")
            )
            self.assertEqual(
                "candidate-a-typo",
                preserved["studies"][0]["outcomes"][0][
                    "observations"
                ][0]["arm"],
            )
            self.assertTrue(
                output.with_name("study.repair-rejected.json").is_file()
            )
            self.assertTrue(
                output.with_name(
                    "study.repair-rejected.unsafe.json"
                ).is_file()
            )

            resume_prompts: list[str] = []

            def safe_resume_run(
                command: list[str],
                **kwargs: object,
            ) -> subprocess.CompletedProcess[str]:
                resume_prompts.append(str(kwargs["input"]))
                output_index = command.index("--output-last-message") + 1
                Path(command[output_index]).write_text(
                    json.dumps(corrected, ensure_ascii=False),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess(
                    command,
                    0,
                    stdout="",
                    stderr="",
                )

            resumed = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=safe_resume_run,
            )
            self.assertEqual(1, len(resume_prompts))
            self.assertIn("candidate-a-typo", resume_prompts[0])
            self.assertEqual(
                corrected["studies"][0]["summary"],
                resumed["studies"][0]["summary"],
            )

    def test_codex_general_contract_repair_resends_focused_source(
        self,
    ) -> None:
        corrected = self.study_draft()
        rejected = json.loads(json.dumps(corrected))
        rejected["studies"][0]["arms"][0]["label"] = ""
        prompts: list[str] = []

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(corrected, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )

        self.assertEqual(1, len(prompts))
        self.assertIn("label is required", prompts[0])
        self.assertIn("FOCUSED SOURCE PACKET", prompts[0])
        self.assertIn("REPAIR OVERRIDE", prompts[0])
        self.assertEqual(
            corrected["studies"][0]["arms"][0]["role"],
            result["studies"][0]["arms"][0]["role"],
        )

    def test_codex_numeric_text_homogeneity_repair_retries_once(
        self,
    ) -> None:
        rejected = self.study_draft()
        observations = rejected["studies"][0]["outcomes"][0][
            "observations"
        ]
        numeric_text = json.loads(json.dumps(observations[0]))
        numeric_text.update(
            {
                "key": "candidate-a-text-number",
                "replicateKey": "position-3",
                "valueNumber": None,
                "valueText": "-0.006",
            }
        )
        observations.append(numeric_text)
        corrected = json.loads(json.dumps(rejected))
        corrected["studies"][0]["outcomes"][0]["observations"][1][
            "valueNumber"
        ] = -0.006
        prompts: list[str] = []

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            response = rejected if len(prompts) == 1 else corrected
            Path(command[output_index]).write_text(
                json.dumps(response, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            result = semantic_ai.run_codex_study_draft(
                source=rejected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=Path(temp_dir) / "study.json",
                run_command=fake_run,
            )

        repaired = result["studies"][0]["outcomes"][0][
            "observations"
        ][1]
        self.assertEqual(1, len(prompts))
        self.assertEqual(-0.006, repaired["valueNumber"])
        self.assertEqual("-0.006", repaired["valueText"])
        self.assertEqual(numeric_text["evidence"], repaired["evidence"])

    def test_numeric_text_repair_preserves_all_non_target_fields(
        self,
    ) -> None:
        baseline = self.study_draft()
        observations = baseline["studies"][0]["outcomes"][0][
            "observations"
        ]
        numeric_text = json.loads(json.dumps(observations[0]))
        numeric_text.update(
            {
                "key": "candidate-a-text-number",
                "replicateKey": "position-3",
                "valueNumber": None,
                "valueText": "-0.010",
            }
        )
        numeric_text["evidence"][0]["note"] = "Exact original note"
        observations.append(numeric_text)
        baseline["studies"][0]["limitations"] = ["Exact original limitation"]

        repaired = (
            semantic_ai._apply_deterministic_numeric_text_homogeneity_repair(
                baseline
            )
        )

        expected = json.loads(json.dumps(baseline))
        expected["studies"][0]["outcomes"][0]["observations"][1][
            "valueNumber"
        ] = -0.01
        self.assertEqual(expected, repaired)

    def test_numeric_text_repair_is_fail_closed_for_categorical_or_other_changes(
        self,
    ) -> None:
        baseline = self.study_draft()
        observations = baseline["studies"][0]["outcomes"][0][
            "observations"
        ]
        numeric_text = json.loads(json.dumps(observations[0]))
        numeric_text.update(
            {
                "key": "candidate-a-text-number",
                "replicateKey": "position-3",
                "valueNumber": None,
                "valueText": "-0.006",
            }
        )
        observations.append(numeric_text)
        unsafe = json.loads(json.dumps(baseline))
        unsafe["studies"][0]["outcomes"][0]["observations"][1][
            "valueNumber"
        ] = -0.006
        unsafe["studies"][0]["summary"] = "Unrelated repair change"

        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "outside the allowed valueNumber paths",
        ):
            semantic_ai._validate_numeric_text_homogeneity_repair(
                baseline,
                unsafe,
            )

        categorical = json.loads(json.dumps(baseline))
        categorical_observation = categorical["studies"][0]["outcomes"][0][
            "observations"
        ][1]
        categorical_observation["valueText"] = "Pass"
        self.assertEqual(
            {},
            semantic_ai._numeric_text_homogeneity_repair_targets(
                categorical
            ),
        )

    def test_incompatible_raw_comparison_repair_projection_is_exact(
        self,
    ) -> None:
        baseline = self.incompatible_raw_comparison_draft()
        preserved = json.loads(
            json.dumps(
                baseline["studies"][2]["comparisons"][0]
            )
        )
        preserved["key"] = "preserved-comparison"
        baseline["studies"][2]["comparisons"].append(preserved)
        error = (
            "ValueError: studies[2].comparisons[0] shared Outcome "
            "'gauss-measurement' RAW representations require compatible "
            "value units and aligned ordered axis identity, shape, and "
            "stratum; omit the invalid Comparison, preserve its "
            "Arms/Outcomes/series, and add a limitation"
        )
        target = (
            semantic_ai._incompatible_raw_comparison_target(error)
        )
        self.assertEqual((2, 0, "gauss-measurement"), target)
        self.assertIsNone(
            semantic_ai._incompatible_raw_comparison_target(
                error.replace(
                    "RAW representations require compatible value units "
                    "and aligned ordered axis identity, shape, and stratum",
                    "scalar representations require a compatible field",
                )
            )
        )

        repaired = (
            semantic_ai
            ._apply_deterministic_incompatible_raw_comparison_repair(
                baseline,
                target,
            )
        )
        expected = json.loads(json.dumps(baseline))
        removed = expected["studies"][2]["comparisons"].pop(0)
        expected["studies"][2]["limitations"].append(
            semantic_ai._incompatible_raw_comparison_limitation(
                removed,
                "gauss-measurement",
            )
        )
        self.assertEqual(expected, repaired)
        self.assertEqual(
            [preserved],
            repaired["studies"][2]["comparisons"],
        )

        unsafe = json.loads(json.dumps(repaired))
        unsafe["studies"][2]["arms"][0]["label"] = (
            "Synthetic repaired arm"
        )
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "outside the one validator-identified Comparison",
        ):
            semantic_ai._validate_incompatible_raw_comparison_repair(
                baseline,
                unsafe,
                target,
            )
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "does not match declared Outcome",
        ):
            semantic_ai._apply_deterministic_incompatible_raw_comparison_repair(
                baseline,
                (2, 0, "other-outcome"),
            )

    def test_incompatible_raw_comparison_repair_uses_no_ai_call(
        self,
    ) -> None:
        baseline = self.incompatible_raw_comparison_draft()
        study_before = baseline["studies"][2]
        comparison_error = (
            "studies[2].comparisons[0] shared Outcome "
            "'gauss-measurement' RAW representations require compatible "
            "value units and aligned ordered axis identity, shape, and "
            "stratum; omit the invalid Comparison, preserve its "
            "Arms/Outcomes/series, and add a limitation"
        )

        def comparison_validator(draft: dict) -> None:
            if draft["studies"][2]["comparisons"]:
                raise ValueError(comparison_error)

        exact_prompt = "exact budgeted B09 prompt"
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(baseline, ensure_ascii=False),
                encoding="utf-8",
            )
            unsafe_artifact = json.loads(json.dumps(baseline))
            unsafe_artifact["studies"][2]["arms"].append(
                {
                    "key": "synthetic-repair-arm",
                    "role": "OTHER",
                    "label": "Synthetic",
                    "condition": "Synthetic",
                    "sampleSize": None,
                    "sampleBasis": "",
                    "matchingBasis": "",
                    "factorValues": [],
                    "evidence": json.loads(
                        json.dumps(study_before["evidence"])
                    ),
                }
            )
            output.with_name(
                "study.repair-rejected.json"
            ).write_text(
                json.dumps(unsafe_artifact, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=baseline["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                additional_validator=comparison_validator,
                exact_prompt_text=exact_prompt,
                expected_prompt_sha256=hashlib.sha256(
                    exact_prompt.encode("utf-8")
                ).hexdigest(),
                run_command=mock.Mock(
                    side_effect=AssertionError(
                        "Deterministic comparison repair must not call AI"
                    )
                ),
            )

        repaired_study = result["studies"][2]
        self.assertEqual([], repaired_study["comparisons"])
        for field in (
            "arms",
            "factors",
            "outcomes",
            "measurementSeries",
            "conclusions",
        ):
            self.assertEqual(
                study_before[field],
                repaired_study[field],
            )
        self.assertEqual(
            study_before["limitations"]
            + [
                semantic_ai._incompatible_raw_comparison_limitation(
                    study_before["comparisons"][0],
                    "gauss-measurement",
                )
            ],
            repaired_study["limitations"],
        )
        self.assertNotIn(
            "synthetic-repair-arm",
            {
                arm["key"]
                for arm in repaired_study["arms"]
            },
        )

    def test_split_cell_reference_repair_is_exact_and_safe(
        self,
    ) -> None:
        baseline, focused_chunks = self.synthetic_split_arm_draft()
        validation_error = (
            "ValueError: studies[2].arms[2].role REFERENCE requires "
            "directly cited captured full Normal, Reference, Standard, "
            "Spec, or equivalent reference wording matching the Arm "
            "label or condition, or at least two exact ordered distinct "
            "full reference #N identity cells for a descriptive grouped "
            "Arm; a bare abbreviation such as ST or mixed Test/Normal "
            "evidence is not reference semantics"
        )
        target = semantic_ai._synthetic_reference_arm_target(
            validation_error
        )
        self.assertEqual((2, 2), target)
        self.assertIsNone(
            semantic_ai._synthetic_reference_arm_target(
                validation_error.replace(
                    "REFERENCE requires directly cited captured full",
                    "REFERENCE has unrelated validation wording for",
                )
            )
        )
        plan = semantic_ai._synthetic_arm_repair_plan(
            baseline,
            study_index=2,
            focused_chunks=focused_chunks,
        )
        self.assertEqual(
            {
                0: "Test",
                1: "Test",
                2: "Normal",
                3: "Normal",
            },
            plan,
        )

        repaired = (
            semantic_ai._apply_deterministic_synthetic_reference_arm_repair(
                baseline,
                target=target,
                focused_chunks=focused_chunks,
            )
        )
        arms = repaired["studies"][2]["arms"]
        self.assertEqual(
            [
                ("TEST", "Test", "Test"),
                ("TEST", "Test", "Test"),
                ("REFERENCE", "Normal", "Normal"),
                ("REFERENCE", "Normal", "Normal"),
            ],
            [
                (arm["role"], arm["label"], arm["condition"])
                for arm in arms[:4]
            ],
        )
        self.assertEqual(
            baseline["studies"][2]["arms"][4:],
            arms[4:],
        )
        self.assertEqual(
            [
                arm["factorValues"]
                for arm in baseline["studies"][2]["arms"][:4]
            ],
            [arm["factorValues"] for arm in arms[:4]],
        )

        unsafe = json.loads(json.dumps(repaired))
        unsafe["studies"][2]["arms"][2]["factorValues"][1][
            "value"
        ] = "Synthetic"
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "outside the planned exact source label",
        ):
            semantic_ai._validate_synthetic_reference_arm_repair(
                baseline,
                unsafe,
                target=target,
                focused_chunks=focused_chunks,
            )
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "not an exact split-cell unsupported REFERENCE",
        ):
            semantic_ai._apply_deterministic_synthetic_reference_arm_repair(
                baseline,
                target=(2, 4),
                focused_chunks=focused_chunks,
            )

    def test_newer_rejected_split_reference_repair_precedes_stale_target(
        self,
    ) -> None:
        baseline, focused_chunks = self.synthetic_split_arm_draft()
        validation_error = (
            "studies[2].arms[2].role REFERENCE requires directly cited "
            "captured full Normal, Reference, Standard, Spec, or "
            "equivalent reference wording matching the Arm label or "
            "condition, or at least two exact ordered distinct full "
            "reference #N identity cells for a descriptive grouped Arm; "
            "a bare abbreviation such as ST or mixed Test/Normal evidence "
            "is not reference semantics"
        )

        def source_validator(draft: dict) -> None:
            arms = draft["studies"][2]["arms"]
            if arms[2]["label"] != "Normal":
                raise ValueError(validation_error)
            identities = {
                arm["condition"]
                for arm in arms[:4]
            }
            if not {"Test", "Normal"}.issubset(identities):
                raise ValueError(
                    "Test!E30 and Test!E32 semantic identities are missing"
                )

        exact_prompt = "exact split-cell deterministic repair prompt"
        run_command = mock.Mock(
            side_effect=AssertionError(
                "Deterministic split-cell repair must not call AI"
            )
        )
        ai_call_observer = mock.Mock()
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            rejected = output.with_name("study.rejected.json")
            output.write_text(
                json.dumps({"stale": True}),
                encoding="utf-8",
            )
            rejected.write_text(
                json.dumps(baseline, ensure_ascii=False),
                encoding="utf-8",
            )
            stale_time = 1_700_000_000_000_000_000
            current_time = 1_700_000_001_000_000_000
            os.utime(output, ns=(stale_time, stale_time))
            os.utime(rejected, ns=(current_time, current_time))
            result = semantic_ai.run_codex_study_draft(
                source=baseline["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=focused_chunks,
                content_complete=True,
                output_path=output,
                additional_validator=source_validator,
                exact_prompt_text=exact_prompt,
                expected_prompt_sha256=hashlib.sha256(
                    exact_prompt.encode("utf-8")
                ).hexdigest(),
                run_command=run_command,
                ai_call_observer=ai_call_observer,
            )

        self.assertEqual(0, run_command.call_count)
        self.assertEqual(0, ai_call_observer.call_count)
        arms = result["studies"][2]["arms"]
        self.assertEqual("Test", arms[0]["label"])
        self.assertEqual("Normal", arms[2]["label"])
        self.assertEqual("REFERENCE", arms[2]["role"])
        self.assertEqual(
            baseline["studies"][2]["arms"][4:],
            arms[4:],
        )

    def test_a1_union_evidence_repair_splits_only_exact_ranges(
        self,
    ) -> None:
        baseline = json.loads(json.dumps(self.study_draft()))
        evidence = baseline["studies"][0]["outcomes"][0][
            "evidence"
        ][0]
        evidence.update(
            {
                "range": "H7,H9,H11",
                "sourceText": "Exact source text",
                "note": "Exact note",
            }
        )
        repaired = (
            semantic_ai._apply_deterministic_a1_union_evidence_repair(
                baseline
            )
        )
        expected = json.loads(json.dumps(baseline))
        expected["studies"][0]["outcomes"][0]["evidence"] = [
            {
                **json.loads(json.dumps(evidence)),
                "range": address,
            }
            for address in ("H7", "H9", "H11")
        ]
        self.assertEqual(expected, repaired)
        self.assertTrue(
            semantic_ai._a1_union_evidence_repair_applicable(
                "SemanticAiError: invalid canonical study draft: "
                "studies[0].outcomes[0].evidence[0].range must be an "
                "A1 cell or range",
                baseline,
            )
        )

        unsafe = json.loads(json.dumps(repaired))
        unsafe["studies"][0]["summary"] = "Unrelated change"
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "outside the exact evidence range split",
        ):
            semantic_ai._validate_a1_union_evidence_repair(
                baseline,
                unsafe,
            )
        invalid = json.loads(json.dumps(baseline))
        invalid["studies"][0]["outcomes"][0]["evidence"][0][
            "range"
        ] = "H7,not-a-cell"
        self.assertFalse(
            semantic_ai._a1_union_evidence_repair_applicable(
                "studies[0].outcomes[0].evidence[0].range must be an "
                "A1 cell or range",
                invalid,
            )
        )

    def test_a1_union_evidence_repair_uses_no_ai_call(
        self,
    ) -> None:
        baseline = json.loads(json.dumps(self.study_draft()))
        baseline["studies"][0]["outcomes"][0]["evidence"][0][
            "range"
        ] = "H7,H9,H11"
        exact_prompt = "exact budgeted B15 prompt"
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(baseline, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=baseline["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                exact_prompt_text=exact_prompt,
                expected_prompt_sha256=hashlib.sha256(
                    exact_prompt.encode("utf-8")
                ).hexdigest(),
                run_command=mock.Mock(
                    side_effect=AssertionError(
                        "Deterministic A1 split must not call AI"
                    )
                ),
            )

        self.assertEqual(
            ["H7", "H9", "H11"],
            [
                item["range"]
                for item in result["studies"][0]["outcomes"][0][
                    "evidence"
                ]
            ],
        )

    def test_codex_unsupported_rate_pair_repair_only_clears_pair(
        self,
    ) -> None:
        rejected = self.study_draft()
        observation = rejected["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        observation.update(
            {
                "valueNumber": 5,
                "valueText": "5.0%",
                "numerator": 3,
                "denominator": 60,
                "sampleSize": 60,
            }
        )
        corrected = (
            semantic_ai._apply_deterministic_unsupported_rate_pair_repair(
                rejected,
                (0, 0, 0),
            )
        )
        prompts: list[str] = []

        def source_validator(draft: dict) -> None:
            candidate = draft["studies"][0]["outcomes"][0][
                "observations"
            ][0]
            if candidate["numerator"] is not None:
                raise ValueError(
                    "studies[0].outcomes[0].observations[0].numerator=3 "
                    "is not present in its cited Capture v2 cells"
                )

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            result = semantic_ai.run_codex_study_draft(
                source=rejected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=Path(temp_dir) / "study.json",
                additional_validator=source_validator,
                run_command=fake_run,
            )

        result_observation = result["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        self.assertEqual(1, len(prompts))
        self.assertIsNone(result_observation["numerator"])
        self.assertIsNone(result_observation["denominator"])
        for field in (
            "valueNumber",
            "valueText",
            "sampleSize",
            "evidence",
        ):
            self.assertEqual(observation[field], result_observation[field])

        unsafe = json.loads(json.dumps(corrected))
        unsafe["studies"][0]["summary"] = "Unsafe unrelated mutation"
        with self.assertRaisesRegex(
            semantic_ai.SemanticAiError,
            "outside the allowed numerator and denominator paths",
        ):
            semantic_ai._validate_unsupported_rate_pair_repair(
                rejected,
                unsafe,
                (0, 0, 0),
            )

    def test_rate_pair_repair_discards_unsafe_repair_artifact(
        self,
    ) -> None:
        baseline = self.study_draft()
        observation = baseline["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        observation.update(
            {
                "valueNumber": 5,
                "valueText": "5.0%",
                "numerator": 3,
                "denominator": 60,
                "sampleSize": 60,
            }
        )
        unsafe_rewrite = json.loads(json.dumps(baseline))
        unsafe_rewrite["studies"][0]["summary"] = "Unsafe model rewrite"

        def source_validator(draft: dict) -> None:
            candidate = draft["studies"][0]["outcomes"][0][
                "observations"
            ][0]
            if candidate["numerator"] is not None:
                raise ValueError(
                    "studies[0].outcomes[0].observations[0].numerator=3 "
                    "is not present in its cited Capture v2 cells"
                )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(baseline, ensure_ascii=False),
                encoding="utf-8",
            )
            output.with_name("study.repair-rejected.json").write_text(
                json.dumps(unsafe_rewrite, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=baseline["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                additional_validator=source_validator,
                run_command=mock.Mock(
                    side_effect=AssertionError(
                        "Deterministic repair must not call the model"
                    )
                ),
            )

        repaired_observation = result["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        self.assertIsNone(repaired_observation["numerator"])
        self.assertIsNone(repaired_observation["denominator"])
        self.assertEqual(
            baseline["studies"][0]["summary"],
            result["studies"][0]["summary"],
        )

    def test_latest_valid_repair_checkpoint_is_promoted_without_ai_call(
        self,
    ) -> None:
        checkpoint = self.study_draft()
        newer_invalid = self.study_draft()
        newer_invalid["studies"][0]["arms"][0]["label"] = ""

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.repair-rejected.json").write_text(
                json.dumps(checkpoint, ensure_ascii=False),
                encoding="utf-8",
            )
            output.with_name("study.rejected.json").write_text(
                json.dumps(newer_invalid, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=checkpoint["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=mock.Mock(
                    side_effect=AssertionError(
                        "A valid checkpoint must not call the model"
                    )
                ),
            )
            persisted = json.loads(output.read_text(encoding="utf-8"))

        self.assertEqual(checkpoint, result)
        self.assertEqual(checkpoint, persisted)

    def test_rate_pair_repair_uses_trusted_repair_baseline_without_ai(
        self,
    ) -> None:
        rejected = self.study_draft()
        rejected["studies"][0]["arms"][0]["label"] = ""
        repair_baseline = self.study_draft()
        observation = repair_baseline["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        observation.update(
            {
                "valueNumber": 5,
                "valueText": "5.0%",
                "numerator": 3,
                "denominator": 60,
                "sampleSize": 60,
            }
        )
        second_outcome = json.loads(
            json.dumps(repair_baseline["studies"][0]["outcomes"][0])
        )
        second_outcome["key"] = "second-rate-outcome"
        second_outcome["originalLabel"] = "Second rate"
        second_observation = second_outcome["observations"][0]
        second_observation["key"] = "second-rate-observation"
        second_observation.update(
            {
                "valueNumber": 10,
                "valueText": "10.0%",
                "numerator": 6,
                "denominator": 60,
                "sampleSize": 60,
            }
        )
        repair_baseline["studies"][0]["outcomes"].append(second_outcome)
        validation_calls = 0

        def source_validator(draft: dict) -> None:
            nonlocal validation_calls
            validation_calls += 1
            for outcome_index, outcome in enumerate(
                draft["studies"][0]["outcomes"]
            ):
                candidate = outcome["observations"][0]
                if candidate["denominator"] is not None:
                    raise ValueError(
                        "studies[0].outcomes"
                        f"[{outcome_index}].observations[0]"
                        ".denominator=60 is not present in its cited "
                        "Capture v2 cells"
                    )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            rejected_path = output.with_name("study.rejected.json")
            repair_path = output.with_name(
                "study.repair-rejected.json"
            )
            rejected_path.write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            repair_path.write_text(
                json.dumps(repair_baseline, ensure_ascii=False),
                encoding="utf-8",
            )
            rejected_mtime = rejected_path.stat().st_mtime_ns
            os.utime(
                repair_path,
                ns=(
                    rejected_mtime + 1_000_000_000,
                    rejected_mtime + 1_000_000_000,
                ),
            )
            result = semantic_ai.run_codex_study_draft(
                source=repair_baseline["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                additional_validator=source_validator,
                unsupported_rate_pair_paths=lambda _draft: [
                    (0, 0, 0),
                    (0, 1, 0),
                ],
                run_command=mock.Mock(
                    side_effect=AssertionError(
                        "Trusted repair baseline must be fixed locally"
                    )
                ),
            )

        repaired = result["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        self.assertIsNone(repaired["numerator"])
        self.assertIsNone(repaired["denominator"])
        second_repaired = result["studies"][0]["outcomes"][1][
            "observations"
        ][0]
        self.assertIsNone(second_repaired["numerator"])
        self.assertIsNone(second_repaired["denominator"])
        self.assertEqual(2, validation_calls)

    def test_codex_dense_series_source_failure_uses_general_repair(
        self,
    ) -> None:
        rejected = self.study_draft()
        series = {
            "key": "unseen-profile",
            "seriesRole": "RAW",
            "aggregationFunction": "",
            "aggregateOfSeries": [],
            "outcome": "unseen-response",
            "arm": "candidate-a",
            "sheet": "Unfamiliar data",
            "headerRange": "C3",
            "valueRange": "C4:C5",
            "rowIdentityRange": "E4:E5",
            "aggregateReplicateRanges": [],
            "axisSource": "ROW_IDENTITY",
            "axisLabel": "Source axis",
            "axisUnit": "",
            "valueUnit": "",
            "stratumKey": "",
            "verificationStatus": "NEEDS_REVIEW",
        }
        rejected["studies"][0]["measurementSeries"] = [series]
        corrected = json.loads(json.dumps(rejected))
        corrected["studies"][0]["measurementSeries"][0][
            "rowIdentityRange"
        ] = "B4:B5"
        prompts: list[str] = []

        def source_validator(draft: dict) -> None:
            row_identity = draft["studies"][0]["measurementSeries"][0][
                "rowIdentityRange"
            ]
            if row_identity == "E4:E5":
                raise ValueError(
                    "studies[0].measurementSeries[0].rowIdentityRange "
                    "source cell E4 has no value"
                )

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(corrected, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                additional_validator=source_validator,
                run_command=fake_run,
            )

        self.assertEqual(1, len(prompts))
        self.assertIn(
            "rowIdentityRange source cell E4 has no value",
            prompts[0],
        )
        self.assertIn("FOCUSED SOURCE PACKET", prompts[0])
        self.assertEqual(
            "B4:B5",
            result["studies"][0]["measurementSeries"][0][
                "rowIdentityRange"
            ],
        )

    def test_codex_incomplete_count_pair_uses_grounded_general_repair(
        self,
    ) -> None:
        rejected = self.study_draft()
        observation = rejected["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        observation["valueNumber"] = 1
        observation["valueText"] = "1"
        observation["numerator"] = 1
        observation["denominator"] = None
        corrected = json.loads(json.dumps(rejected))
        corrected_observation = corrected["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        corrected_observation["numerator"] = None
        corrected_observation["denominator"] = None
        prompts: list[str] = []

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            prompts.append(str(kwargs["input"]))
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(corrected, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(
                command,
                0,
                stdout="",
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "study.json"
            output.with_name("study.rejected.json").write_text(
                json.dumps(rejected, ensure_ascii=False),
                encoding="utf-8",
            )
            result = semantic_ai.run_codex_study_draft(
                source=corrected["source"],
                workbook={"semanticCellCoverageComplete": True},
                locator_results=[self.result()],
                focused_chunks=[self.chunk()],
                content_complete=True,
                output_path=output,
                run_command=fake_run,
            )

        repaired_observation = result["studies"][0]["outcomes"][0][
            "observations"
        ][0]
        self.assertEqual(1, len(prompts))
        self.assertIn(
            "numerator and denominator must be supplied together",
            prompts[0],
        )
        self.assertIn("FOCUSED SOURCE PACKET", prompts[0])
        self.assertIn(
            "never invent the missing value",
            prompts[0],
        )
        self.assertEqual(1, repaired_observation["valueNumber"])
        self.assertEqual("1", repaired_observation["valueText"])
        self.assertIsNone(repaired_observation["numerator"])
        self.assertIsNone(repaired_observation["denominator"])
        self.assertEqual(
            observation["evidence"],
            repaired_observation["evidence"],
        )

    def test_codex_study_runner_applies_additional_numeric_validator_before_write(self) -> None:
        draft = self.study_draft()

        def fake_run(command: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            output_index = command.index("--output-last-message") + 1
            Path(command[output_index]).write_text(
                json.dumps(draft, ensure_ascii=False),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "rejected.json"
            with self.assertRaisesRegex(ValueError, "numeric evidence mismatch"):
                semantic_ai.run_codex_study_draft(
                    source=draft["source"],
                    workbook={"semanticCellCoverageComplete": True},
                    locator_results=[self.result()],
                    focused_chunks=[self.chunk()],
                    content_complete=True,
                    output_path=output,
                    additional_validator=lambda _draft: (_ for _ in ()).throw(
                        ValueError("numeric evidence mismatch")
                    ),
                    run_command=fake_run,
                )
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
