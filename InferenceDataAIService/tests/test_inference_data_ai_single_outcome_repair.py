from __future__ import annotations

import copy
import unittest

import inference_data_ai_single_outcome_repair as repair


class SingleOutcomeRepairTests(unittest.TestCase):
    def baseline(self) -> dict:
        outcome_keys = [
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
        ]
        outcomes = [
            {"key": key, "observations": []}
            for key in outcome_keys
        ]
        outcomes[6] = {
            "key": "hearing_plus_1v_noise_ng",
            "originalLabel": "Noise ",
            "metricType": "defect_count",
            "unit": "",
            "favorableDirection": "UNKNOWN",
            "evidence": [
                self.evidence("L4:L6", "Hearing; Noise; 0")
            ],
            "observations": [
                {
                    "key": "hearing_plus_1v_noise_ng_condition_test",
                    "arm": "condition_test",
                    "valueNumber": 0,
                    "valueText": "0",
                    "numerator": 0,
                    "denominator": 57,
                    "sampleSize": 57,
                    "evidence": [
                        self.evidence("L6", "0"),
                        self.evidence("F6", "57"),
                    ],
                }
            ],
        }
        outcomes.extend(
            [
                self.percentage_outcome(
                    "sigma_spl_percentage",
                    "H7",
                    0,
                    "0.0%",
                ),
                self.percentage_outcome(
                    "sigma_thd_percentage",
                    "I7",
                    0,
                    "0.0%",
                ),
                self.percentage_outcome(
                    "sigma_spl_thd_percentage",
                    "J7",
                    0,
                    "0.0%",
                ),
                self.percentage_outcome(
                    "sigma_spl_thd_f0_percentage",
                    "K7",
                    0,
                    "0.0%",
                ),
                self.percentage_outcome(
                    "hearing_plus_1v_touch_percentage",
                    "M7",
                    100,
                    "100.0%",
                ),
            ]
        )
        study_keys = [
            "function_lot_test",
            "nti_mask_profile",
            "spl_numeric_profile",
            "thd_numeric_profile",
            "imp_numeric_profile",
            "fo_result_check",
        ]
        studies = [
            {"key": key, "arms": [], "outcomes": []}
            for key in study_keys
        ]
        studies[0] = {
            "key": "function_lot_test",
            "summary": "must remain unchanged",
            "arms": [
                {
                    "key": "condition_test",
                    "role": "TEST",
                    "label": "Condition test ",
                    "condition": "Condition test ",
                    "sampleSize": 57,
                }
            ],
            "outcomes": outcomes,
        }
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "revisionUid": repair.B22_REVISION_UID,
                "contentSha256": repair.B22_CONTENT_SHA256,
            },
            "workbookAnalysis": {"key": "unchanged"},
            "studies": studies,
        }

    def focused_chunks(self) -> list[dict]:
        cells = [
            self.text_cell("H4", "Sigma ", merge_range="H4:K4"),
            self.text_cell("H5", "SPL "),
            self.text_cell("I5", "THD "),
            self.text_cell("J5", "SPL+THD "),
            self.text_cell("K5", "SPL+THD+F0"),
            self.text_cell(
                "L4",
                "Hearing  ( + 1V ) ",
                merge_range="L4:O4",
            ),
            self.text_cell("L5", "Noise "),
            self.text_cell("M5", "Touch "),
            self.number_cell("F6", 57, merge_range="F6:F7"),
            self.number_cell("L6", 0),
            self.number_cell("M6", 9),
            self.number_cell("H7", 0, number_format="0.0%"),
            self.number_cell("I7", 0, number_format="0.0%"),
            self.number_cell("J7", 0, number_format="0.0%"),
            self.number_cell("K7", 0, number_format="0.0%"),
            self.number_cell("L7", 0, number_format="0.0%"),
            self.number_cell("M7", 1, number_format="0.0%"),
        ]
        return [
            {
                "sheet": {
                    "title": "NG function",
                    "sheetIndex": 1,
                },
                "sourceRevision": {
                    "revisionUid": repair.B22_REVISION_UID,
                    "contentSha256": repair.B22_CONTENT_SHA256,
                },
                "mergedRanges": [
                    {"address": "H4:K4"},
                    {"address": "L4:O4"},
                ],
                "cells": cells,
            }
        ]

    def evidence(self, address: str, source_text: str) -> dict:
        return {
            "sheet": "NG function",
            "range": address,
            "role": "SOURCE",
            "sourceText": source_text,
            "note": "",
        }

    def percentage_outcome(
        self,
        key: str,
        coordinate: str,
        value_number: int,
        value_text: str,
    ) -> dict:
        return {
            "key": key,
            "metricType": "percentage",
            "unit": "%",
            "observations": [
                {
                    "key": key + "_condition_test",
                    "arm": "condition_test",
                    "valueNumber": value_number,
                    "valueText": value_text,
                    "sampleSize": 57,
                    "evidence": [
                        self.evidence(coordinate, value_text),
                        self.evidence("F6", "57"),
                    ],
                }
            ],
        }

    def text_cell(
        self,
        coordinate: str,
        value: str,
        *,
        merge_range: str | None = None,
    ) -> dict:
        return {
            "coordinate": coordinate,
            "dataType": "s",
            "rawValue": value,
            "displayValue": value,
            "numberFormat": "General",
            "mergeRange": merge_range,
            "mergeRole": "anchor" if merge_range else "none",
            "primary": True,
            "sourceCellKey": (
                f"{repair.B22_REVISION_UID}:1:{coordinate}"
            ),
        }

    def number_cell(
        self,
        coordinate: str,
        value: int,
        *,
        number_format: str = "General",
        merge_range: str | None = None,
    ) -> dict:
        return {
            "coordinate": coordinate,
            "dataType": "n",
            "rawValue": value,
            "displayValue": value,
            "numberFormat": number_format,
            "mergeRange": merge_range,
            "mergeRole": "anchor" if merge_range else "none",
            "primary": True,
            "sourceCellKey": (
                f"{repair.B22_REVISION_UID}:1:{coordinate}"
            ),
        }

    def cell(self, chunks: list[dict], coordinate: str) -> dict:
        return next(
            cell
            for cell in chunks[0]["cells"]
            if cell["coordinate"] == coordinate
        )

    def apply(self, baseline: dict, chunks: list[dict]) -> dict:
        return repair.apply_deterministic_single_outcome_repair(
            baseline,
            validation_error=repair.B22_VALIDATION_ERROR,
            focused_chunks=chunks,
        )

    def test_exact_repair_inserts_only_source_derived_outcome(self) -> None:
        baseline = self.baseline()
        original = copy.deepcopy(baseline)

        repaired = self.apply(baseline, self.focused_chunks())

        self.assertEqual(original, baseline)
        outcomes = repaired["studies"][0]["outcomes"]
        self.assertEqual(20, len(outcomes))
        self.assertEqual(
            repair.B22_OUTCOME_KEY,
            outcomes[18]["key"],
        )
        self.assertEqual(
            "hearing_plus_1v_touch_percentage",
            outcomes[19]["key"],
        )
        self.assertEqual(
            {
                "key": repair.B22_OUTCOME_KEY,
                "originalLabel": (
                    "Hearing  ( + 1V ) Noise percentage"
                ),
                "metricType": "defect_rate_percent",
                "unit": "%",
                "favorableDirection": "UNKNOWN",
                "evidence": [
                    self.evidence("L7", "0.0%"),
                ],
                "observations": [
                    {
                        "key": repair.B22_OBSERVATION_KEY,
                        "arm": "condition_test",
                        "valueNumber": 0.0,
                        "valueText": "0.0%",
                        "numerator": 0,
                        "denominator": 57,
                        "ratePpm": None,
                        "min": None,
                        "max": None,
                        "average": None,
                        "sampleSize": 57,
                        "evidence": [
                            self.evidence("F6", "57"),
                            self.evidence("L6:L7", "0; 0.0%"),
                        ],
                    }
                ],
            },
            outcomes[18],
        )
        stripped = copy.deepcopy(repaired)
        del stripped["studies"][0]["outcomes"][18]
        self.assertEqual(baseline, stripped)

    def test_repair_is_idempotent(self) -> None:
        chunks = self.focused_chunks()
        repaired = self.apply(self.baseline(), chunks)

        repeated = self.apply(repaired, chunks)

        self.assertEqual(repaired, repeated)
        self.assertEqual(
            1,
            sum(
                outcome["key"] == repair.B22_OUTCOME_KEY
                for outcome in repeated["studies"][0]["outcomes"]
            ),
        )

    def test_projection_rejects_any_unrelated_change(self) -> None:
        baseline = self.baseline()
        chunks = self.focused_chunks()
        repaired = self.apply(baseline, chunks)
        repaired["studies"][0]["summary"] = "unrelated mutation"

        with self.assertRaisesRegex(
            repair.SingleOutcomeRepairError,
            "outside the exact B22",
        ):
            repair.validate_deterministic_single_outcome_repair(
                baseline,
                repaired,
                validation_error=repair.B22_VALIDATION_ERROR,
                focused_chunks=chunks,
            )

    def test_only_the_exact_single_l7_error_is_applicable(self) -> None:
        baseline = self.baseline()
        chunks = self.focused_chunks()
        self.assertTrue(
            repair.single_outcome_repair_applicable(
                baseline,
                validation_error=repair.B22_VALIDATION_ERROR,
                focused_chunks=chunks,
            )
        )
        invalid_errors = [
            repair.B22_VALIDATION_ERROR.replace("1 quantitative", "2 quantitative"),
            repair.B22_VALIDATION_ERROR.replace("L7", "M7"),
            repair.B22_VALIDATION_ERROR
            + "; 1 semantic label cell(s): NG function!C4",
            repair.B22_VALIDATION_ERROR.removeprefix(
                "ContentCoverageError: "
            ),
        ]
        for invalid_error in invalid_errors:
            with self.subTest(error=invalid_error):
                self.assertFalse(
                    repair.single_outcome_repair_applicable(
                        baseline,
                        validation_error=invalid_error,
                        focused_chunks=chunks,
                    )
                )
                with self.assertRaisesRegex(
                    repair.SingleOutcomeRepairError,
                    "exact single B22 L7",
                ):
                    repair.apply_deterministic_single_outcome_repair(
                        baseline,
                        validation_error=invalid_error,
                        focused_chunks=chunks,
                    )

    def test_source_geometry_mismatches_fail_closed(self) -> None:
        mutations = [
            ("L7 raw", "L7", "rawValue", 1),
            ("L7 format", "L7", "numberFormat", "General"),
            ("L4 merge", "L4", "mergeRange", None),
            ("H7 neighbor", "H7", "rawValue", 1),
            ("M7 neighbor", "M7", "rawValue", 0),
            ("source key", "L7", "sourceCellKey", "wrong"),
        ]
        for name, coordinate, field, value in mutations:
            with self.subTest(name=name):
                chunks = self.focused_chunks()
                self.cell(chunks, coordinate)[field] = value
                self.assertFalse(
                    repair.single_outcome_repair_applicable(
                        self.baseline(),
                        validation_error=repair.B22_VALIDATION_ERROR,
                        focused_chunks=chunks,
                    )
                )
                with self.assertRaises(repair.SingleOutcomeRepairError):
                    self.apply(self.baseline(), chunks)

        chunks = self.focused_chunks()
        chunks[0]["cells"] = [
            cell
            for cell in chunks[0]["cells"]
            if cell["coordinate"] != "L7"
        ]
        with self.assertRaisesRegex(
            repair.SingleOutcomeRepairError,
            "lack required",
        ):
            self.apply(self.baseline(), chunks)

    def test_manifest_key_arm_and_value_mismatches_fail_closed(self) -> None:
        baselines = []

        wrong_arm = self.baseline()
        wrong_arm["studies"][0]["arms"][0]["key"] = "other"
        baselines.append(("arm", wrong_arm))

        wrong_count = self.baseline()
        wrong_count["studies"][0]["outcomes"][6]["observations"][0][
            "numerator"
        ] = 1
        baselines.append(("count", wrong_count))

        wrong_neighbor = self.baseline()
        wrong_neighbor["studies"][0]["outcomes"][18]["observations"][0][
            "valueNumber"
        ] = 99
        baselines.append(("neighbor", wrong_neighbor))

        duplicate_target = self.baseline()
        duplicate_target["studies"][0]["outcomes"][0]["key"] = (
            repair.B22_OUTCOME_KEY
        )
        baselines.append(("target key", duplicate_target))

        precovered = self.baseline()
        precovered["studies"][0]["outcomes"][0]["evidence"] = [
            self.evidence("L7", "0.0%")
        ]
        baselines.append(("precovered L7", precovered))

        wrong_source = self.baseline()
        wrong_source["source"]["contentSha256"] = "0" * 64
        baselines.append(("source", wrong_source))

        for name, baseline in baselines:
            with self.subTest(name=name):
                self.assertFalse(
                    repair.single_outcome_repair_applicable(
                        baseline,
                        validation_error=repair.B22_VALIDATION_ERROR,
                        focused_chunks=self.focused_chunks(),
                    )
                )
                with self.assertRaises(repair.SingleOutcomeRepairError):
                    self.apply(baseline, self.focused_chunks())


if __name__ == "__main__":
    unittest.main()
