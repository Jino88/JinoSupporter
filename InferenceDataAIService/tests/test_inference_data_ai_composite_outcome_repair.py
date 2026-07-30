from __future__ import annotations

import copy
import unittest

import inference_data_ai_composite_outcome_repair as repair


class CompositeOutcomeRepairTests(unittest.TestCase):
    awf_count_rows = [25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45]

    def chunks(self) -> list[dict]:
        test_cells = [
            self.text("Test", 1, "F23", "AWF", "F23:F24"),
            self.text("Test", 1, "G23", "SIGMA", "G23:K23"),
            self.text("Test", 1, "L23", "HEARING", "L23:O23"),
        ]
        for coordinate, value in {
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
            test_cells.append(self.text("Test", 1, coordinate, value))
        for row in self.awf_count_rows:
            run_value: str | int = (
                "Total"
                if row in {35, 45}
                else (row - 23) // 2 if row < 37
                else (row - 35) // 2
            )
            test_cells.append(self.value("Test", 1, f"F{row}", run_value))
            sigma_sample = 300 if row == 35 else 240 if row == 45 else 60
            hearing_sample = (
                300 if row == 35 else 239 if row == 45 else 60
            )
            test_cells.append(self.number("Test", 1, f"G{row}", sigma_sample))
            test_cells.append(self.number("Test", 1, f"L{row}", hearing_sample))
            for column in "HIJKMNO":
                count = 1 if column == "J" and row in {39, 45} else 0
                test_cells.append(
                    self.number("Test", 1, f"{column}{row}", count)
                )
                rate = count / sigma_sample if column in "HIJK" else 0
                test_cells.append(
                    self.number(
                        "Test",
                        1,
                        f"{column}{row + 1}",
                        rate,
                        "0.0%",
                    )
                )

        separate_cells = [
            self.text(
                "Separate mold VP",
                4,
                "D13",
                "Mold VP",
                "D13:D14",
            ),
            self.text(
                "Separate mold VP",
                4,
                "E13",
                "IR",
                "E13:E14",
            ),
            self.text(
                "Separate mold VP",
                4,
                "F13",
                "Input",
                "F13:F14",
            ),
            self.text(
                "Separate mold VP",
                4,
                "H13",
                "Sigma",
                "H13:J13",
            ),
            self.text("Separate mold VP", 4, "H14", "SPL"),
            self.text("Separate mold VP", 4, "I14", "SPL+THD"),
            self.text("Separate mold VP", 4, "J14", "SPL+THD+F0"),
        ]
        for row, label, sample in [
            (15, "#5", 490),
            (17, "#9", 587),
            (19, "#12", 492),
            (21, "Total", 1569),
        ]:
            separate_cells.extend(
                [
                    self.value("Separate mold VP", 4, f"D{row}", label),
                    self.number(
                        "Separate mold VP",
                        4,
                        f"E{row}",
                        250418008,
                    ),
                    self.number("Separate mold VP", 4, f"F{row}", sample),
                ]
            )
            for column in "HIJ":
                separate_cells.append(
                    self.number("Separate mold VP", 4, f"{column}{row}", 0)
                )
                if row < 21:
                    separate_cells.append(
                        self.number(
                            "Separate mold VP",
                            4,
                            f"{column}{row + 1}",
                            0,
                            "0.0%",
                        )
                    )
        return [
            self.chunk("Test", 1, test_cells),
            self.chunk("Separate mold VP", 4, separate_cells),
        ]

    def baseline(self) -> dict:
        awf_outcomes = [
            {"key": "awf_value", "observations": []},
            {"key": "awf_sigma_input", "observations": []},
            self.composite("awf_sigma_counts", "G", "K", rate=False),
            self.composite("awf_sigma_rates", "H", "K", rate=True),
            {"key": "awf_hearing_input", "observations": []},
            self.composite("awf_hearing_counts", "L", "O", rate=False),
            self.composite("awf_hearing_rates", "M", "O", rate=True),
            {"key": "awf_total_rate", "observations": []},
        ]
        separate_keys = [
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
        separate_outcomes = [
            {"key": key, "observations": []}
            for key in separate_keys
        ]
        for index, column in zip([2, 3, 4], "HIJ", strict=True):
            separate_outcomes[index]["observations"] = [
                {
                    "key": f"{separate_keys[index]}_{row}",
                    "arm": arm,
                    "valueNumber": 0,
                    "valueText": "0",
                    "sampleSize": sample,
                    "evidence": [
                        self.evidence("Separate mold VP", f"{column}{row}", "0")
                    ],
                }
                for row, arm, sample in [
                    (15, "sep_vp5", 490),
                    (17, "sep_vp9", 587),
                    (19, "sep_vp12", 492),
                    (21, "sep_total", 1569),
                ]
            ]
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "revisionUid": repair.B16_REVISION_UID,
                "contentSha256": repair.B16_CONTENT_SHA256,
            },
            "workbookAnalysis": {"key": "preserved"},
            "studies": [
                {
                    "key": "uc_vp_coil_function_by_head",
                    "arms": [],
                    "outcomes": [],
                },
                {
                    "key": "production_function_by_line_date_mold_vp",
                    "arms": [],
                    "outcomes": [],
                },
                {
                    "key": "awf_function_test_blocks",
                    "summary": "preserved",
                    "arms": [
                        {"key": "awf_test1"},
                        {"key": "awf_test2"},
                    ],
                    "outcomes": awf_outcomes,
                },
                {
                    "key": "separate_all_mold_vp_function",
                    "arms": [
                        {"key": "sep_vp5"},
                        {"key": "sep_vp9"},
                        {"key": "sep_vp12"},
                        {"key": "sep_total"},
                    ],
                    "outcomes": separate_outcomes,
                },
            ],
        }

    def composite(
        self,
        key: str,
        start_column: str,
        end_column: str,
        *,
        rate: bool,
    ) -> dict:
        rows = [
            row + 1 if rate else row
            for row in self.awf_count_rows
        ]
        observations = []
        for index, row in enumerate(rows):
            count_row = row - 1 if rate else row
            arm = "awf_test1" if index < 6 else "awf_test2"
            sample = (
                300
                if count_row == 35
                else 239
                if count_row == 45 and start_column in {"L", "M"}
                else 240
                if count_row == 45
                else 60
            )
            observations.append(
                {
                    "key": f"{key}_{index}",
                    "arm": arm,
                    "valueNumber": None,
                    "valueText": "composite",
                    "sampleSize": sample,
                    "evidence": [
                        self.evidence(
                            "Test",
                            f"{start_column}{row}:{end_column}{row}",
                            "composite",
                        )
                    ],
                }
            )
        return {"key": key, "observations": observations}

    def inventory(self, metadata: dict) -> dict:
        target_items = [
            {
                "sourceCellKey": target["sourceCellKey"],
                "chunkId": "separate",
                "sheet": target["sheet"],
                "coordinate": target["coordinate"],
                "row": int(target["coordinate"][1:]),
                "column": 5,
                "numericValue": 250418008.0,
                "numberFormat": "General",
                "sourceRole": "RESULT",
                "fieldRole": "",
                "classification": "REQUIRED_RESULT",
                "exclusionReason": "",
            }
            for target in metadata["targets"]
        ]
        other = {
            "sourceCellKey": f"{repair.B16_REVISION_UID}:4:F17",
            "sheet": "Separate mold VP",
            "coordinate": "F17",
            "numericValue": 587.0,
            "classification": "REQUIRED_RESULT",
            "exclusionReason": "",
        }
        numeric = [other, *target_items]
        return {
            "schemaVersion": "study-content-coverage-v1",
            "numericCells": copy.deepcopy(numeric),
            "numericCellCount": 4,
            "requiredCells": copy.deepcopy(numeric),
            "requiredCellCount": 4,
            "excludedCells": [],
            "excludedCellCount": 0,
        }

    def apply(self, baseline: dict, chunks: list[dict]) -> dict:
        return repair.apply_deterministic_composite_outcome_repair(
            baseline,
            validation_error=repair.B16_VALIDATION_ERROR,
            focused_chunks=chunks,
        )

    def test_expands_exactly_163_numeric_leaves_and_is_idempotent(self) -> None:
        baseline = self.baseline()
        original = copy.deepcopy(baseline)
        chunks = self.chunks()

        repaired = self.apply(baseline, chunks)

        self.assertEqual(original, baseline)
        awf = repaired["studies"][2]
        separate = repaired["studies"][3]
        self.assertEqual(18, len(awf["outcomes"]))
        self.assertEqual(
            154,
            sum(
                len(outcome["observations"])
                for outcome in awf["outcomes"]
                if outcome["key"].endswith(("_count", "_rate"))
                and outcome["key"].startswith(
                    ("awf_sigma_", "awf_hearing_")
                )
            ),
        )
        self.assertEqual(14, len(separate["outcomes"]))
        self.assertEqual(
            9,
            sum(
                len(outcome["observations"])
                for outcome in separate["outcomes"]
                if outcome["key"] in {
                    "sep_sigma_spl_rate",
                    "sep_sigma_spl_thd_rate",
                    "sep_sigma_spl_thd_f0_rate",
                }
            ),
        )
        self.assertEqual(repaired, self.apply(repaired, chunks))

    def test_projection_rejects_unrelated_changes(self) -> None:
        baseline = self.baseline()
        chunks = self.chunks()
        repaired = self.apply(baseline, chunks)
        repaired["studies"][2]["summary"] = "changed"

        with self.assertRaisesRegex(
            repair.CompositeOutcomeRepairError,
            "outside the exact",
        ):
            repair.validate_deterministic_composite_outcome_repair(
                baseline,
                repaired,
                validation_error=repair.B16_VALIDATION_ERROR,
                focused_chunks=chunks,
            )

    def test_error_geometry_and_composite_mismatches_fail_closed(self) -> None:
        baseline = self.baseline()
        chunks = self.chunks()
        self.assertFalse(
            repair.composite_outcome_repair_applicable(
                baseline,
                validation_error=repair.B16_VALIDATION_ERROR.replace(
                    "163 quantitative",
                    "166 quantitative",
                ),
                focused_chunks=chunks,
            )
        )

        wrong_geometry = self.chunks()
        self.find_cell(wrong_geometry, "Test", "H25")["rawValue"] = "0"
        with self.assertRaises(repair.CompositeOutcomeRepairError):
            self.apply(self.baseline(), wrong_geometry)

        wrong_composite = self.baseline()
        wrong_composite["studies"][2]["outcomes"][2]["observations"][0][
            "valueNumber"
        ] = 0
        with self.assertRaises(repair.CompositeOutcomeRepairError):
            self.apply(wrong_composite, self.chunks())

        wrong_arm = self.baseline()
        wrong_arm["studies"][2]["arms"][0]["key"] = "changed"
        with self.assertRaises(repair.CompositeOutcomeRepairError):
            self.apply(wrong_arm, self.chunks())

    def test_ir_exclusion_metadata_is_exact_and_idempotent(self) -> None:
        chunks = self.chunks()
        metadata = repair.build_b16_inventory_exclusion_metadata(
            self.baseline(),
            focused_chunks=chunks,
        )
        inventory = self.inventory(metadata)

        excluded = repair.apply_b16_inventory_exclusions(
            inventory,
            metadata,
        )

        self.assertEqual(1, excluded["requiredCellCount"])
        self.assertEqual(3, excluded["excludedCellCount"])
        self.assertEqual(
            ["E17", "E19", "E21"],
            [item["coordinate"] for item in excluded["excludedCells"]],
        )
        self.assertEqual(
            excluded,
            repair.apply_b16_inventory_exclusions(excluded, metadata),
        )

        unsafe = copy.deepcopy(excluded)
        unsafe["requiredCells"][0]["numericValue"] = 999
        with self.assertRaisesRegex(
            repair.CompositeOutcomeRepairError,
            "outside E17/E19/E21",
        ):
            repair.validate_b16_inventory_exclusions(
                inventory,
                unsafe,
                metadata,
            )

    def chunk(self, sheet: str, sheet_index: int, cells: list[dict]) -> dict:
        return {
            "sheet": {"title": sheet, "sheetIndex": sheet_index},
            "sourceRevision": {
                "revisionUid": repair.B16_REVISION_UID,
                "contentSha256": repair.B16_CONTENT_SHA256,
            },
            "cells": cells,
        }

    def text(
        self,
        sheet: str,
        sheet_index: int,
        coordinate: str,
        value: str,
        merge_range: str | None = None,
    ) -> dict:
        return {
            "coordinate": coordinate,
            "rawValue": value,
            "displayValue": value,
            "dataType": "s",
            "mergeRange": merge_range,
            "mergeRole": "anchor" if merge_range else "none",
            "primary": True,
            "sourceCellKey": (
                f"{repair.B16_REVISION_UID}:{sheet_index}:{coordinate}"
            ),
        }

    def value(
        self,
        sheet: str,
        sheet_index: int,
        coordinate: str,
        value: str | int,
    ) -> dict:
        if isinstance(value, int):
            return self.number(sheet, sheet_index, coordinate, value)
        return self.text(sheet, sheet_index, coordinate, value)

    def number(
        self,
        sheet: str,
        sheet_index: int,
        coordinate: str,
        value: int | float,
        number_format: str = "General",
    ) -> dict:
        return {
            "coordinate": coordinate,
            "rawValue": value,
            "displayValue": value,
            "dataType": "n",
            "numberFormat": number_format,
            "mergeRange": None,
            "mergeRole": "none",
            "primary": True,
            "sourceCellKey": (
                f"{repair.B16_REVISION_UID}:{sheet_index}:{coordinate}"
            ),
        }

    def evidence(self, sheet: str, address: str, text: str) -> dict:
        return {
            "sheet": sheet,
            "range": address,
            "role": "SOURCE",
            "sourceText": text,
            "note": "",
        }

    def find_cell(
        self,
        chunks: list[dict],
        sheet: str,
        coordinate: str,
    ) -> dict:
        return next(
            cell
            for chunk in chunks
            if chunk["sheet"]["title"] == sheet
            for cell in chunk["cells"]
            if cell["coordinate"] == coordinate
        )


if __name__ == "__main__":
    unittest.main()
