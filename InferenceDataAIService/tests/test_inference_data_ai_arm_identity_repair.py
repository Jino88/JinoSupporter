from __future__ import annotations

import copy
import unittest

import inference_data_ai_arm_identity_repair as repair


class ArmIdentityRepairTests(unittest.TestCase):
    def cell(self, coordinate: str, value: object) -> dict:
        return {
            "coordinate": coordinate,
            "rawValue": value,
            "primary": True,
            "contextOnly": False,
        }

    def chunk(
        self,
        sheet: str,
        values: list[tuple[str, object]],
    ) -> dict:
        return {
            "sheet": {"title": sheet},
            "cells": [
                self.cell(coordinate, value)
                for coordinate, value in values
            ],
        }

    def source(
        self,
        revision_uid: str,
        content_sha256: str,
        file_name: str,
    ) -> dict:
        return {
            "contentComplete": True,
            "revisionUid": revision_uid,
            "contentSha256": content_sha256,
            "sourcePath": rf"D:\Input\{file_name}",
        }

    def keyed(self, keys: list[str], **extra: object) -> list[dict]:
        return [{"key": key, **copy.deepcopy(extra)} for key in keys]

    def b05(self) -> tuple[dict, list[dict]]:
        manifest = {
            "source": self.source(
                repair.B05_REVISION_UID,
                repair.B05_CONTENT_SHA256,
                repair.B05_FILE_NAME,
            ),
            "studies": [
                {
                    "key": "led_uv_peak_total",
                    "arms": self.keyed(
                        [
                            "test_395nm",
                            "normal_365nm_led_uv_1st",
                            "normal_365nm_led_uv_2nd",
                            "normal_365nm_led_uv_3rd",
                        ],
                        label="composite",
                        condition="composite",
                    ),
                },
                {
                    "key": "vision_vp_cd_bond_results",
                    "arms": [],
                },
                {
                    "key": "vp_cd_assembly_tension",
                    "arms": self.keyed(
                        [
                            "ve_562850_test_dry_uv_395nm",
                            "ve_562850_normal_dry_uv_365nm",
                            "ea_16116_test_dry_uv_395nm",
                            "ea_16116_normal_dry_uv_365nm",
                        ],
                        label="composite",
                        condition="composite",
                    ),
                },
            ],
        }
        chunks = [
            self.chunk(
                "161016",
                [
                    ("D22", "Test ( 395nm)"),
                    ("D24", "Normal ( 365nm)"),
                    (
                        "E24",
                        "Led UV 1st\nPeak: 600~900mW/cm²\n"
                        "Total: 2500~3800mW/cm",
                    ),
                    (
                        "E26",
                        "Led UV 2nd\nPeak: 780~900mW/cm²\n"
                        "Total: 2500~3500mW/cm²",
                    ),
                    (
                        "E28",
                        "Led UV 3rd\nPeak: 780~900mW/cm²\n"
                        "Total: 2500~3500mW/cm²",
                    ),
                    ("E38", "Test Dry UV 395nm "),
                    ("E39", "Normal ( Dry UV 365nm)"),
                    ("E40", "Test Dry UV 395nm "),
                    ("E41", "Normal ( Dry UV 365nm)"),
                ],
            )
        ]
        return manifest, chunks

    def b19(self) -> tuple[dict, list[dict]]:
        manifest = {
            "source": self.source(
                repair.B19_REVISION_UID,
                repair.B19_CONTENT_SHA256,
                repair.B19_FILE_NAME,
            ),
            "studies": [
                {"key": "spot_welding_test_normal", "arms": []},
                {
                    "key": "d3_function_test_normal",
                    "arms": self.keyed(
                        [
                            "function_line_r_test",
                            "function_line_r_normal",
                            "function_line_l_test",
                            "function_line_l_normal",
                        ],
                        label="composite",
                        condition="composite",
                    ),
                },
            ],
        }
        chunks = [
            self.chunk(
                "Test",
                [
                    ("E22", "Line R"),
                    ("F22", "Test"),
                    ("F24", "Normal"),
                    ("E26", "Line L"),
                    ("F26", "Test"),
                    ("F28", "Normal"),
                ],
            )
        ]
        return manifest, chunks

    def b27(self) -> tuple[dict, list[dict]]:
        manifest = {
            "source": self.source(
                repair.B27_REVISION_UID,
                repair.B27_CONTENT_SHA256,
                repair.B27_FILE_NAME,
            ),
            "studies": [
                {
                    "key": "condition_matrix_types_1_to_4",
                    "factors": self.keyed(
                        [
                            "c_mg_setting",
                            "s_mg_new_jig_003_setting",
                        ]
                    ),
                    "arms": self.keyed(
                        [
                            "type_1_normal_normal",
                            "type_2_07_normal",
                            "type_3_normal_07",
                            "type_4_07_07",
                        ],
                        role="REFERENCE",
                        factorValues=[],
                    ),
                },
                {"key": "test_normal_actual_dimension"},
                {"key": "sheet2_quantity_check"},
            ],
        }
        values = [
            ("D6", "Type"),
            ("D7", 1),
            ("E7", "Normal"),
            ("F7", "Normal"),
            ("G7", 100),
            ("D8", 2),
            ("E8", 0.7),
            ("F8", "Normal"),
            ("G8", 100),
            ("D9", 3),
            ("E9", "Normal"),
            ("F9", 0.7),
            ("G9", 100),
            ("D10", 4),
            ("E10", 0.7),
            ("F10", 0.7),
            ("G10", 100),
        ]
        return manifest, [self.chunk("Sheet1", values)]

    def test_all_three_exact_projections(self) -> None:
        validation_error = (
            "ValueError: studies[0].arms[0].role REFERENCE requires "
            "directly cited captured full Normal"
        )
        for expected, fixture in (
            ("B05", self.b05),
            ("B19", self.b19),
            ("B27", self.b27),
        ):
            with self.subTest(expected):
                baseline, chunks = fixture()
                untouched = copy.deepcopy(baseline)
                target = repair.arm_identity_repair_target(
                    validation_error,
                    baseline,
                    chunks,
                )
                self.assertEqual(expected, target)
                repaired = repair.apply_arm_identity_repair(
                    baseline,
                    target,
                )
                self.assertEqual(untouched, baseline)
                self.assertNotEqual(baseline, repaired)

        b27, _chunks = self.b27()
        repaired_b27 = repair.apply_arm_identity_repair(b27, "B27")
        study = repaired_b27["studies"][0]
        self.assertEqual("condition_type", study["factors"][0]["key"])
        self.assertEqual("OTHER", study["arms"][0]["role"])
        self.assertEqual(
            [1, 2, 3, 4],
            [
                arm["factorValues"][0]["valueNumber"]
                for arm in study["arms"]
            ],
        )

    def test_non_exact_source_or_geometry_is_not_targeted(self) -> None:
        baseline, chunks = self.b19()
        validation_error = (
            "ValueError: studies[1].arms[1].role REFERENCE requires "
            "directly cited captured full Normal"
        )
        baseline["source"]["contentSha256"] = "0" * 64
        self.assertIsNone(
            repair.arm_identity_repair_target(
                validation_error,
                baseline,
                chunks,
            )
        )

        baseline, chunks = self.b19()
        chunks[0]["cells"][2]["rawValue"] = "Changed"
        self.assertIsNone(
            repair.arm_identity_repair_target(
                validation_error,
                baseline,
                chunks,
            )
        )


if __name__ == "__main__":
    unittest.main()
