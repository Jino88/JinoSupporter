from __future__ import annotations

import copy
import unittest

import inference_data_ai_b17_report_table_repair as repair


class B17ReportTableRepairTests(unittest.TestCase):
    def cell(
        self,
        coordinate: str,
        value: object,
        merge_range: str = "",
    ) -> dict:
        return {
            "coordinate": coordinate,
            "rawValue": value,
            "primary": True,
            "contextOnly": False,
            "mergeRange": merge_range or None,
        }

    def baseline(self) -> dict:
        return {
            "source": {
                "contentComplete": True,
                "revisionUid": repair.B17_REVISION_UID,
                "contentSha256": repair.B17_CONTENT_SHA256,
                "sourcePath": rf"D:\Input\{repair.B17_FILE_NAME}",
            },
            "studies": [
                {
                    "key": "report-new-lot-cd-checks",
                    "conclusions": [
                        {"key": "dyne-pen-all-lots-ok"},
                        {
                            "key": "tension-lots-cannot-use",
                            "evidence": [
                                {
                                    "sheet": "Report",
                                    "range": "C110",
                                    "role": "SOURCE",
                                    "sourceText": (
                                        " + Lot 18,21,27,30,31,34,41  "
                                        "happen tension => Can not use "
                                    ),
                                    "note": "",
                                }
                            ],
                        },
                        {
                            "key": "separation-lots-cannot-use",
                            "evidence": [
                                {
                                    "sheet": "Report",
                                    "range": "C112",
                                    "role": "SOURCE",
                                    "sourceText": (
                                        " + Lot 2,4,7,8,13,17,18,21,24,"
                                        "25,26,27,30,31,33,34,35, 37,38,"
                                        "40,41,42,44   happen separate "
                                        "VP/CD  => Can not use "
                                    ),
                                    "note": "",
                                }
                            ],
                        },
                        {
                            "key": (
                                "listed-lots-require-second-sample-"
                                "function-test"
                            )
                        },
                    ],
                },
                {"key": "vp-cd-assembly-tension-tests"},
            ],
        }

    def chunks(self) -> list[dict]:
        cells = [
            self.cell("C15", "Date test", "C15:C17"),
            self.cell("D15", "LOT TEST", "D15:D17"),
            self.cell("E15", "LOT", "E15:E17"),
            self.cell("F16", "Dyne", "F16:F17"),
            self.cell("G16", "Input", "G16:G17"),
            self.cell("H16", "NG / NG rate", "H16:H17"),
            self.cell("I16", "Input", "I16:I17"),
            self.cell("J16", "Tension TEST", "J16:J17"),
            self.cell("K16", "Input", "K16:K17"),
            self.cell("L16", "OK", "L16:L17"),
            self.cell("M16", "VP/CD separate", "M16:M17"),
            self.cell("N16", "Total NG", "N16:N17"),
            self.cell("O16", "NG rate", "O16:O17"),
            self.cell("F18", 40, "F18:F107"),
        ]
        for lot, top_row in enumerate(range(18, 108, 2), start=1):
            lower_row = top_row + 1
            cells.extend(
                [
                    self.cell(
                        f"D{top_row}",
                        lot,
                        f"D{top_row}:D{lower_row}",
                    ),
                    self.cell(
                        f"G{top_row}",
                        10,
                        f"G{top_row}:G{lower_row}",
                    ),
                    self.cell(f"H{top_row}", 0),
                    self.cell(f"H{lower_row}", 0),
                    self.cell(
                        f"I{top_row}",
                        5,
                        f"I{top_row}:I{lower_row}",
                    ),
                    self.cell(f"J{top_row}", 0),
                    self.cell(f"J{lower_row}", 0),
                    self.cell(
                        f"K{top_row}",
                        50,
                        f"K{top_row}:K{lower_row}",
                    ),
                    self.cell(
                        f"L{top_row}",
                        50,
                        f"L{top_row}:L{lower_row}",
                    ),
                    self.cell(f"M{top_row}", 0),
                    self.cell(f"M{lower_row}", 0),
                    self.cell(
                        f"N{top_row}",
                        0,
                        f"N{top_row}:N{lower_row}",
                    ),
                    self.cell(
                        f"O{top_row}",
                        0,
                        f"O{top_row}:O{lower_row}",
                    ),
                ]
            )
        cells.extend(
            [
                self.cell(
                    "C109",
                    " => Result check tension  :",
                    "C109:O109",
                ),
                self.cell(
                    "C110",
                    (
                        " + Lot 18,21,27,30,31,34,41  happen tension "
                        "=> Can not use "
                    ),
                    "C110:O110",
                ),
                self.cell(
                    "C111",
                    " => Result check separate VP/CD :",
                    "C111:O111",
                ),
                self.cell(
                    "C112",
                    (
                        " + Lot 2,4,7,8,13,17,18,21,24,25,26,27,"
                        "30,31,33,34,35, 37,38,40,41,42,44   happen "
                        "separate VP/CD  => Can not use "
                    ),
                    "C112:O112",
                ),
            ]
        )
        return [{"sheet": {"title": "Report"}, "cells": cells}]

    def test_exact_projection_adds_540_observations(self) -> None:
        baseline = self.baseline()
        untouched = copy.deepcopy(baseline)
        chunks = self.chunks()
        error = (
            "ContentCoverageError: Source content coverage is incomplete; "
            "586 quantitative cell(s)"
        )

        self.assertTrue(
            repair.b17_report_table_repair_applicable(
                baseline,
                validation_error=error,
                focused_chunks=chunks,
            )
        )
        repaired = repair.apply_b17_report_table_repair(
            baseline,
            focused_chunks=chunks,
        )

        self.assertEqual(untouched, baseline)
        self.assertEqual(
            "report-lot-test-results",
            repaired["studies"][1]["key"],
        )
        self.assertEqual(45, len(repaired["studies"][1]["arms"]))
        self.assertEqual(
            540,
            sum(
                len(outcome["observations"])
                for outcome in repaired["studies"][1]["outcomes"]
            ),
        )
        self.assertEqual(
            "C109:C110",
            repaired["studies"][0]["conclusions"][1][
                "evidence"
            ][0]["range"],
        )

    def test_changed_source_geometry_fails_closed(self) -> None:
        chunks = self.chunks()
        chunks[0]["cells"][0]["rawValue"] = "Changed"
        self.assertFalse(
            repair.b17_report_table_repair_applicable(
                self.baseline(),
                validation_error=(
                    "Source content coverage is incomplete; "
                    "586 quantitative cell(s)"
                ),
                focused_chunks=chunks,
            )
        )
        with self.assertRaises(repair.B17ReportTableRepairError):
            repair.apply_b17_report_table_repair(
                self.baseline(),
                focused_chunks=chunks,
            )


if __name__ == "__main__":
    unittest.main()
