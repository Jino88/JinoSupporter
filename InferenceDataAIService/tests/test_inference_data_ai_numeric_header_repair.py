from __future__ import annotations

import copy
import unittest

import inference_data_ai_numeric_header_repair as repair


class NumericHeaderSeriesRepairTests(unittest.TestCase):
    title = (
        "RESULT HEIGHT CHECK MATERIAL C-MG, S-MG  "
        "( Spec : 0.66~0.70mm )"
    )
    arm_rows = (
        ("height-c-mg", "C-MG", 7, 8, [0, 232, 251, 17, 0]),
        (
            "height-s-mg-long",
            "S-MG Long",
            9,
            10,
            [311, 187, 2, 0, 0],
        ),
        (
            "height-s-mg-short",
            "S-MG Short",
            11,
            12,
            [0, 184, 292, 24, 0],
        ),
    )
    columns = "FGHIJ"
    suffixes = ("066", "067", "068", "069", "070")
    headers = (0.66, 0.67, 0.68, 0.69, 0.70)

    def cell(
        self,
        coordinate: str,
        value: object,
        *,
        number_format: str = "General",
        merge_range: str = "",
        merge_role: str = "none",
    ) -> dict:
        return {
            "sourceCellKey": f"revision-1:1:{coordinate}",
            "coordinate": coordinate,
            "rawValue": value,
            "displayValue": value,
            "numberFormat": number_format,
            "mergeRange": merge_range or None,
            "mergeRole": merge_role,
        }

    def chunks(self) -> list[dict]:
        cells = [
            self.cell("C4", self.title),
            self.cell("D6", "Type"),
            self.cell("E6", "Q'ty check"),
        ]
        cells.extend(
            self.cell(
                f"{column}6",
                value,
                number_format="0.00",
            )
            for column, value in zip(self.columns, self.headers)
        )
        for _arm_key, label, row, rate_row, counts in self.arm_rows:
            cells.extend(
                [
                    self.cell(
                        f"D{row}",
                        label,
                        merge_range=f"D{row}:D{rate_row}",
                        merge_role="anchor",
                    ),
                    self.cell(
                        f"E{row}",
                        500,
                        merge_range=f"E{row}:E{rate_row}",
                        merge_role="anchor",
                    ),
                ]
            )
            cells.extend(
                self.cell(
                    f"{column}{row}",
                    count,
                    number_format="0",
                )
                for column, count in zip(self.columns, counts)
            )
            cells.extend(
                self.cell(
                    f"{column}{rate_row}",
                    count / 500,
                    number_format="0.0%",
                )
                for column, count in zip(self.columns, counts)
            )
        return [
            {
                "chunkId": "b09-height-1",
                "sheet": {"title": "Sheet1"},
                "cells": cells[:24],
            },
            {
                "chunkId": "b09-height-2",
                "sheet": {"title": "Sheet1"},
                "cells": cells[24:],
            },
        ]

    def evidence(
        self,
        coordinate: str,
        source_text: str,
    ) -> list[dict]:
        return [
            {
                "sheet": "Sheet1",
                "range": coordinate,
                "role": "SOURCE",
                "sourceText": source_text,
                "note": "",
            }
        ]

    def baseline(self) -> dict:
        arms = []
        input_observations = []
        for arm_key, label, row, _rate_row, _counts in self.arm_rows:
            arms.append(
                {
                    "key": arm_key,
                    "role": "OTHER",
                    "label": label,
                    "condition": label,
                    "sampleSize": 500,
                    "sampleBasis": "Q'ty check",
                    "matchingBasis": "",
                    "factorValues": [
                        {
                            "factor": "height-material-type",
                            "value": label,
                            "valueNumber": None,
                            "unit": "",
                            "isBaseline": False,
                            "heldConstant": False,
                        }
                    ],
                    "evidence": [
                        *self.evidence(f"D{row}", label),
                        *self.evidence(f"E{row}", "500"),
                    ],
                }
            )
            input_observations.append(
                {
                    "key": f"input-{arm_key}",
                    "arm": arm_key,
                    "valueNumber": 500,
                    "valueText": "500",
                    "numerator": None,
                    "denominator": None,
                    "ratePpm": None,
                    "min": None,
                    "max": None,
                    "average": None,
                    "sampleSize": 500,
                    "evidence": self.evidence(f"E{row}", "500"),
                }
            )

        outcomes = [
            {
                "key": "height-input",
                "originalLabel": "Q'ty check",
                "metricType": "sample_size",
                "unit": "",
                "favorableDirection": "UNKNOWN",
                "evidence": self.evidence("E6", "Q'ty check"),
                "observations": input_observations,
            }
        ]
        for column, suffix, header, category_index in zip(
            self.columns,
            self.suffixes,
            self.headers,
            range(len(self.headers)),
        ):
            count_observations = []
            rate_observations = []
            for arm_key, _label, row, rate_row, counts in self.arm_rows:
                count = counts[category_index]
                rate_percent = round(count / 500 * 100, 10)
                count_observations.append(
                    {
                        "key": f"{suffix}-{arm_key}-count",
                        "arm": arm_key,
                        "valueNumber": count,
                        "valueText": str(count),
                        "numerator": count,
                        "denominator": 500,
                        "ratePpm": None,
                        "min": None,
                        "max": None,
                        "average": None,
                        "sampleSize": 500,
                        "evidence": [
                            *self.evidence(
                                f"{column}{row}",
                                str(count),
                            ),
                            *self.evidence(f"E{row}", "500"),
                        ],
                    }
                )
                rate_observations.append(
                    {
                        "key": f"{suffix}-{arm_key}-rate",
                        "arm": arm_key,
                        "valueNumber": rate_percent,
                        "valueText": f"{rate_percent:.1f}%",
                        "numerator": count,
                        "denominator": 500,
                        "ratePpm": None,
                        "min": None,
                        "max": None,
                        "average": None,
                        "sampleSize": 500,
                        "evidence": [
                            *self.evidence(
                                f"{column}{rate_row}",
                                f"{rate_percent:.1f}%",
                            ),
                            *self.evidence(
                                f"{column}{row}",
                                str(count),
                            ),
                            *self.evidence(f"E{row}", "500"),
                        ],
                    }
                )
            header_text = f"{header:.2f}"
            outcomes.extend(
                [
                    {
                        "key": f"height-{suffix}-count",
                        "originalLabel": f"{header_text} count",
                        "metricType": "height_category_count",
                        "unit": "",
                        "favorableDirection": "UNKNOWN",
                        "evidence": self.evidence(
                            f"{column}6",
                            header_text,
                        ),
                        "observations": count_observations,
                    },
                    {
                        "key": f"height-{suffix}-rate",
                        "originalLabel": f"{header_text} percentage",
                        "metricType": "height_category_rate",
                        "unit": "%",
                        "favorableDirection": "UNKNOWN",
                        "evidence": self.evidence(
                            f"{column}6",
                            header_text,
                        ),
                        "observations": rate_observations,
                    },
                ]
            )

        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "dataset": "InputDataFinish",
                "sourcePath": r"D:\input\b09.xlsx",
                "revisionUid": "capture_revision_b09",
                "contentSha256": "a" * 64,
                "contentComplete": True,
            },
            "workbookAnalysis": {
                "key": "b09",
                "title": self.title,
                "summary": "B09",
            },
            "studies": [
                {
                    "key": "height-category-distribution",
                    "title": self.title,
                    "limitations": [],
                    "evidence": self.evidence("C4", self.title),
                    "contexts": [],
                    "factors": [
                        {
                            "key": "height-material-type",
                            "originalLabel": "Type",
                            "evidence": self.evidence("D6", "Type"),
                        }
                    ],
                    "arms": arms,
                    "outcomes": outcomes,
                    "measurementSeries": [],
                    "comparisons": [],
                    "conclusions": [],
                },
                {
                    "key": "speaker-gauss-check",
                    "title": "Gauss",
                    "limitations": [],
                    "comparisons": [
                        {"key": "gauss-cmg-vs-normal"}
                    ],
                },
            ],
        }

    def target(
        self,
        baseline: dict | None = None,
        chunks: list[dict] | None = None,
        error: str | None = None,
    ) -> dict | None:
        return repair.numeric_header_series_repair_target(
            (
                repair.B09_NUMERIC_HEADER_COVERAGE_ERROR
                if error is None
                else error
            ),
            self.baseline() if baseline is None else baseline,
            self.chunks() if chunks is None else chunks,
        )

    def find_cell(
        self,
        chunks: list[dict],
        coordinate: str,
    ) -> dict:
        return next(
            cell
            for chunk in chunks
            for cell in chunk["cells"]
            if cell["coordinate"] == coordinate
        )

    def test_exact_target_appends_one_outcome_and_three_series(self) -> None:
        baseline = self.baseline()
        target = self.target(baseline)
        self.assertIsNotNone(target)
        assert target is not None

        repaired = repair.apply_numeric_header_series_repair(
            baseline,
            target,
        )
        repair.validate_numeric_header_series_repair(
            baseline,
            repaired,
            target,
        )

        self.assertEqual(11, len(baseline["studies"][0]["outcomes"]))
        self.assertEqual(12, len(repaired["studies"][0]["outcomes"]))
        self.assertEqual(
            "height-category-count-series",
            repaired["studies"][0]["outcomes"][-1]["key"],
        )
        self.assertEqual(
            [
                (
                    "height-cmg-count-series",
                    "F6:J6",
                    "F7:J7",
                    "D7:D7",
                ),
                (
                    "height-smg-long-count-series",
                    "F6:J6",
                    "F9:J9",
                    "D9:D9",
                ),
                (
                    "height-smg-short-count-series",
                    "F6:J6",
                    "F11:J11",
                    "D11:D11",
                ),
            ],
            [
                (
                    series["key"],
                    series["headerRange"],
                    series["valueRange"],
                    series["rowIdentityRange"],
                )
                for series in repaired["studies"][0][
                    "measurementSeries"
                ]
            ],
        )
        self.assertEqual([], baseline["studies"][0]["measurementSeries"])

    def test_apply_and_target_are_idempotent(self) -> None:
        baseline = self.baseline()
        target = self.target(baseline)
        assert target is not None
        once = repair.apply_numeric_header_series_repair(
            baseline,
            target,
        )
        twice = repair.apply_numeric_header_series_repair(once, target)
        self.assertEqual(once, twice)

        repeated_target = self.target(once)
        self.assertIsNotNone(repeated_target)
        assert repeated_target is not None
        self.assertEqual(
            once,
            repair.apply_numeric_header_series_repair(
                once,
                repeated_target,
            ),
        )

    def test_projection_rejects_any_unrelated_mutation(self) -> None:
        baseline = self.baseline()
        target = self.target(baseline)
        assert target is not None
        repaired = repair.apply_numeric_header_series_repair(
            baseline,
            target,
        )
        repaired["workbookAnalysis"]["summary"] = "changed"
        with self.assertRaisesRegex(
            repair.NumericHeaderRepairError,
            "outside the exact",
        ):
            repair.validate_numeric_header_series_repair(
                baseline,
                repaired,
                target,
            )

    def test_target_composes_with_other_study_repair(self) -> None:
        baseline = self.baseline()
        target = self.target(baseline)
        assert target is not None
        after_gauss_repair = copy.deepcopy(baseline)
        after_gauss_repair["studies"][1]["comparisons"] = []
        after_gauss_repair["studies"][1]["limitations"].append(
            "gauss comparison omitted"
        )

        repaired = repair.apply_numeric_header_series_repair(
            after_gauss_repair,
            target,
        )
        repair.validate_numeric_header_series_repair(
            after_gauss_repair,
            repaired,
            target,
        )
        self.assertEqual([], repaired["studies"][1]["comparisons"])
        self.assertEqual(
            ["gauss comparison omitted"],
            repaired["studies"][1]["limitations"],
        )

    def test_validation_error_must_be_exact(self) -> None:
        exact_without_type = (
            repair.B09_NUMERIC_HEADER_COVERAGE_ERROR.split(": ", 1)[1]
        )
        self.assertIsNotNone(
            self.target(error=exact_without_type)
        )
        for error in (
            "",
            repair.B09_NUMERIC_HEADER_COVERAGE_ERROR
            + "; 1 semantic label cell(s): Sheet1!D6",
            repair.B09_NUMERIC_HEADER_COVERAGE_ERROR.replace(
                "Sheet1!G6, Sheet1!H6",
                "Sheet1!H6, Sheet1!G6",
            ),
            repair.B09_NUMERIC_HEADER_COVERAGE_ERROR.replace(
                "5 quantitative",
                "6 quantitative",
            ),
        ):
            with self.subTest(error=error):
                self.assertIsNone(self.target(error=error))

    def test_source_geometry_fail_closed(self) -> None:
        mutations = {}

        wrong_header = self.chunks()
        self.find_cell(wrong_header, "F6")["rawValue"] = 0.65
        mutations["wrong header"] = wrong_header

        wrong_format = self.chunks()
        self.find_cell(wrong_format, "G6")["numberFormat"] = "General"
        mutations["wrong header format"] = wrong_format

        missing_cell = self.chunks()
        missing_cell[0]["cells"] = [
            cell
            for cell in missing_cell[0]["cells"]
            if cell["coordinate"] != "H6"
        ]
        mutations["missing header"] = missing_cell

        wrong_sum = self.chunks()
        self.find_cell(wrong_sum, "G7")["rawValue"] = 233
        mutations["count sum"] = wrong_sum

        wrong_rate = self.chunks()
        self.find_cell(wrong_rate, "G8")["rawValue"] = 0.465
        mutations["rate arithmetic"] = wrong_rate

        wrong_title = self.chunks()
        self.find_cell(wrong_title, "C4")["rawValue"] = (
            "RESULT HEIGHT CHECK"
        )
        mutations["title"] = wrong_title

        wrong_merge = self.chunks()
        self.find_cell(wrong_merge, "D7")["mergeRange"] = "D7:D9"
        mutations["arm merge"] = wrong_merge

        duplicate = self.chunks()
        duplicate[1]["cells"].append(
            copy.deepcopy(self.find_cell(duplicate, "F6"))
        )
        mutations["duplicate coordinate"] = duplicate

        for name, chunks in mutations.items():
            with self.subTest(name=name):
                self.assertIsNone(self.target(chunks=chunks))

    def test_manifest_geometry_fail_closed(self) -> None:
        mutations = {}

        wrong_arm = self.baseline()
        wrong_arm["studies"][0]["arms"][0]["sampleSize"] = 499
        mutations["arm sample size"] = wrong_arm

        wrong_count = self.baseline()
        wrong_count["studies"][0]["outcomes"][1]["observations"][0][
            "valueNumber"
        ] = 1
        mutations["count observation"] = wrong_count

        wrong_factor = self.baseline()
        wrong_factor["studies"][0]["factors"][0]["evidence"][0][
            "range"
        ] = "C6"
        mutations["factor evidence"] = wrong_factor

        incomplete = self.baseline()
        target = self.target(incomplete)
        assert target is not None
        incomplete["studies"][0]["outcomes"].append(
            copy.deepcopy(target["outcome"])
        )
        mutations["partial repair"] = incomplete

        conflicting = self.baseline()
        conflicting["studies"][0]["measurementSeries"].append(
            {
                "key": "height-cmg-count-series",
                "headerRange": "F6:I6",
            }
        )
        mutations["conflicting series"] = conflicting

        for name, baseline in mutations.items():
            with self.subTest(name=name):
                self.assertIsNone(self.target(baseline=baseline))

    def test_apply_rejects_changed_projection_and_tampered_target(
        self,
    ) -> None:
        baseline = self.baseline()
        target = self.target(baseline)
        assert target is not None

        changed = copy.deepcopy(baseline)
        changed["studies"][0]["limitations"].append("changed")
        with self.assertRaisesRegex(
            repair.NumericHeaderRepairError,
            "protected repair projection",
        ):
            repair.apply_numeric_header_series_repair(changed, target)

        tampered = copy.deepcopy(target)
        tampered["measurementSeries"][0]["valueRange"] = "F8:J8"
        with self.assertRaisesRegex(
            repair.NumericHeaderRepairError,
            "target is not exact",
        ):
            repair.apply_numeric_header_series_repair(
                baseline,
                tampered,
            )


if __name__ == "__main__":
    unittest.main()
