from __future__ import annotations

import copy
import unittest

import inference_data_ai_content_coverage as coverage


class ContentCoverageTests(unittest.TestCase):
    def cell(
        self,
        coordinate: str,
        value: object,
        *,
        sheet_index: int = 1,
        data_type: str | None = None,
        number_format: str = "General",
        formula: str = "",
        cached_value: object = None,
    ) -> dict:
        return {
            "sourceCellKey": (
                f"revision-1:{sheet_index}:{coordinate}"
            ),
            "coordinate": coordinate,
            "rawValue": value,
            "cachedValue": cached_value,
            "displayValue": (
                cached_value if formula else value
            ),
            "dataType": (
                data_type
                if data_type is not None
                else ("n" if isinstance(value, (int, float)) else "s")
            ),
            "cachedDataType": "n" if formula else None,
            "numberFormat": number_format,
            "formula": formula,
        }

    def chunk(
        self,
        *,
        sheet: str = "Data",
        sheet_index: int = 1,
        cells: list[dict],
    ) -> dict:
        return {
            "chunkId": f"chunk-{sheet_index}",
            "sheet": {
                "sheetIndex": sheet_index,
                "title": sheet,
            },
            "sectionIndex": 1,
            "cells": cells,
            "contextCells": [],
        }

    def locator(
        self,
        chunk: dict,
        *,
        evidence: list[dict] | None = None,
    ) -> dict:
        return {
            "chunkId": chunk["chunkId"],
            "status": "CANDIDATES",
            "candidates": [
                {
                    "key": f"candidate-{chunk['chunkId']}",
                    "evidence": evidence or [
                        {
                            "sheet": chunk["sheet"]["title"],
                            "range": "A1:Z999",
                            "role": "CANDIDATE_REGION",
                        }
                    ],
                }
            ],
        }

    def manifest(self, studies: list[dict]) -> dict:
        return {
            "source": {"contentComplete": True},
            "workbookAnalysis": {
                "summary": "",
            },
            "studies": studies,
        }

    def observation_study(
        self,
        observations: list[dict],
    ) -> dict:
        return {
            "measurementSeries": [],
            "outcomes": [
                {
                    "observations": observations,
                }
            ],
            "conclusions": [],
        }

    def inventory(
        self,
        chunks: list[dict],
        locators: list[dict] | None = None,
    ) -> dict:
        return coverage.build_content_coverage_inventory(
            chunks=chunks,
            locator_results=locators
            or [self.locator(chunk) for chunk in chunks],
            expected_source_cell_keys=[
                cell["sourceCellKey"]
                for chunk in chunks
                for cell in chunk["cells"]
            ],
        )

    def test_b09_empty_outcomes_cannot_hide_numeric_panel(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Gauss"),
                self.cell("B2", 410),
                self.cell("C2", 412),
            ]
        )
        inventory = self.inventory([chunk])
        manifest = self.manifest(
            [
                {
                    "measurementSeries": [],
                    "outcomes": [],
                    "conclusions": [],
                }
            ]
        )

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "2 quantitative cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=manifest,
                inventory=inventory,
                require_complete=True,
            )

    def test_numeric_arm_label_is_source_identity_not_measurement(
        self,
    ) -> None:
        chunk = self.chunk(cells=[self.cell("B6", 1.2)])
        inventory = self.inventory([chunk])
        manifest = self.manifest(
            [
                {
                    "arms": [
                        {
                            "key": "row-1-2",
                            "label": "1.2",
                            "condition": "1.2",
                            "sampleSize": None,
                            "factorValues": [],
                            "evidence": [
                                {
                                    "sheet": "Data",
                                    "range": "B6",
                                }
                            ],
                        }
                    ],
                    "measurementSeries": [],
                    "outcomes": [],
                    "conclusions": [],
                }
            ]
        )

        report = coverage.validate_content_manifest_coverage(
            manifest=manifest,
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(0, report["uncoveredCellCount"])
        self.assertEqual(
            "ARM_NUMERIC_IDENTITY",
            report["coverageBySourceCellKey"]["revision-1:1:B6"],
        )

    def test_isolated_page_ordinal_above_merged_title_is_structural(
        self,
    ) -> None:
        title = self.cell("B2", "TITLE")
        title["mergeRange"] = "B2:C4"
        chunk = self.chunk(
            cells=[
                self.cell("C1", 1),
                title,
                self.cell("D2", "REPORT TEST SEMI YOKE"),
                self.cell("B5", "I. Purpose."),
                self.cell("C18", 452),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(
            {"C1": "SHEET_LAYOUT_ORDINAL"},
            {
                item["coordinate"]: item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )
        self.assertEqual(
            {"C18"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )

        second_ordinal = copy.deepcopy(chunk)
        second_ordinal["cells"].insert(1, self.cell("D1", 2))
        inventory = self.inventory([second_ordinal])
        self.assertEqual(
            {"C1", "D1", "C18"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )

    def test_blank_formula_label_above_error_summary_is_structural(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell(
                    "CI3",
                    None,
                    formula="=D7",
                    cached_value=0,
                ),
                self.cell(
                    "CI4",
                    None,
                    formula="=SUM(A1:A2)/0",
                    cached_value="#DIV/0!",
                ),
                self.cell(
                    "CI5",
                    None,
                    formula="=SUM(A1:A2)/0",
                    cached_value="#DIV/0!",
                ),
                self.cell("CI7", "AVG"),
                self.cell(
                    "CI8",
                    None,
                    formula='=D7&"_AVG"',
                    cached_value="_AVG",
                ),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(
            {
                "CI3": "HIDDEN_FORMULA_WITHOUT_SOURCE_INPUT",
            },
            {
                item["coordinate"]: item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

        no_error_summary = copy.deepcopy(chunk)
        no_error_summary["cells"] = [
            no_error_summary["cells"][0],
            *no_error_summary["cells"][3:],
        ]
        inventory = self.inventory([no_error_summary])
        self.assertEqual(
            {"CI3"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )

    def test_horizontal_sample_numbers_are_replicate_headers(
        self,
    ) -> None:
        cells = [self.cell("H7", "AVG")]
        for column, value in zip(
            ("I", "J", "K", "L", "M"),
            range(1, 6),
            strict=True,
        ):
            cells.append(self.cell(f"{column}7", value))
            cells.append(self.cell(f"{column}8", value * 10))
        chunk = self.chunk(cells=cells)

        inventory = self.inventory([chunk])

        excluded = {
            item["coordinate"]: item["exclusionReason"]
            for item in inventory["excludedCells"]
        }
        required = {
            item["coordinate"]
            for item in inventory["requiredCells"]
        }
        self.assertEqual(
            {
                "I7": "HORIZONTAL_REPLICATE_IDENTIFIER",
                "J7": "HORIZONTAL_REPLICATE_IDENTIFIER",
                "K7": "HORIZONTAL_REPLICATE_IDENTIFIER",
                "L7": "HORIZONTAL_REPLICATE_IDENTIFIER",
                "M7": "HORIZONTAL_REPLICATE_IDENTIFIER",
            },
            excluded,
        )
        self.assertEqual(
            {"I8", "J8", "K8", "L8", "M8"},
            required,
        )

    def test_formula_labeled_empty_sample_columns_keep_header_identity(
        self,
    ) -> None:
        cells = [self.cell("H7", "AVG")]
        for column, value in zip(
            ("I", "J", "K", "L", "M"),
            range(1, 6),
            strict=True,
        ):
            cells.append(self.cell(f"{column}7", value))
            cells.append(
                self.cell(
                    f"{column}8",
                    None,
                    formula=f'=$D$2&" #"&{column}$7',
                    cached_value=f"STD #{value}",
                )
            )
        inventory = self.inventory([self.chunk(cells=cells)])

        self.assertEqual(
            {"I7", "J7", "K7", "L7", "M7"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"]
                == "HORIZONTAL_REPLICATE_IDENTIFIER"
            },
        )

    def test_b13_scalar_summary_cannot_consume_50_raw_gauss_cells(
        self,
    ) -> None:
        cells = [self.cell("A1", "Gauss raw measurements")]
        for row in range(2, 52):
            cells.append(self.cell(f"B{row}", 400 + row))
        chunk = self.chunk(cells=cells)
        inventory = self.inventory([chunk])
        manifest = self.manifest(
            [
                self.observation_study(
                    [
                        {
                            "average": 426.5,
                            "evidence": [
                                {
                                    "sheet": "Data",
                                    "range": "B2:B51",
                                }
                            ],
                        }
                    ]
                )
            ]
        )

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "50 quantitative cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=manifest,
                inventory=inventory,
                require_complete=True,
            )

    def test_b14_partial_sheet_manifest_fails_exact_union(self) -> None:
        chunks: list[dict] = []
        locators: list[dict] = []
        for sheet_index in range(1, 23):
            chunk = self.chunk(
                sheet=f"Panel {sheet_index:02d}",
                sheet_index=sheet_index,
                cells=[
                    self.cell(
                        "B2",
                        400 + sheet_index,
                        sheet_index=sheet_index,
                    )
                ],
            )
            chunks.append(chunk)
            locators.append(self.locator(chunk))
        inventory = self.inventory(chunks, locators)
        observations = [
            {
                "valueNumber": 400 + sheet_index,
                "evidence": [
                    {
                        "sheet": f"Panel {sheet_index:02d}",
                        "range": "B2",
                    }
                ],
            }
            for sheet_index in range(1, 4)
        ]

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "19 quantitative cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study(observations)]
                ),
                inventory=inventory,
                require_complete=True,
            )

    def test_exact_scalar_and_count_observations_pass(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B2", 12),
                self.cell("B3", 3),
                self.cell("C3", 60),
            ]
        )
        manifest = self.manifest(
            [
                self.observation_study(
                    [
                        {
                            "valueNumber": 12,
                            "evidence": [
                                {"sheet": "Data", "range": "B2"}
                            ],
                        },
                        {
                            "valueNumber": 5,
                            "numerator": 3,
                            "denominator": 60,
                            "evidence": [
                                {"sheet": "Data", "range": "B3:C3"}
                            ],
                        },
                    ]
                )
            ]
        )

        report = coverage.validate_content_manifest_coverage(
            manifest=manifest,
            inventory=self.inventory([chunk]),
            require_complete=True,
        )

        self.assertEqual(3, report["coveredCellCount"])
        self.assertEqual(0, report["uncoveredCellCount"])

    def test_raw_and_aggregate_series_preserve_formula_results(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Sample No."),
                self.cell("B1", 100),
                self.cell("C1", 200),
                self.cell("A2", 1),
                self.cell("A3", 2),
                self.cell("B2", 410),
                self.cell("C2", 411),
                self.cell("B3", 412),
                self.cell("C3", 413),
                self.cell("D1", "Test date"),
                self.cell(
                    "D2",
                    45600,
                    number_format="yyyy-mm-dd",
                ),
                self.cell("E1", "Average"),
                self.cell("E2", 411.5),
                self.cell(
                    "F2",
                    None,
                    formula="=AVERAGE(B2:C3)",
                    cached_value=411.5,
                ),
            ]
        )
        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B1:C1",
                "valueRange": "B2:C3",
                "rowIdentityRange": "A2:A3",
                "axisSource": "HEADER",
            },
            {
                "seriesRole": "AGGREGATE",
                "sheet": "Data",
                "headerRange": "E1:F1",
                "valueRange": "E2:F2",
                "rowIdentityRange": "A2",
                "axisSource": "HEADER",
            },
        ]
        inventory = self.inventory([chunk])

        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(0, report["uncoveredCellCount"])
        self.assertEqual(
            {
                "DATE_FORMAT",
                "SEQUENCE_LABEL",
            },
            {
                item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

    def test_formula_and_explicit_aggregate_results_cannot_be_omitted(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Average"),
                self.cell("B1", 411.5),
                self.cell(
                    "C1",
                    None,
                    formula="=AVERAGE(D1:G1)",
                    cached_value=412.5,
                ),
            ]
        )
        inventory = self.inventory([chunk])

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "2 quantitative cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

    def test_strict_numeric_text_is_quantitative_source_content(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B2", "5.0", data_type="s"),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(1, inventory["requiredCellCount"])

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B2",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

    def test_formula_without_cached_result_fails_closed(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell(
                    "B2",
                    None,
                    formula="=SUM(C2:D2)",
                    cached_value=None,
                ),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(
            1,
            inventory["unresolvedFormulaCellCount"],
        )

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "1 unresolved formula cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

    def test_hidden_blank_input_formulas_and_excel_errors_are_not_results(
        self,
    ) -> None:
        ok_formula = self.cell(
            "H23",
            None,
            formula="=G23-N23",
            cached_value=0,
        )
        total_formula = self.cell(
            "N23",
            None,
            formula="=SUM(I23:M23)",
            cached_value=0,
        )
        for cell in (ok_formula, total_formula):
            cell["hidden"] = {
                "row": True,
                "column": False,
                "sheet": False,
            }
        error_formula = self.cell(
            "I24",
            None,
            formula="=I23/G23",
            cached_value=-2146826281,
        )
        error_formula["displayValue"] = "#DIV/0!"
        error_formula["hidden"] = {
            "row": True,
            "column": False,
            "sheet": False,
        }

        inventory = self.inventory(
            [self.chunk(cells=[
                ok_formula,
                total_formula,
                error_formula,
            ])]
        )

        self.assertEqual(0, inventory["requiredCellCount"])
        self.assertEqual(
            {"H23", "N23"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
            },
        )
        self.assertEqual(
            {"HIDDEN_FORMULA_WITHOUT_SOURCE_INPUT"},
            {
                item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )
        self.assertNotIn(
            "I24",
            {
                item["coordinate"]
                for item in inventory["numericCells"]
            },
        )

    def test_visible_zero_formula_without_source_input_remains_required(
        self,
    ) -> None:
        formula = self.cell(
            "H23",
            None,
            formula="=G23",
            cached_value=0,
        )

        inventory = self.inventory([self.chunk(cells=[formula])])

        self.assertEqual(1, inventory["requiredCellCount"])

    def test_adjacent_error_only_axis_tail_is_excluded(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A100", 14000),
                self.cell("B100", 82),
                self.cell("A175", 19000),
                self.cell("B175", "#REF!", data_type="e"),
                self.cell("A176", 20000),
                self.cell("B176", "#REF!", data_type="e"),
            ]
        )
        inventory = self.inventory([chunk])
        by_coordinate = {
            item["coordinate"]: item
            for item in inventory["numericCells"]
        }
        self.assertEqual(
            "REQUIRED_RESULT",
            by_coordinate["A100"]["classification"],
        )
        self.assertEqual(
            "REQUIRED_RESULT",
            by_coordinate["B100"]["classification"],
        )
        self.assertEqual(
            {"ERROR_ONLY_AXIS_TAIL"},
            {
                item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

    def test_isolated_error_only_axis_fragment_is_excluded(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A175", 19000),
                self.cell("B175", "#REF!", data_type="e"),
                self.cell("A176", 20000),
                self.cell("B176", "#REF!", data_type="e"),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(0, inventory["requiredCellCount"])
        self.assertEqual(2, inventory["excludedCellCount"])
        self.assertEqual(
            {"ERROR_ONLY_AXIS_TAIL"},
            {
                item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

    def test_small_two_row_values_next_to_errors_remain_required(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A175", 19),
                self.cell("B175", "#REF!", data_type="e"),
                self.cell("A176", 20),
                self.cell("B176", "#REF!", data_type="e"),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(2, inventory["requiredCellCount"])
        self.assertEqual(0, inventory["excludedCellCount"])

    def test_detached_blank_only_axis_tail_is_excluded(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A10", 14000),
                self.cell("B10", 82),
                self.cell("A30", 19000),
                self.cell("A31", 20000),
            ]
        )
        inventory = self.inventory([chunk])
        by_coordinate = {
            item["coordinate"]: item
            for item in inventory["numericCells"]
        }

        self.assertEqual(
            "REQUIRED_RESULT",
            by_coordinate["A10"]["classification"],
        )
        self.assertEqual(
            "REQUIRED_RESULT",
            by_coordinate["B10"]["classification"],
        )
        for coordinate in ("A30", "A31"):
            self.assertEqual(
                "EXCLUDED_NON_RESULT",
                by_coordinate[coordinate]["classification"],
            )
            self.assertEqual(
                "ISOLATED_BLANK_AXIS_TAIL",
                by_coordinate[coordinate]["exclusionReason"],
            )

    def test_adjacent_blank_numeric_rows_are_not_an_axis_tail(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A10", 14000),
                self.cell("B10", 82),
                self.cell("A11", 19000),
                self.cell("A12", 20000),
            ]
        )
        inventory = self.inventory([chunk])

        self.assertEqual(
            {"A10", "B10", "A11", "A12"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )

    def test_separate_rate_count_evidence_ranges_pass(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("H16", 60),
                self.cell("J16", 3),
                self.cell("J17", 0.05, number_format="0.0%"),
            ]
        )
        observation = {
            "valueNumber": 5,
            "numerator": 3,
            "denominator": 60,
            "evidence": [
                {"sheet": "Data", "range": "J17"},
                {"sheet": "Data", "range": "J16"},
                {"sheet": "Data", "range": "H16"},
            ],
        }

        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest(
                [self.observation_study([observation])]
            ),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )

        self.assertEqual(3, report["coveredCellCount"])

    def test_numeric_factor_values_are_source_content(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("F4", 4),
                self.cell("F12", 12),
                self.cell("F20", 20),
            ]
        )
        evidence = [
            {"sheet": "Data", "range": "F4"},
            {"sheet": "Data", "range": "F12"},
            {"sheet": "Data", "range": "F20"},
        ]
        study = self.observation_study([])
        study["factors"] = [
            {
                "key": "condition",
                "evidence": evidence,
            }
        ]
        study["arms"] = [
            {
                "evidence": evidence,
                "factorValues": [
                    {
                        "factor": "condition",
                        "valueNumber": value,
                    }
                    for value in (4, 12, 20)
                ],
            }
        ]

        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )

        self.assertEqual(3, report["coveredCellCount"])

    def test_categorical_status_cells_are_not_source_conclusions(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Decision"),
                self.cell("B2", "PASSED"),
                self.cell("B3", "FAILED"),
            ]
        )
        inventory = self.inventory(
            [chunk],
            [
                self.locator(
                    chunk,
                    evidence=[
                        {
                            "sheet": "Data",
                            "range": "A1:B3",
                            "role": "DECISION_CONCLUSION_REGION",
                        }
                    ],
                )
            ],
        )

        self.assertEqual(
            0,
            inventory["narrativeConclusionCellCount"],
        )
        self.assertEqual(
            2,
            inventory["categoricalStatusCellCount"],
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "2 categorical status cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

    def test_status_column_header_is_not_a_status_observation(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("E4", "OK"),
                self.cell("E6", 169),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(
            0,
            inventory["categoricalStatusCellCount"],
        )

    def test_sparse_merged_ok_column_header_is_not_observation(
        self,
    ) -> None:
        ok_header = self.cell("F28", "OK")
        ok_header.update(
            {
                "mergeRole": "anchor",
                "mergeRange": "F28:F29",
            }
        )
        chunk = self.chunk(
            cells=[
                self.cell("C28", "Date"),
                self.cell("D28", "Type"),
                self.cell("E28", "Input"),
                ok_header,
                self.cell("G28", "Sigma"),
                self.cell("D30", "TEST"),
                self.cell("M30", 0),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(
            0,
            inventory["categoricalStatusCellCount"],
        )

    def test_exact_categorical_observations_are_injective(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B2", "PASSED"),
                self.cell("B3", "PASSED"),
            ]
        )
        inventory = self.inventory([chunk])
        broad_observation = {
            "valueText": "PASSED",
            "evidence": [
                {"sheet": "Data", "range": "B2:B3"}
            ],
        }
        broad_report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest(
                [
                    self.observation_study(
                        [broad_observation]
                    )
                ]
            ),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual(
            2,
            broad_report["coveredCategoricalStatusCellCount"],
        )

        exact = self.observation_study(
            [
                {
                    "valueText": "PASSED",
                    "evidence": [
                        {"sheet": "Data", "range": "B2"}
                    ],
                },
                {
                    "valueText": "PASSED",
                    "evidence": [
                        {"sheet": "Data", "range": "B3"}
                    ],
                },
            ]
        )
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([exact]),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual(
            2,
            report["coveredCategoricalStatusCellCount"],
        )

    def test_one_point_series_cannot_launder_result_as_axis(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B2", 999),
                self.cell("B3", 1),
                self.cell("A3", "row-1"),
            ]
        )
        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B2",
                "valueRange": "B3",
                "rowIdentityRange": "A3",
                "axisSource": "HEADER",
            }
        ]

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B2",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([study]),
                inventory=self.inventory([chunk]),
                require_complete=True,
            )

    def test_repeated_scalar_values_require_distinct_claim_slots(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", 0),
                self.cell("A2", 0),
                self.cell("B1", 3),
                self.cell("C1", 60),
            ]
        )
        observation = {
            "valueNumber": 0,
            "numerator": 3,
            "denominator": 60,
            "evidence": [
                {"sheet": "Data", "range": "A1:A2"},
                {"sheet": "Data", "range": "B1"},
                {"sheet": "Data", "range": "C1"},
            ],
        }

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "1 quantitative cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [
                        self.observation_study(
                            [observation]
                        )
                    ]
                ),
                inventory=self.inventory([chunk]),
                require_complete=True,
            )

    def test_factor_value_cannot_consume_repeated_evidence_cells(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B2", 5),
                self.cell("C2", 5),
            ]
        )
        evidence = [
            {"sheet": "Data", "range": "B2"},
            {"sheet": "Data", "range": "C2"},
        ]
        study = self.observation_study([])
        study["factors"] = [
            {"key": "setting", "evidence": evidence}
        ]
        study["arms"] = [
            {
                "evidence": evidence,
                "factorValues": [
                    {
                        "factor": "setting",
                        "valueNumber": 5,
                    }
                ],
            }
        ]

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "1 quantitative cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([study]),
                inventory=self.inventory([chunk]),
                require_complete=True,
            )

    def test_error_pairs_without_detached_tail_are_results(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", 1),
                self.cell("B1", "#N/A", data_type="e"),
                self.cell("A2", 2),
                self.cell("B2", "#N/A", data_type="e"),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(
            {"A1", "A2"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )

    def test_distant_sequence_label_does_not_cross_empty_gap(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Sample No."),
                self.cell("A100", 42),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(
            ["A100"],
            [
                item["coordinate"]
                for item in inventory["requiredCells"]
            ],
        )

    def test_sequence_header_survives_merged_record_row_gaps(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("C15", "No"),
                self.cell("C17", 1),
                self.cell("E17", "First condition"),
                self.cell("C19", 2),
                self.cell("E19", "Second condition"),
                self.cell("C23", 3),
                self.cell("E23", "Normal"),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(
            {"C17", "C19", "C23"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"] == "SEQUENCE_LABEL"
            },
        )

    def test_sequence_values_do_not_hide_their_text_header(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B3", "No."),
                *[
                    self.cell(f"B{row}", row - 3)
                    for row in range(4, 34)
                ],
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(
            {f"B{row}" for row in range(4, 34)},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"] == "SEQUENCE_LABEL"
            },
        )

    def test_decimal_row_numbers_under_no_header_are_structural(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B3", "No"),
                self.cell("B5", 1.1),
                self.cell("B6", 1.2),
                self.cell("B7", 1.3),
                self.cell("B8", 2),
                self.cell("F5", 100),
                self.cell("F6", 101),
                self.cell("F7", 102),
                self.cell("F8", 103),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(
            {"B5", "B6", "B7", "B8"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"] == "SEQUENCE_LABEL"
            },
        )

    def test_merged_numeric_row_identifiers_allow_skipped_numbers(
        self,
    ) -> None:
        cells = [
            self.cell("C14", "No"),
            self.cell("C17", 1),
            self.cell("H17", 379),
            self.cell("C19", 2),
            self.cell("H19", 378),
            self.cell("C21", 4),
            self.cell("H21", 296),
            self.cell("C23", 5),
            self.cell("H23", 417),
        ]
        cells[0].update(
            {"mergeRole": "anchor", "mergeRange": "C14:C16"}
        )
        for index, coordinate in enumerate(
            ("C17", "C19", "C21", "C23"),
            start=1,
        ):
            cells[index * 2 - 1].update(
                {
                    "mergeRole": "anchor",
                    "mergeRange": (
                        f"{coordinate}:C{int(coordinate[1:]) + 1}"
                    ),
                }
            )

        inventory = self.inventory([self.chunk(cells=cells)])

        self.assertEqual(
            {"C17", "C19", "C21", "C23"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"]
                == "MERGED_ROW_IDENTIFIER"
            },
        )

    def test_merged_numeric_group_identifiers_are_not_results(
        self,
    ) -> None:
        cells = [
            self.cell("F4", 1),
            self.cell("G4", "2-A"),
            self.cell("H4", 1),
            self.cell("G5", "2-B"),
            self.cell("H5", 0),
            self.cell("G6", "2-C"),
            self.cell("H6", 1),
        ]
        cells[0].update(
            {"mergeRole": "anchor", "mergeRange": "F4:F6"}
        )

        inventory = self.inventory([self.chunk(cells=cells)])

        self.assertIn(
            "F4",
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"]
                == "MERGED_GROUP_IDENTIFIER"
            },
        )

    def test_merged_numeric_column_levels_beneath_parent_header(
        self,
    ) -> None:
        cells = [self.cell("H14", "Position")]
        cells[0].update(
            {"mergeRole": "anchor", "mergeRange": "H14:K14"}
        )
        for offset, column in enumerate("HIJK", start=1):
            level = self.cell(f"{column}15", offset)
            level.update(
                {
                    "mergeRole": "anchor",
                    "mergeRange": f"{column}15:{column}16",
                }
            )
            cells.extend([level, self.cell(f"{column}17", 300 + offset)])

        inventory = self.inventory([self.chunk(cells=cells)])

        self.assertEqual(
            {"H15", "I15", "J15", "K15"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"]
                == "MERGED_COLUMN_LEVEL"
            },
        )

    def test_conclusion_heading_row_metadata_is_not_conclusion(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Decision"),
                self.cell("B1", "Model X"),
                self.cell("B2", "The reviewed change can be used."),
            ]
        )
        inventory = self.inventory(
            [chunk],
            [
                self.locator(
                    chunk,
                    evidence=[
                        {
                            "sheet": "Data",
                            "range": "A1:B2",
                            "role": "DECISION_CONCLUSION_REGION",
                        }
                    ],
                )
            ],
        )
        self.assertEqual(
            ["B2"],
            [
                item["coordinate"]
                for item in inventory["narrativeConclusionCells"]
            ],
        )

    def test_mixed_condition_measurement_conclusion_region_uses_result_column(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("C3", "Item"),
                self.cell("D3", "Q'ty"),
                self.cell("G3", "Result"),
                self.cell(
                    "C4",
                    "Increase bond assy to max spec ~10mg",
                ),
                self.cell("D4", 1272),
                self.cell("G4", "NG bako reduced."),
            ]
        )
        inventory = self.inventory(
            [chunk],
            [
                self.locator(
                    chunk,
                    evidence=[
                        {
                            "sheet": "Data",
                            "range": "C4:G4",
                            "role": (
                                "condition, measurements, and conclusion"
                            ),
                        }
                    ],
                )
            ],
        )

        self.assertEqual(
            ["G4"],
            [
                item["coordinate"]
                for item in inventory["narrativeConclusionCells"]
            ],
        )
        self.assertEqual(
            "",
            next(
                item["fieldRole"]
                for item in inventory["requiredCells"]
                if item["coordinate"] == "D4"
            ),
        )

    def test_aggregate_outcomes_and_conclusion_uses_total_result_column(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("G2", "Total Q'ty"),
                self.cell("H2", "OK rate"),
                self.cell("M2", "Total result"),
                self.cell(
                    "M3",
                    "SPK Test not affect to function SET result",
                ),
            ]
        )
        inventory = self.inventory(
            [chunk],
            [
                self.locator(
                    chunk,
                    evidence=[
                        {
                            "sheet": "Data",
                            "range": "G2:M3",
                            "role": "aggregate outcomes and conclusion",
                        }
                    ],
                )
            ],
        )

        self.assertEqual(
            ["M3"],
            [
                item["coordinate"]
                for item in inventory["narrativeConclusionCells"]
            ],
        )

    def test_single_source_conclusion_cell_cannot_be_omitted(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell(
                    "B26",
                    "The changed yoke condition was accepted.",
                )
            ]
        )
        inventory = self.inventory(
            [chunk],
            [
                self.locator(
                    chunk,
                    evidence=[
                        {
                            "sheet": "Data",
                            "range": "B26",
                            "role": "DECISION_CONCLUSION_REGION",
                        }
                    ],
                )
            ],
        )

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B26",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

        study = self.observation_study([])
        study["conclusions"] = [
            self.source_conclusion(
                text=(
                    "The changed yoke condition was accepted."
                ),
                evidence_range="B26",
            )
        ]
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual(
            1,
            report["coveredNarrativeConclusionCellCount"],
        )

    def narrative_fixture(
        self,
    ) -> tuple[dict, dict]:
        chunk = self.chunk(
            cells=[
                self.cell("B24", "IV. Decision."),
                self.cell(
                    "B25",
                    "Gauss values remained within the reviewed range.",
                ),
                self.cell(
                    "B26",
                    "The changed yoke condition was accepted.",
                ),
            ]
        )
        locator = self.locator(
            chunk,
            evidence=[
                {
                    "sheet": "Data",
                    "range": "B24:B26",
                    "role": "DECISION_CONCLUSION_REGION",
                }
            ],
        )
        return chunk, self.inventory([chunk], [locator])

    def source_conclusion(
        self,
        *,
        text: str,
        evidence_range: str,
    ) -> dict:
        return {
            "claimType": "SOURCE_CONCLUSION",
            "text": text,
            "evidence": [
                {
                    "sheet": "Data",
                    "range": evidence_range,
                }
            ],
        }

    def test_b08_joined_or_separate_source_conclusions_preserve_all_cells(
        self,
    ) -> None:
        _chunk, inventory = self.narrative_fixture()
        joined = self.manifest(
            [
                {
                    "measurementSeries": [],
                    "outcomes": [],
                    "conclusions": [
                        self.source_conclusion(
                            text=(
                                "Gauss values remained within the reviewed "
                                "range. The changed yoke condition was "
                                "accepted."
                            ),
                            evidence_range="B25:B26",
                        )
                    ],
                }
            ]
        )
        separate = copy.deepcopy(joined)
        separate["studies"][0]["conclusions"] = [
            self.source_conclusion(
                text="Gauss values remained within the reviewed range.",
                evidence_range="B25",
            ),
            self.source_conclusion(
                text="The changed yoke condition was accepted.",
                evidence_range="B26",
            ),
        ]

        joined_report = coverage.validate_content_manifest_coverage(
            manifest=joined,
            inventory=inventory,
            require_complete=True,
        )
        separate_report = coverage.validate_content_manifest_coverage(
            manifest=separate,
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(
            2,
            joined_report["coveredNarrativeConclusionCellCount"],
        )
        self.assertEqual(
            2,
            separate_report["coveredNarrativeConclusionCellCount"],
        )
        self.assertEqual(
            ["B25", "B26"],
            [
                item["coordinate"]
                for item in inventory["narrativeConclusionCells"]
            ],
        )
        reversed_join = copy.deepcopy(joined)
        reversed_join["studies"][0]["conclusions"][0]["text"] = (
            "The changed yoke condition was accepted. "
            "Gauss values remained within the reviewed range."
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B26",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=reversed_join,
                inventory=inventory,
                require_complete=True,
            )

    def test_b08_one_cited_cell_cannot_hide_second_summary_conclusion(
        self,
    ) -> None:
        _chunk, inventory = self.narrative_fixture()
        manifest = self.manifest(
            [
                {
                    "measurementSeries": [],
                    "outcomes": [],
                    "conclusions": [
                        self.source_conclusion(
                            text=(
                                "Gauss values remained within the reviewed "
                                "range."
                            ),
                            evidence_range="B25",
                        )
                    ],
                }
            ]
        )
        manifest["workbookAnalysis"]["summary"] = (
            "The changed yoke condition was accepted."
        )

        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B26",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=manifest,
                inventory=inventory,
                require_complete=True,
            )

    def test_semantic_label_table_requires_exact_role_binding(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Assembly method"),
                self.cell("A2", "Method A"),
                self.cell("A3", "Method B"),
                self.cell("B1", "FUNCTION NG"),
                self.cell("B2", "9.5 s"),
                self.cell("C2", "3/60 pcs"),
            ]
        )
        no_candidate = {
            "chunkId": chunk["chunkId"],
            "status": "NO_CANDIDATE",
            "candidates": [],
        }
        inventory = self.inventory([chunk], [no_candidate])
        self.assertEqual(
            {
                "A1": "FACTOR_LABEL",
                "A2": "FACTOR_LEVEL",
                "A3": "FACTOR_LEVEL",
                "B1": "OUTCOME_LABEL",
                "B2": "UNIT_QUANTITY",
                "C2": "COUNT_RATIO",
            },
            {
                item["coordinate"]: item["semanticRoles"][0]
                for item in inventory["semanticLabelCells"]
            },
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "semantic label cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

        study = self.observation_study(
            [
                {
                    "valueText": "9.5 s",
                    "evidence": [
                        {"sheet": "Data", "range": "B2"}
                    ],
                },
                {
                    "valueText": "3/60 pcs",
                    "numerator": 3,
                    "denominator": 60,
                    "evidence": [
                        {"sheet": "Data", "range": "C2"}
                    ],
                },
            ]
        )
        study["factors"] = [
            {
                "key": "method",
                "originalLabel": "Assembly method",
                "evidence": [
                    {"sheet": "Data", "range": "A1"}
                ],
            }
        ]
        study["arms"] = [
            {
                "label": label,
                "condition": label,
                "evidence": [
                    {"sheet": "Data", "range": coordinate}
                ],
                "factorValues": [
                    {
                        "factor": "method",
                        "value": label,
                    }
                ],
            }
            for label, coordinate in (
                ("Method A", "A2"),
                ("Method B", "A3"),
            )
        ]
        study["outcomes"][0].update(
            {
                "originalLabel": "FUNCTION NG",
                "evidence": [
                    {"sheet": "Data", "range": "B1"}
                ],
            }
        )
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual(
            6,
            report["coveredSemanticLabelCellCount"],
        )

        wrong_outcome = copy.deepcopy(study)
        wrong_outcome["outcomes"][0][
            "originalLabel"
        ] = "Different outcome"
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B1",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([wrong_outcome]),
                inventory=inventory,
                require_complete=True,
            )

    def test_exact_study_title_and_scoped_arm_evidence_cover_labels(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell(
                    "C34",
                    "2.RESULT TENSION TEST C-MG+YOKE, S-MG+YOKE",
                ),
                self.cell(
                    "C37",
                    46171,
                    number_format="dd-mmm",
                ),
                self.cell("F27", "Normal"),
            ]
        )
        inventory = self.inventory([chunk])
        study = self.observation_study([])
        study.update(
            {
                "title": (
                    "2.RESULT TENSION TEST C-MG+YOKE, S-MG+YOKE"
                ),
                "evidence": [
                    {
                        "sheet": "Data",
                        "range": "C34",
                        "sourceText": (
                            "2.RESULT TENSION TEST C-MG+YOKE, "
                            "S-MG+YOKE"
                        ),
                    },
                    {
                        "sheet": "Data",
                        "range": "C26:H27",
                        "sourceText": (
                            "Assembly-specific Test and Normal labels"
                        ),
                    },
                ],
                "arms": [
                    {
                        "label": "Normal",
                        "condition": "Normal",
                        "evidence": [
                            {
                                "sheet": "Data",
                                "range": "E24",
                            }
                        ],
                        "factorValues": [],
                    }
                ],
            }
        )

        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(
            {
                "revision-1:1:C34": "STUDY_TITLE",
                "revision-1:1:F27": (
                    "ARM_LABEL_STUDY_EVIDENCE"
                ),
            },
            report["semanticCoverageBySourceCellKey"],
        )

        broad = copy.deepcopy(study)
        broad["evidence"][1]["range"] = "A1:Z999"
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!F27",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([broad]),
                inventory=inventory,
                require_complete=True,
            )

    def test_process_ng_rate_section_is_outcome_not_factor_label(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B15", "1. Check Process NG rate"),
                self.cell("C18", 14),
                self.cell("D18", 2),
            ]
        )
        inventory = self.inventory([chunk])

        semantic = next(
            item
            for item in inventory["semanticLabelCells"]
            if item["coordinate"] == "B15"
        )
        self.assertEqual(["OUTCOME_LABEL"], semantic["semanticRoles"])
        study = self.observation_study(
            [
                {
                    "valueNumber": 14,
                    "evidence": [
                        {"sheet": "Data", "range": "C18"}
                    ],
                },
                {
                    "valueNumber": 2,
                    "evidence": [
                        {"sheet": "Data", "range": "D18"}
                    ],
                }
            ]
        )
        study.update(
            {
                "title": "1. Check Process NG rate",
                "evidence": [
                    {
                        "sheet": "Data",
                        "range": "B15",
                        "sourceText": "1. Check Process NG rate",
                    }
                ],
            }
        )

        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(0, report["uncoveredSemanticLabelCellCount"])

    def test_merged_test_heading_is_not_a_distinct_arm_label(
        self,
    ) -> None:
        heading = self.cell("D19", "TEST")
        heading.update(
            {
                "mergeRole": "anchor",
                "mergeRange": "D19:E19",
            }
        )
        arm_cells = []
        for coordinate, label, merge_range in (
            ("D20", "TEST 1", "D20:E20"),
            ("D21", "TEST 2", "D21:E21"),
            ("D22", "TEST 3", "D22:E22"),
        ):
            cell = self.cell(coordinate, label)
            cell.update(
                {
                    "mergeRole": "anchor",
                    "mergeRange": merge_range,
                }
            )
            arm_cells.append(cell)
        chunk = self.chunk(cells=[heading, *arm_cells])

        inventory = self.inventory([chunk])

        semantic_by_coordinate = {
            item["coordinate"]: item["semanticRoles"]
            for item in inventory["semanticLabelCells"]
        }
        self.assertNotIn("D19", semantic_by_coordinate)

    def test_merged_type_heading_is_not_a_factor_level(
        self,
    ) -> None:
        cells = []
        for coordinate, label, merge_range in (
            ("D13", "Process", "D13:E13"),
            ("D14", "Vision Bond BP/SM", "D14:E14"),
            ("D15", "Vision Bond MG/PT", "D15:E15"),
            ("D16", "Vision Bond Yoke", "D16:E16"),
            ("D18", "Type", "D18:F18"),
            ("D19", "Type 1: NG Over bond BP/SM", "D19:F22"),
        ):
            cell = self.cell(coordinate, label)
            cell.update(
                {
                    "mergeRole": "anchor",
                    "mergeRange": merge_range,
                }
            )
            cells.append(cell)
        chunk = self.chunk(cells=cells)

        inventory = self.inventory([chunk])

        semantic_by_coordinate = {
            item["coordinate"]: item["semanticRoles"]
            for item in inventory["semanticLabelCells"]
        }
        self.assertEqual(["FACTOR_LABEL"], semantic_by_coordinate["D13"])
        self.assertNotIn("D18", semantic_by_coordinate)

    def test_source_roles_block_header_laundering_and_allow_one_point(
        self,
    ) -> None:
        laundering = self.chunk(
            cells=[
                self.cell("B2", 999),
                self.cell("C2", 998),
                self.cell("B3", 1),
                self.cell("C3", 2),
            ]
        )
        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B2:C2",
                "valueRange": "B3:C3",
                "rowIdentityRange": "A3",
                "axisSource": "HEADER",
            }
        ]
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B2",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([study]),
                inventory=self.inventory([laundering]),
                require_complete=True,
            )

        one_point = self.chunk(
            cells=[
                self.cell("B1", "100 Hz"),
                self.cell("B2", "80 dB"),
            ]
        )
        one_point_study = self.observation_study([])
        one_point_study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B1",
                "valueRange": "B2",
                "rowIdentityRange": "A2",
                "axisSource": "HEADER",
            }
        ]
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([one_point_study]),
            inventory=self.inventory([one_point]),
            require_complete=True,
        )
        self.assertEqual(
            2,
            report["coveredSemanticLabelCellCount"],
        )

    def test_field_swaps_and_reused_scalar_source_fail(self) -> None:
        ratio_chunk = self.chunk(
            cells=[
                self.cell("A1", "NG Count"),
                self.cell("B1", 3),
                self.cell("A2", "Sample Size"),
                self.cell("B2", 60),
            ]
        )
        swapped = self.observation_study(
            [
                {
                    "numerator": 60,
                    "denominator": 3,
                    "evidence": [
                        {"sheet": "Data", "range": "B1:B2"}
                    ],
                }
            ]
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "field/source binding",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([swapped]),
                inventory=self.inventory([ratio_chunk]),
                require_complete=True,
            )
        correct = self.observation_study(
            [
                {
                    "numerator": 3,
                    "denominator": 60,
                    "evidence": [
                        {"sheet": "Data", "range": "B1:B2"}
                    ],
                }
            ]
        )
        coverage.validate_content_manifest_coverage(
            manifest=self.manifest([correct]),
            inventory=self.inventory([ratio_chunk]),
            require_complete=True,
        )

        one_source = self.chunk(cells=[self.cell("B2", 5)])
        duplicated = self.observation_study(
            [
                {
                    "valueNumber": 5,
                    "evidence": [
                        {"sheet": "Data", "range": "B2"}
                    ],
                },
                {
                    "valueNumber": 5,
                    "evidence": [
                        {"sheet": "Data", "range": "B2"}
                    ],
                },
            ]
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "field/source binding",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([duplicated]),
                inventory=self.inventory([one_source]),
                require_complete=True,
            )

        shared_sample_size = copy.deepcopy(duplicated)
        shared_sample_size["outcomes"][0][
            "metricType"
        ] = "sample_size"
        coverage.validate_content_manifest_coverage(
            manifest=self.manifest([shared_sample_size]),
            inventory=self.inventory([one_source]),
            require_complete=True,
        )

    def test_composite_outcome_preserves_subheader_in_exact_evidence(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Total NG"),
                self.cell("B1", "NG rate"),
                self.cell("A2", 5),
                self.cell(
                    "B2",
                    0.05,
                    number_format="0.00%",
                ),
            ]
        )
        study = self.observation_study(
            [
                {
                    "valueNumber": 5,
                    "evidence": [
                        {"sheet": "Data", "range": "A2"}
                    ],
                },
                {
                    "valueNumber": 5,
                    "evidence": [
                        {"sheet": "Data", "range": "B2"}
                    ],
                },
            ]
        )
        study["outcomes"][0].update(
            {
                "originalLabel": "Total NG function voltage normal",
                "metricType": "defect_count_and_rate",
                "evidence": [
                    {
                        "sheet": "Data",
                        "range": "A1",
                        "sourceText": "Total NG",
                    },
                    {
                        "sheet": "Data",
                        "range": "B1",
                        "sourceText": "NG rate",
                    },
                ],
            }
        )
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )
        self.assertEqual(
            1,
            report["coveredSemanticLabelCellCount"],
        )

    def test_min_max_and_baseline_changed_are_field_specific(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Minimum value"),
                self.cell("B1", 1),
                self.cell("A2", "Maximum value"),
                self.cell("B2", 9),
                self.cell("A3", "Baseline value"),
                self.cell("B3", 10),
                self.cell("A4", "Changed value"),
                self.cell("B4", 20),
            ]
        )
        study = self.observation_study(
            [
                {
                    "min": 1,
                    "max": 9,
                    "evidence": [
                        {"sheet": "Data", "range": "B1:B2"}
                    ],
                }
            ]
        )
        study["factors"] = [
            {
                "key": "condition",
                "baselineCondition": 10,
                "changedCondition": 20,
                "evidence": [
                    {"sheet": "Data", "range": "B3:B4"}
                ],
            }
        ]
        coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )
        swapped = copy.deepcopy(study)
        swapped["outcomes"][0]["observations"][0].update(
            {"min": 9, "max": 1}
        )
        swapped["factors"][0].update(
            {"baselineCondition": 20, "changedCondition": 10}
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "field/source binding",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([swapped]),
                inventory=self.inventory([chunk]),
                require_complete=True,
            )

    def test_bare_stage_labels_do_not_retype_every_row_metric(
        self,
    ) -> None:
        self.assertEqual("", coverage._field_role("Before", ""))
        self.assertEqual("", coverage._field_role("After", ""))
        self.assertEqual(
            "BASELINE",
            coverage._field_role("Before value", ""),
        )
        self.assertEqual(
            "CHANGED",
            coverage._field_role("After value", ""),
        )

    def test_no_candidate_cannot_hide_adjacent_source_conclusion(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Conclusion"),
                self.cell("A2", "Bond increase raised NG."),
            ]
        )
        inventory = self.inventory(
            [chunk],
            [
                {
                    "chunkId": chunk["chunkId"],
                    "status": "NO_CANDIDATE",
                    "candidates": [],
                }
            ],
        )
        self.assertEqual(
            ["A2"],
            [
                item["coordinate"]
                for item in inventory["narrativeConclusionCells"]
            ],
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "source conclusion cell",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest(
                    [self.observation_study([])]
                ),
                inventory=inventory,
                require_complete=True,
            )

    def test_status_legend_is_excluded_but_actual_row_is_required(
        self,
    ) -> None:
        legend = self.chunk(
            cells=[
                self.cell("A1", "Legend"),
                self.cell("B1", "PASS"),
                self.cell("C1", "FAIL"),
            ]
        )
        self.assertEqual(
            0,
            self.inventory([legend])[
                "categoricalStatusCellCount"
            ],
        )

        actual = self.chunk(
            cells=[
                self.cell("A1", "Specimen"),
                self.cell("B1", "Status"),
                self.cell("A2", "#1"),
                self.cell("B2", "PASS"),
                self.cell("C2", 1),
            ]
        )
        inventory = self.inventory([actual])
        self.assertEqual(
            ["B2"],
            [
                item["coordinate"]
                for item in inventory["categoricalStatusCells"]
            ],
        )
        study = self.observation_study(
            [
                {
                    "valueText": "PASS",
                    "evidence": [
                        {"sheet": "Data", "range": "B2"}
                    ],
                },
                {
                    "valueNumber": 1,
                    "evidence": [
                        {"sheet": "Data", "range": "C2"}
                    ],
                },
            ]
        )
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual(
            1,
            report["coveredCategoricalStatusCellCount"],
        )

    def test_vertical_result_rows_are_not_reclassified_as_axes(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B1", "Run"),
                self.cell("B2", 10),
                self.cell("B3", 11),
                self.cell("B4", 12),
            ]
        )
        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B1",
                "valueRange": "B2:B4",
                "rowIdentityRange": "A2:A4",
                "axisSource": "HEADER",
            }
        ]
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )
        self.assertEqual(3, report["coveredCellCount"])

    def test_numeric_header_axis_requires_two_aligned_value_rows(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B1", 1),
                self.cell("C1", 2),
                self.cell("B2", 10),
                self.cell("C2", 11),
                self.cell("B3", 12),
                self.cell("C3", 13),
            ]
        )
        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B1:C1",
                "valueRange": value_range,
                "rowIdentityRange": row_range,
                "axisSource": "HEADER",
            }
            for value_range, row_range in (
                ("B2:C2", "A2"),
                ("B3:C3", "A3"),
            )
        ]
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )
        self.assertEqual(6, report["coveredCellCount"])

        empty_rows = self.chunk(
            cells=[
                self.cell("B1", 1),
                self.cell("C1", 2),
            ]
        )
        with self.assertRaisesRegex(
            coverage.ContentCoverageError,
            "Data!B1",
        ):
            coverage.validate_content_manifest_coverage(
                manifest=self.manifest([study]),
                inventory=self.inventory([empty_rows]),
                require_complete=True,
            )

    def test_changed_row_label_does_not_block_generic_observation(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Changed bonding lot"),
                self.cell("B1", 3),
            ]
        )
        study = self.observation_study(
            [
                {
                    "valueNumber": 3,
                    "evidence": [
                        {"sheet": "Data", "range": "B1"}
                    ],
                }
            ]
        )
        coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )

    def test_prior_total_ng_does_not_cross_replicate_header_boundary(
        self,
    ) -> None:
        cells = [
            self.cell("P16", "Total NG"),
            self.cell("P24", "#8"),
        ]
        coordinates: list[str] = []
        for row, base in ((25, 0.507), (26, 0.450)):
            for offset, column in enumerate("GHIJKLMNOP"):
                coordinate = f"{column}{row}"
                coordinates.append(coordinate)
                cells.append(
                    self.cell(
                        coordinate,
                        round(base + offset / 1000, 3),
                    )
                )
        chunk = self.chunk(cells=cells)
        inventory = self.inventory([chunk])
        roles = {
            item["coordinate"]: item["sourceRole"]
            for item in inventory["requiredCells"]
        }
        self.assertEqual("RESULT", roles["P25"])
        self.assertEqual("RESULT", roles["P26"])

        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "G24:P24",
                "valueRange": "G25:P26",
                "rowIdentityRange": "F25:F26",
                "axisSource": "HEADER",
            }
        ]
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual(
            len(coordinates),
            report["coveredCellCount"],
        )

    def test_bare_test_row_identity_is_not_changed_field_role(
        self,
    ) -> None:
        cells = [self.cell("F18", "Test")]
        observations: list[dict] = []
        numeric_coordinates: list[str] = []
        for offset, column in enumerate("GHIJKLMNOP"):
            coordinate = f"{column}18"
            value = 80 + offset
            numeric_coordinates.append(coordinate)
            cells.append(self.cell(coordinate, value))
            observations.append(
                {
                    "valueNumber": value,
                    "evidence": [
                        {"sheet": "Data", "range": coordinate}
                    ],
                }
            )
        cells.append(self.cell("F30", "Test-Auto"))
        for coordinate, value in (
            ("G30", 1.1),
            ("I30", 1.2),
            ("J30", 1.3),
        ):
            numeric_coordinates.append(coordinate)
            cells.append(self.cell(coordinate, value))
            observations.append(
                {
                    "valueNumber": value,
                    "evidence": [
                        {"sheet": "Data", "range": coordinate}
                    ],
                }
            )
        chunk = self.chunk(cells=cells)
        inventory = self.inventory([chunk])
        field_roles = {
            item["coordinate"]: item["fieldRole"]
            for item in inventory["requiredCells"]
        }
        self.assertTrue(
            all(
                field_roles[coordinate] == ""
                for coordinate in numeric_coordinates
            )
        )

        study = self.observation_study(observations)
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "G17:P17",
                "valueRange": "G18:P18",
                "rowIdentityRange": "F18",
                "axisSource": "ROW_IDENTITY",
            },
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "G29:J29",
                "valueRange": "G30:J30",
                "rowIdentityRange": "F30",
                "axisSource": "ROW_IDENTITY",
            },
        ]
        study["arms"] = [
            {
                "label": "Test",
                "condition": "Test",
                "evidence": [
                    {"sheet": "Data", "range": "F18"}
                ],
                "factorValues": [],
            }
        ]
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=inventory,
            require_complete=True,
        )
        self.assertEqual([], report["bindingErrors"])

        explicit = self.chunk(
            sheet_index=2,
            cells=[
                self.cell(
                    "A1",
                    "Changed value",
                    sheet_index=2,
                ),
                self.cell("B1", 20, sheet_index=2),
                self.cell(
                    "A2",
                    "Test value",
                    sheet_index=2,
                ),
                self.cell("B2", 30, sheet_index=2),
            ],
        )
        explicit_roles = {
            item["coordinate"]: item["fieldRole"]
            for item in self.inventory([explicit])["requiredCells"]
        }
        self.assertEqual("CHANGED", explicit_roles["B1"])
        self.assertEqual("CHANGED", explicit_roles["B2"])

    def test_count_source_can_support_its_derived_rate(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "NG Count"),
                self.cell("B1", 3),
                self.cell("A2", "Sample Size"),
                self.cell("B2", 60),
                self.cell("A3", "NG Rate"),
                self.cell("B3", 0.05),
            ]
        )
        study = self.observation_study(
            [
                {
                    "valueNumber": 3,
                    "numerator": 3,
                    "denominator": 60,
                    "evidence": [
                        {"sheet": "Data", "range": "B1:B2"}
                    ],
                },
                {
                    "valueNumber": 0.05,
                    "numerator": 3,
                    "denominator": 60,
                    "evidence": [
                        {"sheet": "Data", "range": "B1:B3"}
                    ],
                },
            ]
        )
        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )
        self.assertEqual(3, report["coveredCellCount"])

    def test_bare_no_is_sequence_but_no_sound_is_an_outcome(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "No"),
                self.cell("A2", 1),
                self.cell("C1", "No sound"),
                self.cell("D1", 0),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertEqual(
            ["A2"],
            [
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"] == "SEQUENCE_LABEL"
            ],
        )
        self.assertIn(
            "D1",
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )

    def test_header_axis_series_cover_exact_replicate_identities(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A2", 1),
                self.cell("A3", 2),
                self.cell("B1", "Position 1"),
                self.cell("C1", "Position 2"),
                self.cell("B2", 410),
                self.cell("C2", 411),
                self.cell("B3", 412),
                self.cell("C3", 413),
            ]
        )
        study = self.observation_study([])
        study["measurementSeries"] = [
            {
                "seriesRole": "RAW",
                "sheet": "Data",
                "headerRange": "B1:C1",
                "valueRange": f"B{row}:C{row}",
                "rowIdentityRange": f"A{row}",
                "axisSource": "HEADER",
            }
            for row in (2, 3)
        ]

        report = coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([chunk]),
            require_complete=True,
        )

        self.assertEqual(6, report["coveredCellCount"])
        self.assertEqual(0, report["uncoveredCellCount"])

    def test_repeated_ir_identifier_preserves_first_and_excludes_duplicates(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("E13", "IR"),
                self.cell("F13", "Input"),
                self.cell("E15", 250418008),
                self.cell("F15", 490),
                self.cell("E17", 250418008),
                self.cell("F17", 587),
                self.cell("E19", 250418008),
                self.cell("F19", 492),
                self.cell("E21", 250418008),
                self.cell("F21", 1569),
            ]
        )
        inventory = self.inventory([chunk])
        self.assertIn(
            "E15",
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )
        self.assertEqual(
            ["E17", "E19", "E21"],
            [
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"]
                == "DUPLICATE_IDENTIFIER"
            ],
        )

        non_identifier = self.chunk(
            sheet_index=2,
            cells=[
                self.cell(
                    "E13",
                    "Result",
                    sheet_index=2,
                ),
                self.cell(
                    "E15",
                    250418008,
                    sheet_index=2,
                ),
                self.cell(
                    "F15",
                    1,
                    sheet_index=2,
                ),
                self.cell(
                    "E17",
                    250418008,
                    sheet_index=2,
                ),
                self.cell(
                    "F17",
                    2,
                    sheet_index=2,
                ),
            ],
        )
        inventory = self.inventory([non_identifier])
        self.assertNotIn(
            "DUPLICATE_IDENTIFIER",
            {
                item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

    def test_compound_outcome_component_and_table_headers(
        self,
    ) -> None:
        compound = self.chunk(
            cells=[
                self.cell("F1", "Sigma"),
                self.cell("G2", "THD"),
                self.cell("G3", 3),
            ]
        )
        study = self.observation_study(
            [
                {
                    "valueNumber": 3,
                    "evidence": [
                        {"sheet": "Data", "range": "G3"}
                    ],
                }
            ]
        )
        study["outcomes"][0].update(
            {
                "originalLabel": "Sigma THD",
                "evidence": [
                    {"sheet": "Data", "range": "F1:G2"}
                ],
            }
        )
        coverage.validate_content_manifest_coverage(
            manifest=self.manifest([study]),
            inventory=self.inventory([compound]),
            require_complete=True,
        )

        headings = self.chunk(
            cells=[
                self.cell("A1", "III. Result"),
                self.cell("B2", "1. Result checking function"),
                self.cell("D1", "Row"),
                self.cell("E1", "OK"),
                self.cell("E3", 10),
            ]
        )
        inventory = self.inventory([headings])
        self.assertEqual(0, inventory["semanticLabelCellCount"])
        self.assertEqual(0, inventory["categoricalStatusCellCount"])

    def test_semantic_labels_do_not_cross_word_or_blank_boundaries(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("A1", "Assembly method"),
                self.cell("A2", "Method A"),
                self.cell("A10", "Unrelated note"),
                self.cell("C1", "Improve VP"),
                self.cell("C2", 3),
            ]
        )
        inventory = self.inventory([chunk])
        roles = {
            (item["coordinate"], role)
            for item in inventory["semanticLabelCells"]
            for role in item["semanticRoles"]
        }
        self.assertIn(("A2", "FACTOR_LEVEL"), roles)
        self.assertNotIn(("A10", "FACTOR_LEVEL"), roles)
        self.assertFalse(
            any(
                coordinate == "C1"
                for coordinate, _role in roles
            )
        )

    def test_merged_factor_group_header_defers_to_leaf_grid(
        self,
    ) -> None:
        parent = self.cell("C4", "Condition test")
        parent.update(
            {
                "mergeRange": "C4:E4",
                "mergeRole": "anchor",
            }
        )
        cells = [
            parent,
            self.cell("C5", "Bonding amount"),
            self.cell("D5", "AWF machine"),
            self.cell("E5", "Frame"),
            self.cell("C6", "Bonding amount 1.8~2.0mg"),
            self.cell("D6", "AWF MC : #3 ( Coil : 0.096 )"),
            self.cell("E6", "Frame normal"),
        ]
        inventory = self.inventory([self.chunk(cells=cells)])
        roles = {
            (item["coordinate"], role)
            for item in inventory["semanticLabelCells"]
            for role in item["semanticRoles"]
        }
        self.assertNotIn(("C4", "FACTOR_LABEL"), roles)
        self.assertIn(("C5", "FACTOR_LABEL"), roles)

        explicit_locator = self.locator(
            self.chunk(cells=cells),
            evidence=[
                {
                    "sheet": "Data",
                    "range": "C4",
                    "role": "FACTOR_LABEL",
                }
            ],
        )
        explicit_inventory = self.inventory(
            [self.chunk(cells=cells)],
            [explicit_locator],
        )
        self.assertFalse(
            any(
                item["coordinate"] == "C4"
                and "FACTOR_LABEL" in item["semanticRoles"]
                for item in explicit_inventory["semanticLabelCells"]
            )
        )

    def test_merged_factor_matrix_leafs_bind_levels_structurally(
        self,
    ) -> None:
        parent = self.cell("J8", "S-MG")
        parent.update(
            {
                "mergeRange": "J8:K8",
                "mergeRole": "anchor",
            }
        )
        normal_1 = self.cell("J10", "Normal")
        normal_1.update(
            {
                "mergeRange": "J10:K10",
                "mergeRole": "anchor",
            }
        )
        normal_2 = self.cell("J11", "Normal")
        normal_2.update(
            {
                "mergeRange": "J11:K11",
                "mergeRole": "anchor",
            }
        )
        chunk = self.chunk(
            cells=[
                parent,
                self.cell("J9", "Spec"),
                self.cell("K9", "Supplier"),
                normal_1,
                normal_2,
                self.cell("J12", "0.69~0.70"),
                self.cell("K12", "Maglong"),
            ]
        )

        roles = {
            (item["coordinate"], role)
            for item in self.inventory([chunk])[
                "semanticLabelCells"
            ]
            for role in item["semanticRoles"]
        }

        self.assertNotIn(("J9", "FACTOR_LABEL"), roles)
        self.assertNotIn(("K9", "FACTOR_LABEL"), roles)
        self.assertIn(("J10", "FACTOR_LEVEL"), roles)
        self.assertIn(("J11", "FACTOR_LEVEL"), roles)
        self.assertNotIn(("J10", "ARM_LABEL"), roles)

        bare = self.chunk(
            sheet_index=2,
            cells=[
                self.cell("A1", "Spec", sheet_index=2),
                self.cell("A2", "Normal", sheet_index=2),
            ],
        )
        bare_roles = {
            (item["coordinate"], role)
            for item in self.inventory([bare])[
                "semanticLabelCells"
            ]
            for role in item["semanticRoles"]
        }
        self.assertNotIn(("A1", "FACTOR_LABEL"), bare_roles)

    def test_merged_leaf_and_single_factor_labels_remain_required(
        self,
    ) -> None:
        merged_leaf = self.cell("A1", "Assembly method")
        merged_leaf.update(
            {
                "mergeRange": "A1:C1",
                "mergeRole": "anchor",
            }
        )
        chunk = self.chunk(
            cells=[
                merged_leaf,
                self.cell("A2", "Method A"),
                self.cell("B2", "Method B"),
                self.cell("C2", "Method C"),
                self.cell("E1", "Process condition"),
                self.cell("E2", "Normal"),
            ]
        )
        roles = {
            (item["coordinate"], role)
            for item in self.inventory([chunk])["semanticLabelCells"]
            for role in item["semanticRoles"]
        }
        self.assertIn(("A1", "FACTOR_LABEL"), roles)
        self.assertIn(("E1", "FACTOR_LABEL"), roles)

    def test_typed_mg_matrix_headers_bind_normal_levels_below(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("D6", "Type"),
                self.cell("E6", "C-MG"),
                self.cell("F6", "S-MG (New JIG 0.03)"),
                self.cell("D7", 1),
                self.cell("E7", "Normal"),
                self.cell("F7", "Normal"),
                self.cell("D8", 2),
                self.cell("E8", 0.7),
                self.cell("F8", "Normal"),
                self.cell("D9", 3),
                self.cell("E9", "Normal"),
                self.cell("F9", 0.7),
            ]
        )
        roles = {
            (item["coordinate"], role)
            for item in self.inventory([chunk])["semanticLabelCells"]
            for role in item["semanticRoles"]
        }

        for coordinate in ("E7", "F7", "F8", "E9"):
            self.assertIn((coordinate, "FACTOR_LEVEL"), roles)
            self.assertNotIn((coordinate, "ARM_LABEL"), roles)

    def test_empty_min_max_normal_summary_row_is_not_an_arm(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("C10", "Min"),
                self.cell("C11", "Max"),
                self.cell("C12", "Normal"),
            ]
        )

        roles = {
            (item["coordinate"], role)
            for item in self.inventory([chunk])["semanticLabelCells"]
            for role in item["semanticRoles"]
        }

        self.assertNotIn(("C12", "ARM_LABEL"), roles)

    def test_context_kind_covers_exact_semantic_factor_label(self) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("E7", "CD LOT Date"),
            ]
        )
        locator = self.locator(
            chunk,
            evidence=[
                {
                    "sheet": "Data",
                    "range": "E7",
                    "role": "FACTOR_LABEL",
                }
            ],
        )
        inventory = self.inventory([chunk], [locator])
        manifest = self.manifest(
            [
                {
                    "contexts": [
                        {
                            "key": "cd-lot-date",
                            "kind": "CD LOT Date",
                            "originalValue": "2025-03-17",
                            "evidence": [
                                {
                                    "sheet": "Data",
                                    "range": "E7",
                                }
                            ],
                        }
                    ],
                    "factors": [],
                    "arms": [],
                    "outcomes": [],
                    "measurementSeries": [],
                    "conclusions": [],
                }
            ]
        )

        report = coverage.validate_content_manifest_coverage(
            manifest=manifest,
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(1, report["coveredSemanticLabelCellCount"])

    def test_com_error_variant_and_hidden_companion_are_not_results(
        self,
    ) -> None:
        left_error = self.cell(
            "A2",
            None,
            formula="=+#REF!-#REF!",
            cached_value=-2146826265,
        )
        helper = self.cell("B2", 0)
        right_error = self.cell(
            "C2",
            None,
            formula="=B2/A2",
            cached_value=-2146826265,
        )
        for cell in (left_error, helper, right_error):
            cell["hidden"] = {
                "row": False,
                "column": True,
                "sheet": False,
            }
        visible = self.cell("D2", 5)

        inventory = self.inventory(
            [self.chunk(cells=[
                left_error,
                helper,
                right_error,
                visible,
            ])]
        )

        self.assertEqual(
            {"D2"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )
        self.assertEqual(
            {"B2": "HIDDEN_ERROR_GRID_INPUT"},
            {
                item["coordinate"]: item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

    def test_quoted_date_and_gapped_no_column_are_structural(
        self,
    ) -> None:
        chunk = self.chunk(
            cells=[
                self.cell("B3", "NO."),
                self.cell("B5", 1),
                self.cell("B6", 2),
                self.cell("B7", 4),
                self.cell("W4", "Start"),
                self.cell(
                    "W5",
                    46217,
                    number_format='mm"-"dd',
                ),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertEqual(0, inventory["requiredCellCount"])
        self.assertEqual(
            {
                "B5": "SEQUENCE_LABEL",
                "B6": "SEQUENCE_LABEL",
                "B7": "SEQUENCE_LABEL",
                "W5": "DATE_FORMAT",
            },
            {
                item["coordinate"]: item["exclusionReason"]
                for item in inventory["excludedCells"]
            },
        )

    def test_series_axes_and_observation_strata_preserve_labels(
        self,
    ) -> None:
        series_chunk = self.chunk(
            sheet="Series",
            sheet_index=1,
            cells=[
                self.cell("A1", "Process", sheet_index=1),
                self.cell("A2", "Material Lower", sheet_index=1),
                self.cell("A3", "Ultrasonic", sheet_index=1),
                self.cell("B1", "Input", sheet_index=1),
                self.cell("B2", 10, sheet_index=1),
                self.cell("B3", 20, sheet_index=1),
                self.cell("E2", "Row", sheet_index=1),
                self.cell("M1", "18/7", sheet_index=1),
                self.cell("N1", "17/7", sheet_index=1),
                self.cell("M2", 1, sheet_index=1),
                self.cell("N2", 2, sheet_index=1),
            ],
        )
        scalar_chunk = self.chunk(
            sheet="Scalar",
            sheet_index=2,
            cells=[
                self.cell("D1", "Process", sheet_index=2),
                self.cell("D2", "Bonding", sheet_index=2),
                self.cell("F2", 360, sheet_index=2),
            ],
        )
        manifest = self.manifest(
            [
                {
                    "contexts": [],
                    "factors": [
                        {
                            "key": "process",
                            "originalLabel": "Process",
                            "evidence": [
                                {
                                    "sheet": "Series",
                                    "range": "A1",
                                }
                            ],
                        }
                    ],
                    "arms": [],
                    "outcomes": [],
                    "measurementSeries": [
                        {
                            "seriesRole": "RAW",
                            "sheet": "Series",
                            "headerRange": "B1",
                            "valueRange": "B2:B3",
                            "rowIdentityRange": "A2:A3",
                            "axisSource": "ROW_IDENTITY",
                        },
                        {
                            "seriesRole": "RAW",
                            "sheet": "Series",
                            "headerRange": "M1:N1",
                            "valueRange": "M2:N2",
                            "rowIdentityRange": "E2",
                            "axisSource": "HEADER",
                        },
                    ],
                    "conclusions": [],
                },
                {
                    "contexts": [],
                    "factors": [
                        {
                            "key": "process",
                            "originalLabel": "Process",
                            "evidence": [
                                {
                                    "sheet": "Scalar",
                                    "range": "D1",
                                }
                            ],
                        }
                    ],
                    "arms": [],
                    "outcomes": [
                        {
                            "observations": [
                                {
                                    "stratumKey": "Bonding",
                                    "valueNumber": 360,
                                    "evidence": [
                                        {
                                            "sheet": "Scalar",
                                            "range": "D2",
                                        },
                                        {
                                            "sheet": "Scalar",
                                            "range": "F2",
                                        },
                                    ],
                                }
                            ]
                        }
                    ],
                    "measurementSeries": [],
                    "conclusions": [],
                },
            ]
        )
        inventory = self.inventory([series_chunk, scalar_chunk])

        report = coverage.validate_content_manifest_coverage(
            manifest=manifest,
            inventory=inventory,
            require_complete=True,
        )

        self.assertEqual(0, report["uncoveredSemanticLabelCellCount"])
        self.assertEqual(0, report["uncoveredCellCount"])

    def test_wide_merged_field_headers_are_not_semantic_claims(
        self,
    ) -> None:
        headers = [
            self.cell("B3", "NO."),
            self.cell("C3", "Line"),
            self.cell("D3", "Process"),
            self.cell("T3", "4M"),
        ]
        for header in headers:
            column = header["coordinate"][0]
            header.update(
                {
                    "mergeRole": "anchor",
                    "mergeRange": f"{column}3:{column}4",
                }
            )
        chunk = self.chunk(
            cells=[
                *headers,
                self.cell("B5", 1),
                self.cell("B6", 2),
                self.cell("C5", "L1"),
                self.cell("C6", "L1"),
                self.cell("D5", "NG Function"),
                self.cell("D6", "Visual final"),
                self.cell("T5", "Machine"),
                self.cell("T6", "Man"),
            ]
        )

        inventory = self.inventory([chunk])

        self.assertNotIn(
            "D3",
            {
                item["coordinate"]
                for item in inventory["semanticLabelCells"]
            },
        )
        self.assertNotIn(
            "T3",
            {
                item["coordinate"]
                for item in inventory["semanticLabelCells"]
            },
        )

    def test_merged_status_header_above_excel_error_is_not_result(
        self,
    ) -> None:
        status_header = self.cell("H13", "OK")
        status_header.update(
            {
                "mergeRole": "anchor",
                "mergeRange": "H13:H14",
            }
        )
        peer_headers = [
            self.cell("B13", "Date"),
            self.cell("D13", "Item"),
            self.cell("F13", "Input"),
        ]
        error = self.cell(
            "H15",
            None,
            formula="=+#REF!",
            cached_value=-2146826265,
        )

        inventory = self.inventory(
            [self.chunk(cells=[
                *peer_headers,
                status_header,
                error,
            ])]
        )

        self.assertEqual([], inventory["categoricalStatusCells"])

    def test_exact_inventory_conclusion_is_attached_to_covering_study(
        self,
    ) -> None:
        source_text = (
            "- After improve => NG reduce from 2% => 0.4%"
        )
        inventory = {
            "schemaVersion": coverage.CONTENT_COVERAGE_SCHEMA_VERSION,
            "requiredCells": [],
            "semanticLabelCells": [],
            "categoricalStatusCells": [],
            "unresolvedFormulaCells": [],
            "narrativeConclusionCells": [
                {
                    "sourceCellKey": "revision-1:1:B33",
                    "sheet": "SUM",
                    "coordinate": "B33",
                    "row": 33,
                    "column": 2,
                    "sourceText": source_text,
                }
            ],
            "excludedCellCount": 0,
        }
        manifest = self.manifest(
            [
                {
                    "evidence": [
                        {
                            "sheet": "SUM",
                            "range": "B32:B33",
                            "sourceText": "IV. Decision",
                        }
                    ],
                    "contexts": [],
                    "factors": [],
                    "arms": [],
                    "outcomes": [],
                    "measurementSeries": [],
                    "conclusions": [],
                }
            ]
        )

        augmented = coverage.augment_exact_source_conclusions(
            manifest=manifest,
            inventory=inventory,
        )
        repeated = coverage.augment_exact_source_conclusions(
            manifest=augmented,
            inventory=inventory,
        )

        self.assertEqual([], manifest["studies"][0]["conclusions"])
        self.assertEqual(augmented, repeated)
        conclusion = augmented["studies"][0]["conclusions"][0]
        self.assertEqual(source_text, conclusion["text"])
        self.assertEqual("SOURCE_CONCLUSION", conclusion["claimType"])
        coverage.validate_content_manifest_coverage(
            manifest=augmented,
            inventory=inventory,
            require_complete=True,
        )

    def test_formula_label_lookup_and_layout_helpers_are_structural(
        self,
    ) -> None:
        cells = [
            self.cell("N1", 200),
            self.cell(
                "N2",
                None,
                formula='="THD_% ("&N1&"Hz)"',
                cached_value="THD_% (200Hz)",
            ),
            self.cell("N3", 300),
            self.cell("A8", 2),
            self.cell("A9", 32),
            self.cell("A10", 54),
            self.cell(
                "A11",
                None,
                formula="=A10+22",
                cached_value=76,
            ),
            self.cell(
                "A12",
                None,
                formula="=A11+22",
                cached_value=98,
            ),
            self.cell(
                "A13",
                None,
                formula="=A12+22",
                cached_value=120,
            ),
            self.cell(
                "B8",
                None,
                formula="=MATCH(C8,D1:D5,0)",
                cached_value=2,
            ),
            self.cell("C8", "STD"),
            self.cell("C9", "Test OK"),
            self.cell("C10", "Test NG"),
            self.cell("C11", "Normal"),
            self.cell(
                "F8",
                None,
                formula="=INDEX(D1:D5,B8)",
                cached_value=10,
            ),
            self.cell(
                "G8",
                None,
                formula="=INDEX(E1:E5,B8)",
                cached_value=11,
            ),
        ]

        inventory = self.inventory([self.chunk(cells=cells)])
        exclusion_by_coordinate = {
            item["coordinate"]: item["exclusionReason"]
            for item in inventory["excludedCells"]
        }

        self.assertEqual(
            "FORMULA_LABEL_INPUT",
            exclusion_by_coordinate["N1"],
        )
        self.assertEqual(
            "FORMULA_LOOKUP_INDEX",
            exclusion_by_coordinate["B8"],
        )
        self.assertEqual(
            {
                "A8",
                "A9",
                "A10",
                "A11",
                "A12",
                "A13",
            },
            {
                coordinate
                for coordinate, reason
                in exclusion_by_coordinate.items()
                if reason == "FORMULA_LAYOUT_SEQUENCE"
            },
        )
        self.assertTrue(
            {"N3", "F8", "G8"}.issubset(
                {
                    item["coordinate"]
                    for item in inventory["requiredCells"]
                }
            )
        )

    def test_direct_ratio_operands_are_not_formula_label_inputs(
        self,
    ) -> None:
        ratio = self.cell(
            "I42",
            None,
            number_format="0.00%",
            formula="=+I41/G41",
            cached_value=0.05,
        )
        ratio["displayValue"] = "5.00%"
        inventory = self.inventory(
            [
                self.chunk(
                    cells=[
                        self.cell("G41", 100),
                        self.cell("I41", 5),
                        ratio,
                    ]
                )
            ]
        )

        self.assertEqual(
            {"G41", "I41", "I42"},
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )
        self.assertNotIn(
            "I41",
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"] == "FORMULA_LABEL_INPUT"
            },
        )

    def test_index_selector_inputs_rendered_as_row_labels_are_structural(
        self,
    ) -> None:
        inventory = self.inventory(
            [
                self.chunk(
                    cells=[
                        self.cell("A53", 1),
                        self.cell(
                            "B53",
                            None,
                            formula="=INDEX($H$8:$DV$8,1,A53)",
                            cached_value="STD_AVG",
                        ),
                        self.cell("A54", 2),
                        self.cell(
                            "B54",
                            None,
                            formula="=INDEX($H$8:$DV$8,1,A54)",
                            cached_value="STD #1",
                        ),
                        self.cell("C54", 528.27),
                    ]
                )
            ]
        )

        self.assertEqual(
            {"A53", "A54"},
            {
                item["coordinate"]
                for item in inventory["excludedCells"]
                if item["exclusionReason"] == "FORMULA_LABEL_INPUT"
            },
        )
        self.assertIn(
            "C54",
            {
                item["coordinate"]
                for item in inventory["requiredCells"]
            },
        )


if __name__ == "__main__":
    unittest.main()
