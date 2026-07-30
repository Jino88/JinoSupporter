from __future__ import annotations

import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import inference_data_ai_cli as cli
from inference_data_ai_table_first import (
    ANALYSIS_SCHEMA_VERSION,
    PROMPT_VERSION,
    TableFirstError,
    _compact_repeated_table_templates,
    _expand_repeated_table_analysis,
    build_table_first_prompt,
    build_table_first_request,
    normalize_table_first_analysis,
    project_table_first_analysis,
    run_codex_table_first_analysis,
    table_first_prompt_stats,
    validate_table_first_analysis,
)
from inference_data_ai_term_dictionary import TermDictionaryAdapter


def source_cell(
    row: int,
    column: int,
    value: object,
    *,
    display: object | None = None,
    number_format: str = "General",
) -> dict:
    column_label = ""
    number = column
    while number:
        number, remainder = divmod(number - 1, 26)
        column_label = chr(ord("A") + remainder) + column_label
    coordinate = f"{column_label}{row}"
    return {
        "sourceCellKey": f"revision_test:1:{coordinate}",
        "row": row,
        "column": column,
        "coordinate": coordinate,
        "rawValue": value,
        "displayValue": value if display is None else display,
        "cachedValue": None,
        "formula": None,
        "dataType": "n" if isinstance(value, (int, float)) else "s",
        "cachedDataType": None,
        "numberFormat": number_format,
        "mergeRange": None,
        "primary": True,
        "valueSource": "RAW",
    }


def packet(cells: list[dict], *, chunks: list[dict] | None = None) -> dict:
    if chunks is None:
        chunks = [
            {
                "sheet": {"sheetIndex": 1, "title": "Result"},
                "sourceRevision": {
                    "revisionUid": "revision_test",
                    "contentSha256": "a" * 64,
                },
                "sectionIndex": 1,
                "cells": cells,
            }
        ]
    return {
        "schemaVersion": "semantic-source-packet-v1",
        "inventory": {
            "sourceRevision": {
                "revisionUid": "revision_test",
                "contentSha256": "a" * 64,
                "fileName": "test.xlsx",
                "sourcePath": r"D:\input\test.xlsx",
            },
            "workbook": {
                "status": "CAPTURED",
                "sheetCount": 1,
            },
            "coverage": {"status": "COMPLETE"},
        },
        "chunks": chunks,
        "terminalPackets": [],
    }


def reviewed_term_adapter() -> TermDictionaryAdapter:
    return TermDictionaryAdapter.from_rows(
        [
            {
                "term_raw": "F0",
                "normalized_name": "Resonance Frequency",
                "definition_status": "DEFINED",
                "source_count": "10",
            },
            {
                "term_raw": "FO",
                "normalized_name": "Resonance Frequency",
                "definition_status": "DEFINED",
                "source_count": "5",
            },
            {
                "term_raw": "VP-CD",
                "normalized_name": "VP-CD Assembly",
                "definition_status": "DEFINED",
                "source_count": "10",
            },
            {
                "term_raw": "VP+CD",
                "normalized_name": "VP-CD Assembly",
                "definition_status": "DEFINED",
                "source_count": "5",
            },
            {
                "term_raw": "SPL",
                "normalized_name": "Sound Pressure Level",
                "definition_status": "DEFINED",
                "source_count": "10",
            },
            {
                "term_raw": "FOLLOW",
                "normalized_name": "Generic Word",
                "definition_status": "IGNORE",
                "source_count": "1",
            },
            {
                "term_raw": "UNCONFIRMED",
                "normalized_name": "Unknown",
                "definition_status": "NEEDS_DEFINITION",
                "source_count": "1",
            },
        ],
        source_path="reviewed-terms.csv",
        content_sha256="1" * 64,
    )


def valid_analysis(request: dict) -> dict:
    tables = []
    for index, table in enumerate(request["tables"], start=1):
        axis_refs = (
            [table["numericColumns"][0]["columnId"]]
            if table["numericColumns"]
            else []
        )
        tables.append(
            {
                "tableId": table["tableId"],
                "title": table["titleCandidates"][0]
                if table["titleCandidates"]
                else f"Table {index}",
                "type": "COMPARISON",
                "studyGroup": "study_height",
                "groups": [
                    {
                        "label": "Normal",
                        "role": "REFERENCE",
                        "basis": "Original table label",
                    },
                    {
                        "label": "Test",
                        "role": "TEST",
                        "basis": "Original table label",
                    },
                ],
                "metrics": [
                    {
                        "name": "Height",
                        "unit": "%",
                        "axisRefs": axis_refs,
                    }
                ],
                "comparisonRelations": [
                    {
                        "leftGroup": "Test",
                        "rightGroup": "Normal",
                        "basis": "Rows in the same table",
                    }
                ],
                "textLinks": list(table["nearbyTextIds"]),
                "relatedTableIds": [
                    other["tableId"]
                    for other in request["tables"]
                    if other["tableId"] != table["tableId"]
                ],
                "confidence": "MEDIUM",
                "limitations": [],
            }
        )
    return {
        "schemaVersion": ANALYSIS_SCHEMA_VERSION,
        "promptVersion": PROMPT_VERSION,
        "requestId": request["requestId"],
        "revisionUid": request["source"]["revisionUid"],
        "status": "ANALYZED" if tables else "NO_TABLES",
        "workbookSummary": "Height comparison tables.",
        "tables": tables,
        "notes": [],
    }


class TableFirstRequestTests(unittest.TestCase):
    def test_builds_table_text_inventory_and_code_owned_statistics(self) -> None:
        cells = [
            source_cell(1, 1, "Height test"),
            source_cell(2, 1, "Type"),
            source_cell(2, 2, "Height"),
            source_cell(3, 1, "Normal"),
            source_cell(3, 2, 0.22, display="22.00%", number_format="0.00%"),
            source_cell(4, 1, "Test"),
            source_cell(4, 2, 0.30, display="30.00%", number_format="0.00%"),
            source_cell(6, 1, "Test result is higher"),
        ]

        request = build_table_first_request(packet(cells))

        self.assertEqual("table-first-request-v1", request["schemaVersion"])
        self.assertEqual(1, len(request["tables"]))
        self.assertEqual(1, len(request["textBlocks"]))
        table = request["tables"][0]
        self.assertEqual("A1:B4", table["range"])
        self.assertEqual(["Height test", "Type | Height"], table["titleCandidates"])
        self.assertEqual(
            [request["textBlocks"][0]["textId"]],
            table["nearbyTextIds"],
        )
        numeric = table["numericColumns"][0]
        self.assertEqual(2, numeric["numericCount"])
        self.assertAlmostEqual(0.22, numeric["min"])
        self.assertAlmostEqual(0.30, numeric["max"])
        self.assertAlmostEqual(0.26, numeric["average"])
        self.assertEqual(
            "22.00%",
            numeric["displaySamples"][0]["normalizedDisplay"],
        )
        self.assertEqual(
            "PERCENT",
            numeric["displaySamples"][0]["displayScale"],
        )
        self.assertTrue(request["policy"]["statisticsAreCodeOwned"])
        self.assertEqual(request["requestBytes"], len(json.dumps(
            request,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")) + 1)

    def test_large_numeric_table_is_bounded_instead_of_dumping_every_cell(self) -> None:
        cells = [
            source_cell(1, 1, "Raw measurements"),
            source_cell(2, 1, "Sample"),
            source_cell(2, 2, "Value"),
        ]
        for row in range(3, 103):
            cells.append(source_cell(row, 1, f"S{row - 2}"))
            cells.append(source_cell(row, 2, row / 10))

        request = build_table_first_request(packet(cells))
        table = request["tables"][0]
        serialized = json.dumps(request, ensure_ascii=False)

        self.assertLessEqual(len(table["previewRows"]), 12)
        self.assertEqual(100, table["numericColumns"][0]["numericCount"])
        self.assertEqual(3, len(table["numericColumns"][0]["displaySamples"]))
        self.assertNotIn('"coordinate": "B50"', serialized)
        self.assertEqual(203, request["workbook"]["capturedPrimaryCellCount"])

    def test_very_wide_matrix_uses_one_compact_numeric_series(self) -> None:
        cells = []
        for column in range(1, 81):
            cells.append(source_cell(1, column, f"Point {column}"))
            cells.append(source_cell(2, column, column / 10))

        request = build_table_first_request(packet(cells))
        table = request["tables"][0]

        self.assertEqual(80, table["numericColumnCount"])
        self.assertLessEqual(len(table["numericColumns"]), 24)
        self.assertEqual(1, len(table["numericSeries"]))
        self.assertEqual(80, table["numericSeries"][0]["numericColumnCount"])
        self.assertEqual(80, table["numericSeries"][0]["numericCellCount"])

    def test_excludes_vertical_spl_frequency_response_from_ai_tables(self) -> None:
        frequencies = [
            100,
            150,
            200,
            300,
            500,
            800,
            1000,
            1500,
            2000,
            3000,
            5000,
            8000,
        ]
        cells = [
            source_cell(1, 1, "Sample"),
            source_cell(1, 2, "Test #1"),
        ]
        for row, frequency in enumerate(frequencies, start=2):
            cells.append(source_cell(row, 1, frequency))
            cells.append(source_cell(row, 2, 70.123456 + row))
        source_packet = packet(cells)
        source_packet["chunks"][0]["sheet"]["title"] = "SPL DATA_(NTI)"

        request = build_table_first_request(source_packet)

        self.assertEqual([], request["tables"])
        exclusions = request["codeOwnedExclusions"][
            "rawFrequencyResponseTables"
        ]
        self.assertEqual(1, len(exclusions))
        exclusion = exclusions[0]
        self.assertEqual("RAW_FREQUENCY_RESPONSE_DATA", exclusion["reason"])
        self.assertEqual(["SPL"], exclusion["metricFamilies"])
        self.assertEqual("A1:B13", exclusion["range"])
        self.assertTrue(exclusion["sourceStorage"]["coordinatesPreserved"])
        self.assertTrue(exclusion["sourceStorage"]["valuesPreserved"])
        self.assertEqual(
            "COLUMN",
            exclusion["frequencyAxes"][0]["orientation"],
        )
        workbook_exclusion = request["codeOwnedExclusions"][
            "workbookLearningExclusion"
        ]
        self.assertTrue(workbook_exclusion["excluded"])
        self.assertEqual(
            "WORKBOOK_CONTAINS_SPL_THD_IMP_RAW_FREQUENCY_DATA",
            workbook_exclusion["reason"],
        )
        self.assertEqual(["SPL"], workbook_exclusion["metricFamilies"])
        self.assertEqual(
            "EXCLUDED_RAW_FREQUENCY_RESPONSE_WORKBOOK",
            request["workbook"]["learningStatus"],
        )
        self.assertEqual(0, request["policy"]["aiCallBudget"])
        self.assertFalse(
            request["policy"]["workbookSemanticLearningEnabled"]
        )
        prompt = build_table_first_prompt(request)
        self.assertNotIn("codeOwnedExclusions", prompt)
        self.assertNotIn("70.123456", prompt)
        self.assertEqual("NO_TABLES", cli._no_table_first_analysis(request)["status"])

    def test_excludes_entire_workbook_when_raw_frequency_data_is_present(
        self,
    ) -> None:
        frequencies = [
            100,
            150,
            200,
            300,
            500,
            800,
            1000,
            1500,
            2000,
            3000,
            5000,
            8000,
        ]
        cells = [
            source_cell(1, 1, "Sample"),
            source_cell(1, 2, "Test #1"),
        ]
        for row, frequency in enumerate(frequencies, start=2):
            cells.append(source_cell(row, 1, frequency))
            cells.append(source_cell(row, 2, 70.0 + row))
        cells.extend(
            [
                source_cell(16, 1, "Height test"),
                source_cell(17, 1, "Type"),
                source_cell(17, 2, "Height"),
                source_cell(18, 1, "Normal"),
                source_cell(18, 2, 0.22),
                source_cell(19, 1, "Test"),
                source_cell(19, 2, 0.30),
                source_cell(22, 1, "Operator memo"),
            ]
        )
        source_packet = packet(cells)
        source_packet["chunks"][0]["sheet"]["title"] = "THD DATA_(NTI)"

        request = build_table_first_request(source_packet)

        self.assertEqual([], request["tables"])
        self.assertEqual([], request["textBlocks"])
        exclusion = request["codeOwnedExclusions"][
            "workbookLearningExclusion"
        ]
        self.assertEqual(["THD"], exclusion["metricFamilies"])
        self.assertEqual(1, exclusion["triggerTableCount"])
        self.assertEqual(1, exclusion["excludedNonRawTableCount"])
        self.assertEqual(1, exclusion["excludedTextBlockCount"])
        self.assertEqual(
            request["workbook"]["capturedPrimaryCellCount"],
            exclusion["excludedCapturedPrimaryCellCount"],
        )
        self.assertEqual(0, request["workbook"]["tableCount"])
        self.assertEqual(0, request["workbook"]["textBlockCount"])
        self.assertEqual("NO_TABLES", cli._no_table_first_analysis(request)["status"])

    def test_imp_raw_frequency_workbook_is_excluded_from_learning(
        self,
    ) -> None:
        frequencies = [
            100,
            150,
            200,
            300,
            500,
            800,
            1000,
            1500,
            2000,
            3000,
            5000,
            8000,
        ]
        cells = [
            source_cell(1, 1, "Frequency"),
            source_cell(1, 2, "Sample"),
        ]
        for row, frequency in enumerate(frequencies, start=2):
            cells.append(source_cell(row, 1, frequency))
            cells.append(source_cell(row, 2, 8.0 + row / 10))
        source_packet = packet(cells)
        source_packet["chunks"][0]["sheet"]["title"] = "IMP RAW DATA"

        request = build_table_first_request(source_packet)

        exclusion = request["codeOwnedExclusions"][
            "workbookLearningExclusion"
        ]
        self.assertTrue(exclusion["excluded"])
        self.assertEqual(["IMP"], exclusion["metricFamilies"])
        self.assertEqual([], request["tables"])

    def test_excludes_horizontal_hz_frequency_response(self) -> None:
        cells = [
            source_cell(1, 1, "Frequency response [dBSPL]"),
            source_cell(1, 2, "100.00Hz"),
            source_cell(1, 3, "500.00Hz"),
            source_cell(1, 4, "1.00kHz"),
            source_cell(1, 5, "2.00kHz"),
        ]
        for row in range(2, 5):
            cells.append(source_cell(row, 1, f"Sample {row - 1}"))
            for column in range(2, 6):
                cells.append(source_cell(row, column, 80 + row + column / 10))

        request = build_table_first_request(packet(cells))

        self.assertEqual([], request["tables"])
        exclusion = request["codeOwnedExclusions"][
            "rawFrequencyResponseTables"
        ][0]
        self.assertEqual(["SPL"], exclusion["metricFamilies"])
        self.assertEqual(
            "ROW",
            exclusion["frequencyAxes"][0]["orientation"],
        )

    def test_keeps_function_ng_summary_with_spl_and_thd_labels(self) -> None:
        cells = [
            source_cell(1, 1, "RESULT CHECKING FUNCTION"),
            source_cell(2, 1, "Type"),
            source_cell(2, 2, "Input"),
            source_cell(2, 3, "SPL"),
            source_cell(2, 4, "THD"),
            source_cell(2, 5, "SPL+THD+F0"),
            source_cell(3, 1, "Normal"),
            source_cell(3, 2, 200),
            source_cell(3, 3, 2),
            source_cell(3, 4, 1),
            source_cell(3, 5, 0),
            source_cell(4, 1, "Test"),
            source_cell(4, 2, 200),
            source_cell(4, 3, 0),
            source_cell(4, 4, 1),
            source_cell(4, 5, 1),
        ]

        request = build_table_first_request(packet(cells))

        self.assertEqual(1, len(request["tables"]))
        self.assertEqual(
            [],
            request["codeOwnedExclusions"]["rawFrequencyResponseTables"],
        )
        self.assertIsNone(
            request["codeOwnedExclusions"]["workbookLearningExclusion"]
        )
        self.assertTrue(
            request["policy"]["workbookSemanticLearningEnabled"]
        )

    def test_applies_reviewed_aliases_without_dropping_source_labels(self) -> None:
        adapter = reviewed_term_adapter()
        request = build_table_first_request(
            packet(
                [
                    source_cell(1, 1, "F0"),
                    source_cell(1, 2, 930.8),
                    source_cell(1, 3, 936.9),
                ]
            ),
            term_dictionary_adapter=adapter,
        )
        result = valid_analysis(request)
        result["tables"][0]["groups"] = [
            {"label": "VP-CD", "role": "REFERENCE", "basis": "Source"},
            {"label": "VP+CD", "role": "TEST", "basis": "Source alias"},
            {"label": "FOLLOW", "role": "OTHER", "basis": "Noise"},
        ]
        result["tables"][0]["metrics"] = [
            {"name": "FO", "unit": "Hz", "axisRefs": []},
        ]
        result["tables"][0]["comparisonRelations"] = [
            {
                "leftGroup": "VP+CD",
                "rightGroup": "VP-CD",
                "basis": "Source aliases",
            }
        ]

        normalized = normalize_table_first_analysis(result, request=request)

        table = normalized["tables"][0]
        self.assertEqual(
            ["VP-CD", "FOLLOW"],
            [group["label"] for group in table["groups"]],
        )
        self.assertEqual([], table["comparisonRelations"])
        self.assertEqual("DESCRIPTIVE", table["type"])
        self.assertEqual(["F0"], [metric["name"] for metric in table["metrics"]])
        self.assertEqual(
            request["tables"][0]["metricHints"][0]["axisRefs"],
            table["metrics"][0]["axisRefs"],
        )
        self.assertEqual(
            adapter.semantic_key("F0"),
            adapter.semantic_key("Resonance Frequency"),
        )
        self.assertNotEqual(
            adapter.semantic_key("SPL"),
            adapter.semantic_key("SPL Average"),
        )
        self.assertNotIn("codeOwnedTermDictionary", build_table_first_prompt(request))
        self.assertEqual(
            2,
            request["codeOwnedTermDictionary"]["aliasGroupCount"],
        )
        self.assertEqual(
            ["FOLLOW"],
            request["codeOwnedTermDictionary"]["ignoredTerms"],
        )
        self.assertTrue(adapter.is_ignored("FOLLOW"))
        self.assertEqual("follow", adapter.semantic_key("FOLLOW"))

    def test_derives_supported_formula_values_without_changing_source(self) -> None:
        formula = source_cell(
            3,
            3,
            None,
        )
        formula.update(
            {
                "formula": "=B3/A3",
                "dataType": "f",
                "rawValue": None,
                "displayValue": None,
                "valueSource": "FORMULA_NO_CACHE",
            }
        )
        cells = [
            source_cell(1, 1, "NG result"),
            source_cell(2, 1, "Input"),
            source_cell(2, 2, "NG"),
            source_cell(2, 3, "NG Rate"),
            source_cell(3, 1, 10),
            source_cell(3, 2, 2),
            formula,
        ]
        source_packet = packet(cells)

        request = build_table_first_request(source_packet)

        self.assertEqual("DERIVED", request["formulaDerivation"]["status"])
        self.assertEqual(1, request["formulaDerivation"]["numericCount"])
        rate_column = next(
            column
            for column in request["tables"][0]["numericColumns"]
            if column["column"] == "C"
        )
        self.assertEqual(0.2, rate_column["displaySamples"][0]["rawNumber"])
        self.assertEqual(
            "20.00%",
            rate_column["displaySamples"][0]["normalizedDisplay"],
        )
        self.assertEqual(
            "PERCENT_FROM_RATE_FORMULA",
            rate_column["displaySamples"][0]["displayScale"],
        )
        self.assertIsNone(formula["cachedValue"])

    def test_prompt_compacts_only_code_owned_numeric_details(self) -> None:
        request = build_table_first_request(
            packet(
                [
                    source_cell(1, 1, "Height test"),
                    source_cell(2, 1, "Type"),
                    source_cell(2, 2, "Height"),
                    source_cell(3, 1, "Normal"),
                    source_cell(3, 2, 0.22),
                    source_cell(4, 1, "Test"),
                    source_cell(4, 2, 0.30),
                ]
            )
        )
        table = request["tables"][0]
        table["aggregateChecks"] = [
            {"status": "MATCH", "rawRange": "B3:B4", "rawCount": 2}
        ]
        table["numericColumns"][0]["displaySamples"][0][
            "normalizedDisplay"
        ] = "22%"
        table["numericColumns"][0]["displaySamples"][0][
            "displayScale"
        ] = "PERCENT_FORMAT"
        original = json.dumps(request, ensure_ascii=False, sort_keys=True)

        prompt = build_table_first_prompt(request)
        prompt_request = json.loads(prompt.split("REQUEST_JSON:\n", 1)[1])
        prompt_table = prompt_request["tables"][0]
        prompt_column = prompt_table["numericColumns"][0]

        self.assertEqual(original, json.dumps(request, ensure_ascii=False, sort_keys=True))
        self.assertNotIn("aggregateChecks", prompt_table)
        self.assertEqual(
            {"count": 1, "statusCounts": {"MATCH": 1}},
            prompt_table["aggregateCheckSummary"],
        )
        self.assertNotIn("average", prompt_column)
        self.assertNotIn("min", prompt_column)
        self.assertNotIn("max", prompt_column)
        self.assertNotIn("numericCount", prompt_column)
        self.assertNotIn("sourceRange", prompt_column)
        self.assertNotIn("rawNumber", prompt_column["displaySamples"][0])
        self.assertIn("normalizedDisplay", prompt_column["displaySamples"][0])
        self.assertEqual(table["rowLabels"], prompt_table["rowLabels"])
        self.assertLess(
            len(prompt_table["previewRows"]),
            len(table["previewRows"]),
        )
        self.assertEqual(
            len(prompt.encode("utf-8")),
            table_first_prompt_stats(request)["promptBytes"],
        )

    def test_classifies_text_formula_without_reporting_numeric_error(self) -> None:
        formula = source_cell(2, 2, None)
        formula.update(
            {
                "formula": '=A2&" #1"',
                "dataType": "f",
                "rawValue": None,
                "displayValue": None,
                "valueSource": "FORMULA_NO_CACHE",
            }
        )
        source_packet = packet(
            [
                source_cell(1, 1, "Label"),
                source_cell(2, 1, "sample"),
                formula,
            ]
        )

        request = build_table_first_request(source_packet)

        self.assertEqual(
            "CLASSIFIED_NON_NUMERIC",
            request["formulaDerivation"]["status"],
        )
        self.assertEqual(0, request["formulaDerivation"]["numericCount"])
        self.assertEqual(1, request["formulaDerivation"]["nonNumericCount"])
        self.assertEqual(0, request["formulaDerivation"]["errorCount"])
        self.assertIsNone(formula["cachedValue"])

    def test_verifies_row_and_block_min_max_average_from_raw_values(self) -> None:
        row_cells = [
            source_cell(1, 1, "Tension result"),
            source_cell(2, 1, "Type"),
            source_cell(2, 2, "Min"),
            source_cell(2, 3, "Max"),
            source_cell(2, 4, "Average"),
            source_cell(2, 5, "#1"),
            source_cell(2, 6, "#2"),
            source_cell(2, 7, "#3"),
            source_cell(3, 1, "Test"),
            source_cell(3, 2, 1),
            source_cell(3, 3, 3),
            source_cell(3, 4, 2),
            source_cell(3, 5, 1),
            source_cell(3, 6, 2),
            source_cell(3, 7, 3),
        ]
        row_request = build_table_first_request(packet(row_cells))
        row_table = row_request["tables"][0]

        self.assertEqual(
            ["AGGREGATE_MIN", "AGGREGATE_MAX", "AGGREGATE_AVERAGE",
             "MEASURE_VALUE", "MEASURE_VALUE", "MEASURE_VALUE"],
            [column["columnRole"] for column in row_table["numericColumns"]],
        )
        self.assertEqual("MATCH", row_table["aggregateChecks"][0]["status"])
        self.assertEqual("ROW", row_table["aggregateChecks"][0]["mode"])
        self.assertEqual("E3:G3", row_table["aggregateChecks"][0]["rawRange"])

        block_cells = [
            source_cell(1, 1, "Height result"),
            source_cell(2, 1, "No"),
            source_cell(2, 2, "Measurement"),
            source_cell(2, 5, "Min"),
            source_cell(2, 6, "Max"),
            source_cell(2, 7, "Average"),
            source_cell(3, 1, 1),
            source_cell(3, 2, 1),
            source_cell(3, 3, 2),
            source_cell(3, 4, 3),
            source_cell(3, 5, 1),
            source_cell(3, 6, 6),
            source_cell(3, 7, 3.5),
            source_cell(4, 1, 2),
            source_cell(4, 2, 4),
            source_cell(4, 3, 5),
            source_cell(4, 4, 6),
        ]
        block_request = build_table_first_request(packet(block_cells))
        block_check = block_request["tables"][0]["aggregateChecks"][0]

        self.assertEqual("MATCH", block_check["status"])
        self.assertEqual("BLOCK", block_check["mode"])
        self.assertEqual("B3:D4", block_check["rawRange"])
        self.assertEqual(6, block_check["rawCount"])

        result_column_cells = [
            source_cell(1, 1, "Tension result"),
            source_cell(2, 1, "Type"),
            source_cell(2, 2, "Min"),
            source_cell(2, 3, "Max"),
            source_cell(2, 4, "Average"),
            source_cell(2, 5, "Sample 1"),
            source_cell(2, 6, "Sample 2"),
            source_cell(2, 7, "Result"),
            source_cell(3, 1, "Test"),
            source_cell(3, 2, 1),
            source_cell(3, 3, 2),
            source_cell(3, 4, 1.5),
            source_cell(3, 5, 1),
            source_cell(3, 6, 2),
            source_cell(3, 7, 2),
        ]
        result_column_request = build_table_first_request(
            packet(result_column_cells)
        )
        result_column_check = result_column_request["tables"][0][
            "aggregateChecks"
        ][0]

        self.assertEqual("MATCH", result_column_check["status"])
        self.assertEqual("E3:F3", result_column_check["rawRange"])
        self.assertEqual(2, result_column_check["rawCount"])

    def test_aggregate_checks_separate_triplets_and_basis_columns(self) -> None:
        cells = [
            source_cell(1, 1, "After Plasma"),
            source_cell(1, 2, "Sample #1"),
            source_cell(1, 3, "Sample #2"),
            source_cell(1, 4, "Min"),
            source_cell(1, 5, "Max"),
            source_cell(1, 6, "Average"),
            source_cell(1, 7, "OK #1"),
            source_cell(1, 8, "OK #2"),
            source_cell(1, 9, "Min"),
            source_cell(1, 10, "Max"),
            source_cell(1, 11, "Average"),
            source_cell(2, 1, 300),
            source_cell(2, 2, 1),
            source_cell(2, 3, 2),
            source_cell(2, 4, 1),
            source_cell(2, 5, 2),
            source_cell(2, 6, 1.5),
            source_cell(2, 7, 10),
            source_cell(2, 8, 20),
            source_cell(2, 9, 10),
            source_cell(2, 10, 20),
            source_cell(2, 11, 15),
        ]

        request = build_table_first_request(packet(cells))
        checks = request["tables"][0]["aggregateChecks"]

        self.assertEqual(2, len(checks))
        self.assertEqual(["MATCH", "MATCH"], [check["status"] for check in checks])
        self.assertEqual(["B2:C2", "G2:H2"], [check["rawRange"] for check in checks])

    def test_aggregate_checks_exclude_total_ng_from_sample_values(self) -> None:
        cells = [
            source_cell(1, 1, "Sample 1"),
            source_cell(1, 2, "Sample 2"),
            source_cell(1, 3, "Total NG"),
            source_cell(1, 4, "Max"),
            source_cell(1, 5, "Min"),
            source_cell(1, 6, "Avg"),
            source_cell(2, 1, 8.75),
            source_cell(2, 2, 14.85),
            source_cell(2, 3, 0),
            source_cell(2, 4, 14.85),
            source_cell(2, 5, 8.75),
            source_cell(2, 6, 11.8),
        ]

        request = build_table_first_request(packet(cells))
        check = request["tables"][0]["aggregateChecks"][0]

        self.assertEqual("MATCH", check["status"])
        self.assertEqual("A2:B2", check["rawRange"])
        self.assertEqual(2, check["rawCount"])

    def test_aggregate_checks_bound_single_column_blocks_by_summary_rows(self) -> None:
        cells = [
            source_cell(1, 1, "Vertical measurement"),
            source_cell(2, 2, "Measurement"),
            source_cell(2, 3, "Min"),
            source_cell(2, 4, "Max"),
            source_cell(2, 5, "Average"),
            source_cell(3, 2, 1),
            source_cell(3, 3, 1),
            source_cell(3, 4, 3),
            source_cell(3, 5, 2),
            source_cell(4, 2, 2),
            source_cell(5, 2, 3),
            source_cell(6, 2, 10),
            source_cell(6, 3, 10),
            source_cell(6, 4, 30),
            source_cell(6, 5, 20),
            source_cell(7, 2, 20),
            source_cell(8, 2, 30),
        ]

        request = build_table_first_request(packet(cells))
        checks = request["tables"][0]["aggregateChecks"]

        self.assertEqual(2, len(checks))
        self.assertEqual(["MATCH", "MATCH"], [check["status"] for check in checks])
        self.assertEqual(["BLOCK", "BLOCK"], [check["mode"] for check in checks])
        self.assertEqual(["B3:B5", "B6:B8"], [check["rawRange"] for check in checks])

    def test_splits_numbered_logical_result_sections_without_ai(self) -> None:
        cells = [
            source_cell(1, 1, "III. Result."),
            source_cell(2, 1, "1. Result checking height"),
            source_cell(3, 1, "Type"),
            source_cell(3, 2, "Value"),
            source_cell(4, 1, "Test"),
            source_cell(4, 2, 1.2),
            source_cell(5, 1, "2. Function NG Rate"),
            source_cell(6, 1, "Type"),
            source_cell(6, 2, "NG"),
            source_cell(7, 1, "Normal"),
            source_cell(7, 2, 0),
        ]

        request = build_table_first_request(packet(cells))

        self.assertEqual(2, len(request["tables"]))
        self.assertEqual(["A2:B4", "A5:B7"], [
            table["range"] for table in request["tables"]
        ])
        self.assertEqual(1, len(request["textBlocks"]))

    def test_splits_adjacent_single_row_metrics_into_independent_tables(self) -> None:
        cells = [
            source_cell(102, 14, "SPL Average"),
            source_cell(102, 15, 115.6),
            source_cell(102, 16, 116.4),
            source_cell(102, 59, "F0"),
            source_cell(102, 60, 930.8),
            source_cell(102, 61, 936.9),
        ]

        request = build_table_first_request(packet(cells))

        self.assertEqual(
            ["N102:P102", "BG102:BI102"],
            [table["range"] for table in request["tables"]],
        )
        self.assertEqual(
            [["SPL Average"], ["F0"]],
            [
                [hint["name"] for hint in table["metricHints"]]
                for table in request["tables"]
            ],
        )
        self.assertEqual(
            [],
            request["codeOwnedExclusions"]["rawFrequencyResponseTables"],
        )

    def test_adapts_stored_analysis_when_only_raw_tables_were_removed(self) -> None:
        request = build_table_first_request(
            packet(
                [
                    source_cell(1, 1, "Height test"),
                    source_cell(2, 1, "Type"),
                    source_cell(2, 2, "Height"),
                    source_cell(3, 1, "Normal"),
                    source_cell(3, 2, 0.22),
                    source_cell(4, 1, "Test"),
                    source_cell(4, 2, 0.30),
                ]
            )
        )
        request["codeOwnedExclusions"]["rawFrequencyResponseTables"] = [
            {"sourceTableId": "table_removed_raw"}
        ]
        existing = valid_analysis(request)
        existing["requestId"] = "table_request_before_raw_exclusion"
        extra = json.loads(json.dumps(existing["tables"][0]))
        extra["tableId"] = "table_removed_raw"
        extra["relatedTableIds"] = []
        extra["metrics"] = []
        existing["tables"].append(extra)

        adapted = cli._adapt_reusable_table_first_analysis(
            existing,
            request=request,
        )
        normalized = normalize_table_first_analysis(adapted, request=request)

        validated = validate_table_first_analysis(normalized, request=request)
        self.assertEqual(request["requestId"], validated["requestId"])
        self.assertEqual(
            [request["tables"][0]["tableId"]],
            [table["tableId"] for table in validated["tables"]],
        )

    def test_adapts_v3_analysis_when_retained_request_payload_is_identical(
        self,
    ) -> None:
        request = build_table_first_request(
            packet(
                [
                    source_cell(1, 1, "Height test"),
                    source_cell(2, 1, "Type"),
                    source_cell(2, 2, "Height"),
                    source_cell(3, 1, "Normal"),
                    source_cell(3, 2, 0.22),
                    source_cell(4, 1, "Test"),
                    source_cell(4, 2, 0.30),
                ]
            )
        )
        previous_request = json.loads(json.dumps(request))
        previous_request["builderVersion"] = "table-first-builder-v3"
        previous_request["requestId"] = "table_request_v3"
        existing = valid_analysis(request)
        existing["requestId"] = previous_request["requestId"]

        adapted = cli._adapt_reusable_table_first_analysis(
            existing,
            request=request,
            previous_request=previous_request,
        )

        validated = validate_table_first_analysis(
            normalize_table_first_analysis(adapted, request=request),
            request=request,
        )
        self.assertEqual(request["requestId"], validated["requestId"])

    def test_splits_top_fo_summary_from_following_response_matrix(self) -> None:
        cells = [
            source_cell(1, 7, "18 Kpa"),
            source_cell(2, 3, "Fo"),
            source_cell(2, 4, 654.8),
            source_cell(2, 7, 645.3),
            source_cell(2, 8, 681.7),
            source_cell(3, 3, "Sample"),
            source_cell(3, 4, "REF"),
            source_cell(3, 7, 1),
            source_cell(3, 8, 2),
            source_cell(4, 3, 100),
            source_cell(4, 4, 6.54),
            source_cell(4, 7, 6.62),
            source_cell(4, 8, 6.89),
            source_cell(5, 3, 106),
            source_cell(5, 4, 6.60),
            source_cell(5, 7, 6.67),
            source_cell(5, 8, 6.95),
        ]

        request = build_table_first_request(packet(cells))

        self.assertEqual(2, len(request["tables"]))
        summary, response = request["tables"]
        self.assertEqual("C1:H2", summary["range"])
        self.assertEqual(["Fo"], [hint["name"] for hint in summary["metricHints"]])
        self.assertEqual("C1:H5", response["range"])
        self.assertNotIn(
            "Fo",
            [
                label
                for row in response["rowLabels"]
                for label in row["labels"]
            ],
        )

    def test_normalizer_excludes_tiny_ref_error_fragment_as_text(self) -> None:
        cells = [
            source_cell(175, 1, 19000),
            source_cell(175, 2, "#REF!"),
            source_cell(176, 1, 20000),
            source_cell(176, 2, "#REF!"),
        ]

        request = build_table_first_request(packet(cells))
        result = valid_analysis(request)
        normalized = normalize_table_first_analysis(result, request=request)

        self.assertEqual(1, len(request["tables"]))
        self.assertEqual("TEXT", normalized["tables"][0]["type"])
        self.assertEqual([], normalized["tables"][0]["metrics"])
        self.assertEqual("HIGH", normalized["tables"][0]["confidence"])

    def test_cli_exposes_non_ai_prepare_and_single_call_analyze_commands(self) -> None:
        prepare = cli.build_parser().parse_args(
            ["table-first-request", "--packet", "packet.json"]
        )
        analyze = cli.build_parser().parse_args(
            ["table-first-analyze", "--request", "request.json"]
        )
        batch = cli.build_parser().parse_args(
            ["table-first-batch", "--packet-dir", "packets"]
        )
        request_batch = cli.build_parser().parse_args(
            ["table-first-request-batch", "--packet-dir", "packets"]
        )

        self.assertIs(prepare.func, cli.cmd_table_first_request)
        self.assertIs(analyze.func, cli.cmd_table_first_analyze)
        self.assertIs(batch.func, cli.cmd_table_first_batch)
        self.assertIs(request_batch.func, cli.cmd_table_first_request_batch)
        self.assertEqual("low", analyze.reasoning_effort)
        self.assertEqual(3, batch.workers)
        self.assertEqual(3, batch.max_value_samples)
        self.assertEqual(240000, request_batch.oversized_request_bytes)

    def test_request_batch_audits_packets_without_calling_ai(self) -> None:
        table_packet = packet(
            [
                source_cell(1, 1, "Height test"),
                source_cell(2, 1, "Type"),
                source_cell(2, 2, "Height"),
                source_cell(3, 1, "Normal"),
                source_cell(3, 2, 0.22),
                source_cell(4, 1, "Test"),
                source_cell(4, 2, 0.30),
            ]
        )
        text_packet = packet([source_cell(1, 1, "memo only")])
        text_packet["inventory"]["sourceRevision"].update(
            {
                "revisionUid": "revision_text",
                "contentSha256": "b" * 64,
                "fileName": "text.xlsx",
            }
        )
        text_packet["chunks"][0]["sourceRevision"].update(
            {
                "revisionUid": "revision_text",
                "contentSha256": "b" * 64,
            }
        )
        text_packet["chunks"][0]["cells"][0][
            "sourceCellKey"
        ] = "revision_text:1:A1"

        with tempfile.TemporaryDirectory(dir=cli.SERVICE_DIR) as temp:
            root = Path(temp)
            packet_dir = root / "packets"
            output_dir = root / "output"
            packet_dir.mkdir()
            (packet_dir / "one.json").write_text(
                json.dumps(table_packet, ensure_ascii=False),
                encoding="utf-8",
            )
            (packet_dir / "two.json").write_text(
                json.dumps(text_packet, ensure_ascii=False),
                encoding="utf-8",
            )
            args = cli.build_parser().parse_args(
                [
                    "table-first-request-batch",
                    "--packet-dir",
                    str(packet_dir),
                    "--out-dir",
                    str(output_dir),
                    "--workers",
                    "2",
                    "--checkpoint-every",
                    "1",
                ]
            )

            with (
                mock.patch.object(
                    cli,
                    "run_codex_table_first_analysis",
                    side_effect=AssertionError("AI must not be called"),
                ) as ai_call,
                mock.patch.object(cli, "print_json"),
            ):
                self.assertEqual(0, args.func(args))

            ai_call.assert_not_called()
            report = json.loads(
                (output_dir / "request-batch-report.json").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual("ok", report["status"])
            self.assertEqual(2, report["selected"])
            self.assertEqual(2, report["succeeded"])
            self.assertEqual(0, report["failed"])
            self.assertEqual(0, report["aiCalls"])
            self.assertEqual(1, report["plannedAiCalls"])
            self.assertEqual(1, report["noTables"])
            self.assertIn("promptBytes", report)
            self.assertTrue(
                all("promptBytes" in item for item in report["items"])
            )
            self.assertEqual(
                ["NO_TABLES"],
                next(
                    item["outlierReasons"]
                    for item in report["items"]
                    if item["fileName"] == "text.xlsx"
                ),
            )
            self.assertEqual(2, len(list((output_dir / "requests").glob("*.json"))))

    def test_batch_runs_one_ai_call_per_tabular_workbook_and_resumes(self) -> None:
        table_packet = packet(
            [
                source_cell(1, 1, "Height test"),
                source_cell(2, 1, "Type"),
                source_cell(2, 2, "Height"),
                source_cell(3, 1, "Normal"),
                source_cell(3, 2, 0.22),
                source_cell(4, 1, "Test"),
                source_cell(4, 2, 0.30),
            ]
        )
        text_packet = packet([source_cell(1, 1, "memo only")])
        text_packet["inventory"]["sourceRevision"].update(
            {
                "revisionUid": "revision_text",
                "contentSha256": "b" * 64,
                "fileName": "text.xlsx",
            }
        )
        text_packet["chunks"][0]["sourceRevision"].update(
            {
                "revisionUid": "revision_text",
                "contentSha256": "b" * 64,
            }
        )
        text_packet["chunks"][0]["cells"][0][
            "sourceCellKey"
        ] = "revision_text:1:A1"

        with tempfile.TemporaryDirectory(dir=cli.SERVICE_DIR) as temp:
            root = Path(temp)
            packet_dir = root / "packets"
            output_dir = root / "output"
            packet_dir.mkdir()
            (packet_dir / "one.json").write_text(
                json.dumps(table_packet, ensure_ascii=False),
                encoding="utf-8",
            )
            (packet_dir / "two.json").write_text(
                json.dumps(text_packet, ensure_ascii=False),
                encoding="utf-8",
            )
            args = cli.build_parser().parse_args(
                [
                    "table-first-batch",
                    "--packet-dir",
                    str(packet_dir),
                    "--out-dir",
                    str(output_dir),
                    "--workers",
                    "2",
                ]
            )
            calls: list[str] = []

            def fake_analysis(*, request, output_path, **_kwargs):
                calls.append(request["requestId"])
                result = valid_analysis(request)
                Path(output_path).write_bytes(cli.table_first_json_bytes(result))
                return result

            with (
                mock.patch.object(
                    cli,
                    "run_codex_table_first_analysis",
                    side_effect=fake_analysis,
                ),
                mock.patch.object(cli, "print_json"),
            ):
                self.assertEqual(0, args.func(args))
                self.assertEqual(1, len(calls))
                first_report = json.loads(
                    (output_dir / "batch-report.json").read_text(
                        encoding="utf-8"
                    )
                )
                self.assertEqual(1, first_report["newAnalyses"])
                self.assertEqual(1, first_report["noTables"])
                self.assertEqual(1, first_report["aiCalls"])

                self.assertEqual(0, args.func(args))
                self.assertEqual(1, len(calls))
                resumed_report = json.loads(
                    (output_dir / "batch-report.json").read_text(
                        encoding="utf-8"
                    )
                )
                self.assertEqual(2, resumed_report["reused"])
                self.assertEqual(0, resumed_report["aiCalls"])

    def test_compacts_three_repeated_templates_and_expands_to_source_order(
        self,
    ) -> None:
        chunks = []
        for sheet_index in range(1, 4):
            cells = [
                source_cell(1, 1, "Height test"),
                source_cell(2, 1, "Type"),
                source_cell(2, 2, "Height"),
                source_cell(3, 1, "Normal"),
                source_cell(3, 2, 0.22 + sheet_index / 100),
                source_cell(4, 1, "Test"),
                source_cell(4, 2, 0.30 + sheet_index / 100),
            ]
            for item in cells:
                item["sourceCellKey"] = (
                    f"revision_test:{sheet_index}:{item['coordinate']}"
                )
            chunks.append(
                {
                    "sheet": {
                        "sheetIndex": sheet_index,
                        "title": f"Day {sheet_index}",
                    },
                    "sourceRevision": {
                        "revisionUid": "revision_test",
                        "contentSha256": "a" * 64,
                    },
                    "sectionIndex": 1,
                    "cells": cells,
                }
            )
        request = build_table_first_request(packet([], chunks=chunks))

        prompt_request, families = (
            _compact_repeated_table_templates(request)
        )
        self.assertEqual(3, len(request["tables"]))
        self.assertEqual(1, len(prompt_request["tables"]))
        self.assertEqual(1, len(families))
        self.assertEqual(
            3,
            prompt_request["tables"][0]["templateOccurrenceCount"],
        )

        representative_analysis = valid_analysis(prompt_request)
        expanded = _expand_repeated_table_analysis(
            representative_analysis,
            request=request,
            prompt_request=prompt_request,
            families=families,
        )
        validated = validate_table_first_analysis(expanded, request=request)
        self.assertEqual(
            [table["tableId"] for table in request["tables"]],
            [table["tableId"] for table in validated["tables"]],
        )
        self.assertTrue(
            all(table["metrics"][0]["axisRefs"] for table in validated["tables"])
        )


class TableFirstAnalysisTests(unittest.TestCase):
    def setUp(self) -> None:
        first = [
            source_cell(1, 1, "Height test"),
            source_cell(2, 1, "Type"),
            source_cell(2, 2, "Height"),
            source_cell(3, 1, "Normal"),
            source_cell(3, 2, 0.22, display="22%"),
            source_cell(4, 1, "Test"),
            source_cell(4, 2, 0.30, display="30%"),
        ]
        second = [
            source_cell(10, 1, "Height repeat"),
            source_cell(11, 1, "Normal"),
            source_cell(11, 2, 0.24, display="24%"),
            source_cell(12, 1, "Test"),
            source_cell(12, 2, 0.31, display="31%"),
        ]
        self.request = build_table_first_request(packet([*first, *second]))
        self.analysis = valid_analysis(self.request)

    def test_validates_request_bound_references_and_projects_review_state(self) -> None:
        validated = validate_table_first_analysis(
            self.analysis,
            request=self.request,
        )
        projection = project_table_first_analysis(self.request, validated)

        self.assertEqual(2, len(validated["tables"]))
        self.assertEqual(1, len(projection["studies"]))
        self.assertEqual("NEEDS_REVIEW", projection["verificationStatus"])
        self.assertEqual(
            "NOT_ELIGIBLE_UNTIL_CANONICAL_REVIEW",
            projection["queryEligibility"],
        )
        facts = projection["studies"][0]["deterministicNumericFacts"]
        self.assertEqual(2, len(facts))
        self.assertTrue(
            all(
                fact["calculationAuthority"]
                == "CODE_FROM_CAPTURED_RAW_VALUES"
                for fact in facts
            )
        )

    def test_rejects_an_axis_reference_not_present_in_the_request(self) -> None:
        self.analysis["tables"][0]["metrics"][0]["axisRefs"] = ["invented_column"]

        with self.assertRaisesRegex(TableFirstError, "unknown ids"):
            validate_table_first_analysis(
                self.analysis,
                request=self.request,
            )

    def test_runner_uses_one_codex_process_and_writes_only_valid_output(self) -> None:
        calls: list[list[str]] = []

        def fake_run(command, **kwargs):
            calls.append(list(command))
            message_path = Path(
                command[command.index("--output-last-message") + 1]
            )
            message_path.write_text(
                json.dumps(self.analysis, ensure_ascii=False),
                encoding="utf-8",
            )
            self.assertIn("REQUEST_JSON:", kwargs["input"])
            return subprocess.CompletedProcess(command, 0, "", "")

        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "analysis.json"
            result = run_codex_table_first_analysis(
                request=self.request,
                output_path=output,
                codex_command=["fake-codex"],
                run_command=fake_run,
            )

            self.assertEqual(1, len(calls))
            self.assertEqual(self.analysis, result)
            self.assertTrue(output.is_file())

    def test_guard_removes_identity_comparisons_and_aggregate_metrics(self) -> None:
        first_table = self.request["tables"][0]
        aggregate_id = first_table["numericColumns"][0]["columnId"]
        first_table["numericColumns"][0]["columnRole"] = "AGGREGATE_MIN"
        result = valid_analysis(self.request)
        result["tables"][0]["groups"] = [
            {"label": "1", "role": "UNASSESSED", "basis": "Position"},
            {"label": "2", "role": "UNASSESSED", "basis": "Position"},
        ]
        result["tables"][0]["comparisonRelations"] = [
            {"leftGroup": "1", "rightGroup": "2", "basis": "Position"}
        ]
        result["tables"][0]["metrics"] = [
            {"name": "Min", "unit": "", "axisRefs": [aggregate_id]},
            {"name": "Height", "unit": "", "axisRefs": []},
        ]

        guarded = normalize_table_first_analysis(result, request=self.request)

        self.assertEqual("DESCRIPTIVE", guarded["tables"][0]["type"])
        self.assertEqual([], guarded["tables"][0]["groups"])
        self.assertEqual([], guarded["tables"][0]["comparisonRelations"])
        self.assertEqual(
            ["Height"],
            [metric["name"] for metric in guarded["tables"][0]["metrics"]],
        )

    def test_guard_deduplicates_repeated_group_labels(self) -> None:
        result = valid_analysis(self.request)
        result["tables"][0]["groups"].append(
            {
                "label": " Test ",
                "role": "TEST",
                "basis": "Repeated model output",
            }
        )

        guarded = normalize_table_first_analysis(
            result,
            request=self.request,
        )

        self.assertEqual(
            ["Normal", "Test"],
            [group["label"] for group in guarded["tables"][0]["groups"]],
        )

    def test_guard_uses_source_metric_hint_and_updates_partial_formula_text(
        self,
    ) -> None:
        source_table = self.request["tables"][0]
        source_table["metricHints"] = [
            {
                "name": "SPL Average",
                "axisRefs": [source_table["numericColumns"][0]["columnId"]],
            }
        ]
        self.request["formulaDerivation"]["status"] = "PARTIALLY_DERIVED"
        result = valid_analysis(self.request)
        result["tables"][0]["metrics"].append(
            {
                "name": "SPL",
                "unit": "dB",
                "axisRefs": [],
            }
        )
        result["tables"][0]["metrics"].append(
            {
                "name": "SPL Average",
                "unit": "",
                "axisRefs": [],
            }
        )
        result["tables"][0]["limitations"] = [
            "Formula derivation was skipped, and formula cells are not recalculated here."
        ]

        guarded = normalize_table_first_analysis(result, request=self.request)

        spl_metrics = [
            metric
            for metric in guarded["tables"][0]["metrics"]
            if metric["name"] == "SPL Average"
        ]
        self.assertEqual(1, len(spl_metrics))
        self.assertEqual("", spl_metrics[0]["unit"])
        self.assertEqual(
            ["SPL Average"],
            [metric["name"] for metric in guarded["tables"][0]["metrics"]],
        )
        self.assertEqual(
            source_table["metricHints"][0]["axisRefs"],
            spl_metrics[0]["axisRefs"],
        )
        self.assertEqual(
            [
                "Formula derivation was partial: supported formulas were derived, "
                "while unsupported formulas remain unresolved."
            ],
            guarded["tables"][0]["limitations"],
        )
        audit = cli._table_first_item_audit(self.request, guarded)
        self.assertEqual([], audit["reviewReasons"])
        self.assertFalse(audit["reviewRecommended"])

    def test_guard_verifies_ng_rate_as_ppm_from_input_and_ng_samples(self) -> None:
        cells = [
            source_cell(1, 1, "Input"),
            source_cell(1, 2, "NG"),
            source_cell(1, 3, "NG Rate"),
            source_cell(2, 1, 250),
            source_cell(2, 2, 24),
            source_cell(2, 3, 96000),
            source_cell(3, 1, 140),
            source_cell(3, 2, 11),
            source_cell(3, 3, 78571.42857142857),
            source_cell(4, 1, 188),
            source_cell(4, 2, 25),
            source_cell(4, 3, 132978.7234042553),
        ]
        request = build_table_first_request(packet(cells))
        result = valid_analysis(request)
        table = result["tables"][0]
        table.update(
            {
                "type": "DESCRIPTIVE",
                "groups": [],
                "comparisonRelations": [],
                "confidence": "LOW",
                "metrics": [
                    {
                        "name": "NG Rate",
                        "unit": "",
                        "axisRefs": [
                            request["tables"][0]["numericColumns"][2]["columnId"]
                        ],
                    }
                ],
                "limitations": [
                    "NG Rate uses a non-percent format and its scale is unclear."
                ],
            }
        )

        normalized = normalize_table_first_analysis(result, request=request)
        normalized_table = normalized["tables"][0]

        self.assertEqual("PPM", normalized_table["metrics"][0]["unit"])
        self.assertEqual("MEDIUM", normalized_table["confidence"])
        self.assertEqual([], normalized_table["limitations"])

    def test_guard_does_not_require_review_for_reliability_raw_data(self) -> None:
        cells = [
            source_cell(1, 1, "SPL RAW DATA"),
            source_cell(2, 1, "Before"),
            source_cell(2, 2, "After"),
            source_cell(3, 1, 115.0),
            source_cell(3, 2, 116.0),
        ]
        request = build_table_first_request(packet(cells))
        request["source"]["fileName"] = "Reliability Test result.xlsx"
        result = valid_analysis(request)
        result["tables"][0]["confidence"] = "LOW"

        normalized = normalize_table_first_analysis(result, request=request)
        audit = cli._table_first_item_audit(request, normalized)

        self.assertEqual("MEDIUM", normalized["tables"][0]["confidence"])
        self.assertFalse(audit["reviewRecommended"])


if __name__ == "__main__":
    unittest.main()
