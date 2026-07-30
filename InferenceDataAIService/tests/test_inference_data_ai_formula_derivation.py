from __future__ import annotations

import copy
import json
import sqlite3
import unittest

import inference_data_ai_content_coverage as content_coverage
import inference_data_ai_formula_derivation as derivation
import inference_data_ai_study_import as study_import


REVISION_UID = "capture_revision_formula_test"
CONTENT_SHA256 = "a" * 64


def _cell(
    coordinate: str,
    value: object = None,
    *,
    formula: str | None = None,
    data_type: str | None = None,
) -> dict:
    source_key = f"{REVISION_UID}:1:{coordinate}"
    return {
        "coordinate": coordinate,
        "row": int("".join(char for char in coordinate if char.isdigit())),
        "column": 1,
        "sourceCellKey": source_key,
        "formula": formula,
        "rawValue": None if formula else value,
        "cachedValue": None,
        "displayValue": None if formula else value,
        "dataType": data_type or ("f" if formula else "n"),
        "cachedDataType": None,
        "numberFormat": "General",
        "valueSource": "FORMULA_NO_CACHE" if formula else "RAW",
    }


def _chunks(cells: list[dict]) -> list[dict]:
    return [
        {
            "chunkId": "source_chunk_test",
            "sheet": {"title": "Data", "sheetIndex": 1},
            "sourceRevision": {
                "revisionUid": REVISION_UID,
                "contentSha256": CONTENT_SHA256,
            },
            "cells": cells,
            "contextCells": [],
        }
    ]


def _sheet_chunk(title: str, sheet_index: int, cells: list[dict]) -> dict:
    chunk = copy.deepcopy(_chunks(cells)[0])
    chunk["chunkId"] = f"source_chunk_test_{sheet_index}"
    chunk["sheet"] = {"title": title, "sheetIndex": sheet_index}
    for cell in chunk["cells"]:
        cell["sourceCellKey"] = (
            f"{REVISION_UID}:{sheet_index}:{cell['coordinate']}"
        )
    return chunk


class FormulaDerivationTests(unittest.TestCase):
    def test_restricted_grammar_derives_numbers_and_div0(self) -> None:
        chunks = _chunks(
            [
                _cell("A1", 10),
                _cell("A2", 5),
                _cell("A3", "ignored", data_type="s"),
                _cell("B1", formula="=+A1-A2*2"),
                _cell("B2", formula="=SUM(A1:A3, 2)"),
                _cell("B3", formula="=AVERAGE(A1:A2)"),
                _cell("B4", formula="=MIN(A1:A2)+MAX(A1:A2)"),
                _cell("B5", formula="=A1/(A2-A2)"),
                _cell("B6", formula="=B5+1"),
            ]
        )

        overlay = derivation.derive_formula_overlay(chunks)

        self.assertEqual(overlay["formulaCount"], 6)
        self.assertEqual(overlay["numericCount"], 4)
        self.assertEqual(overlay["errorCount"], 2)
        self.assertEqual(overlay["errorsByCode"], {"#DIV/0!": 2})
        values = overlay["valuesBySourceCellKey"]
        self.assertEqual(values[f"{REVISION_UID}:1:B1"]["numericValue"], 0)
        self.assertEqual(values[f"{REVISION_UID}:1:B2"]["numericValue"], 17)
        self.assertEqual(values[f"{REVISION_UID}:1:B3"]["numericValue"], 7.5)
        self.assertEqual(values[f"{REVISION_UID}:1:B4"]["numericValue"], 15)

    def test_projection_is_immutable_idempotent_and_preserves_formula(self) -> None:
        chunks = _chunks(
            [
                _cell("A1", 4),
                _cell("B1", formula="=A1/2"),
            ]
        )
        original = copy.deepcopy(chunks)
        overlay = derivation.derive_formula_overlay(chunks)

        projected = derivation.apply_formula_overlay_to_chunks(chunks, overlay)
        projected_again = derivation.apply_formula_overlay_to_chunks(
            projected,
            overlay,
        )

        self.assertEqual(chunks, original)
        self.assertEqual(projected_again, projected)
        formula_cell = projected[0]["cells"][1]
        self.assertEqual(formula_cell["formula"], "=A1/2")
        self.assertEqual(formula_cell["sourceCellKey"], f"{REVISION_UID}:1:B1")
        self.assertEqual(formula_cell["coordinate"], "B1")
        self.assertEqual(formula_cell["cachedValue"], 2)
        self.assertEqual(formula_cell["cachedDataType"], "n")
        self.assertEqual(
            formula_cell["valueSource"],
            derivation.DERIVED_VALUE_SOURCE,
        )
        self.assertIn("provenanceSha256", formula_cell["formulaDerivation"])

    def test_content_inventory_reads_projection_without_changing_source(
        self,
    ) -> None:
        chunks = _chunks(
            [
                _cell("A1", 4),
                _cell("B1", formula="=A1/2"),
            ]
        )
        overlay = derivation.derive_formula_overlay(chunks)
        projected = derivation.apply_formula_overlay_to_chunks(chunks, overlay)

        raw_inventory = content_coverage.build_content_coverage_inventory(
            chunks=chunks,
            locator_results=[],
        )
        derived_inventory = content_coverage.build_content_coverage_inventory(
            chunks=projected,
            locator_results=[],
        )

        self.assertEqual(raw_inventory["numericCellCount"], 1)
        self.assertEqual(raw_inventory["unresolvedFormulaCellCount"], 1)
        self.assertEqual(derived_inventory["numericCellCount"], 2)
        self.assertEqual(derived_inventory["unresolvedFormulaCellCount"], 0)
        self.assertIsNone(chunks[0]["cells"][1]["cachedValue"])

    def test_exact_lookup_rejects_formula_or_revision_mismatch(self) -> None:
        chunks = _chunks([_cell("A1", 4), _cell("B1", formula="=A1/2")])
        overlay = derivation.derive_formula_overlay(chunks)
        entry = derivation.formula_overlay_entry(
            overlay,
            sheet="Data",
            coordinate="B1",
            formula="=A1/2",
            revision_uid=REVISION_UID,
        )
        self.assertEqual(entry["numericValue"], 2)
        with self.assertRaises(derivation.FormulaDerivationError):
            derivation.formula_overlay_entry(
                overlay,
                sheet="Data",
                coordinate="B1",
                formula="=A1/3",
            )
        with self.assertRaises(derivation.FormulaDerivationError):
            derivation.formula_overlay_entry(
                overlay,
                sheet="Data",
                coordinate="B1",
                formula="=A1/2",
                revision_uid="other_revision",
            )

    def test_fail_closed_for_unsupported_external_cycle_and_reference_error(
        self,
    ) -> None:
        cases = [
            [_cell("A1", 1), _cell("B1", formula="=IF(A1,1,0)")],
            [_cell("A1", 1), _cell("B1", formula="='Other'!A1")],
            [_cell("A1", formula="=B1"), _cell("B1", formula="=A1")],
            [
                _cell("A1", formula="=#REF!+1"),
                _cell("B1", formula="=A1+1"),
            ],
            [
                _cell("A1", "#REF!", data_type="e"),
                _cell("B1", formula="=A1+1"),
            ],
        ]
        for cells in cases:
            with self.subTest(formula=cells[-1]["formula"]):
                with self.assertRaises(derivation.FormulaDerivationError):
                    derivation.derive_formula_overlay(_chunks(cells))

    def test_tolerant_overlay_derives_supported_formulas_and_preserves_unsupported(
        self,
    ) -> None:
        chunks = _chunks(
            [
                _cell("A1", formula="=#REF!+1"),
                _cell("B1", formula="=1+1"),
            ]
        )

        overlay = derivation.derive_formula_overlay(
            chunks,
            tolerate_unsupported=True,
        )
        projected = derivation.apply_formula_overlay_to_chunks(chunks, overlay)

        self.assertEqual(1, overlay["numericCount"])
        self.assertEqual(1, overlay["errorCount"])
        self.assertEqual(
            "UNSUPPORTED",
            overlay["valuesBySourceCellKey"][f"{REVISION_UID}:1:A1"][
                "status"
            ],
        )
        self.assertIsNone(projected[0]["cells"][0]["cachedValue"])
        self.assertEqual(2, projected[0]["cells"][1]["cachedValue"])

    def test_direct_cross_sheet_reference_derives_numeric_dependency(self) -> None:
        chunks = [
            _sheet_chunk(
                "Data",
                1,
                [_cell("A1", formula="='Source O''Brien'!$C$2")],
            ),
            _sheet_chunk(
                "Source O'Brien",
                2,
                [
                    _cell("B2", 6.25),
                    _cell("C2", formula="=B2*2"),
                ],
            ),
        ]

        overlay = derivation.derive_formula_overlay(chunks)
        projected = derivation.apply_formula_overlay_to_chunks(chunks, overlay)

        self.assertEqual(2, overlay["numericCount"])
        self.assertEqual(0, overlay["nonNumericCount"])
        self.assertEqual(0, overlay["errorCount"])
        self.assertEqual(
            12.5,
            overlay["valuesBySourceCellKey"][f"{REVISION_UID}:1:A1"][
                "numericValue"
            ],
        )
        self.assertEqual(12.5, projected[0]["cells"][0]["cachedValue"])
        dependency = overlay["valuesBySourceCellKey"][
            f"{REVISION_UID}:1:A1"
        ]["dependencySourceCellKeys"]
        self.assertEqual([f"{REVISION_UID}:2:C2"], dependency)

    def test_text_formulas_are_classified_without_numeric_projection(self) -> None:
        chunks = [
            _sheet_chunk(
                "Data",
                1,
                [
                    _cell("A1", "label", data_type="s"),
                    _cell("A2", 5),
                    _cell("B1", formula="=A1"),
                    _cell("B2", formula='=$A$1&" #"&A2'),
                    _cell(
                        "B3",
                        formula='=IF(AND(A2<=10,A2>=1),"OK","NG")',
                    ),
                    _cell("B4", formula="='Labels'!A1"),
                ],
            ),
            _sheet_chunk("Labels", 2, [_cell("A1", "external label", data_type="s")]),
        ]

        overlay = derivation.derive_formula_overlay(chunks)
        projected = derivation.apply_formula_overlay_to_chunks(chunks, overlay)

        self.assertEqual(4, overlay["formulaCount"])
        self.assertEqual(0, overlay["numericCount"])
        self.assertEqual(4, overlay["nonNumericCount"])
        self.assertEqual(0, overlay["errorCount"])
        self.assertEqual(
            {"DIRECT_REFERENCE_TO_TEXT", "TEXT_CONCATENATION", "TEXT_IF_AND"},
            {
                entry["nonNumericReason"]
                for entry in overlay["valuesBySourceCellKey"].values()
            },
        )
        self.assertTrue(
            all(
                cell["cachedValue"] is None
                for cell in projected[0]["cells"]
                if cell.get("formula")
            )
        )

    def test_cross_sheet_range_remains_fail_closed(self) -> None:
        chunks = [
            _sheet_chunk(
                "Data",
                1,
                [_cell("A1", formula="='Source'!A1:A2")],
            ),
            _sheet_chunk("Source", 2, [_cell("A1", 1), _cell("A2", 2)]),
        ]

        with self.assertRaises(derivation.FormulaDerivationError):
            derivation.derive_formula_overlay(chunks)
        overlay = derivation.derive_formula_overlay(
            chunks,
            tolerate_unsupported=True,
        )
        self.assertEqual(1, overlay["errorCount"])
        self.assertEqual(
            "UNSUPPORTED",
            next(iter(overlay["valuesBySourceCellKey"].values()))["status"],
        )

    def test_overlay_tamper_and_stale_source_are_rejected(self) -> None:
        chunks = _chunks([_cell("A1", 4), _cell("B1", formula="=A1/2")])
        overlay = derivation.derive_formula_overlay(chunks)
        tampered = copy.deepcopy(overlay)
        tampered["valuesBySourceCellKey"][f"{REVISION_UID}:1:B1"][
            "numericValue"
        ] = 999
        with self.assertRaises(derivation.FormulaDerivationError):
            derivation.apply_formula_overlay_to_chunks(chunks, tampered)

        stale = copy.deepcopy(chunks)
        stale[0]["cells"][0]["rawValue"] = 6
        stale[0]["cells"][0]["displayValue"] = 6
        with self.assertRaises(derivation.FormulaDerivationError):
            derivation.validate_formula_overlay(stale, overlay)

    def test_numeric_and_measurement_evidence_use_provenance_overlay_only(
        self,
    ) -> None:
        chunks = _chunks(
            [
                _cell("A1", "condition", data_type="s"),
                _cell("A2", "sample-1", data_type="s"),
                _cell("B2", formula="=1+1"),
            ]
        )
        overlay = derivation.derive_formula_overlay(chunks)
        connection = sqlite3.connect(":memory:")
        connection.row_factory = sqlite3.Row
        connection.executescript(
            """
            CREATE TABLE capture_v2_sheets(
                sheet_id INTEGER PRIMARY KEY,
                revision_id INTEGER NOT NULL,
                title TEXT NOT NULL
            );
            CREATE TABLE capture_v2_cells(
                sheet_id INTEGER NOT NULL,
                row_index INTEGER NOT NULL,
                column_index INTEGER NOT NULL,
                coordinate TEXT NOT NULL,
                raw_value_json TEXT,
                formula_text TEXT,
                cached_value_json TEXT,
                display_value_json TEXT,
                number_format TEXT,
                merge_range TEXT,
                merge_role TEXT
            );
            INSERT INTO capture_v2_sheets(sheet_id, revision_id, title)
            VALUES (1, 7, 'Data');
            """
        )
        rows = [
            (1, 1, 1, "A1", '"condition"', None, None, '"condition"'),
            (1, 2, 1, "A2", '"sample-1"', None, None, '"sample-1"'),
            (1, 2, 2, "B2", None, "=1+1", None, None),
        ]
        connection.executemany(
            """
            INSERT INTO capture_v2_cells(
                sheet_id, row_index, column_index, coordinate,
                raw_value_json, formula_text, cached_value_json,
                display_value_json, number_format, merge_range, merge_role
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'General', NULL, 'none')
            """,
            rows,
        )
        revision = connection.execute(
            """
            SELECT
                1 AS is_current,
                7 AS capture_v2_revision_id,
                ? AS revision_uid,
                ? AS content_sha256
            """,
            (REVISION_UID, CONTENT_SHA256),
        ).fetchone()
        evidence = [{"sheet": "Data", "range": "B2"}]

        self.assertEqual(
            study_import._numeric_cells_from_capture_evidence(
                connection,
                revision,
                evidence,
            ),
            [],
        )
        lookup = study_import._formula_lookup_for_revision(revision, overlay)
        self.assertEqual(
            study_import._numeric_cells_from_capture_evidence(
                connection,
                revision,
                evidence,
                formula_lookup=lookup,
            ),
            [(2.0, False)],
        )
        points = study_import._expand_measurement_series(
            connection,
            revision=revision,
            series={
                "sheet": "Data",
                "headerRange": "A1",
                "valueRange": "B2",
                "rowIdentityRange": "A2",
                "axisSource": "HEADER",
                "seriesRole": "RAW",
                "axisUnit": "",
                "valueUnit": "",
            },
            series_uid="series-test",
            path="series",
            formula_lookup=lookup,
        )
        self.assertEqual(len(points), 1)
        self.assertEqual(points[0]["valueNumber"], 2)
        source_payload = json.loads(points[0]["sourceValueJson"])
        self.assertEqual(source_payload["type"], "float")
        self.assertEqual(source_payload["value"], 2)
        self.assertEqual(
            source_payload["derivation"]["sourceCellKey"],
            f"{REVISION_UID}:1:B2",
        )
        self.assertEqual(
            source_payload["derivation"]["formula"],
            "=1+1",
        )
        self.assertEqual(
            source_payload["derivation"]["overlaySha256"],
            overlay["overlaySha256"],
        )
        self.assertIsNone(
            connection.execute(
                """
                SELECT cached_value_json
                FROM capture_v2_cells
                WHERE coordinate='B2'
                """
            ).fetchone()["cached_value_json"]
        )
        connection.close()


if __name__ == "__main__":
    unittest.main()
