from __future__ import annotations

import importlib.util
import json
import tempfile
import unittest
from unittest import mock
from pathlib import Path


CLI_PATH = Path(__file__).parents[1] / "inference_data_ai_cli.py"
SPEC = importlib.util.spec_from_file_location("inference_data_ai_cli", CLI_PATH)
assert SPEC is not None and SPEC.loader is not None
cli = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(cli)


def cell(row: int, column: int, value: object, merge: dict | None = None) -> dict:
    return {
        "row": row,
        "column": column,
        "colLabel": cli.excel_column_label(column),
        "address": cli.grid_cell_address(row, column),
        "value": value,
        "rawValue": value,
        "merge": merge or {"role": "none"},
    }


class UniversalGridIngestionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=cli.SERVICE_DIR)
        self.root = Path(self.temp.name)
        self.source = self.root / "fixture.xlsx"
        self.source.write_bytes(b"fixture-source")
        self.raw_json = self.root / "fixture.com-grid.json"
        self.db = self.root / "fixture.sqlite"

    def tearDown(self) -> None:
        self.temp.cleanup()

    def payload(self) -> dict:
        stat = self.source.stat()
        merge_address = "C2:D2"
        anchor = {"row": 2, "column": 3}
        anchor_merge = {
            "role": "anchor",
            "address": merge_address,
            "anchor": anchor,
            "anchorValue": "Merged header",
            "coveredCellMode": "blank",
        }
        covered_merge = {
            "role": "covered",
            "address": merge_address,
            "anchor": anchor,
            "anchorValue": "Merged header",
            "coveredCellMode": "blank",
        }
        return {
            "schemaVersion": "input-data-com-grid-v1",
            "extractedAt": "2026-07-11T00:00:00Z",
            "sourcePath": str(self.source.resolve()),
            "fileName": self.source.name,
            "fileSize": stat.st_size,
            "mtimeNs": stat.st_mtime_ns,
            "coveredCellMode": "blank",
            "includeEmptyCells": True,
            "sheets": [
                {
                    "sheetIndex": 1,
                    "sheetName": "Fixture",
                    "visible": True,
                    "usedRange": {"top": 2, "left": 2, "bottom": 3, "right": 4, "rowCount": 2, "columnCount": 3, "address": "B2:D3"},
                    "nonEmptyCells": 4,
                    "mergeCount": 1,
                    "merges": [
                        {
                            "address": merge_address,
                            "top": 2,
                            "left": 3,
                            "bottom": 2,
                            "right": 4,
                            "rowSpan": 1,
                            "columnSpan": 2,
                            "anchor": anchor,
                            "value": "Merged header",
                        }
                    ],
                    "rows": [
                        {
                            "rowNumber": 2,
                            "nonEmptyCount": 2,
                            "cells": [
                                cell(2, 2, "Title"),
                                cell(2, 3, "Merged header", anchor_merge),
                                cell(2, 4, None, covered_merge),
                            ],
                        },
                        {
                            "rowNumber": 3,
                            "nonEmptyCount": 2,
                            "cells": [
                                cell(3, 2, None),
                                cell(3, 3, "Value"),
                                cell(3, 4, 0),
                            ],
                        },
                    ],
                }
            ],
            "totals": {"sheetCount": 1, "rowCount": 2, "cellCount": 6, "nonEmptyCells": 4, "mergeCount": 1},
        }

    def write_payload(self, payload: dict | None = None) -> None:
        self.raw_json.write_text(json.dumps(payload or self.payload(), ensure_ascii=False), encoding="utf-8")

    def analysis_manifest(self) -> dict:
        return {
            "schemaVersion": "universal-analysis-v1",
            "source": {"dataset": "FixtureDataset", "sourcePath": str(self.source.resolve())},
            "report": {
                "key": "fixture-validation",
                "title": "Fixture validation summary",
                "type": "comparative_validation",
                "purpose": "Verify that structured analysis stays linked to the raw grid.",
                "status": "VERIFIED",
                "decision": "CAN_USE",
                "summary": "The test cohort has a lower NG rate.",
                "artifacts": {},
                "evidence": [{"sheet": "Fixture", "range": "B2:D3", "role": "SUMMARY"}],
                "conclusions": [
                    {
                        "key": "final",
                        "verdict": "CAN_USE",
                        "text": "Fixture conclusion is reusable.",
                        "evidence": [{"sheet": "Fixture", "range": "B2:C2", "role": "CONCLUSION"}],
                    }
                ],
            },
            "reviews": [
                {
                    "key": "function-ng",
                    "title": "Function NG rate",
                    "type": "defect_rate_comparison",
                    "status": "VERIFIED",
                    "decision": "IMPROVED",
                    "cohorts": [
                        {"key": "test", "role": "TEST", "label": "Changed condition"},
                        {"key": "control", "role": "CONTROL", "label": "Normal condition"},
                    ],
                    "metrics": [
                        {
                            "key": "total-ng-rate",
                            "label": "Total NG rate",
                            "type": "defect_rate",
                            "unit": "ppm",
                            "evidence": [{"sheet": "Fixture", "range": "B2:D3", "role": "METRIC"}],
                            "values": [
                                {"cohort": "test", "numerator": 1, "denominator": 100, "ratePpm": 10000},
                                {"cohort": "control", "numerator": 2, "denominator": 100, "ratePpm": 20000},
                            ],
                            "comparisons": [
                                {
                                    "key": "test-vs-control",
                                    "comparedCohort": "test",
                                    "controlCohort": "control",
                                    "deltaValue": -10000,
                                    "deltaUnit": "ppm",
                                    "relativeDeltaPercent": -50,
                                    "direction": "IMPROVED",
                                    "status": "IMPROVED",
                                    "summary": "Test NG rate is lower.",
                                    "evidence": [{"sheet": "Fixture", "range": "B2:D3", "role": "COMPARISON"}],
                                }
                            ],
                        }
                    ],
                    "conclusions": [
                        {
                            "key": "function-result",
                            "verdict": "IMPROVED",
                            "text": "The changed condition improves NG rate.",
                            "evidence": [{"sheet": "Fixture", "range": "B2:D3", "role": "CONCLUSION"}],
                        }
                    ],
                }
            ],
        }

    def test_schema_adds_resumable_ingestion_metadata(self) -> None:
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            run_columns = cli.table_columns(conn, "runs")
            self.assertIn("skipped", run_columns)
            self.assertIn("options_json", run_columns)
            self.assertTrue(cli.table_exists(conn, "ingest_items"))
            self.assertTrue(cli.table_exists(conn, "schema_migrations"))

    def test_import_preserves_fixed_coordinates_and_merge_metadata(self) -> None:
        self.write_payload()
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            imported = cli.import_com_json(
                conn,
                "FixtureDataset",
                self.raw_json,
                expected_source=self.source,
                expected_covered_cell_mode="blank",
                verify_after_import=True,
            )
            conn.commit()

            workbook_id = imported["workbookId"]
            self.assertEqual(6, conn.execute("SELECT COUNT(*) FROM grid_sheet_cells WHERE workbook_id=?", (workbook_id,)).fetchone()[0])
            self.assertEqual(2, conn.execute("SELECT COUNT(*) FROM grid_sheet_rows WHERE workbook_id=?", (workbook_id,)).fetchone()[0])
            self.assertEqual(1, conn.execute("SELECT COUNT(*) FROM merge_ranges WHERE workbook_id=?", (workbook_id,)).fetchone()[0])
            covered = conn.execute(
                "SELECT merge_role, merge_address, anchor_row, anchor_col, value_text FROM grid_sheet_cells WHERE workbook_id=? AND address='D2'",
                (workbook_id,),
            ).fetchone()
            self.assertEqual(("covered", "C2:D2", 2, 3, ""), tuple(covered))
            self.assertTrue(imported["verification"]["ok"])
            self.assertIsNotNone(cli.workbook_is_current(conn, "FixtureDataset", self.source))

    def test_invalid_reimport_keeps_the_previous_successful_workbook(self) -> None:
        self.write_payload()
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            first = cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            conn.commit()
            invalid = self.payload()
            invalid["totals"]["cellCount"] = 5
            self.write_payload(invalid)

            with self.assertRaises(ValueError):
                cli.import_com_json(conn, "FixtureDataset", self.raw_json)

            row = conn.execute("SELECT workbook_id, total_cells FROM workbooks WHERE dataset='FixtureDataset'").fetchone()
            self.assertEqual((first["workbookId"], 6), tuple(row))
            self.assertEqual(6, conn.execute("SELECT COUNT(*) FROM grid_sheet_cells WHERE workbook_id=?", (first["workbookId"],)).fetchone()[0])

    def test_changed_source_is_not_skipped(self) -> None:
        self.write_payload()
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            conn.commit()
            self.assertIsNotNone(cli.workbook_is_current(conn, "FixtureDataset", self.source))
            self.source.write_bytes(b"fixture-source-changed")
            self.assertIsNone(cli.workbook_is_current(conn, "FixtureDataset", self.source))

    def test_com_index_records_an_unchanged_workbook_as_skipped_without_starting_com(self) -> None:
        self.write_payload()
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            conn.commit()

        args = cli.build_parser().parse_args(
            [
                "com-index",
                "--input",
                str(self.source),
                "--dataset",
                "FixtureDataset",
                "--db",
                str(self.db),
                "--raw-dir",
                str(self.root / "raw"),
                "--limit",
                "1",
            ]
        )
        self.assertEqual(0, cli.cmd_com_index(args))

        with cli.connect_ro(self.db) as conn:
            run = conn.execute("SELECT succeeded, failed, skipped FROM runs ORDER BY run_id DESC LIMIT 1").fetchone()
            item = conn.execute("SELECT status FROM ingest_items ORDER BY ingest_item_id DESC LIMIT 1").fetchone()
        self.assertEqual((0, 0, 1), tuple(run))
        self.assertEqual("SKIPPED", item["status"])

    def test_analysis_import_stores_cohorts_metrics_comparisons_and_evidence(self) -> None:
        self.write_payload()
        manifest_path = self.root / "fixture.analysis.json"
        manifest_path.write_text(json.dumps(self.analysis_manifest(), ensure_ascii=False), encoding="utf-8")
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            imported = cli.import_analysis_manifest(conn, manifest_path, cli.read_analysis_manifest(manifest_path))
            conn.commit()

            self.assertTrue(imported["verification"]["ok"])
            self.assertEqual(1, imported["reviews"])
            self.assertEqual(2, imported["cohorts"])
            self.assertEqual(1, imported["metrics"])
            self.assertEqual(2, imported["metricValues"])
            self.assertEqual(1, imported["comparisons"])
            self.assertEqual(5, imported["evidence"])
            self.assertEqual(1, conn.execute("SELECT COUNT(*) FROM analysis_reports").fetchone()[0])
            self.assertEqual(1, conn.execute("SELECT COUNT(*) FROM analysis_comparisons").fetchone()[0])

            verification = cli.verify_analysis_report(conn, imported["analysisReportId"])
            self.assertTrue(verification["ok"])
            exported = cli.build_analysis_export(conn, imported["analysisReportId"])
            self.assertEqual("universal-analysis-export-v1", exported["schemaVersion"])
            self.assertEqual("fixture-validation", exported["report"]["key"])
            metric = exported["reviews"][0]["metrics"][0]
            self.assertEqual(["test", "control"], [value["cohort_key"] for value in metric["values"]])
            self.assertEqual(-10000, metric["comparisons"][0]["delta_value"])

    def test_analysis_import_rejects_evidence_outside_the_source_grid(self) -> None:
        self.write_payload()
        manifest = self.analysis_manifest()
        manifest["report"]["evidence"][0]["range"] = "A1:A1"
        manifest_path = self.root / "invalid.analysis.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            with self.assertRaisesRegex(ValueError, "outside the source UsedRange"):
                cli.import_analysis_manifest(conn, manifest_path, cli.read_analysis_manifest(manifest_path))
            self.assertEqual(0, conn.execute("SELECT COUNT(*) FROM analysis_reports").fetchone()[0])

    def test_quick_index_passes_a_single_excel_as_input_file(self) -> None:
        source = self.root / "single.xlsx"
        source.write_bytes(b"fixture")
        args = cli.argparse.Namespace(
            input=str(source), dataset="FixtureDataset", db=None, html=None, log=None,
            limit=0, force=False, no_html=False,
        )
        with mock.patch.object(cli, "run_command", return_value=0) as run:
            self.assertEqual(0, cli.cmd_quick_index(args))
        command = run.call_args.args[0]
        self.assertIn("--input-file", command)
        self.assertNotIn("--input-dir", command)


if __name__ == "__main__":
    unittest.main()
