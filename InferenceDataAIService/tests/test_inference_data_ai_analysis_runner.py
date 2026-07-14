from __future__ import annotations

import argparse
import importlib.util
import json
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path


RUNNER_PATH = Path(__file__).parents[1] / "inference_data_ai_analysis_runner.py"
SPEC = importlib.util.spec_from_file_location("inference_data_ai_analysis_runner", RUNNER_PATH)
assert SPEC is not None and SPEC.loader is not None
runner = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = runner
SPEC.loader.exec_module(runner)


class AnalysisHtmlRendererTests(unittest.TestCase):
    def brs_export(self) -> dict:
        return {
            "report": {
                "title": "BRS validation",
                "summary": "Curated BRS values.",
                "purpose": "Validation",
                "scope": "BRS source",
                "status": "VERIFIED",
                "decision": "CAN_USE",
                "sourcePath": "BRS.xlsx",
                "evidence": [{"sheet_name": "201506", "range_address": "D50:N78"}],
            },
            "reviews": [
                {
                    "title": "Tension 비교",
                    "summary": "Tension values",
                    "cohorts": [
                        {"cohort_key": "new", "label": "New Bond"},
                        {"cohort_key": "normal", "label": "Normal Bond"},
                    ],
                    "metrics": [
                        {
                            "label": "Long S-MG Tension",
                            "metric_type": "tension_measurement",
                            "unit": "N",
                            "spec_text": "45-65 N",
                            "evidence": [{"sheet_name": "201506", "range_address": "D50:E64"}],
                            "values": [
                                {"cohort_key": "new", "min_value": 50.13, "max_value": 61.45, "average_value": 54.99, "result_status": "OK", "details": {}},
                                {"cohort_key": "normal", "min_value": 39.38, "max_value": 66.34, "average_value": 50.72, "result_status": "OK", "details": {}},
                            ],
                            "comparisons": [
                                {
                                    "compared_cohort_key": "new",
                                    "control_cohort_key": "normal",
                                    "summary_text": "New Bond 평균이 높다.",
                                    "status": "TEST_MEAN_HIGHER",
                                    "evidence": [{"sheet_name": "201506", "range_address": "D50:E64"}],
                                }
                            ],
                        }
                    ],
                    "conclusions": [{"conclusion_text": "Tension 확인"}],
                },
                {
                    "title": "Function 불량률 비교",
                    "summary": "Function details",
                    "cohorts": [{"cohort_key": "new", "label": "New Bond"}, {"cohort_key": "normal", "label": "Normal Bond"}],
                    "metrics": [
                        {
                            "label": "Function Total NG Rate",
                            "metric_type": "defect_rate",
                            "unit": "ppm",
                            "evidence": [{"sheet_name": "201506", "range_address": "D76:N78"}],
                            "values": [
                                {
                                    "cohort_key": "new",
                                    "numerator": 89,
                                    "denominator": 2033,
                                    "rate_ppm": 43778,
                                    "result_status": "OK",
                                    "details": {"noise": 24, "touch": 65},
                                },
                                {"cohort_key": "normal", "numerator": 112, "denominator": 2507, "rate_ppm": 44675, "result_status": "REFERENCE", "details": {}},
                            ],
                            "comparisons": [
                                {
                                    "compared_cohort_key": "new",
                                    "control_cohort_key": "normal",
                                    "summary_text": "New Bond NG rate is lower.",
                                    "status": "IMPROVED",
                                    "evidence": [{"sheet_name": "201506", "range_address": "D76:N78"}],
                                }
                            ],
                        }
                    ],
                    "conclusions": [{"conclusion_text": "Function 확인"}],
                },
            ],
        }

    def test_brs_export_renders_measurement_and_function_breakdown(self) -> None:
        rendered = runner.analysis_html(self.brs_export())
        for token in ("50.13", "61.45", "54.99", "Noise", "24", "Touch", "65", "89 / 2,033 (43,778 ppm)", "규격: 45-65 N"):
            self.assertIn(token, rendered)
        self.assertIn("<html lang='ko'>", rendered)
        self.assertNotIn("<td></td>", rendered)
        self.assertEqual(rendered, rendered.encode("utf-8").decode("utf-8"))

    def test_renderer_accepts_camel_case_fields(self) -> None:
        data = self.brs_export()
        metric = data["reviews"][0]["metrics"][0]
        metric["values"] = [
            {
                "cohortKey": "new",
                "valueText": "Measured sample",
                "valueNumber": 3.5,
                "numerator": 1,
                "denominator": 40,
                "ratePpm": 25000,
                "minValue": 2.5,
                "maxValue": 4.5,
                "averageValue": 3.5,
                "resultStatus": "OK",
                "details": {"rejectCount": 2},
            },
            {"cohortKey": "normal", "valueText": "Reference", "resultStatus": "REFERENCE", "details": {}},
        ]
        metric["comparisons"] = [{"comparedCohort": "new", "controlCohort": "normal", "summary": "Camel fields render", "status": "VERIFIED", "evidence": []}]
        rendered = runner.analysis_html(data)
        for token in ("Measured sample", "1 / 40 (25,000 ppm)", "2.5", "4.5", "Reject count", "2", "Camel fields render"):
            self.assertIn(token, rendered)

    def test_small_raw_measurement_samples_survive_normalization_and_rendering(self) -> None:
        value = {
            "cohort": "new",
            "min": 1.0,
            "max": 3.0,
            "average": 2.0,
            "sampleCount": 3,
            "sampleSequence": [1.0, 2.0, 3.0],
            "status": "OK",
        }
        normalized = runner.normalize_manifest({"reviews": [{"metrics": [{"values": [value]}]}]})
        details = normalized["reviews"][0]["metrics"][0]["values"][0]["details"]
        self.assertEqual(3, details["sampleCount"])
        self.assertEqual([1.0, 2.0, 3.0], details["sampleSequence"])

        data = self.brs_export()
        metric = data["reviews"][0]["metrics"][0]
        metric["comparisons"] = []
        metric["values"] = [{"cohort_key": "new", "min_value": 1.0, "max_value": 3.0, "average_value": 2.0, "details": details, "result_status": "OK"}]
        rendered = runner.analysis_html(data)
        self.assertIn("Sample count", rendered)
        self.assertIn("Sample sequence", rendered)
        self.assertIn("[1.0,2.0,3.0]", rendered)


class CompleteRawMeasurementContractTests(unittest.TestCase):
    def packet(self) -> dict:
        return {
            "packetSelection": {"rowTruncated": False, "cellTruncated": False, "dataTruncated": False},
            "sheetRows": [
                {"sheet_index": 1, "sheet_name": "Sheet1", "row_number": 1, "cells": [
                    {"column": 4, "value": "Voltage"}, {"column": 5, "value": "Average"},
                    {"column": 6, "value": "Max"}, {"column": 7, "value": "Min"},
                    {"column": 8, "value": "Sample No"},
                ]},
                {"sheet_index": 1, "sheet_name": "Sheet1", "row_number": 2, "cells": [
                    {"column": 4, "value": "1600 V"}, {"column": 5, "value": 2},
                    {"column": 6, "value": 3}, {"column": 7, "value": 1},
                    {"column": 8, "value": 1}, {"column": 9, "value": 2}, {"column": 10, "value": 3},
                ]},
                {"sheet_index": 1, "sheet_name": "Sheet1", "row_number": 3, "cells": [
                    {"column": 4, "value": "1800 V"}, {"column": 5, "value": 17 / 3},
                    {"column": 6, "value": 7}, {"column": 7, "value": 5},
                    {"column": 8, "value": 5}, {"column": 9, "value": 5}, {"column": 10, "value": 7},
                ]},
            ],
        }

    def manifest(self) -> dict:
        values = []
        for cohort, label, samples, evidence_range in (("voltage-1600", "1600 V", [1, 2, 3], "H2:J2"), ("voltage-1800", "1800 V", [5, 5, 7], "H3:J3")):
            average = sum(samples) / len(samples)
            values.append(
                {
                    "cohort": cohort,
                    "average": average,
                    "min": min(samples),
                    "max": max(samples),
                    "details": {
                        "sampleCount": len(samples),
                        "sampleSequence": samples,
                        "sampleEvidenceRange": evidence_range,
                        "recomputedSummary": {
                            "average": average,
                            "min": min(samples),
                            "max": max(samples),
                            "sampleStandardDeviation": (sum((value - average) ** 2 for value in samples) / (len(samples) - 1)) ** 0.5,
                            "range": max(samples) - min(samples),
                        },
                        "displayedSummaryReconciliation": "MATCH",
                    },
                }
            )
        return {
            "reviews": [{
                "cohorts": [{"key": "voltage-1600", "label": "1600 V"}, {"key": "voltage-1800", "label": "1800 V"}],
                "metrics": [{"values": values}],
            }]
        }

    def test_complete_raw_measurement_requires_and_accepts_reconciled_sequences(self) -> None:
        packet = self.packet()
        expected = runner.packet_complete_raw_measurements(packet)
        self.assertEqual(["1600 V", "1800 V"], [item["label"] for item in expected])
        self.assertEqual([3, 3], [len(item["samples"]) for item in expected])
        runner.validate_complete_raw_measurement_details(packet, self.manifest())

        incomplete = self.manifest()
        del incomplete["reviews"][0]["metrics"][0]["values"][0]["details"]["sampleSequence"]
        with self.assertRaisesRegex(ValueError, "sampleSequence"):
            runner.validate_complete_raw_measurement_details(packet, incomplete)

    def test_deterministic_manifest_keeps_all_complete_cohorts_and_explicit_control_comparison(self) -> None:
        packet = self.packet()
        packet["sheetRows"][1]["cells"][0]["value"] = "Normal"
        packet["sheetRows"][2]["cells"][0]["value"] = "Test 1800"

        manifest = runner.canonical_manifest_from_packet(packet)
        review = manifest["reviews"][0]
        metric = review["metrics"][0]

        self.assertEqual("NEEDS_REVIEW", manifest["report"]["status"])
        self.assertEqual("OBSERVED_ONLY", manifest["report"]["decision"])
        self.assertEqual(["normal", "test-1800"], [cohort["key"] for cohort in review["cohorts"]])
        self.assertEqual("CONTROL", review["cohorts"][0]["role"])
        self.assertEqual(2, len(metric["values"]))
        self.assertEqual([1.0, 2.0, 3.0], metric["values"][0]["details"]["sampleSequence"])
        self.assertEqual(3, metric["values"][0]["details"]["sampleCount"])
        self.assertEqual("H2:J2", metric["values"][0]["details"]["sampleEvidenceRange"])
        self.assertEqual("MATCH", metric["values"][0]["details"]["displayedSummaryReconciliation"])
        self.assertAlmostEqual(1.0, metric["values"][0]["details"]["recomputedSummary"]["sampleStandardDeviation"])
        self.assertEqual(1, len(metric["comparisons"]))
        self.assertEqual("normal", metric["comparisons"][0]["controlCohort"])
        self.assertEqual("test-1800", metric["comparisons"][0]["comparedCohort"])
        self.assertEqual("OBSERVED_ONLY", metric["comparisons"][0]["status"])
        runner.validate_complete_raw_measurement_details(packet, manifest)

    def test_deterministic_fallback_is_packet_evidence_backed_and_needs_review(self) -> None:
        packet = {
            "workbook": {"file_name": "fallback.xlsx"},
            "packetSelection": {"rowTruncated": False, "cellTruncated": False, "dataTruncated": True},
            "sheets": [{"sheet_index": 1, "sheet_name": "Fallback", "used_top": 1, "used_left": 1, "used_bottom": 2, "used_right": 2}],
            "sheetRows": [{"sheet_index": 1, "sheet_name": "Fallback", "row_number": 1, "cells": [{"column": 1, "address": "A1", "value": "Incomplete"}]}],
        }

        manifest = runner.canonical_manifest_from_packet(packet)
        review = manifest["reviews"][0]
        metric = review["metrics"][0]

        self.assertEqual("packet-canonical-needs-review", manifest["report"]["key"])
        self.assertEqual("NEEDS_REVIEW", manifest["report"]["status"])
        self.assertEqual("OBSERVED_ONLY", manifest["report"]["decision"])
        self.assertEqual([{"sheet": "Fallback", "range": "A1", "role": "PACKET"}], manifest["report"]["evidence"])
        self.assertEqual("OBSERVED_ONLY", metric["values"][0]["status"])
        self.assertEqual([], metric["comparisons"])


class CuratedReuseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.service = Path(self.temp.name)
        self.original = self.service / "original.xlsx"
        self.requested = self.service / "isolated-copy.xlsx"
        self.original.write_bytes(b"byte-identical curated workbook")
        self.requested.write_bytes(self.original.read_bytes())
        (self.service / "outputs" / "analysis-manifests").mkdir(parents=True)
        (self.service / "outputs" / "universal-grid").mkdir(parents=True)
        (self.service / "outputs").joinpath("curated.html").write_text("<html><body><h1>CLI baseline</h1></body></html>", encoding="utf-8")
        self.manifest_path = self.service / "outputs" / "analysis-manifests" / "fixture_analysis.json"
        self.manifest_path.write_text(
            json.dumps(
                {
                    "schemaVersion": "universal-analysis-v1",
                    "source": {"dataset": "Fixture", "sourcePath": str(self.original.resolve())},
                    "report": {"key": "fixture", "scope": "Original scope", "artifacts": {"html": "outputs/curated.html"}},
                }
            ),
            encoding="utf-8",
        )
        db = self.service / "outputs" / "universal-grid" / "fixture.sqlite"
        fingerprint = runner.core.file_fingerprint(self.original)
        conn = sqlite3.connect(db)
        try:
            conn.executescript(
                """
                CREATE TABLE workbooks (workbook_id INTEGER PRIMARY KEY, status TEXT, fingerprint TEXT);
                CREATE TABLE analysis_reports (
                    analysis_report_id INTEGER PRIMARY KEY, dataset TEXT, source_path TEXT, overall_status TEXT,
                    manifest_path TEXT, workbook_id INTEGER, workbook_fingerprint TEXT
                );
                """
            )
            conn.execute("INSERT INTO workbooks VALUES (1, 'OK', ?)", (fingerprint,))
            conn.execute(
                "INSERT INTO analysis_reports VALUES (1, 'Fixture', ?, 'VERIFIED', ?, 1, ?)",
                (str(self.original.resolve()), str(self.manifest_path.resolve()), fingerprint),
            )
            conn.commit()
        finally:
            conn.close()

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_byte_identical_verified_baseline_is_rebound_with_provenance(self) -> None:
        reuse = runner.curated_reuse_for_source(self.service, str(self.requested), "Fixture")
        self.assertIsNotNone(reuse)
        assert reuse is not None
        reused_html = runner.write_reused_curated_html(self.service, reuse, self.requested.resolve(), 17)
        clone_path, provenance = runner.rebind_curated_manifest(
            self.service,
            reuse,
            self.requested.resolve(),
            "Fixture",
            {"workbook_id": 17, "fingerprint": "current-workbook-fingerprint"},
            reused_html,
        )
        clone = json.loads(clone_path.read_text(encoding="utf-8"))
        self.assertEqual(str(self.requested.resolve()), clone["source"]["sourcePath"])
        self.assertEqual(17, clone["source"]["workbookId"])
        self.assertEqual("current-workbook-fingerprint", clone["source"]["fingerprint"])
        self.assertEqual(str(self.original.resolve()), provenance["originalSource"])
        self.assertIn(provenance["sha256"], clone["report"]["scope"])
        self.assertIn("큐레이션 기준 재사용", reused_html.read_text(encoding="utf-8"))
        self.assertNotIn("큐레이션 기준 재사용", (self.service / "outputs" / "curated.html").read_text(encoding="utf-8"))

    def test_changed_copy_logs_non_applicable_and_cannot_reuse_the_baseline(self) -> None:
        self.requested.write_bytes(b"different workbook")
        self.assertIsNone(runner.curated_reuse_for_source(self.service, str(self.requested), "Fixture"))
        outcome = runner.curated_reuse_not_applicable(str(self.requested))
        self.assertEqual("curated-reuse-not-applicable", outcome["status"])
        self.assertIn("SHA-256", outcome["reason"])


class ForceAiDraftTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=runner.core.SERVICE_DIR)
        self.service = Path(self.temp.name)
        self.dataset = "ForceFixture"
        self.source = self.service / "force-fixture.xlsx"
        self.source.write_bytes(b"force fixture")
        self.raw_json = self.service / "force-fixture.com-grid.json"
        stat = self.source.stat()
        self.raw_json.write_text(
            json.dumps(
                {
                    "schemaVersion": "input-data-com-grid-v1",
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
                            "usedRange": {"top": 1, "left": 1, "bottom": 1, "right": 1, "rowCount": 1, "columnCount": 1, "address": "A1"},
                            "nonEmptyCells": 1,
                            "mergeCount": 0,
                            "merges": [],
                            "rows": [{"rowNumber": 1, "nonEmptyCount": 1, "cells": [{"row": 1, "column": 1, "colLabel": "A", "address": "A1", "value": "evidence", "rawValue": "evidence", "merge": {"role": "none"}}]}],
                        }
                    ],
                    "totals": {"sheetCount": 1, "rowCount": 1, "cellCount": 1, "nonEmptyCells": 1, "mergeCount": 0},
                }
            ),
            encoding="utf-8",
        )
        self.db = self.service / "force-fixture.sqlite"
        with runner.core.connect_rw(self.db) as conn:
            runner.core.ensure_universal_schema(conn)
            runner.core.import_com_json(conn, self.dataset, self.raw_json, expected_source=self.source, verify_after_import=True)
            conn.commit()
        self.curated_html = self.service / "outputs" / "curated-cli.html"
        self.curated_html.parent.mkdir(parents=True, exist_ok=True)
        self.curated_html.write_text("<html><body>curated CLI artifact</body></html>", encoding="utf-8")
        self.curated_manifest = self.service / "outputs" / "analysis-manifests" / "curated_fixture_analysis.json"
        self.curated_manifest.parent.mkdir(parents=True, exist_ok=True)
        self.curated_manifest.write_text(json.dumps(self.manifest("curated-analysis", {"html": str(self.curated_html)}), ensure_ascii=False), encoding="utf-8")
        with runner.core.connect_rw(self.db) as conn:
            curated = runner.core.import_analysis_manifest(conn, self.curated_manifest, runner.core.read_analysis_manifest(self.curated_manifest), self.dataset)
            conn.commit()
        self.curated_report_id = int(curated["analysisReportId"])

    def tearDown(self) -> None:
        self.temp.cleanup()

    def manifest(self, key: str, artifacts: dict | None = None) -> dict:
        fingerprint = runner.core.file_fingerprint(self.source)
        evidence = [{"sheet": "Fixture", "range": "A1", "role": "SOURCE"}]
        return {
            "schemaVersion": "universal-analysis-v1",
            "source": {"dataset": self.dataset, "sourcePath": str(self.source.resolve()), "fingerprint": fingerprint},
            "report": {
                "key": key,
                "title": "Force fixture",
                "type": "validation",
                "purpose": "Test force isolation.",
                "scope": "Fixture scope",
                "status": "VERIFIED",
                "decision": "CAN_USE",
                "summary": "Fixture summary.",
                "artifacts": artifacts or {},
                "evidence": evidence,
                "conclusions": [{"key": "report", "verdict": "CAN_USE", "text": "Fixture conclusion.", "evidence": evidence}],
            },
            "reviews": [
                {
                    "key": "review",
                    "title": "Fixture review",
                    "type": "validation",
                    "status": "VERIFIED",
                    "decision": "CAN_USE",
                    "cohorts": [{"key": "test", "role": "TEST", "label": "Test"}, {"key": "control", "role": "CONTROL", "label": "Control"}],
                    "metrics": [
                        {
                            "key": "metric",
                            "label": "Fixture metric",
                            "type": "count",
                            "evidence": evidence,
                            "values": [{"cohort": "test", "valueText": "1", "status": "OK"}, {"cohort": "control", "valueText": "1", "status": "OK"}],
                            "comparisons": [],
                        }
                    ],
                    "conclusions": [{"key": "review", "verdict": "CAN_USE", "text": "Fixture review conclusion.", "evidence": evidence}],
                }
            ],
        }

    def test_force_mode_rejects_curated_reuse_even_for_direct_runner_calls(self) -> None:
        args = argparse.Namespace(force_ai_draft=True, reuse_curated=True)
        with self.assertRaisesRegex(SystemExit, "cannot be combined"):
            runner.run(args)

    def test_force_draft_keeps_stale_curated_row_and_artifacts_while_adding_a_unique_report(self) -> None:
        original_manifest_bytes = self.curated_manifest.read_bytes()
        original_html_bytes = self.curated_html.read_bytes()
        with runner.core.connect_rw(self.db) as conn:
            runner.core.import_com_json(conn, self.dataset, self.raw_json, expected_source=self.source, verify_after_import=True)
            curated_before = runner.core.first_dict(conn, "SELECT * FROM analysis_reports WHERE analysis_report_id=?", (self.curated_report_id,))
            conn.commit()
        assert curated_before is not None
        self.assertEqual("STALE", curated_before["overall_status"])

        args = argparse.Namespace(
            service_dir=str(self.service), db=str(self.db), source=str(self.source), dataset=self.dataset,
            row_limit=10, cell_limit=10, replace_auto_draft=False, reuse_curated=False, force_ai_draft=True,
        )
        self.assertFalse(hasattr(runner, "run_codex_command"))
        self.assertEqual(0, runner.run(args))

        with runner.core.connect_ro(self.db) as conn:
            reports = runner.core.dict_rows(conn, "SELECT analysis_report_id, overall_status, analysis_key, manifest_path, dashboard_html_path FROM analysis_reports ORDER BY analysis_report_id")
        self.assertEqual(2, len(reports))
        original = next(report for report in reports if report["analysis_report_id"] == self.curated_report_id)
        fresh = next(report for report in reports if report["analysis_report_id"] != self.curated_report_id)
        self.assertEqual("STALE", original["overall_status"])
        self.assertEqual(str(self.curated_manifest.resolve()), original["manifest_path"])
        self.assertEqual(str(self.curated_html), original["dashboard_html_path"])
        self.assertTrue(fresh["analysis_key"].startswith("packet-canonical-needs-review-force-ai-"))
        self.assertIn("_force_ai_draft_", Path(fresh["manifest_path"]).name)
        self.assertNotEqual(str(self.curated_html), fresh["dashboard_html_path"])
        self.assertEqual(original_manifest_bytes, self.curated_manifest.read_bytes())
        self.assertEqual(original_html_bytes, self.curated_html.read_bytes())


if __name__ == "__main__":
    unittest.main()
