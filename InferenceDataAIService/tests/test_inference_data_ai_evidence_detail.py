from __future__ import annotations

import importlib.util
import json
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SERVICE_DIR = Path(__file__).parents[1]
NOW = "2026-07-17T12:00:00Z"
CONTENT_SHA256 = "a" * 64


def load_module(name: str, file_name: str):
    specification = importlib.util.spec_from_file_location(
        name, SERVICE_DIR / file_name
    )
    assert specification is not None and specification.loader is not None
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


schema = load_module("evidence_detail_schema", "inference_data_ai_schema.py")
capture = load_module(
    "evidence_detail_capture", "inference_data_ai_source_ingest.py"
)
detail = load_module(
    "inference_data_ai_evidence_detail",
    "inference_data_ai_evidence_detail.py",
)


class EvidenceDetailTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.database = self.root / "evidence.sqlite"
        self.source_path = str((self.root / "source.xlsx").resolve())
        self.connection = sqlite3.connect(self.database)
        self.connection.execute(
            """
            CREATE TABLE schema_migrations(
                migration_name TEXT PRIMARY KEY,
                applied_at TEXT NOT NULL
            )
            """
        )
        self._create_legacy_parent_stubs()
        schema.ensure_knowledge_schema(self.connection, lambda: NOW)
        capture.ensure_capture_v2_schema(self.connection)
        self._insert_fixture()
        self.connection.commit()

    def tearDown(self) -> None:
        self.connection.close()
        self.temp.cleanup()

    def _create_legacy_parent_stubs(self) -> None:
        for table, primary_key in [
            ("workbooks", "workbook_id"),
            ("analysis_reports", "analysis_report_id"),
            ("analysis_review_items", "review_item_id"),
            ("analysis_cohorts", "cohort_id"),
            ("analysis_metrics", "metric_id"),
            ("analysis_metric_values", "metric_value_id"),
            ("analysis_comparisons", "comparison_id"),
            ("analysis_evidence", "evidence_id"),
            ("analysis_conclusions", "conclusion_id"),
        ]:
            self.connection.execute(
                f"CREATE TABLE {table}({primary_key} INTEGER PRIMARY KEY)"
            )

    def _insert_fixture(self) -> None:
        self.connection.execute(
            """
            INSERT INTO capture_v2_documents(
                document_id, source_path, file_name, source_kind,
                created_at, updated_at
            ) VALUES (101, ?, 'source.xlsx', 'XLSX', ?, ?)
            """,
            (self.source_path, NOW, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO capture_v2_revisions(
                revision_id, revision_uid, document_id, content_sha256,
                capture_contract, extractor_name, extractor_version,
                size_bytes, mtime_ns, capture_status, is_current,
                capture_json_sha256, captured_at
            ) VALUES (
                201, 'capture_revision_fixture', 101, ?, 'openxml-capture-v2',
                'openpyxl', '3.1', 1234, 5678, 'CAPTURED', 1, ?, ?
            )
            """,
            (CONTENT_SHA256, "b" * 64, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO capture_v2_workbooks(
                revision_id, workbook_status, is_truly_empty, sheet_count,
                nonempty_sheet_count, tabular_sheet_count, metadata_json
            ) VALUES (201, 'CAPTURED', 0, 1, 1, 1, '{}')
            """
        )
        self.connection.execute(
            """
            INSERT INTO capture_v2_sheets(
                sheet_id, revision_id, sheet_index, title, sheet_state,
                capture_status, is_truly_empty, has_tabular_evidence,
                nonempty_cell_count, structural_cell_count, captured_cell_count,
                formula_cell_count, merge_count, used_bounds_json,
                content_bounds_json, freeze_panes, auto_filter, metadata_json
            ) VALUES (
                301, 201, 0, 'Results', 'visible', 'CAPTURED', 0, 1,
                3, 1, 4, 1, 1,
                '{"minRow":1,"minColumn":1,"maxRow":2,"maxColumn":3}',
                '{"minRow":1,"minColumn":1,"maxRow":2,"maxColumn":3}',
                'B2', 'A1:C2', '{"orientation":"landscape"}'
            )
            """
        )
        cells = [
            (
                301,
                1,
                1,
                "A1",
                json.dumps("Merged header"),
                None,
                None,
                json.dumps("Merged header"),
                "s",
                None,
                "General",
                7,
                '{"alignment":{"horizontal":"center"}}',
                "A1:B1",
                "anchor",
            ),
            (
                301,
                1,
                2,
                "B1",
                None,
                None,
                None,
                None,
                "n",
                None,
                "General",
                7,
                '{"alignment":{"horizontal":"center"}}',
                "A1:B1",
                "member",
            ),
            (
                301,
                2,
                2,
                "B2",
                json.dumps("=1/4"),
                "=1/4",
                json.dumps(0.25),
                json.dumps("25.00%"),
                "f",
                "n",
                "0.00%",
                8,
                '{"fill":{"type":"solid","fgColor":"FFFF00"}}',
                None,
                "none",
            ),
            (
                301,
                2,
                3,
                "C2",
                json.dumps(25),
                None,
                None,
                json.dumps(25),
                "n",
                None,
                "0",
                0,
                "{}",
                None,
                "none",
            ),
        ]
        self.connection.executemany(
            """
            INSERT INTO capture_v2_cells(
                sheet_id, row_index, column_index, coordinate,
                raw_value_json, formula_text, cached_value_json,
                display_value_json, data_type, cached_data_type,
                number_format, style_id, style_json, merge_range, merge_role
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            cells,
        )
        self.connection.execute(
            """
            INSERT INTO capture_v2_merged_ranges(
                sheet_id, address, min_row, min_column, max_row, max_column,
                anchor_coordinate
            ) VALUES (301, 'A1:B1', 1, 1, 1, 2, 'A1')
            """
        )
        self.connection.executemany(
            """
            INSERT INTO capture_v2_row_dimensions(
                sheet_id, row_index, height, hidden
            ) VALUES (?,?,?,?)
            """,
            [(301, 1, 22.0, 0), (301, 2, 15.0, 1)],
        )
        self.connection.executemany(
            """
            INSERT INTO capture_v2_column_dimensions(
                sheet_id, dimension_key, min_column, max_column, width, hidden
            ) VALUES (?,?,?,?,?,?)
            """,
            [
                (301, "A", 1, 1, 20.0, 1),
                (301, "B", 2, 2, 12.5, 1),
                (301, "C", 3, 3, 8.0, 0),
            ],
        )

        self.connection.execute(
            """
            INSERT INTO source_documents(
                document_id, document_uid, dataset, source_path,
                original_file_name, source_kind, lifecycle_status,
                created_at, updated_at
            ) VALUES (
                401, 'document_fixture', 'fixture', ?, 'source.xlsx',
                'XLSX', 'ACTIVE', ?, ?
            )
            """,
            (self.source_path, NOW, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO source_revisions(
                revision_id, revision_uid, document_id, source_fingerprint,
                fingerprint_kind, content_sha256, size_bytes, mtime_ns,
                extractor_name, extractor_version, capture_contract,
                capture_status, is_current, captured_at,
                capture_v2_revision_id
            ) VALUES (
                501, 'capture_revision_fixture', 401, ?, 'SHA256', ?,
                1234, 5678, 'openpyxl', '3.1', 'openxml-capture-v2',
                'CAPTURED', 1, ?, 201
            )
            """,
            (CONTENT_SHA256, CONTENT_SHA256, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO workbook_analyses(
                workbook_analysis_id, analysis_uid, public_analysis_id,
                document_id, revision_id, analysis_key, title,
                analysis_status, verification_status, created_at, updated_at
            ) VALUES (
                601, 'analysis_fixture', 'ANALYSIS-FIXTURE', 401, 501,
                'fixture', 'Fixture analysis', 'COMPLETE', 'VERIFIED', ?, ?
            )
            """,
            (NOW, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO knowledge_studies(
                study_id, study_uid, public_data_id, workbook_analysis_id,
                study_key, title, analysis_status, verification_status,
                comparability_status, confounding_status, created_at, updated_at
            ) VALUES (
                701, 'study_fixture', 'DATA-FIXTURE', 601, 'fixture',
                'Fixture study', 'COMPLETE', 'VERIFIED', 'VALID', 'NONE', ?, ?
            )
            """,
            (NOW, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO evidence_items(
                evidence_id, evidence_uid, public_evidence_id, revision_id,
                evidence_kind, sheet_name, start_row, start_col, end_row,
                end_col, range_address, evidence_role, source_text, note,
                content_sha256, verification_status, created_at
            ) VALUES (
                801, 'evidence_fixture', 'EVD-FIXTURE', 501, 'CELL_RANGE',
                'Results', 1, 2, 2, 3, 'B1:C2', 'SOURCE',
                'Exact formula result table', 'fixture note', ?,
                'VERIFIED', ?
            )
            """,
            (CONTENT_SHA256, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id, evidence_role,
                claim_scope
            ) VALUES (
                'STUDY', 'study_fixture', 801, 'SOURCE',
                'observation table'
            )
            """
        )

    def test_returns_exact_current_capture_cells_merges_and_hidden_dimensions(
        self,
    ) -> None:
        result = detail.build_evidence_detail(
            self.connection, "evd-fixture"
        )

        self.assertEqual(
            "canonical-evidence-detail-v1", result["schemaVersion"]
        )
        self.assertEqual("EVD-FIXTURE", result["publicEvidenceId"])
        self.assertTrue(result["trust"]["trusted"])
        self.assertEqual(
            "CURRENT_CAPTURE_VERIFIED", result["trust"]["status"]
        )
        self.assertFalse(result["trust"]["exactRevisionFallbackUsed"])
        self.assertEqual(501, result["revision"]["canonicalRevisionId"])
        self.assertEqual(201, result["revision"]["captureV2"]["revisionId"])
        self.assertEqual("B1:C2", result["evidence"]["range"])
        self.assertEqual("B1:C2", result["preview"]["range"]["a1"])
        self.assertEqual(
            ["B1", "B2", "C2"],
            [cell["coordinate"] for cell in result["preview"]["cells"]],
        )

        formula = next(
            cell
            for cell in result["preview"]["cells"]
            if cell["coordinate"] == "B2"
        )
        self.assertEqual("=1/4", formula["rawValue"])
        self.assertEqual("=1/4", formula["formula"])
        self.assertEqual(0.25, formula["cachedValue"])
        self.assertEqual("25.00%", formula["displayValue"])
        self.assertEqual("0.00%", formula["numberFormat"])
        self.assertEqual("f", formula["dataType"])
        self.assertEqual("n", formula["cachedDataType"])
        self.assertTrue(formula["rowHidden"])
        self.assertTrue(formula["columnHidden"])

        merge = result["preview"]["mergedRanges"][0]
        self.assertEqual("A1:B1", merge["address"])
        self.assertTrue(merge["anchorOutsideEvidenceRange"])
        self.assertEqual("A1", merge["anchorCell"]["coordinate"])
        self.assertEqual("Merged header", merge["anchorCell"]["rawValue"])
        self.assertTrue(merge["anchorCell"]["columnHidden"])
        self.assertEqual(
            [{"entityType": "STUDY", "entityUid": "study_fixture",
              "publicId": "DATA-FIXTURE", "label": "Fixture study",
              "verificationStatus": "VERIFIED", "exists": True,
              "evidenceRole": "SOURCE", "claimScope": "observation table"}],
            result["linkedEntities"],
        )
        self.assertFalse(result["scope"]["imagesAnalyzed"])
        self.assertFalse(result["preview"]["imagesAnalyzed"])

    def test_rejects_invalid_missing_and_case_ambiguous_ids(self) -> None:
        with self.assertRaises(detail.InvalidEvidenceIdError):
            detail.build_evidence_detail(self.connection, "not-an-evd")
        with self.assertRaises(detail.EvidenceNotFoundError):
            detail.build_evidence_detail(self.connection, "EVD-MISSING")

        self.connection.execute(
            """
            INSERT INTO evidence_items(
                evidence_uid, public_evidence_id, revision_id, evidence_kind,
                sheet_name, start_row, start_col, end_row, end_col,
                range_address, created_at
            ) VALUES
                ('evidence_case_one', 'EVD-CASE', 501, 'CELL', 'Results',
                 2, 2, 2, 2, 'B2', ?),
                ('evidence_case_two', 'evd-case', 501, 'CELL', 'Results',
                 2, 3, 2, 3, 'C2', ?)
            """,
            (NOW, NOW),
        )
        with self.assertRaises(detail.AmbiguousEvidenceIdError):
            detail.build_evidence_detail(self.connection, "EVD-CASE")

    def test_rejects_stale_canonical_revision_without_falling_forward(self) -> None:
        self.connection.execute(
            "UPDATE source_revisions SET is_current=0 WHERE revision_id=501"
        )
        self.connection.execute(
            """
            INSERT INTO source_revisions(
                revision_uid, document_id, source_fingerprint,
                fingerprint_kind, content_sha256, capture_contract,
                capture_status, is_current, captured_at
            ) VALUES (
                'new_current_revision', 401, 'new-fingerprint', 'SHA256',
                ?, 'openxml-capture-v2', 'CAPTURED', 1, ?
            )
            """,
            ("c" * 64, NOW),
        )

        with self.assertRaisesRegex(
            detail.EvidenceTrustError, "stale canonical revision"
        ):
            detail.build_evidence_detail(self.connection, "EVD-FIXTURE")

    def test_rejects_stale_capture_bridge_without_falling_forward(self) -> None:
        self.connection.execute(
            "UPDATE capture_v2_revisions SET is_current=0, capture_status='STALE' "
            "WHERE revision_id=201"
        )
        self.connection.execute(
            """
            INSERT INTO capture_v2_revisions(
                revision_id, revision_uid, document_id, content_sha256,
                capture_contract, extractor_name, extractor_version,
                size_bytes, mtime_ns, capture_status, is_current,
                capture_json_sha256, captured_at
            ) VALUES (
                202, 'capture_revision_new', 101, ?, 'openxml-capture-v2',
                'openpyxl', '3.1', 1235, 5679, 'CAPTURED', 1, ?, ?
            )
            """,
            ("c" * 64, "d" * 64, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO capture_v2_workbooks(
                revision_id, workbook_status, is_truly_empty, sheet_count,
                nonempty_sheet_count, tabular_sheet_count, metadata_json
            ) VALUES (202, 'CAPTURED', 0, 1, 1, 1, '{}')
            """
        )

        with self.assertRaisesRegex(
            detail.EvidenceTrustError, "stale Capture v2 revision"
        ):
            detail.build_evidence_detail(self.connection, "EVD-FIXTURE")

    def test_rejects_hash_mismatch_and_does_not_use_other_revision(self) -> None:
        self.connection.execute(
            """
            UPDATE source_revisions
            SET content_sha256=?
            WHERE revision_id=501
            """,
            ("e" * 64,),
        )
        with self.assertRaisesRegex(
            detail.EvidenceTrustError, "contentSha256"
        ):
            detail.build_evidence_detail(self.connection, "EVD-FIXTURE")

    def test_rejects_stored_a1_coordinate_mismatch(self) -> None:
        self.connection.execute(
            """
            UPDATE evidence_items
            SET range_address='A1'
            WHERE public_evidence_id='EVD-FIXTURE'
            """
        )
        with self.assertRaisesRegex(
            detail.EvidenceTrustError, "stored A1 range"
        ):
            detail.build_evidence_detail(self.connection, "EVD-FIXTURE")

    def test_reports_review_state_but_rejects_stale_evidence(self) -> None:
        self.connection.execute(
            """
            UPDATE evidence_items
            SET verification_status='NEEDS_REVIEW'
            WHERE public_evidence_id='EVD-FIXTURE'
            """
        )
        review_detail = detail.build_evidence_detail(
            self.connection, "EVD-FIXTURE"
        )
        self.assertFalse(review_detail["trust"]["trusted"])
        self.assertEqual(
            "CURRENT_CAPTURE_UNVERIFIED", review_detail["trust"]["status"]
        )

        self.connection.execute(
            """
            UPDATE evidence_items
            SET verification_status='STALE'
            WHERE public_evidence_id='EVD-FIXTURE'
            """
        )
        with self.assertRaisesRegex(detail.EvidenceTrustError, "marked STALE"):
            detail.build_evidence_detail(self.connection, "EVD-FIXTURE")

    def test_database_wrapper_and_cli_return_schema_versioned_json(self) -> None:
        self.connection.commit()
        wrapped = detail.build_evidence_detail_from_db(
            self.database, "EVD-FIXTURE"
        )
        self.assertEqual(
            "canonical-evidence-detail-v1", wrapped["schemaVersion"]
        )

        completed = subprocess.run(
            [
                sys.executable,
                str(SERVICE_DIR / "inference_data_ai_cli.py"),
                "evidence-detail",
                "--db",
                str(self.database),
                "--evidence-id",
                "EVD-FIXTURE",
            ],
            cwd=SERVICE_DIR,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        self.assertEqual(0, completed.returncode, completed.stderr)
        cli_result = json.loads(completed.stdout)
        self.assertEqual(
            "canonical-evidence-detail-v1", cli_result["schemaVersion"]
        )
        self.assertEqual("EVD-FIXTURE", cli_result["publicEvidenceId"])


if __name__ == "__main__":
    unittest.main()
