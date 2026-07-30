from __future__ import annotations

import json
import sqlite3
import tempfile
import threading
import time
import unittest
from collections import Counter
from pathlib import Path
from unittest import mock

import inference_data_ai_cli as cli
import inference_data_ai_corpus_workflow as corpus


class CorpusWorkflowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=cli.SERVICE_DIR)
        self.root = Path(self.temp.name)
        self.sources = self.root / "sources"
        self.sources.mkdir()
        self.database = self.root / "knowledge.sqlite"
        self.artifacts = self.root / "artifacts"
        cli.init_universal_db(self.database, "Fixture")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def source(self, relative: str, payload: bytes | None = None) -> Path:
        path = self.sources / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload or ("xlsx:" + relative).encode("utf-8"))
        return path

    def test_parallel_ingest_prepares_database_for_read_write_overlap(
        self,
    ) -> None:
        corpus._prepare_database_for_parallel_ingest(self.database)

        connection = sqlite3.connect(self.database)
        try:
            mode = connection.execute("PRAGMA journal_mode").fetchone()[0]
        finally:
            connection.close()

        self.assertEqual("wal", str(mode).lower())

    def test_windows_system_error_marks_stale_lock_owner(self) -> None:
        lock_path = self.root / "corpus-journal.json.lock"
        lock_path.write_text("424242", encoding="ascii")
        with mock.patch.object(
            corpus.os,
            "kill",
            side_effect=SystemError("invalid Windows PID"),
        ):
            with corpus._batch_lock(lock_path):
                self.assertTrue(lock_path.is_file())
        self.assertFalse(lock_path.exists())

    def test_legacy_drm_release_suffix_is_removed_from_filename_identity(
        self,
    ) -> None:
        self.assertEqual(
            (
                "00. report test new dry machine with material make "
                "press jig 2024.03.26.xlsx"
            ),
            corpus._normalized_workbook_filename(
                "00. Report TEST new dry machine with material make "
                "press JIG 2024.03.26_1778470442_clean.xlsx"
            ),
        )
        self.assertEqual(
            "report_date_250119.xlsx",
            corpus._normalized_workbook_filename(
                "Report_Date_250119_1778470541_clean.xlsx"
            ),
        )
        self.assertEqual(
            "report.xlsx",
            corpus._normalized_workbook_filename("Report_clean.xlsx"),
        )
        self.assertEqual(
            "report.xlsx",
            corpus._normalized_workbook_filename(
                "Report_1778463079_1778470471_clean.xlsx"
            ),
        )
        self.assertEqual(
            "report.xlsx",
            corpus._normalized_workbook_filename(
                "Report_1778470471.xlsx"
            ),
        )

    def test_existing_clean_copy_is_reconciled_by_original_filename(
        self,
    ) -> None:
        clean_path = (
            r"D:\archive\00. Report TEST new dry machine with material "
            r"make press JIG 2024.03.26_1778470442_clean.xlsx"
        )
        original_path = (
            r"D:\incoming\00. Report TEST new dry machine with material "
            r"make press JIG 2024.03.26.xlsx"
        )
        record = {
            "recordId": "incoming-record",
            "sourcePath": original_path,
            "contentSha256": "b" * 64,
            "status": "PENDING",
            "result": None,
            "error": "",
        }
        rows = [
            (
                clean_path,
                Path(clean_path).name,
                "a" * 64,
                "capture-revision-clean",
                "ANALYSIS-CLEAN",
                "COMPLETED",
                "NEEDS_REVIEW",
                2,
            )
        ]
        connection = mock.MagicMock()
        table_cursor = mock.MagicMock()
        table_cursor.__iter__.return_value = iter(
            [
                ("source_documents",),
                ("source_revisions",),
                ("workbook_analyses",),
                ("knowledge_studies",),
            ]
        )
        analysis_cursor = mock.MagicMock()
        analysis_cursor.fetchall.return_value = rows
        connection.execute.side_effect = [
            table_cursor,
            analysis_cursor,
        ]
        with mock.patch.object(
            corpus.sqlite3,
            "connect",
            return_value=connection,
        ):
            reconciled = corpus._reconcile_existing_analyses(
                {"records": [record]},
                database_path=self.database,
                now="2026-07-24T00:00:00+00:00",
            )

        self.assertEqual({"incoming-record"}, reconciled)
        self.assertEqual("COMPLETED", record["status"])
        self.assertEqual(
            "NORMALIZED_FILENAME",
            record["result"]["duplicateMatchKind"],
        )
        self.assertEqual(original_path, record["result"]["sourcePath"])
        self.assertEqual(
            clean_path,
            record["result"]["matchedSourcePath"],
        )
        self.assertEqual("b" * 64, record["result"]["contentSha256"])
        self.assertEqual(
            "a" * 64,
            record["result"]["matchedContentSha256"],
        )
        self.assertFalse(record["result"]["sourceOnlyDuplicate"])
        connection.close.assert_called_once()

    def test_existing_source_without_canonical_analysis_stays_pending(
        self,
    ) -> None:
        clean_path = (
            r"D:\archive\Report_1778463079_1778470471_clean.xlsx"
        )
        original_path = r"D:\incoming\Report.xlsx"
        timestamp = "2026-07-24T00:00:00+00:00"
        with sqlite3.connect(self.database) as connection:
            connection.execute(
                """
                INSERT INTO source_documents(
                    document_uid, dataset, source_path, original_file_name,
                    source_kind, lifecycle_status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, 'XLSX', 'ACTIVE', ?, ?)
                """,
                (
                    "DOCUMENT-SOURCE-ONLY",
                    "Fixture",
                    clean_path,
                    Path(clean_path).name,
                    timestamp,
                    timestamp,
                ),
            )
            document_id = int(
                connection.execute(
                    """
                    SELECT document_id
                    FROM source_documents
                    WHERE document_uid='DOCUMENT-SOURCE-ONLY'
                    """
                ).fetchone()[0]
            )
            connection.execute(
                """
                INSERT INTO source_revisions(
                    revision_uid, document_id, source_fingerprint,
                    fingerprint_kind, content_sha256, size_bytes, mtime_ns,
                    extractor_name, extractor_version, capture_contract,
                    capture_status, is_current, captured_at
                ) VALUES (
                    ?, ?, ?, 'SHA256', ?, 123, 456,
                    'fixture', '1', 'fixture-v1',
                    'CAPTURED', 1, ?
                )
                """,
                (
                    "REVISION-SOURCE-ONLY",
                    document_id,
                    "a" * 64,
                    "a" * 64,
                    timestamp,
                ),
            )
        connection.close()
        record = {
            "recordId": "incoming-source-only",
            "sourcePath": original_path,
            "contentSha256": "b" * 64,
            "status": "PENDING",
            "result": None,
            "error": "",
        }

        reconciled = corpus._reconcile_existing_analyses(
            {"records": [record]},
            database_path=self.database,
            now=timestamp,
        )

        self.assertEqual(set(), reconciled)
        self.assertEqual("PENDING", record["status"])
        self.assertIsNone(record["result"])

    def fake_success(
        self,
        calls: list[str],
        lock: threading.Lock | None = None,
    ):
        def runner(**kwargs: object) -> dict:
            source = Path(kwargs["source_path"])
            if lock is None:
                calls.append(source.name)
            else:
                with lock:
                    calls.append(source.name)
            return {
                "schemaVersion": "incremental-xlsx-ingest-v1",
                "runId": "fixture-" + source.stem,
                "status": "NEEDS_REVIEW",
                "sourcePath": str(source),
                "contentSha256": corpus._sha256_file(source),
                "imagesAnalyzed": False,
                "integrityOk": True,
            }

        return runner

    def test_discovery_is_recursive_sorted_and_excludes_excel_lock_files(
        self,
    ) -> None:
        self.source("z-last.xlsx")
        self.source("Nested/B-first.XLSX")
        self.source("a-first.xlsx")
        self.source("Nested/~$owner.xlsx")
        self.source("ignored.xls")

        found = corpus.discover_xlsx_sources(self.sources)

        self.assertEqual(
            ["a-first.xlsx", "Nested\\B-first.XLSX", "z-last.xlsx"],
            [str(path.relative_to(self.sources)) for path in found],
        )
        com_found = corpus.discover_excel_sources(
            self.sources,
            capture_backend="com",
        )
        self.assertEqual(
            [
                "a-first.xlsx",
                "ignored.xls",
                "Nested\\B-first.XLSX",
                "z-last.xlsx",
            ],
            [str(path.relative_to(self.sources)) for path in com_found],
        )

    def test_offset_limit_concurrency_and_completed_resume_account_exactly(
        self,
    ) -> None:
        originals = {}
        for name in ("a.xlsx", "b.xlsx", "c.xlsx", "d.xlsx"):
            path = self.source(name)
            originals[name] = path.read_bytes()
        calls: list[str] = []
        runner = self.fake_success(calls, threading.Lock())

        first = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            offset=1,
            limit=2,
            workbook_workers=2,
            ingest_runner=runner,
        )
        second = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            offset=1,
            limit=2,
            workbook_workers=2,
            ingest_runner=runner,
        )

        self.assertEqual(Counter({"b.xlsx": 1, "c.xlsx": 1}), Counter(calls))
        self.assertEqual(4, first["summary"]["discoveredSources"])
        self.assertEqual(2, first["summary"]["selectedSources"])
        self.assertEqual(2, first["summary"]["attempted"])
        self.assertEqual(2, first["summary"]["completedThisRun"])
        self.assertEqual(2, first["summary"]["notSelectedSources"])
        self.assertTrue(first["summary"]["integrityOk"])
        self.assertTrue(first["integrity"]["ok"])
        self.assertEqual(0, second["summary"]["attempted"])
        self.assertEqual(2, second["summary"]["skippedCompleted"])
        self.assertTrue(all(not item["imagesAnalyzed"] for item in first["items"]))
        for name, before in originals.items():
            self.assertEqual(before, (self.sources / name).read_bytes())

        journal = json.loads(
            Path(second["journalPath"]).read_text(encoding="utf-8")
        )
        self.assertEqual(4, len(journal["records"]))
        self.assertEqual(2, len(journal["runs"]))
        self.assertFalse(journal["imagesAnalyzed"])
        self.assertTrue(
            all(record["sizeBytes"] > 0 for record in journal["records"])
        )
        self.assertTrue(
            all(record["mtimeNs"] > 0 for record in journal["records"])
        )

    def test_workbook_retry_repairs_before_the_corpus_pass_finishes(
        self,
    ) -> None:
        source = self.source("repair.xlsx")
        calls = 0

        def runner(**kwargs: object) -> dict:
            nonlocal calls
            calls += 1
            if calls < 3:
                raise ValueError("repairable draft validation failure")
            return {
                "status": "NEEDS_REVIEW",
                "contentSha256": corpus._sha256_file(source),
                "sourcePath": str(source),
                "imagesAnalyzed": False,
                "integrityOk": True,
            }

        result = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            workbook_retry_attempts=3,
            ingest_runner=runner,
        )

        self.assertEqual(3, calls)
        self.assertEqual(1, result["summary"]["completedThisRun"])
        self.assertEqual(0, result["summary"]["failedThisRun"])
        self.assertEqual(
            3,
            result["items"][0]["result"]["workflowRetryAttempts"],
        )

    def test_bounded_pipeline_overlaps_files_and_honors_stage_limits(
        self,
    ) -> None:
        for index in range(6):
            self.source(f"{index}.xlsx")
        active = Counter()
        peaks = Counter()
        overlap_seen = {"value": False}
        lock = threading.Lock()
        progress_events: list[dict] = []

        def use_stage(
            gate: object,
            stage: str,
            duration: float,
        ) -> None:
            with gate(stage):
                with lock:
                    active[stage] += 1
                    peaks[stage] = max(peaks[stage], active[stage])
                    if active["COM"] and active["AI"]:
                        overlap_seen["value"] = True
                time.sleep(duration)
                with lock:
                    active[stage] -= 1

        def runner(**kwargs: object) -> dict:
            source = Path(kwargs["source_path"])
            gate = kwargs["pipeline_gate"]
            use_stage(gate, "COM", 0.015)
            use_stage(gate, "PACKET", 0.025)
            use_stage(gate, "AI", 0.075)
            use_stage(gate, "DB", 0.015)
            return {
                "status": "NEEDS_REVIEW",
                "contentSha256": corpus._sha256_file(source),
                "sourcePath": str(source),
                "imagesAnalyzed": False,
                "integrityOk": True,
            }

        result = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            workbook_workers=6,
            com_workers=1,
            packet_workers=2,
            ai_workers=3,
            db_workers=1,
            ingest_options={
                "progress_callback": progress_events.append,
            },
            ingest_runner=runner,
        )

        self.assertEqual(1, peaks["COM"])
        self.assertLessEqual(peaks["PACKET"], 2)
        self.assertGreaterEqual(peaks["PACKET"], 2)
        self.assertLessEqual(peaks["AI"], 3)
        self.assertGreaterEqual(peaks["AI"], 2)
        self.assertEqual(1, peaks["DB"])
        self.assertTrue(overlap_seen["value"])
        self.assertEqual(
            {"COM": 1, "PACKET": 2, "AI": 3, "DB": 1},
            result["pipelineWorkers"],
        )
        corpus_events = [
            event
            for event in progress_events
            if event["stage"] == "CORPUS"
        ]
        self.assertEqual(
            ["RUNNING", "COMPLETED"],
            [event["status"] for event in corpus_events],
        )
        start_detail = json.loads(corpus_events[0]["detail"])
        self.assertEqual(6, start_detail["eligible"])
        self.assertEqual(6, start_detail["selected"])

    def test_failed_record_requires_retry_and_fingerprint_history_is_retained(
        self,
    ) -> None:
        changed = self.source("a.xlsx", b"version-one")
        self.source("b.xlsx", b"stable")
        calls: list[tuple[str, bytes]] = []
        fail_once = {"a.xlsx": True}

        def runner(**kwargs: object) -> dict:
            source = Path(kwargs["source_path"])
            payload = source.read_bytes()
            calls.append((source.name, payload))
            if fail_once.get(source.name):
                fail_once[source.name] = False
                raise RuntimeError("fixture failure")
            return {
                "status": "EXCLUDED",
                "contentSha256": corpus._sha256_file(source),
                "sourcePath": str(source),
                "imagesAnalyzed": False,
            }

        first = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            ingest_runner=runner,
        )
        second = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            ingest_runner=runner,
        )
        third = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            retry_failed=True,
            ingest_runner=runner,
        )
        changed.write_bytes(b"version-two")
        fourth = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            ingest_runner=runner,
        )

        self.assertEqual("COMPLETED_WITH_ERRORS", first["status"])
        self.assertEqual(1, first["summary"]["failedThisRun"])
        self.assertEqual(1, second["summary"]["skippedFailed"])
        self.assertEqual(1, third["summary"]["completedThisRun"])
        self.assertEqual(1, fourth["summary"]["completedThisRun"])
        self.assertEqual(3, fourth["summary"]["trackedFingerprintRecords"])
        self.assertEqual(
            1,
            fourth["summary"]["historicalOrMissingFingerprintRecords"],
        )
        self.assertEqual(
            [
                ("a.xlsx", b"version-one"),
                ("b.xlsx", b"stable"),
                ("a.xlsx", b"version-one"),
                ("a.xlsx", b"version-two"),
            ],
            calls,
        )
        journal = json.loads(
            Path(fourth["journalPath"]).read_text(encoding="utf-8")
        )
        a_records = [
            record
            for record in journal["records"]
            if record["relativePath"] == "a.xlsx"
        ]
        self.assertEqual(2, len(a_records))
        self.assertEqual(
            {False, True},
            {record["presentInLatestDiscovery"] for record in a_records},
        )
        self.assertTrue(
            all(record["status"] == "COMPLETED" for record in a_records)
        )

    def test_invalid_arguments_and_ingest_option_ownership_are_rejected(
        self,
    ) -> None:
        self.source("a.xlsx")
        with self.assertRaises(ValueError):
            corpus.run_corpus_ingest(
                database_path=self.database,
                source_root=self.sources,
                artifact_root=self.artifacts,
                workbook_workers=0,
            )
        with self.assertRaises(ValueError):
            corpus.run_corpus_ingest(
                database_path=self.database,
                source_root=self.sources,
                artifact_root=self.artifacts,
                ingest_options={"source_path": "other.xlsx"},
            )

    def test_inventory_only_freezes_every_source_without_ingest_calls(
        self,
    ) -> None:
        self.source("a.xlsx")
        self.source("b.xlsx")

        result = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            inventory_only=True,
            ingest_runner=lambda **_: self.fail(
                "Inventory-only mode must not call ingest_workbook."
            ),
        )

        self.assertEqual("COMPLETED", result["status"])
        self.assertEqual(2, result["summary"]["discoveredSources"])
        self.assertEqual(0, result["summary"]["selectedSources"])
        self.assertEqual(2, result["summary"]["notSelectedSources"])
        self.assertEqual(0, result["summary"]["attempted"])
        journal = json.loads(
            Path(result["journalPath"]).read_text(encoding="utf-8")
        )
        self.assertEqual(2, len(journal["records"]))
        self.assertTrue(
            journal["runs"][-1]["options"]["inventoryOnly"]
        )

    def test_quarantined_canonical_completion_is_requeued(self) -> None:
        now = "2026-07-18T00:00:00+00:00"
        current_key = ("C:\\sources\\current.xlsx", "a" * 64)
        stale_key = ("C:\\sources\\stale.xlsx", "b" * 64)
        journal = {
            "records": [
                {
                    "recordId": "current",
                    "sourcePath": current_key[0],
                    "contentSha256": current_key[1],
                    "status": "COMPLETED",
                    "result": {
                        "publicAnalysisId": "ANALYSIS-CURRENT",
                    },
                    "error": "",
                },
                {
                    "recordId": "stale",
                    "sourcePath": stale_key[0],
                    "contentSha256": stale_key[1],
                    "status": "COMPLETED",
                    "result": {
                        "publicAnalysisId": "ANALYSIS-STALE",
                    },
                    "error": "",
                },
                {
                    "recordId": "noncanonical-fixture",
                    "sourcePath": "C:\\sources\\fixture.xlsx",
                    "contentSha256": "c" * 64,
                    "status": "COMPLETED",
                    "result": {"status": "NEEDS_REVIEW"},
                    "error": "",
                },
                {
                    "recordId": "legacy-source-only",
                    "sourcePath": "C:\\sources\\captured.xlsx",
                    "contentSha256": "d" * 64,
                    "status": "COMPLETED",
                    "result": {
                        "schemaVersion": "existing-database-source-v1",
                        "publicAnalysisId": "",
                        "sourceOnlyDuplicate": True,
                    },
                    "error": "",
                },
            ]
        }

        downgraded = (
            corpus._downgrade_completed_without_current_analysis(
                journal,
                current_analysis_keys={current_key},
                now=now,
            )
        )

        self.assertEqual(2, downgraded)
        self.assertEqual("COMPLETED", journal["records"][0]["status"])
        self.assertEqual("PENDING", journal["records"][1]["status"])
        self.assertIsNone(journal["records"][1]["result"])
        self.assertIn(
            "semantic re-ingestion is required",
            journal["records"][1]["error"],
        )
        self.assertEqual(
            "COMPLETED",
            journal["records"][2]["status"],
        )
        self.assertEqual("PENDING", journal["records"][3]["status"])
        self.assertIsNone(journal["records"][3]["result"])

    def test_relative_path_selection_is_exact_and_deterministic(self) -> None:
        self.source("a.xlsx")
        self.source("nested/b.xlsx")
        self.source("c.xlsx")
        calls: list[str] = []

        result = corpus.run_corpus_ingest(
            database_path=self.database,
            source_root=self.sources,
            artifact_root=self.artifacts,
            dataset="Fixture",
            include_relative_paths=["nested/b.xlsx", "a.xlsx"],
            ingest_runner=self.fake_success(calls),
        )

        self.assertEqual(["b.xlsx", "a.xlsx"], calls)
        self.assertEqual(2, result["summary"]["selectedSources"])
        with self.assertRaises(corpus.CorpusWorkflowError):
            corpus.run_corpus_ingest(
                database_path=self.database,
                source_root=self.sources,
                artifact_root=self.artifacts,
                dataset="Fixture",
                include_relative_paths=["missing.xlsx"],
                ingest_runner=self.fake_success(calls),
            )


if __name__ == "__main__":
    unittest.main()
