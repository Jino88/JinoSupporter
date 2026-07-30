from __future__ import annotations

import importlib.util
import json
import tempfile
import threading
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


class SemanticLocatorBatchingTests(unittest.TestCase):
    def test_partitions_by_count_and_bytes_without_dropping_oversized_chunks(self) -> None:
        jobs = [
            ({"chunkId": "one", "cells": [{"displayValue": "a" * 40}]}, Path("one")),
            ({"chunkId": "two", "cells": [{"displayValue": "b" * 40}]}, Path("two")),
            ({"chunkId": "huge", "cells": [{"displayValue": "c" * 1000}]}, Path("huge")),
            ({"chunkId": "four", "cells": [{"displayValue": "d"}]}, Path("four")),
        ]
        batches = cli._partition_semantic_locator_jobs(
            jobs,
            batch_size=2,
            batch_max_bytes=250,
        )
        flattened = [
            job[0]["chunkId"]
            for batch in batches
            for job in batch
        ]
        self.assertEqual(["one", "two", "huge", "four"], flattened)
        self.assertTrue(any(batch[0][0]["chunkId"] == "huge" for batch in batches))
        self.assertTrue(all(len(batch) <= 2 for batch in batches))


class CanonicalWorkflowCliTests(unittest.TestCase):
    def test_form_preflight_parser_accepts_configured_output_root(self) -> None:
        parser = cli.build_parser()
        args = parser.parse_args(
            [
                "form-preflight",
                "--db",
                "canonical.sqlite",
                "--input",
                "archive",
                "--out",
                "configured/form-preflight/latest.json",
                "--output-root",
                "configured",
                "--cancel-file",
                "configured/form-preflight/cancel.request",
            ]
        )

        self.assertEqual(args.output_root, "configured")
        self.assertEqual(
            args.cancel_file,
            "configured/form-preflight/cancel.request",
        )

    def test_form_registry_parsers_expose_review_analyze_and_decide(
        self,
    ) -> None:
        parser = cli.build_parser()
        review = parser.parse_args(
            [
                "form-group-review",
                "--db",
                "canonical.sqlite",
                "--report",
                "latest.json",
            ]
        )
        analyze = parser.parse_args(
            [
                "form-family-analyze",
                "--db",
                "canonical.sqlite",
                "--report",
                "latest.json",
                "--family-id",
                "family-123",
            ]
        )
        decide = parser.parse_args(
            [
                "form-family-decide",
                "--db",
                "canonical.sqlite",
                "--report",
                "latest.json",
                "--family-id",
                "family-123",
                "--decision",
                "REGISTER_NEW",
                "--reviewer",
                "tester",
            ]
        )

        self.assertIs(review.func, cli.cmd_form_group_review)
        self.assertIs(analyze.func, cli.cmd_form_family_analyze)
        self.assertEqual("medium", analyze.reasoning_effort)
        self.assertIs(decide.func, cli.cmd_form_family_decide)
        self.assertEqual("tester", decide.reviewer)

    def test_form_pipeline_complete_parser_defaults_to_all_groups(
        self,
    ) -> None:
        args = cli.build_parser().parse_args(
            [
                "form-pipeline-complete",
                "--db",
                "canonical.sqlite",
                "--input",
                "archive",
                "--output-root",
                "configured",
                "--reviewer",
                "cli-user",
            ]
        )

        self.assertIs(args.func, cli.cmd_form_pipeline_complete)
        self.assertEqual(2, args.analysis_workers)
        self.assertEqual("low", args.reasoning_effort)
        self.assertEqual(0, args.max_families)
        self.assertFalse(args.skip_corpus)
        self.assertEqual(400_000, args.draft_monolithic_max_bytes)

    def test_form_family_decision_rebuilds_report_and_review(self) -> None:
        output_root = self.root / "configured-output"
        report = output_root / "form-preflight" / "latest.json"
        report.parent.mkdir(parents=True)
        report.write_text("{}", encoding="utf-8")
        args = cli.build_parser().parse_args(
            [
                "form-family-decide",
                "--db",
                str(self.root / "canonical.sqlite"),
                "--report",
                str(report),
                "--family-id",
                "family-123",
                "--decision",
                "EXCLUDE",
                "--reviewer",
                "tester",
                "--output-root",
                str(output_root),
            ]
        )
        with (
            mock.patch.object(
                cli,
                "decide_form_family",
                return_value={
                    "status": "EXCLUDED",
                    "familyId": "family-123",
                    "memberCount": 2,
                    "linkedFormSignatureId": "",
                },
            ) as decide,
            mock.patch.object(
                cli,
                "reclassify_form_preflight_report",
                return_value={
                    "knownFormManifestPath": "manifest.json",
                    "summary": {"knownForms": 0},
                },
            ) as reclassify,
            mock.patch.object(
                cli,
                "write_form_group_review",
                return_value={"summary": {"pendingCount": 1}},
            ) as review,
            mock.patch.object(cli, "print_json") as print_json,
        ):
            result = args.func(args)

        self.assertEqual(0, result)
        decide.assert_called_once()
        reclassify.assert_called_once()
        review.assert_called_once()
        self.assertEqual(
            "EXCLUDED",
            print_json.call_args.args[0]["status"],
        )

    def test_output_path_under_root_allows_configured_root_only(self) -> None:
        configured_root = Path(self.temp.name) / "configured-output"
        expected = configured_root / "form-preflight" / "latest.json"

        actual = cli.output_path_under_root(
            str(expected),
            expected,
            configured_root,
        )

        self.assertEqual(actual, expected.resolve())
        self.assertTrue(actual.parent.is_dir())
        with self.assertRaises(SystemExit):
            cli.output_path_under_root(
                str(Path(self.temp.name) / "outside" / "latest.json"),
                expected,
                configured_root,
            )

    def test_database_scoped_output_path_allows_configured_output_tree(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as external_temp:
            external = Path(external_temp)
            configured_root = external / "configured-output"
            database = (
                configured_root
                / "universal-grid"
                / "InputDataFinish.sqlite"
            )
            database.parent.mkdir(parents=True)
            database.touch()
            expected = (
                configured_root
                / "wpf-evidence"
                / "old-new.answer.json"
            )

            actual = cli.database_scoped_output_path(
                str(expected),
                cli.OUTPUT_DIR / "evidence-answers" / "default.json",
                database,
            )

            self.assertEqual(expected.resolve(), actual)
            self.assertTrue(actual.parent.is_dir())
            with self.assertRaises(SystemExit):
                cli.database_scoped_output_path(
                    str(external / "outside" / "answer.json"),
                    cli.OUTPUT_DIR / "evidence-answers" / "default.json",
                    database,
                )

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=cli.SERVICE_DIR)
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_evidence_answer_command_writes_and_validates_deterministic_outputs(
        self,
    ) -> None:
        pack = {
            "schemaVersion": "canonical-evidence-pack-v1",
            "question": "새로운 공정 인자가 새로운 결과에 영향을 주나?",
            "normalizedQuestion": "",
            "queryTokens": [],
            "searchTokens": [],
            "queryRoleHints": {
                "outcomeTerms": [],
                "contextOrFactorTerms": [],
                "relationGateApplied": False,
            },
            "studyCandidates": [],
            "answerEligibleEffects": [],
            "excludedCandidates": [],
            "eligibleEffectSummary": [],
            "summary": {
                "relevantStudyCount": 0,
                "answerEligibleEffectCount": 0,
                "excludedCandidateCount": 0,
            },
        }
        pack_path = self.root / "pack.json"
        answer_path = self.root / "answer.json"
        markdown_path = self.root / "answer.md"
        pack_path.write_text(
            json.dumps(pack, ensure_ascii=False),
            encoding="utf-8",
        )
        args = cli.build_parser().parse_args(
            [
                "evidence-answer",
                "--pack",
                str(pack_path),
                "--out-json",
                str(answer_path),
                "--out-markdown",
                str(markdown_path),
            ]
        )
        with mock.patch.object(cli, "print_json"):
            self.assertEqual(0, args.func(args))
        answer = json.loads(answer_path.read_text(encoding="utf-8"))
        self.assertEqual("NO_RELEVANT_DATA", answer["answerStatus"])
        self.assertIn("근거 기반 답변", markdown_path.read_text(encoding="utf-8"))

        validate_args = cli.build_parser().parse_args(
            [
                "evidence-answer-validate",
                "--pack",
                str(pack_path),
                "--answer",
                str(answer_path),
            ]
        )
        with mock.patch.object(cli, "print_json"):
            self.assertEqual(0, validate_args.func(validate_args))

    def test_ingest_workbook_parser_exposes_review_gated_workflow_options(
        self,
    ) -> None:
        args = cli.build_parser().parse_args(
            [
                "ingest-workbook",
                "--db",
                "knowledge.sqlite",
                "--input",
                "new.xlsx",
                "--capture-backend",
                "com",
                "--covered-cell-mode",
                "blank",
                "--dismiss-auth-dialog",
                "--auth-dialog-title",
                "Company Login",
                "--auth-dialog-class",
                "#32770",
                "--auth-dialog-button",
                "Continue",
                "--workers",
                "2",
                "--batch-size",
                "4",
                "--draft-monolithic-max-bytes",
                "234567",
                "--draft-fragment-max-chunks",
                "5",
                "--draft-fragment-max-cells",
                "600",
                "--draft-fragment-max-bytes",
                "123456",
                "--draft-fragment-workers",
                "2",
                "--derive-formula-values",
                "--no-resume",
            ]
        )
        self.assertIs(cli.cmd_ingest_workbook, args.func)
        self.assertEqual(2, args.workers)
        self.assertEqual("com", args.capture_backend)
        self.assertEqual("blank", args.covered_cell_mode)
        self.assertTrue(args.dismiss_auth_dialog)
        self.assertEqual("Company Login", args.auth_dialog_title)
        self.assertEqual("#32770", args.auth_dialog_class)
        self.assertEqual("Continue", args.auth_dialog_button)
        self.assertEqual(4, args.batch_size)
        self.assertEqual(234567, args.draft_monolithic_max_bytes)
        self.assertEqual(5, args.draft_fragment_max_chunks)
        self.assertEqual(600, args.draft_fragment_max_cells)
        self.assertEqual(123456, args.draft_fragment_max_bytes)
        self.assertEqual(2, args.draft_fragment_workers)
        self.assertTrue(args.derive_formula_values)
        self.assertTrue(args.no_resume)

    def test_openxml_index_parser_exposes_parallel_reader_count(self) -> None:
        args = cli.build_parser().parse_args(
            [
                "openxml-index",
                "--input",
                "corpus",
                "--workers",
                "4",
            ]
        )
        self.assertIs(cli.cmd_openxml_index, args.func)
        self.assertEqual(4, args.workers)

    def test_parallel_capture_extraction_preserves_source_order_and_errors(
        self,
    ) -> None:
        sources = [
            self.root / "a.xlsx",
            self.root / "b.xlsx",
            self.root / "c.xlsx",
        ]
        first_pair_started = threading.Barrier(2)

        def extract(source: Path) -> dict:
            if source.name in {"a.xlsx", "b.xlsx"}:
                first_pair_started.wait(timeout=2)
            if source.name == "b.xlsx":
                raise ValueError("broken fixture")
            return {"source": {"fileName": source.name}}

        with mock.patch.object(
            cli,
            "extract_openxml_workbook",
            side_effect=extract,
        ):
            results = list(
                cli._ordered_capture_v2_extractions(
                    sources,
                    workers=2,
                )
            )

        self.assertEqual(sources, [item[0] for item in results])
        self.assertEqual("a.xlsx", results[0][1]["source"]["fileName"])
        self.assertIsNone(results[0][2])
        self.assertIsNone(results[1][1])
        self.assertIsInstance(results[1][2], ValueError)
        self.assertEqual("c.xlsx", results[2][1]["source"]["fileName"])
        self.assertIsNone(results[2][2])

    def test_related_studies_parser_is_domain_neutral_and_bounded(
        self,
    ) -> None:
        args = cli.build_parser().parse_args(
            [
                "related-studies",
                "--db",
                "knowledge.sqlite",
                "--target",
                "DATA-ARBITRARY",
                "--limit",
                "7",
            ]
        )
        self.assertIs(cli.cmd_related_studies, args.func)
        self.assertEqual("DATA-ARBITRARY", args.target)
        self.assertEqual(7, args.limit)

    def test_concept_curation_cli_lists_and_rejects_with_json(
        self,
    ) -> None:
        database = self.root / "concept-curation.sqlite"
        candidate_uid = "schema_candidate_cli_fixture"
        with cli.connect_rw(database) as connection:
            cli.ensure_universal_schema(connection)
            connection.execute(
                """
                INSERT INTO knowledge_schema_candidates(
                    candidate_uid, candidate_kind, normalized_value,
                    original_value, suggested_canonical_name,
                    occurrence_count, status, first_seen_at, last_seen_at
                ) VALUES (
                    ?, 'CONCEPT:ARBITRARY_CONTEXT', 'fixture context',
                    'Fixture context', '', 1, 'OPEN',
                    '2026-07-18T00:00:00Z',
                    '2026-07-18T00:00:00Z'
                )
                """,
                (candidate_uid,),
            )
            connection.commit()

        candidates = cli.build_parser().parse_args(
            [
                "concept-candidates",
                "--db",
                str(database),
                "--kind",
                "CONCEPT:ARBITRARY_CONTEXT",
                "--query",
                "fixture",
            ]
        )
        printed: list[dict] = []
        with mock.patch.object(
            cli,
            "print_json",
            side_effect=printed.append,
        ):
            self.assertEqual(0, candidates.func(candidates))
        self.assertEqual("concept-candidate-list-v1", printed[0]["schemaVersion"])
        self.assertEqual(candidate_uid, printed[0]["candidates"][0]["candidateUid"])

        concepts = cli.build_parser().parse_args(
            [
                "concept-list",
                "--db",
                str(database),
                "--kind",
                "OUTCOME",
                "--query",
                "function ng",
            ]
        )
        printed.clear()
        with mock.patch.object(
            cli,
            "print_json",
            side_effect=printed.append,
        ):
            self.assertEqual(0, concepts.func(concepts))
        self.assertEqual("canonical-concept-list-v1", printed[0]["schemaVersion"])
        self.assertGreaterEqual(printed[0]["count"], 1)

        reject = cli.build_parser().parse_args(
            [
                "concept-resolve",
                "--db",
                str(database),
                "--candidate-uid",
                candidate_uid,
                "--action",
                "REJECT",
                "--reviewer",
                "cli-human",
                "--note",
                "Workbook-local context.",
            ]
        )
        printed.clear()
        with mock.patch.object(
            cli,
            "print_json",
            side_effect=printed.append,
        ):
            self.assertEqual(0, reject.func(reject))
        self.assertEqual("concept-resolution-v1", printed[0]["schemaVersion"])
        self.assertEqual("REJECTED", printed[0]["candidate"]["status"])

        alias = cli.build_parser().parse_args(
            [
                "concept-alias-upsert",
                "--db",
                str(database),
                "--concept-uid",
                "concept_fixture",
                "--alias",
                "fixture alias",
                "--reviewer",
                "cli-human",
                "--note",
                "Checked source.",
            ]
        )
        self.assertIs(cli.cmd_concept_alias_upsert, alias.func)

    def test_review_and_golden_acceptance_parsers_expose_safe_workflows(
        self,
    ) -> None:
        review = cli.build_parser().parse_args(
            [
                "review-decide",
                "--db",
                "knowledge.sqlite",
                "--comparison-id",
                "CMP-ARBITRARY",
                "--decision",
                "APPROVE",
                "--reviewer",
                "human-1",
                "--reason",
                "Checked source and pairing.",
                "--study-comparability",
                "VALID",
                "--study-confounding",
                "NONE",
                "--comparison-validity",
                "VALID",
                "--comparison-confounding",
                "NONE",
                "--matching-basis",
                "same unit and period",
            ]
        )
        self.assertIs(cli.cmd_review_decide, review.func)
        self.assertEqual("CMP-ARBITRARY", review.comparison_id)
        self.assertEqual("VALID", review.study_comparability)
        self.assertEqual("NONE", review.comparison_confounding)

        acceptance = cli.build_parser().parse_args(
            [
                "golden-acceptance",
                "--db",
                "knowledge.sqlite",
                "--out-dir",
                "outputs/golden-acceptance/test",
            ]
        )
        self.assertIs(cli.cmd_golden_acceptance, acceptance.func)
        self.assertTrue(
            acceptance.manifest.endswith("representative-pilot-v1.json")
        )

    def test_ingest_corpus_parser_exposes_durable_chunk_and_retry_options(
        self,
    ) -> None:
        args = cli.build_parser().parse_args(
            [
                "ingest-corpus",
                "--db",
                "knowledge.sqlite",
                "--input",
                "corpus",
                "--source-manifest",
                "pilot.json",
                "--offset",
                "50",
                "--limit",
                "25",
                "--workbook-workers",
                "4",
                "--com-workers",
                "1",
                "--packet-workers",
                "2",
                "--ai-workers",
                "3",
                "--db-workers",
                "1",
                "--locator-workers",
                "3",
                "--draft-fragment-max-chunks",
                "7",
                "--draft-monolithic-max-bytes",
                "345678",
                "--draft-fragment-workers",
                "4",
                "--derive-formula-values",
                "--retry-failed",
            ]
        )
        self.assertIs(cli.cmd_ingest_corpus, args.func)
        self.assertEqual(50, args.offset)
        self.assertEqual(25, args.limit)
        self.assertEqual(4, args.workbook_workers)
        self.assertEqual(1, args.com_workers)
        self.assertEqual(2, args.packet_workers)
        self.assertEqual(3, args.ai_workers)
        self.assertEqual(1, args.db_workers)
        self.assertEqual(3, args.locator_workers)
        self.assertEqual(7, args.draft_fragment_max_chunks)
        self.assertEqual(345678, args.draft_monolithic_max_bytes)
        self.assertEqual(4, args.draft_fragment_workers)
        self.assertTrue(args.derive_formula_values)
        self.assertTrue(args.retry_failed)
        self.assertEqual("pilot.json", args.source_manifest)


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
            self.assertTrue(cli.table_exists(conn, "knowledge_studies"))
            self.assertTrue(cli.table_exists(conn, "knowledge_effects"))
            self.assertTrue(cli.table_exists(conn, "evidence_items"))
            self.assertEqual(
                1,
                conn.execute(
                    "SELECT COUNT(*) FROM schema_migrations WHERE migration_name='canonical-study-evidence-v1'"
                ).fetchone()[0],
            )

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
            self.assertEqual(1, conn.execute("SELECT COUNT(*) FROM knowledge_studies").fetchone()[0])
            self.assertEqual(2, conn.execute("SELECT COUNT(*) FROM knowledge_arms").fetchone()[0])
            self.assertEqual(1, conn.execute("SELECT COUNT(*) FROM knowledge_outcomes").fetchone()[0])
            self.assertEqual(2, conn.execute("SELECT COUNT(*) FROM knowledge_observations").fetchone()[0])
            self.assertEqual(1, conn.execute("SELECT COUNT(*) FROM knowledge_comparisons").fetchone()[0])
            self.assertEqual(2, conn.execute("SELECT COUNT(*) FROM knowledge_effects").fetchone()[0])
            self.assertEqual(5, conn.execute("SELECT COUNT(*) FROM evidence_items").fetchone()[0])
            public_data_id = conn.execute("SELECT public_data_id FROM knowledge_studies").fetchone()[0]
            self.assertRegex(public_data_id, r"^DATA-[0-9A-F]{12}$")
            self.assertEqual(imported["canonical"]["analysisUid"], conn.execute("SELECT analysis_uid FROM workbook_analyses").fetchone()[0])

            verification = cli.verify_analysis_report(conn, imported["analysisReportId"])
            self.assertTrue(verification["ok"])
            exported = cli.build_analysis_export(conn, imported["analysisReportId"])
            self.assertEqual("universal-analysis-export-v1", exported["schemaVersion"])
            self.assertEqual("fixture-validation", exported["report"]["key"])
            metric = exported["reviews"][0]["metrics"][0]
            self.assertEqual(["test", "control"], [value["cohort_key"] for value in metric["values"]])
            self.assertEqual(-10000, metric["comparisons"][0]["delta_value"])

    def test_canonical_projection_is_idempotent_and_public_ids_are_stable(self) -> None:
        self.write_payload()
        manifest_path = self.root / "fixture.analysis.json"
        manifest_path.write_text(json.dumps(self.analysis_manifest(), ensure_ascii=False), encoding="utf-8")
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            first = cli.import_analysis_manifest(conn, manifest_path, cli.read_analysis_manifest(manifest_path))
            first_ids = {
                "data": conn.execute("SELECT public_data_id FROM knowledge_studies").fetchone()[0],
                "comparison": conn.execute("SELECT public_comparison_id FROM knowledge_comparisons").fetchone()[0],
                "evidence": [
                    row[0]
                    for row in conn.execute("SELECT public_evidence_id FROM evidence_items ORDER BY public_evidence_id")
                ],
            }
            second = cli.import_analysis_manifest(conn, manifest_path, cli.read_analysis_manifest(manifest_path))
            second_ids = {
                "data": conn.execute("SELECT public_data_id FROM knowledge_studies").fetchone()[0],
                "comparison": conn.execute("SELECT public_comparison_id FROM knowledge_comparisons").fetchone()[0],
                "evidence": [
                    row[0]
                    for row in conn.execute("SELECT public_evidence_id FROM evidence_items ORDER BY public_evidence_id")
                ],
            }
            counts = cli.knowledge_counts(conn)
            integrity = cli.validate_knowledge_integrity(conn)

        self.assertNotEqual(first["analysisReportId"], second["analysisReportId"])
        self.assertEqual(first_ids, second_ids)
        self.assertEqual(5, len(second_ids["evidence"]))
        self.assertEqual(1, counts["knowledge_studies"])
        self.assertEqual(1, counts["knowledge_comparisons"])
        self.assertEqual(5, counts["evidence_items"])
        self.assertTrue(integrity["ok"], integrity)

    def test_aggregation_guard_rejects_unreviewed_legacy_effect(self) -> None:
        self.write_payload()
        manifest_path = self.root / "fixture.analysis.json"
        manifest_path.write_text(json.dumps(self.analysis_manifest(), ensure_ascii=False), encoding="utf-8")
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            cli.import_analysis_manifest(conn, manifest_path, cli.read_analysis_manifest(manifest_path))
            effect_id = int(conn.execute("SELECT effect_id FROM knowledge_effects ORDER BY effect_id LIMIT 1").fetchone()[0])
            comparison_id = int(conn.execute("SELECT comparison_id FROM knowledge_comparisons").fetchone()[0])
            with self.assertRaisesRegex(Exception, "aggregation-eligible effect"):
                conn.execute("UPDATE knowledge_effects SET aggregation_eligible=1 WHERE effect_id=?", (effect_id,))
            conn.execute(
                """
                UPDATE knowledge_comparisons
                SET validity_status='VALID', confounding_status='NONE',
                    aggregation_eligible=1, verification_status='VERIFIED'
                WHERE comparison_id=?
                """,
                (comparison_id,),
            )
            conn.execute(
                "UPDATE knowledge_effects SET verification_status='VERIFIED', aggregation_eligible=1 WHERE effect_id=?",
                (effect_id,),
            )
            self.assertEqual(1, conn.execute("SELECT aggregation_eligible FROM knowledge_effects WHERE effect_id=?", (effect_id,)).fetchone()[0])

    def test_canonical_schema_represents_vp_cd_context_and_changed_factor(self) -> None:
        self.write_payload()
        manifest_path = self.root / "fixture.analysis.json"
        manifest_path.write_text(json.dumps(self.analysis_manifest(), ensure_ascii=False), encoding="utf-8")
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            cli.import_analysis_manifest(conn, manifest_path, cli.read_analysis_manifest(manifest_path))
            study_id = int(conn.execute("SELECT study_id FROM knowledge_studies").fetchone()[0])
            process_concept_id = int(
                conn.execute(
                    "SELECT concept_id FROM knowledge_concepts WHERE canonical_name='VP+CD assembly'"
                ).fetchone()[0]
            )
            factor_concept_id = int(
                conn.execute(
                    "SELECT concept_id FROM knowledge_concepts WHERE canonical_name='Bonding amount'"
                ).fetchone()[0]
            )
            conn.execute(
                """
                INSERT INTO knowledge_study_contexts(
                    context_uid, study_id, context_kind, concept_id,
                    original_value, normalized_value, verification_status
                ) VALUES ('context_fixture_vp_cd', ?, 'PROCESS', ?,
                          'VP+CD 조립', 'vp+cd assembly', 'VERIFIED')
                """,
                (study_id, process_concept_id),
            )
            conn.execute(
                """
                INSERT INTO knowledge_factors(
                    factor_uid, study_id, concept_id, factor_key, factor_domain,
                    original_label, baseline_condition, changed_condition,
                    change_direction, isolation_status, verification_status
                ) VALUES ('factor_fixture_bond_amount', ?, ?, 'vp-cd-bonding-amount',
                          'ASSEMBLY', 'VP+CD 본드량', 'Normal', '1.5 mg',
                          'INCREASED', 'ISOLATED', 'VERIFIED')
                """,
                (study_id, factor_concept_id),
            )
            context = conn.execute(
                """
                SELECT c.original_value, k.canonical_name
                FROM knowledge_study_contexts c
                JOIN knowledge_concepts k ON k.concept_id=c.concept_id
                """
            ).fetchone()
            factor = conn.execute(
                """
                SELECT f.baseline_condition, f.changed_condition, k.canonical_name
                FROM knowledge_factors f
                JOIN knowledge_concepts k ON k.concept_id=f.concept_id
                """
            ).fetchone()

        self.assertEqual(("VP+CD 조립", "VP+CD assembly"), tuple(context))
        self.assertEqual(("Normal", "1.5 mg", "Bonding amount"), tuple(factor))

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

    def test_universal_packet_compacts_empty_grid_cells_without_losing_zero_or_coordinates(self) -> None:
        self.write_payload()
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            imported = cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            packet = cli.build_universal_packet(conn, imported["workbookId"], row_limit=20, cell_limit=20)

        self.assertEqual("inference-data-ai-reviewcase-packet-v2", packet["schemaVersion"])
        self.assertNotIn("sheetCells", packet)
        self.assertFalse(packet["packetSelection"]["dataTruncated"])
        cells = [cell for row in packet["sheetRows"] for cell in row["cells"]]
        self.assertIn("C3", [cell["address"] for cell in cells])
        self.assertIn("0", [str(cell["value"]) for cell in cells])
        self.assertTrue(all(cell["value"] not in (None, "") for cell in cells))

    def test_universal_packet_marks_limited_source_as_needs_review_only(self) -> None:
        self.write_payload()
        with cli.connect_rw(self.db) as conn:
            cli.ensure_universal_schema(conn)
            imported = cli.import_com_json(conn, "FixtureDataset", self.raw_json)
            packet = cli.build_universal_packet(conn, imported["workbookId"], row_limit=1, cell_limit=1)

        selection = packet["packetSelection"]
        self.assertTrue(selection["dataTruncated"])
        self.assertTrue(selection["rowTruncated"] or selection["cellTruncated"])
        self.assertEqual(1, selection["includedCells"])
        self.assertIn("packetSelection.dataTruncated", packet["reviewCaseContract"]["packetCompletenessRule"])
        self.assertTrue(any("NEEDS_REVIEW" in note for note in packet["notes"]))

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
