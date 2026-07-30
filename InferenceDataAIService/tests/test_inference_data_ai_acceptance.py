from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

import inference_data_ai_acceptance as acceptance


class GoldenQuestionAcceptanceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.source_root = self.root / "sources"
        self.source_root.mkdir()
        self.database = self.root / "knowledge.sqlite"
        self.output = self.root / "acceptance"
        self.connection = sqlite3.connect(self.database)
        self.connection.executescript(
            """
            CREATE TABLE source_documents(
                document_id INTEGER PRIMARY KEY,
                source_path TEXT NOT NULL,
                original_file_name TEXT NOT NULL,
                lifecycle_status TEXT NOT NULL
            );
            CREATE TABLE source_revisions(
                revision_id INTEGER PRIMARY KEY,
                revision_uid TEXT NOT NULL,
                document_id INTEGER NOT NULL,
                content_sha256 TEXT NOT NULL,
                is_current INTEGER NOT NULL
            );
            CREATE TABLE workbook_analyses(
                workbook_analysis_id INTEGER PRIMARY KEY,
                analysis_uid TEXT NOT NULL,
                public_analysis_id TEXT NOT NULL,
                document_id INTEGER NOT NULL,
                revision_id INTEGER NOT NULL,
                analysis_status TEXT NOT NULL,
                verification_status TEXT NOT NULL
            );
            CREATE TABLE knowledge_studies(
                study_id INTEGER PRIMARY KEY,
                public_data_id TEXT NOT NULL,
                workbook_analysis_id INTEGER NOT NULL
            );
            """
        )
        self.connection.commit()

    def tearDown(self) -> None:
        self.connection.close()
        self.temporary.cleanup()

    def _manifest(
        self,
        *,
        workbooks: list[tuple[str, str]],
        questions: list[dict],
    ) -> Path:
        path = self.root / "representative-pilot-v1.json"
        path.write_text(
            json.dumps(
                {
                    "schemaVersion": "representative-pilot-v1",
                    "sourceRoot": str(self.source_root),
                    "workbooks": [
                        {"id": pilot_id, "relativePath": relative_path}
                        for pilot_id, relative_path in workbooks
                    ],
                    "goldenQuestions": questions,
                }
            ),
            encoding="utf-8",
        )
        return path

    def _insert_analysis(
        self,
        relative_path: str,
        *,
        suffix: int,
        with_study: bool,
        analysis_status: str = "COMPLETE",
    ) -> str:
        source_path = str((self.source_root / relative_path).resolve())
        self.connection.execute(
            """
            INSERT INTO source_documents
            VALUES (?, ?, ?, 'ACTIVE')
            """,
            (suffix, source_path, relative_path),
        )
        self.connection.execute(
            """
            INSERT INTO source_revisions
            VALUES (?, ?, ?, ?, 1)
            """,
            (
                suffix,
                f"revision-{suffix}",
                suffix,
                str(suffix) * 64,
            ),
        )
        self.connection.execute(
            """
            INSERT INTO workbook_analyses
            VALUES (?, ?, ?, ?, ?, ?, 'VERIFIED')
            """,
            (
                suffix,
                f"analysis-{suffix}",
                f"ANALYSIS-{suffix}",
                suffix,
                suffix,
                analysis_status,
            ),
        )
        if with_study:
            self.connection.execute(
                """
                INSERT INTO knowledge_studies
                VALUES (?, ?, ?)
                """,
                (suffix, f"DATA-{suffix}", suffix),
            )
        self.connection.commit()
        return source_path

    @staticmethod
    def _answer_builder(pack: dict) -> dict:
        return {
            "question": pack["question"],
            "deterministic": True,
        }

    @staticmethod
    def _answer_validator(answer: dict, pack: dict) -> dict:
        if answer["question"] != pack["question"]:
            raise ValueError("answer mismatch")
        return answer

    def test_exact_study_and_source_exclusion_representations_pass(self) -> None:
        study_path = self._insert_analysis(
            "study.xlsx",
            suffix=1,
            with_study=True,
        )
        excluded_path = self._insert_analysis(
            "excluded.xlsx",
            suffix=2,
            with_study=False,
            analysis_status="NO_TABULAR_EVIDENCE",
        )
        manifest = self._manifest(
            workbooks=[
                ("P01", "study.xlsx"),
                ("P02", "excluded.xlsx"),
            ],
            questions=[
                {
                    "id": "GQ01",
                    "question": "arbitrary factor and outcome",
                    "primaryPilotIds": ["P01", "P02"],
                    "requiredBehavior": [
                        "Separate every factor.",
                        "Show exact evidence.",
                    ],
                }
            ],
        )

        def pack_builder(_database: Path, question: str) -> dict:
            return {
                "question": question,
                "studyCandidates": [
                    {
                        "publicDataId": "DATA-1",
                        "source": {"sourcePath": study_path},
                        "analysis": {"publicAnalysisId": "ANALYSIS-1"},
                    }
                ],
                "sourceExclusions": [
                    {
                        "publicAnalysisId": "ANALYSIS-2",
                        "sourcePath": excluded_path,
                    }
                ],
                "answerEligibleEffects": [
                    {
                        "publicDataId": "DATA-1",
                        "publicComparisonId": "CMP-1",
                        "publicEffectId": "EFF-1",
                        "publicEvidenceIds": ["EVD-1", "EVD-2"],
                    },
                    {
                        "publicDataId": "DATA-1",
                        "publicComparisonId": "CMP-1",
                        "publicEffectId": "EFF-2",
                        "publicEvidenceIds": ["EVD-2"],
                    },
                ],
            }

        report = acceptance.run_golden_question_acceptance(
            self.database,
            manifest,
            self.output,
            pack_builder=pack_builder,
            answer_builder=self._answer_builder,
            answer_validator=self._answer_validator,
        )

        self.assertEqual("PASS", report["overallStatus"])
        question = report["questions"][0]
        self.assertEqual("PASS", question["status"])
        self.assertEqual(
            ["REPRESENTED", "REPRESENTED"],
            [source["status"] for source in question["primarySources"]],
        )
        self.assertEqual(
            ["DATA-1"],
            question["primarySources"][0]["representedThrough"][
                "studyCandidates"
            ]["publicDataIds"],
        )
        self.assertEqual(
            [],
            question["primarySources"][0]["representedThrough"][
                "sourceExclusions"
            ]["publicAnalysisIds"],
        )
        self.assertEqual(
            ["ANALYSIS-2"],
            question["primarySources"][1]["representedThrough"][
                "sourceExclusions"
            ]["publicAnalysisIds"],
        )
        self.assertEqual(
            {
                "eligibleEffectCount": 2,
                "eligibleDataCount": 1,
                "eligibleComparisonCount": 1,
                "eligibleEvidenceCount": 2,
            },
            question["counts"],
        )
        self.assertTrue(
            all(
                item["status"] == "MANUAL_REVIEW_REQUIRED"
                for item in question["requiredBehavior"]
            )
        )
        self.assertFalse(report["imagesAnalyzed"])
        self.assertTrue(
            (self.output / "questions" / "GQ01.pack.json").is_file()
        )
        self.assertTrue(
            (self.output / "questions" / "GQ01.answer.json").is_file()
        )
        on_disk = json.loads(
            (self.output / "acceptance-report.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(report, on_disk)

    def test_declarative_required_behaviors_verify_confounded_answer(
        self,
    ) -> None:
        source_path = self._insert_analysis(
            "confounded.xlsx",
            suffix=6,
            with_study=True,
        )
        factor_differences = [
            {
                "factorLabel": "Mold temperature",
                "controlValue": "180 C",
                "comparedValue": "190 C",
                "controlValueRecorded": True,
                "comparedValueRecorded": True,
            },
            {
                "factorLabel": "Vulcanizing agent",
                "controlValue": "",
                "comparedValue": "10%",
                "controlValueRecorded": False,
                "comparedValueRecorded": True,
            },
        ]
        manifest = self._manifest(
            workbooks=[("P06", "confounded.xlsx")],
            questions=[
                {
                    "id": "GQ06",
                    "question": "confounded factor question",
                    "primaryPilotIds": ["P06"],
                    "requiredBehavior": [
                        "List every changed factor.",
                        "Return the multi-factor code.",
                        "Do not calculate an isolated effect.",
                    ],
                    "requiredBehaviorAssertions": [
                        {
                            "type": "CONFOUNDED_FACTORS_COMPLETE",
                            "minimumFactorCount": 2,
                            "minimumComparisonCount": 1,
                        },
                        {
                            "type": "REQUIRED_ANSWER_CODE",
                            "code": "CONFOUNDED_MULTI_FACTOR",
                        },
                        {
                            "type": "MAX_ELIGIBLE_EFFECT_COUNT",
                            "maximum": 0,
                        },
                    ],
                }
            ],
        )

        def pack_builder(_database: Path, question: str) -> dict:
            return {
                "question": question,
                "studyCandidates": [
                    {
                        "publicDataId": "DATA-6",
                        "source": {"sourcePath": source_path},
                        "analysis": {"publicAnalysisId": "ANALYSIS-6"},
                    }
                ],
                "sourceExclusions": [],
                "answerEligibleEffects": [],
                "excludedCandidates": [
                    {
                        "publicDataId": "DATA-6",
                        "publicComparisonId": "CMP-6",
                        "comparison": {
                            "confoundingStatus": "CONFOUNDED",
                            "factorDifferences": factor_differences,
                        },
                    }
                ],
            }

        def answer_builder(pack: dict) -> dict:
            return {
                "question": pack["question"],
                "deterministic": True,
                "quantitativeGroups": [],
                "excludedRecords": [
                    {
                        "dataId": "DATA-6",
                        "comparisonId": "CMP-6",
                        "reasonCodes": [
                            "CONFOUNDED_MULTI_FACTOR"
                        ],
                        "comparisonAssessment": {
                            "code": "CONFOUNDED_MULTI_FACTOR",
                            "factorDifferences": factor_differences,
                        },
                    }
                ],
                "limitations": [
                    {"code": "CONFOUNDED_MULTI_FACTOR"}
                ],
                "directAnswer": {
                    "templateCode": "NO_VALID_COMPARISON"
                },
            }

        report = acceptance.run_golden_question_acceptance(
            self.database,
            manifest,
            self.output,
            pack_builder=pack_builder,
            answer_builder=answer_builder,
            answer_validator=self._answer_validator,
        )

        self.assertEqual("PASS", report["overallStatus"])
        behaviors = report["questions"][0]["requiredBehavior"]
        self.assertEqual(["PASS", "PASS", "PASS"], [
            item["status"] for item in behaviors
        ])
        self.assertEqual(
            3,
            report["summary"]["automatedRequiredBehaviorCount"],
        )
        self.assertEqual(
            3,
            report["summary"]["passedAutomatedRequiredBehaviorCount"],
        )
        self.assertEqual(
            0,
            report["summary"]["manualRequiredBehaviorCount"],
        )

    def test_pending_ingest_is_not_reported_as_retrieval_miss(self) -> None:
        manifest = self._manifest(
            workbooks=[("P03", "not-yet-ingested.xlsx")],
            questions=[
                {
                    "id": "GQ02",
                    "question": "new source question",
                    "primaryPilotIds": ["P03"],
                    "requiredBehavior": [],
                }
            ],
        )

        report = acceptance.run_golden_question_acceptance(
            self.database,
            manifest,
            self.output,
            pack_builder=lambda _database, question: {
                "question": question,
                "studyCandidates": [],
                "sourceExclusions": [],
                "answerEligibleEffects": [],
            },
            answer_builder=self._answer_builder,
            answer_validator=self._answer_validator,
        )

        self.assertEqual("BLOCKED_PENDING_INGEST", report["overallStatus"])
        source = report["questions"][0]["primarySources"][0]
        self.assertEqual("PENDING_INGEST", source["status"])
        self.assertEqual([], source["currentCanonicalAnalyses"])

    def test_ingested_but_unretrieved_primary_source_fails(self) -> None:
        self._insert_analysis(
            "retrieval-miss.xlsx",
            suffix=4,
            with_study=True,
        )
        manifest = self._manifest(
            workbooks=[("P04", "retrieval-miss.xlsx")],
            questions=[
                {
                    "id": "GQ03",
                    "question": "expected source terms",
                    "primaryPilotIds": ["P04"],
                    "requiredBehavior": ["Do not invent an effect."],
                }
            ],
        )

        report = acceptance.run_golden_question_acceptance(
            self.database,
            manifest,
            self.output,
            pack_builder=lambda _database, question: {
                "question": question,
                "studyCandidates": [],
                "sourceExclusions": [],
                "answerEligibleEffects": [],
            },
            answer_builder=self._answer_builder,
            answer_validator=self._answer_validator,
        )

        self.assertEqual("FAIL", report["overallStatus"])
        question = report["questions"][0]
        self.assertEqual("FAIL", question["status"])
        self.assertEqual(
            "RETRIEVAL_MISS",
            question["primarySources"][0]["status"],
        )
        self.assertEqual("PASS", question["answerValidation"]["status"])

    def test_answer_validation_failure_is_an_acceptance_failure(self) -> None:
        source_path = self._insert_analysis(
            "validation.xlsx",
            suffix=5,
            with_study=True,
        )
        manifest = self._manifest(
            workbooks=[("P05", "validation.xlsx")],
            questions=[
                {
                    "id": "GQ04",
                    "question": "validation question",
                    "primaryPilotIds": ["P05"],
                    "requiredBehavior": [],
                }
            ],
        )

        def failing_validator(_answer: dict, _pack: dict) -> dict:
            raise ValueError("tampered answer")

        report = acceptance.run_golden_question_acceptance(
            self.database,
            manifest,
            self.output,
            pack_builder=lambda _database, question: {
                "question": question,
                "studyCandidates": [
                    {
                        "publicDataId": "DATA-5",
                        "source": {"sourcePath": source_path},
                        "analysis": {"publicAnalysisId": "ANALYSIS-5"},
                    }
                ],
                "sourceExclusions": [],
                "answerEligibleEffects": [],
            },
            answer_builder=self._answer_builder,
            answer_validator=failing_validator,
        )

        self.assertEqual("FAIL", report["overallStatus"])
        self.assertEqual(
            "FAIL",
            report["questions"][0]["answerValidation"]["status"],
        )


if __name__ == "__main__":
    unittest.main()
