from __future__ import annotations

import json
import sqlite3
import subprocess
import tempfile
import unittest
from pathlib import Path

from inference_data_ai_relevance_query import (
    AI_SCHEMA_VERSION,
    PROMPT_VERSION,
    REQUEST_SCHEMA_VERSION,
    RESULT_SCHEMA_VERSION,
    RelevanceQueryError,
    build_relevance_prompt,
    build_relevance_result,
    run_codex_relevance_query,
    validate_relevance_ai_response,
)


class RelevanceQueryTests(unittest.TestCase):
    @staticmethod
    def _request() -> dict:
        return {
            "schemaVersion": REQUEST_SCHEMA_VERSION,
            "promptVersion": PROMPT_VERSION,
            "question": "VP+CD 조립과 히어링 불량 관련 보고서",
            "database": "history.sqlite",
            "retrieval": {
                "candidateStudyCount": 3,
                "candidateWorkbookCount": 3,
                "indexedStudyCount": 3710,
                "candidateLimit": 200,
                "warning": "후보",
            },
            "candidates": [
                {
                    "retrievalRank": 1,
                    "studyId": "TF-STU-1",
                    "workbookId": "TF-WBK-1",
                    "date": "2025-01-01",
                    "fileName": "vp-cd-bonding.xlsx",
                    "sourcePath": "D:/source/vp-cd-bonding.xlsx",
                    "workbookSummary": "VP+CD bonding and hearing result report",
                    "studyGroup": "VP+CD bonding",
                    "titles": ["Function result"],
                    "groups": [{"label": "VP+CD bonding", "role": "TEST"}],
                    "metrics": [{"name": "Hearing", "unit": "%"}],
                    "comparisonRelations": [],
                    "limitations": [],
                    "matchedQueryTerms": ["vp", "cd", "hearing"],
                    "retrievalScore": 100,
                    "verificationStatus": "NEEDS_REVIEW",
                    "evidenceIds": ["TF-EVD-1", "TF-EVD-2"],
                },
                {
                    "retrievalRank": 2,
                    "studyId": "TF-STU-2",
                    "workbookId": "TF-WBK-2",
                    "date": "2025-02-01",
                    "fileName": "vp-dimension.xlsx",
                    "sourcePath": "D:/source/vp-dimension.xlsx",
                    "workbookSummary": "VP dimension only",
                    "studyGroup": "VP dimension",
                    "titles": ["Dimension"],
                    "groups": [{"label": "VP", "role": "TEST"}],
                    "metrics": [{"name": "Dimension", "unit": "mm"}],
                    "comparisonRelations": [],
                    "limitations": [],
                    "matchedQueryTerms": ["vp"],
                    "retrievalScore": 50,
                    "verificationStatus": "NEEDS_REVIEW",
                    "evidenceIds": ["TF-EVD-3"],
                },
                {
                    "retrievalRank": 3,
                    "studyId": "TF-STU-3",
                    "workbookId": "TF-WBK-3",
                    "date": "2025-03-01",
                    "fileName": "cd-separation.xlsx",
                    "sourcePath": "D:/source/cd-separation.xlsx",
                    "workbookSummary": "VP+CD separation and hearing check",
                    "studyGroup": "VP+CD separation",
                    "titles": ["Hearing check"],
                    "groups": [{"label": "VP+CD separate", "role": "TEST"}],
                    "metrics": [{"name": "Hearing", "unit": "count"}],
                    "comparisonRelations": [],
                    "limitations": [],
                    "matchedQueryTerms": ["cd", "hearing"],
                    "retrievalScore": 45,
                    "verificationStatus": "NEEDS_REVIEW",
                    "evidenceIds": ["TF-EVD-4"],
                },
            ],
            "evidenceRegistry": [
                {
                    "evidenceId": f"TF-EVD-{index}",
                    "studyId": "TF-STU-1" if index < 3 else f"TF-STU-{index - 1}",
                    "sourcePath": f"D:/source/{index}.xlsx",
                    "sheet": "Result",
                    "range": f"A{index}:D{index + 3}",
                    "tableId": f"table-{index}",
                    "verificationStatus": "NEEDS_REVIEW",
                }
                for index in range(1, 5)
            ],
            "requestSha256": "abc123",
        }

    @staticmethod
    def _response(request: dict) -> dict:
        return {
            "schemaVersion": AI_SCHEMA_VERSION,
            "promptVersion": PROMPT_VERSION,
            "question": request["question"],
            "queryInterpretation": {
                "documentNeed": "VP+CD 조립 조건과 히어링 지표를 함께 기록한 보고서",
                "subjects": ["VP", "CD"],
                "conditions": ["조립", "본딩", "분리"],
                "metrics": ["Hearing"],
            },
            "selectedStudies": [
                {
                    "studyId": "TF-STU-1",
                    "relevanceReason": "VP+CD 본딩 조건과 Hearing 지표를 같은 Study에 포함합니다.",
                    "matchedAspects": ["VP+CD 본딩", "Hearing"],
                },
                {
                    "studyId": "TF-STU-3",
                    "relevanceReason": "VP+CD 분리 조건과 Hearing 확인 항목을 포함합니다.",
                    "matchedAspects": ["VP+CD 분리", "Hearing"],
                },
            ],
        }

    def test_prompt_limits_ai_to_document_relevance(self) -> None:
        prompt = build_relevance_prompt(self._request())
        self.assertIn("오직 문서 관련성 판정", prompt)
        self.assertIn("효과가 있다/없다를 판단하지 마십시오", prompt)
        self.assertIn("사용자 질문에 답하지 마십시오", prompt)

    def test_result_contains_selected_documents_without_result_judgment(self) -> None:
        request = self._request()
        result = build_relevance_result(self._response(request), request)
        self.assertEqual(RESULT_SCHEMA_VERSION, result["schemaVersion"])
        self.assertEqual(2, result["coverage"]["relevantStudyCount"])
        self.assertEqual(3, result["coverage"]["citationCount"])
        self.assertEqual(
            ["TF-STU-1", "TF-STU-3"],
            [item["studyId"] for item in result["studies"]],
        )
        for forbidden in ("directAnswer", "findings", "trendRows", "facts"):
            self.assertNotIn(forbidden, result)

    def test_validation_rejects_unknown_or_duplicate_study(self) -> None:
        request = self._request()
        unknown = self._response(request)
        unknown["selectedStudies"][0]["studyId"] = "TF-STU-NOT-FOUND"
        with self.assertRaises(RelevanceQueryError):
            validate_relevance_ai_response(unknown, request)
        duplicate = self._response(request)
        duplicate["selectedStudies"][1]["studyId"] = "TF-STU-1"
        with self.assertRaises(RelevanceQueryError):
            validate_relevance_ai_response(duplicate, request)

    def test_raw_source_values_are_attached_without_ai_interpretation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            captured_request = root / "captured-request.json"
            captured_request.write_text(
                json.dumps(
                    {
                        "tables": [
                            {
                                "tableId": "table-1",
                                "sheet": "Result",
                                "range": "A1:D2",
                                "numericColumns": [
                                    {
                                        "column": "B",
                                        "columnId": "table-1_col_B",
                                        "columnRole": "IDENTIFIER_OR_BASIS",
                                        "headerTexts": ["Input"],
                                        "displaySamples": [
                                            {
                                                "coordinate": "B2",
                                                "sourceDisplay": "100",
                                            }
                                        ],
                                    },
                                    {
                                        "column": "C",
                                        "columnId": "table-1_col_C",
                                        "columnRole": "MEASURE_VALUE",
                                        "headerTexts": ["Total NG"],
                                        "displaySamples": [
                                            {
                                                "coordinate": "C2",
                                                "sourceDisplay": "5",
                                            }
                                        ],
                                    },
                                    {
                                        "column": "D",
                                        "columnId": "table-1_col_D",
                                        "columnRole": "MEASURE_VALUE",
                                        "headerTexts": ["NG rate"],
                                        "displaySamples": [],
                                    },
                                ],
                                "previewRows": [
                                    {
                                        "row": 1,
                                        "cells": [
                                            {"coordinate": "A1", "kind": "TEXT", "value": "Condition"},
                                            {"coordinate": "B1", "kind": "TEXT", "value": "Input"},
                                            {"coordinate": "C1", "kind": "TEXT", "value": "Total NG"},
                                            {"coordinate": "D1", "kind": "TEXT", "value": "NG rate"},
                                        ],
                                    },
                                    {
                                        "row": 2,
                                        "cells": [
                                            {"coordinate": "A2", "kind": "TEXT", "value": "Test"},
                                            {"coordinate": "B2", "kind": "NUMBER", "value": "100"},
                                            {"coordinate": "C2", "kind": "NUMBER", "value": "5"},
                                            {
                                                "coordinate": "D2",
                                                "kind": "NUMBER",
                                                "value": "0.05",
                                                "numberFormat": "0.0%",
                                            },
                                        ],
                                    },
                                ],
                            }
                        ]
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            database = root / "history.sqlite"
            connection = sqlite3.connect(database)
            connection.executescript(
                """
                CREATE TABLE history_workbooks (
                    workbook_id INTEGER PRIMARY KEY,
                    request_path TEXT NOT NULL
                );
                CREATE TABLE history_studies (
                    workbook_id INTEGER NOT NULL,
                    public_study_id TEXT NOT NULL,
                    payload_json TEXT NOT NULL
                );
                """
            )
            payload = json.dumps(
                {
                    "metrics": [
                        {
                            "name": "Total NG",
                            "unit": "count",
                            "axisRefs": ["table-1_col_C"],
                        },
                        {
                            "name": "NG rate",
                            "unit": "%",
                            "axisRefs": ["table-1_col_D"],
                        },
                    ]
                }
            )
            connection.execute(
                "INSERT INTO history_workbooks VALUES (?, ?)",
                (1, str(captured_request)),
            )
            connection.execute(
                "INSERT INTO history_studies VALUES (?, ?, ?)",
                (1, "TF-STU-1", payload),
            )
            connection.commit()
            connection.close()
            request = self._request()
            request["database"] = str(database)
            result = build_relevance_result(self._response(request), request)
            points = result["studies"][0]["rawDataPoints"]
            self.assertEqual(
                {"100", "5", "5.0%"},
                {point["displayValue"] for point in points},
            )
            self.assertTrue(all(point["condition"] == "A2=Test" for point in points))
            self.assertEqual(3, result["coverage"]["rawDataPointCount"])
            self.assertNotIn("directAnswer", result)

    def test_runner_uses_read_only_codex_and_writes_result(self) -> None:
        request = self._request()
        response = self._response(request)
        observed: dict[str, object] = {}

        def fake_run(command: list[str], **kwargs: object):
            observed["command"] = command
            observed["input"] = kwargs["input"]
            output_path = Path(command[command.index("--output-last-message") + 1])
            output_path.write_text(
                json.dumps(response, ensure_ascii=False), encoding="utf-8"
            )
            return subprocess.CompletedProcess(command, 0, "", "")

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "answer.json"
            result = run_codex_relevance_query(
                request=request,
                output_path=output_path,
                codex_command=["codex-test"],
                run_command=fake_run,
            )
            self.assertTrue(output_path.is_file())
            self.assertEqual(2, result["coverage"]["relevantStudyCount"])
            self.assertIn("--sandbox", observed["command"])
            self.assertIn("read-only", observed["command"])
            self.assertNotIn("directAnswer", str(observed["input"]))


if __name__ == "__main__":
    unittest.main()
