from __future__ import annotations

import copy
import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from inference_data_ai_contextual_query import (
    CONTEXT_AI_SCHEMA_VERSION,
    CONTEXT_PROMPT_VERSION,
    ContextualQueryError,
    build_contextual_prompt,
    build_contextual_query_request,
    finalize_contextual_answer,
    run_codex_contextual_query,
    validate_contextual_ai_response,
)
from inference_data_ai_table_first_history import build_history_index


def _write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


class ContextualQueryTests(unittest.TestCase):
    def _batch(self, root: Path) -> Path:
        batch = root / "batch"
        items = []
        fixtures = [
            (
                "revision_january",
                "VP CD Assy Hearing NG rate 2025.01.10.xlsx",
                "VP+CD 조립 조건의 Hearing 불량률 결과.",
                "VP+CD assembly Hearing NG rate",
                "2025-01-10",
                0.10,
                "10.00%",
                True,
            ),
            (
                "revision_february",
                "VP CD Assy Hearing NG rate 2025.02.10.xlsx",
                "VP+CD 조립 조건의 Hearing 불량률 후속 결과.",
                "VP+CD assembly Hearing NG rate",
                "2025-02-10",
                0.08,
                "8.00%",
                True,
            ),
            (
                "revision_distractor",
                "VP dimension and CD memo Hearing 2025.03.10.xlsx",
                "VP 치수 결과와 별도 CD 메모 및 Hearing 불량률 참고 자료.",
                "VP dimension investigation",
                "2025-03-10",
                0.25,
                "25.00%",
                False,
            ),
        ]
        aliases = {
            "status": "LOADED",
            "aliasGroups": [
                {
                    "canonicalTerm": "VP-CD",
                    "normalizedName": "VP-CD Assembly",
                    "terms": ["VP+CD", "VP CD", "VP-CD", "조립"],
                },
                {
                    "canonicalTerm": "Hearing NG rate",
                    "normalizedName": "Hearing NG rate",
                    "terms": ["Hearing", "NG rate", "불량률"],
                },
            ],
        }
        for index, fixture in enumerate(fixtures, start=1):
            (
                revision,
                file_name,
                summary,
                study_group,
                _date,
                raw_rate,
                display_rate,
                direct_relation,
            ) = fixture
            request_id = f"context_request_{index}"
            table_id = f"table_context_{index}"
            source = {
                "revisionUid": revision,
                "contentSha256": f"context-sha-{index}",
                "fileName": file_name,
                "sourcePath": str(root / file_name),
            }
            request = {
                "schemaVersion": "table-first-request-v1",
                "requestId": request_id,
                "source": source,
                "codeOwnedTermDictionary": aliases,
                "tables": [
                    {
                        "tableId": table_id,
                        "sheet": "Result",
                        "range": "B2:E4",
                        "bounds": {
                            "minRow": 2,
                            "minColumn": 2,
                            "maxRow": 4,
                            "maxColumn": 5,
                        },
                        "previewRows": [
                            {
                                "row": 2,
                                "cells": [
                                    {
                                        "coordinate": "B2",
                                        "kind": "TEXT",
                                        "value": "Condition",
                                    },
                                    {
                                        "coordinate": "E2",
                                        "kind": "TEXT",
                                        "value": "Hearing NG rate",
                                    },
                                ],
                            },
                            {
                                "row": 3,
                                "cells": [
                                    {
                                        "coordinate": "B3",
                                        "kind": "TEXT",
                                        "value": (
                                            "VP+CD Assy"
                                            if direct_relation
                                            else "VP dimension"
                                        ),
                                    },
                                    {
                                        "coordinate": "E3",
                                        "kind": "NUMBER",
                                        "value": str(raw_rate),
                                    },
                                ],
                            },
                        ],
                        "numericColumns": [
                            {
                                "columnId": f"{table_id}_col_E",
                                "column": "E",
                                "columnRole": "MEASURE_VALUE",
                                "headerTexts": ["Hearing", "NG rate"],
                                "displaySamples": [
                                    {
                                        "coordinate": "E3",
                                        "rawNumber": raw_rate,
                                        "displayScale": "PERCENT",
                                        "normalizedDisplay": display_rate,
                                    }
                                ],
                                "numericCount": 1,
                                "min": raw_rate,
                                "max": raw_rate,
                                "average": raw_rate,
                                "sourceRange": "E3",
                            }
                        ],
                    }
                ],
                "textBlocks": [],
            }
            analysis = {
                "schemaVersion": "table-first-analysis-v1",
                "requestId": request_id,
                "revisionUid": revision,
                "status": "NEEDS_REVIEW",
                "workbookSummary": summary,
                "tables": [],
            }
            groups = [
                {
                    "label": "VP+CD Assy" if direct_relation else "VP dimension",
                    "role": "TEST",
                    "basis": "source row",
                }
            ]
            projection = {
                "schemaVersion": "table-first-projection-v1",
                "requestId": request_id,
                "source": source,
                "analysisStatus": "NEEDS_REVIEW",
                "verificationStatus": "NEEDS_REVIEW",
                "queryEligibility": "NOT_ELIGIBLE_UNTIL_CANONICAL_REVIEW",
                "studies": [
                    {
                        "studyGroup": study_group,
                        "titles": ["Hearing NG rate result"],
                        "tableTypes": ["COMPARISON"],
                        "groups": groups,
                        "metrics": [
                            {
                                "name": "Hearing NG rate",
                                "unit": "%",
                                "axisRefs": [f"{table_id}_col_E"],
                            }
                        ],
                        "comparisonRelations": (
                            [
                                {
                                    "leftGroup": "VP+CD Assy",
                                    "rightGroup": "Normal",
                                    "basis": "source comparison",
                                }
                            ]
                            if direct_relation
                            else []
                        ),
                        "deterministicNumericFacts": [],
                        "deterministicNumericSeries": [],
                        "limitations": ["review required"],
                        "verificationStatus": "NEEDS_REVIEW",
                        "evidence": [
                            {
                                "tableId": table_id,
                                "sheet": "Result",
                                "range": "B2:E4",
                            }
                        ],
                    }
                ],
                "textBlocks": [],
            }
            for kind, value in (
                ("requests", request),
                ("analyses", analysis),
                ("projections", projection),
            ):
                _write(batch / kind / f"{revision}.json", value)
            items.append(
                {
                    "index": index,
                    "fileName": file_name,
                    "request": str(batch / "requests" / f"{revision}.json"),
                    "analysis": str(batch / "analyses" / f"{revision}.json"),
                    "projection": str(batch / "projections" / f"{revision}.json"),
                }
            )
        _write(
            batch / "batch-report.json",
            {
                "schemaVersion": "table-first-batch-report-v1",
                "status": "ok",
                "builderVersion": "table-first-builder-v7",
                "promptVersion": "table-first-analysis-prompt-v4",
                "items": items,
            },
        )
        return batch

    def _request(self, root: Path) -> dict:
        database = root / "history.sqlite"
        build_history_index(self._batch(root), database)
        return build_contextual_query_request(
            database,
            "VP CD 조립에 따른 Hearing 불량률 추이",
            candidate_limit=10,
            detail_candidate_limit=10,
        )

    @staticmethod
    def _response(request: dict) -> dict:
        direct_candidates = [
            item
            for item in request["candidates"]
            if item["studyGroup"] == "VP+CD assembly Hearing NG rate"
        ]
        facts = [
            item
            for item in request["factRegistry"]
            if item["studyId"] in {value["studyId"] for value in direct_candidates}
        ]
        facts.sort(key=lambda item: str(item["date"]))
        return {
            "schemaVersion": CONTEXT_AI_SCHEMA_VERSION,
            "promptVersion": CONTEXT_PROMPT_VERSION,
            "question": request["question"],
            "intent": {
                "answerMode": "TREND",
                "subject": "VP+CD 조립",
                "conditions": ["VP+CD Assy"],
                "metrics": ["Hearing NG rate"],
                "comparison": "날짜별 비교",
                "timeScope": "2025-01-10~2025-02-10",
            },
            "relevanceAssessment": (
                "VP와 CD 단어가 따로 존재하는 치수 자료는 직접 관계가 없어 제외했습니다."
            ),
            "evidenceStatus": "ANSWERED",
            "confidence": "MEDIUM",
            "directAnswer": "비교 가능한 두 관측에서 Hearing 불량률이 낮아졌습니다.",
            "findings": [
                {
                    "statement": "두 날짜의 동일 지표가 직접 연결됩니다.",
                    "significance": "후속 관측값이 더 낮습니다.",
                    "evidenceIds": [item["evidenceId"] for item in facts],
                    "factIds": [item["factId"] for item in facts],
                }
            ],
            "trendRows": [
                {
                    "date": str(item["date"]),
                    "condition": "VP+CD Assy",
                    "metric": "Hearing NG rate",
                    "value": item["displayValue"],
                    "interpretation": "원본 표의 관측값",
                    "evidenceIds": [item["evidenceId"]],
                    "factIds": [item["factId"]],
                }
                for item in facts
            ],
            "limitations": ["NEEDS_REVIEW 자료입니다."],
            "usedStudyIds": [item["studyId"] for item in direct_candidates],
        }

    def test_request_exposes_candidates_but_registers_exact_numeric_facts(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            request = self._request(Path(directory))
            self.assertEqual(3, request["retrieval"]["candidateStudyCount"])
            self.assertEqual(
                {"10.00%", "8.00%", "25.00%"},
                {item["displayValue"] for item in request["factRegistry"]},
            )
            self.assertTrue(
                all(item["rowContext"] for item in request["factRegistry"])
            )
            prompt = build_contextual_prompt(request)
            self.assertIn("단어가 겹친다는 이유만으로", prompt)
            self.assertIn("최소 2개의 날짜/순서 관측", prompt)
            self.assertIn("모든 직접 관련 후보", prompt)
            self.assertIn("관련성 판정과 인과 충분성 판정을 분리", prompt)

    def test_korean_hearing_alias_retrieves_english_hearing_studies(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = root / "history.sqlite"
            build_history_index(self._batch(root), database)
            request = build_contextual_query_request(
                database,
                "VP CD 조립에 따른 히어링 불량률 추이",
                candidate_limit=10,
                detail_candidate_limit=10,
            )
            self.assertEqual(3, request["retrieval"]["candidateStudyCount"])

    def test_finalize_keeps_only_ai_selected_direct_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            request = self._request(Path(directory))
            answer = finalize_contextual_answer(self._response(request), request)
            self.assertEqual("CONTEXTUAL_AI_ANSWERED", answer["answerStatus"])
            self.assertEqual(2, answer["coverage"]["relevantStudyCount"])
            self.assertEqual(1, answer["coverage"]["excludedCandidateCount"])
            self.assertEqual(2, answer["coverage"]["citationCount"])
            self.assertEqual(2, answer["coverage"]["relatedEvidenceCount"])
            self.assertEqual(2, len(answer["relatedCitations"]))
            self.assertEqual(2, len(answer["trendRows"]))
            self.assertNotIn("25.00%", answer["markdown"])

    def test_finalize_separates_related_study_evidence_from_core_claims(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            request = self._request(Path(directory))
            response = self._response(request)
            response["usedStudyIds"] = [
                item["studyId"] for item in request["candidates"]
            ]
            answer = finalize_contextual_answer(response, request)
            self.assertEqual(3, answer["coverage"]["relevantStudyCount"])
            self.assertEqual(2, answer["coverage"]["citationCount"])
            self.assertEqual(3, answer["coverage"]["relatedEvidenceCount"])
            self.assertEqual(3, len(answer["relatedCitations"]))

    def test_validation_rejects_unregistered_or_changed_numeric_claim(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            request = self._request(Path(directory))
            response = self._response(request)
            changed = copy.deepcopy(response)
            changed["trendRows"][0]["value"] = "99.00%"
            with self.assertRaises(ContextualQueryError):
                validate_contextual_ai_response(changed, request)
            unknown = copy.deepcopy(response)
            unknown["trendRows"][0]["factIds"] = ["TF-FCT-NOT-REGISTERED"]
            with self.assertRaises(ContextualQueryError):
                validate_contextual_ai_response(unknown, request)
            invented = copy.deepcopy(response)
            invented["findings"][0]["statement"] += " 99.00%"
            with self.assertRaises(ContextualQueryError):
                validate_contextual_ai_response(invented, request)

    def test_answered_trend_requires_two_observations(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            request = self._request(Path(directory))
            response = self._response(request)
            response["trendRows"] = response["trendRows"][:1]
            with self.assertRaises(ContextualQueryError):
                validate_contextual_ai_response(response, request)

    def test_codex_runner_uses_schema_and_finalizes_response(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            request = self._request(root)
            response = self._response(request)
            observed: dict[str, object] = {}

            def fake_run(command: list[str], **kwargs: object):
                observed["command"] = command
                observed["input"] = kwargs["input"]
                output_path = Path(
                    command[command.index("--output-last-message") + 1]
                )
                _write(output_path, response)
                return subprocess.CompletedProcess(command, 0, "", "")

            output_path = root / "answer.json"
            answer = run_codex_contextual_query(
                request=request,
                output_path=output_path,
                codex_command=["codex-test"],
                run_command=fake_run,
            )
            self.assertTrue(output_path.is_file())
            self.assertIn("--output-schema", observed["command"])
            self.assertIn("--sandbox", observed["command"])
            self.assertIn("--skip-git-repo-check", observed["command"])
            self.assertEqual("CONTEXTUAL_AI_ANSWERED", answer["answerStatus"])

    def test_codex_runner_retries_rejected_numeric_claim(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            request = self._request(root)
            valid_response = self._response(request)
            rejected_response = copy.deepcopy(valid_response)
            rejected_response["findings"][0]["statement"] += " 99.00%"
            prompts: list[str] = []

            def fake_run(command: list[str], **kwargs: object):
                prompts.append(str(kwargs["input"]))
                output_path = Path(
                    command[command.index("--output-last-message") + 1]
                )
                _write(
                    output_path,
                    rejected_response if len(prompts) == 1 else valid_response,
                )
                return subprocess.CompletedProcess(command, 0, "", "")

            output_path = root / "answer.json"
            answer = run_codex_contextual_query(
                request=request,
                output_path=output_path,
                codex_command=["codex-test"],
                run_command=fake_run,
            )
            self.assertEqual(2, len(prompts))
            self.assertIn("이전 초안은 근거 결속 검증에서 거절", prompts[1])
            self.assertTrue(
                (root / "answer.ai-response.attempt-1.json").is_file()
            )
            self.assertEqual("CONTEXTUAL_AI_ANSWERED", answer["answerStatus"])

    def test_codex_runner_returns_safe_answer_after_repeated_rejection(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            request = self._request(root)
            rejected_response = self._response(request)
            rejected_response["findings"][0]["statement"] += " 99.00%"
            attempts = 0

            def fake_run(command: list[str], **_: object):
                nonlocal attempts
                attempts += 1
                output_path = Path(
                    command[command.index("--output-last-message") + 1]
                )
                _write(output_path, rejected_response)
                return subprocess.CompletedProcess(command, 0, "", "")

            answer = run_codex_contextual_query(
                request=request,
                output_path=root / "answer.json",
                codex_command=["codex-test"],
                run_command=fake_run,
            )
            self.assertEqual(2, attempts)
            self.assertEqual(
                "CONTEXTUAL_AI_INSUFFICIENT_EVIDENCE", answer["answerStatus"]
            )
            self.assertEqual("AI_RESPONSE_REJECTED", answer["guardrail"]["status"])
            self.assertEqual([], answer["citations"])
            self.assertEqual(2, answer["coverage"]["relevantStudyCount"])
            self.assertEqual(2, len(answer["relatedCitations"]))
            self.assertNotIn("99.00%", answer["directAnswer"])


if __name__ == "__main__":
    unittest.main()
