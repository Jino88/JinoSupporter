from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from inference_data_ai_table_recipe_proposal import (
    _semantic_header_sha256,
    semantic_header_signature,
)
from inference_data_ai_terminal_review_queue import (
    NON_QUANTITATIVE_REVIEW,
    NON_QUANTITATIVE_TABLE,
    ONE_OFF_REVIEW,
    ONE_OFF_TABLE,
    REGISTERED_CONFLICT,
    REPEATED_REVIEW,
    REPEATED_TABLE,
    build_terminal_review_queue,
)


def _captured_table(table_id: str, header: str) -> dict:
    return {
        "tableId": table_id,
        "bounds": {
            "minRow": 1,
            "minColumn": 1,
            "maxRow": 2,
            "maxColumn": 2,
        },
        "numericColumns": [
            {
                "column": "B",
                "columnRole": "MEASURE_VALUE",
                "headerTexts": [header],
            }
        ],
    }


class TerminalReviewQueueTests(unittest.TestCase):
    def _fixture(
        self,
        root: Path,
        *,
        repeated_outcome: str,
    ) -> dict:
        repeated = _captured_table("table-repeated", "Hearing NG")
        supporting = _captured_table("table-supporting", "")
        one_off = _captured_table("table-one-off", "Q'ty")
        first_request = root / "request-first.json"
        second_request = root / "request-second.json"
        first_request.write_text(
            json.dumps({"tables": [repeated, supporting]}),
            encoding="utf-8",
        )
        second_request.write_text(
            json.dumps({"tables": [one_off]}),
            encoding="utf-8",
        )
        repeated_semantic_sha256 = _semantic_header_sha256(
            semantic_header_signature(repeated)
        )
        canonical_audit = {
            "summary": {
                "actionCounts": {
                    "IMPORT_NEEDS_REVIEW_TERMINAL": 2,
                }
            },
            "actions": [
                {
                    "action": "IMPORT_NEEDS_REVIEW_TERMINAL",
                    "reason": "No replay-verified recipe result.",
                    "source": {
                        "captureRevisionId": 1,
                        "revisionUid": "capture-first",
                        "contentSha256": "a" * 64,
                        "fileName": "MSM-X526TOP report.xlsx",
                        "sourcePath": "first.xlsx",
                    },
                },
                {
                    "action": "IMPORT_NEEDS_REVIEW_TERMINAL",
                    "reason": "No replay-verified recipe result.",
                    "source": {
                        "captureRevisionId": 2,
                        "revisionUid": "capture-second",
                        "contentSha256": "b" * 64,
                        "fileName": "MSM-X626BOTTOM report.xlsx",
                        "sourcePath": "second.xlsx",
                    },
                },
            ],
        }
        table_match_report = {
            "workbooks": [
                {
                    "captureRevisionId": 1,
                    "fileName": "MSM-X526TOP report.xlsx",
                    "requestPath": str(first_request),
                    "tableCount": 2,
                    "tables": [
                        {
                            "tableId": "table-repeated",
                            "fingerprintSha256": "1" * 64,
                            "sheet": "Result",
                            "range": "A1:B2",
                            "numericCellCount": 2,
                        },
                        {
                            "tableId": "table-supporting",
                            "fingerprintSha256": "2" * 64,
                            "sheet": "Note",
                            "range": "A1:B2",
                            "numericCellCount": 0,
                        },
                    ],
                },
                {
                    "captureRevisionId": 2,
                    "fileName": "MSM-X626BOTTOM report.xlsx",
                    "requestPath": str(second_request),
                    "tableCount": 1,
                    "tables": [
                        {
                            "tableId": "table-one-off",
                            "fingerprintSha256": "3" * 64,
                            "sheet": "Result",
                            "range": "A1:B2",
                            "numericCellCount": 1,
                        }
                    ],
                },
            ]
        }
        table_structure_catalog = {
            "structures": [
                {
                    "tableStructureId": "table-structure-repeated",
                    "fingerprintSha256": "1" * 64,
                    "quantitative": True,
                },
                {
                    "tableStructureId": "table-structure-supporting",
                    "fingerprintSha256": "2" * 64,
                    "quantitative": False,
                },
                {
                    "tableStructureId": "table-structure-one-off",
                    "fingerprintSha256": "3" * 64,
                    "quantitative": True,
                },
            ]
        }
        priority_report = {
            "queue": [
                {
                    "tableStructureId": "table-structure-repeated",
                    "baseTableStructureId": "table-structure-repeated",
                    "semanticHeaderSha256": repeated_semantic_sha256,
                    "rank": 1,
                    "priorityScore": 0.9,
                    "safetyReasons": [],
                }
            ]
        }
        completion_state = {
            "status": "COMPLETED",
            "summary": {"unresolvedStructureCount": 0},
            "outcomes": {
                "table-structure-repeated": {
                    "status": repeated_outcome,
                }
            },
        }
        return {
            "canonical_audit": canonical_audit,
            "table_match_report": table_match_report,
            "table_structure_catalog": table_structure_catalog,
            "priority_report": priority_report,
            "completion_state": completion_state,
        }

    def test_terminal_workbooks_are_partitioned_without_ai(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            report = build_terminal_review_queue(
                **self._fixture(
                    Path(temporary),
                    repeated_outcome=(
                        "QUARANTINED_RECIPE_CONTRACT_FAILURE"
                    ),
                ),
                generated_at="2026-07-28T00:00:00+00:00",
            )

        self.assertEqual("READY_FOR_HUMAN_REVIEW", report["status"])
        self.assertTrue(all(report["invariants"].values()))
        self.assertEqual(2, report["summary"]["terminalWorkbookCount"])
        self.assertEqual(3, report["summary"]["terminalTableCount"])
        self.assertEqual(
            {
                ONE_OFF_REVIEW: 1,
                REPEATED_REVIEW: 1,
            },
            report["summary"]["reviewClassCounts"],
        )
        self.assertEqual(
            1,
            report["summary"]["tableClassCounts"][REPEATED_TABLE],
        )
        self.assertEqual(
            1,
            report["summary"]["tableClassCounts"][ONE_OFF_TABLE],
        )
        self.assertEqual(
            "QUARANTINED_RECIPE_CONTRACT_FAILURE",
            report["repeatedStructureGroups"][0]["outcomeStatus"],
        )
        self.assertEqual(
            {"X526": 1, "X626B": 1},
            report["summary"]["modelFamilyCounts"],
        )
        self.assertEqual(1, report["items"][0]["reviewPriority"])
        self.assertEqual(0, report["policy"]["aiCalls"])
        self.assertFalse(report["policy"]["numericValuesRead"])

    def test_registered_recipe_in_terminal_queue_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            report = build_terminal_review_queue(
                **self._fixture(
                    Path(temporary),
                    repeated_outcome="REGISTERED_AI_REPLAY",
                )
            )

        self.assertEqual("INCOMPLETE", report["status"])
        self.assertFalse(
            report["invariants"]["noRegisteredRecipeConflict"]
        )
        self.assertEqual(
            1,
            report["summary"]["tableClassCounts"][REGISTERED_CONFLICT],
        )

    def test_source_owned_non_metric_quarantine_is_supporting(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            fixture = self._fixture(
                Path(temporary),
                repeated_outcome="QUARANTINED_AI_FAILURE_NO_RETRY",
            )
            report = build_terminal_review_queue(
                **fixture,
                source_owned_decisions=[
                    {
                        "targetTableStructureId": (
                            "table-structure-repeated"
                        ),
                        "decision": "QUARANTINE",
                        "semanticContract": {"metricColumns": []},
                    }
                ],
            )

        first = next(
            item
            for item in report["items"]
            if item["captureRevisionId"] == 1
        )
        repeated = next(
            table
            for table in first["tables"]
            if table["tableId"] == "table-repeated"
        )
        self.assertEqual("READY_FOR_HUMAN_REVIEW", report["status"])
        self.assertEqual(NON_QUANTITATIVE_REVIEW, first["reviewClass"])
        self.assertEqual(7, first["reviewPriority"])
        self.assertEqual(NON_QUANTITATIVE_TABLE, repeated["classification"])
        self.assertEqual(
            "SOURCE_OWNED_NON_METRIC_QUARANTINE",
            repeated["classificationBasis"],
        )
        self.assertEqual(0, len(report["repeatedStructureGroups"]))
        self.assertEqual(
            1,
            report["summary"]["sourceOwnedNonMetricTableCount"],
        )

    def test_date_only_administrative_table_is_supporting(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            fixture = self._fixture(
                Path(temporary),
                repeated_outcome="QUARANTINED_AI_FAILURE_NO_RETRY",
            )
            request_path = Path(
                fixture["table_match_report"]["workbooks"][0][
                    "requestPath"
                ]
            )
            request = json.loads(request_path.read_text(encoding="utf-8"))
            repeated = request["tables"][0]
            repeated["numericColumns"][0]["headerTexts"] = ["Maker"]
            repeated["numericColumns"][0]["numberFormats"] = ["dd-mmm"]
            repeated["rowLabels"] = [
                {"relativeRow": 0, "labels": ["Finish Date", "Checker"]}
            ]
            request_path.write_text(json.dumps(request), encoding="utf-8")
            fixture["priority_report"]["queue"][0][
                "semanticHeaderSha256"
            ] = _semantic_header_sha256(
                semantic_header_signature(repeated)
            )
            report = build_terminal_review_queue(**fixture)

        first = next(
            item
            for item in report["items"]
            if item["captureRevisionId"] == 1
        )
        repeated_item = next(
            table
            for table in first["tables"]
            if table["tableId"] == "table-repeated"
        )
        self.assertEqual(NON_QUANTITATIVE_REVIEW, first["reviewClass"])
        self.assertEqual(
            "DATE_ONLY_METADATA",
            repeated_item["classificationBasis"],
        )
        self.assertEqual(
            1,
            report["summary"]["dateOnlyMetadataTableCount"],
        )


if __name__ == "__main__":
    unittest.main()
