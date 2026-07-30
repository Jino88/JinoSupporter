from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from inference_data_ai_table_first_history import (
    TableFirstHistoryError,
    build_history_answer,
    build_history_detail,
    build_history_index,
    build_history_pack,
    run_history_acceptance,
    validate_history_answer,
)


def _write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


class TableFirstHistoryTests(unittest.TestCase):
    def _batch(self, root: Path, *, status: str = "ok") -> Path:
        batch = root / "batch"
        items = []
        fixtures = [
            (
                "capture_revision_bonding",
                "009.MSU Report bonding VP+CD 2025.06.05.xlsx",
                "VP+CD bonding condition and Function NG rate history.",
                {
                    "studyGroup": "VP+CD bonding Function NG Rate",
                    "titles": ["Function NG Rate"],
                    "tableTypes": ["COMPARISON"],
                    "groups": [
                        {"label": "Bond 1.5mg", "role": "TEST", "basis": "source"},
                        {"label": "Normal", "role": "REFERENCE", "basis": "source"},
                    ],
                    "metrics": [
                        {"name": "Total NG Rate", "unit": "%", "axisRefs": ["c1"]}
                    ],
                    "comparisonRelations": [
                        {
                            "leftGroup": "Bond 1.5mg",
                            "rightGroup": "Normal",
                            "basis": "source-authored comparison",
                        }
                    ],
                    "deterministicNumericFacts": [],
                    "deterministicNumericSeries": [],
                    "limitations": ["review required"],
                    "verificationStatus": "NEEDS_REVIEW",
                    "evidence": [
                        {"tableId": "table_bond", "sheet": "Result", "range": "B2:F5"}
                    ],
                },
            ),
            (
                "capture_revision_gauss",
                "BRS 161014S08ZZ Gauss supplier result 2025.07.01.xlsx",
                "Magnetic flux density supplier result.",
                {
                    "studyGroup": "Gauss supplier comparison",
                    "titles": ["Gauss"],
                    "tableTypes": ["COMPARISON"],
                    "groups": [
                        {"label": "A", "role": "TEST", "basis": "source"},
                        {"label": "B", "role": "REFERENCE", "basis": "source"},
                    ],
                    "metrics": [{"name": "Gauss", "unit": "G", "axisRefs": ["c2"]}],
                    "comparisonRelations": [],
                    "deterministicNumericFacts": [],
                    "deterministicNumericSeries": [],
                    "limitations": [],
                    "verificationStatus": "NEEDS_REVIEW",
                    "evidence": [
                        {"tableId": "table_gauss", "sheet": "Data", "range": "A1:C3"}
                    ],
                },
            ),
        ]
        aliases = {
            "status": "LOADED",
            "aliasGroups": [
                {
                    "canonicalTerm": "VP-CD",
                    "normalizedName": "VP-CD Assembly",
                    "terms": ["VP+CD", "VP-CD", "VP/CD"],
                },
                {
                    "canonicalTerm": "Bond",
                    "normalizedName": "Bonding",
                    "terms": ["Bond", "Bonding"],
                },
                {
                    "canonicalTerm": "NG",
                    "normalizedName": "NG",
                    "terms": ["NG"],
                },
            ],
        }
        for index, (stem, file_name, summary, study) in enumerate(fixtures, start=1):
            request_id = f"request_{index}"
            source = {
                "revisionUid": stem,
                "contentSha256": f"sha{index}",
                "fileName": file_name,
                "sourcePath": str(root / file_name),
            }
            table_id = study["evidence"][0]["tableId"]
            request = {
                "schemaVersion": "table-first-request-v1",
                "requestId": request_id,
                "source": source,
                "codeOwnedTermDictionary": aliases,
                "tables": [
                    {
                        "tableId": table_id,
                        "sheet": study["evidence"][0]["sheet"],
                        "range": study["evidence"][0]["range"],
                        "bounds": {
                            "minRow": 2,
                            "minColumn": 2,
                            "maxRow": 5,
                            "maxColumn": 6,
                        },
                        "previewRows": [
                            {
                                "row": 2,
                                "cells": [
                                    {"coordinate": "B2", "kind": "TEXT", "value": "Test"},
                                    {"coordinate": "C2", "kind": "TEXT", "value": "NG Rate"},
                                ],
                            }
                        ],
                    }
                ],
                "textBlocks": [],
            }
            analysis = {
                "schemaVersion": "table-first-analysis-v1",
                "requestId": request_id,
                "revisionUid": stem,
                "status": "NEEDS_REVIEW",
                "workbookSummary": summary,
                "tables": [],
            }
            projection = {
                "schemaVersion": "table-first-projection-v1",
                "requestId": request_id,
                "source": source,
                "analysisStatus": "NEEDS_REVIEW",
                "verificationStatus": "NEEDS_REVIEW",
                "queryEligibility": "NOT_ELIGIBLE_UNTIL_CANONICAL_REVIEW",
                "studies": [study],
                "textBlocks": [],
            }
            for kind, value in (
                ("requests", request),
                ("analyses", analysis),
                ("projections", projection),
            ):
                _write(batch / kind / f"{stem}.json", value)
            items.append(
                {
                    "index": index,
                    "fileName": file_name,
                    "request": str(batch / "requests" / f"{stem}.json"),
                    "analysis": str(batch / "analyses" / f"{stem}.json"),
                    "projection": str(batch / "projections" / f"{stem}.json"),
                }
            )
        _write(
            batch / "batch-report.json",
            {
                "schemaVersion": "table-first-batch-report-v1",
                "status": status,
                "builderVersion": "table-first-builder-v7",
                "promptVersion": "table-first-analysis-prompt-v4",
                "items": items,
            },
        )
        return batch

    def test_build_query_answer_and_detail(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            batch = self._batch(root)
            database = root / "history.sqlite"
            report = build_history_index(batch, database)
            self.assertEqual(report["workbookCount"], 2)
            self.assertEqual(report["studyCount"], 2)
            self.assertEqual(report["evidenceCount"], 2)

            pack = build_history_pack(
                database,
                "VP+CD 본딩 FUNCTION NG율 이력을 설명해줘",
                limit=10,
            )
            self.assertEqual(pack["summary"]["relevantWorkbookCount"], 1)
            self.assertEqual(pack["studies"][0]["date"], "2025-06-05")
            self.assertIn("VP+CD", pack["studies"][0]["fileName"])
            self.assertEqual(pack["summary"]["eligibleEffectCount"], 0)

            supplier_pack = build_history_pack(
                database,
                "Gauss 공급처",
                limit=10,
            )
            self.assertIn(
                "Gauss supplier",
                supplier_pack["studies"][0]["fileName"],
            )
            identifier_pack = build_history_pack(
                database,
                "BRS-161014S08ZZ",
                limit=10,
            )
            self.assertIn(
                "BRS 161014S08ZZ",
                identifier_pack["studies"][0]["fileName"],
            )

            answer = build_history_answer(pack)
            validate_history_answer(answer, pack)
            self.assertEqual(answer["answerStatus"], "REVIEW_GATED_HISTORY_FOUND")
            self.assertIn("승인된 효과", answer["markdown"])
            evidence_id = answer["citations"][0]["evidenceId"]
            self.assertTrue(evidence_id.startswith("TF-EVD-"))

            detail = build_history_detail(database, evidence_id)
            self.assertEqual(detail["evidence"]["sheet"], "Result")
            self.assertEqual(detail["preview"]["capturedCellCountInRange"], 2)
            self.assertFalse(detail["trust"]["trusted"])

    def test_answer_validation_rejects_changed_citations(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = root / "history.sqlite"
            build_history_index(self._batch(root), database)
            pack = build_history_pack(database, "Gauss supplier", limit=5)
            answer = build_history_answer(pack)
            answer["citations"] = []
            with self.assertRaises(TableFirstHistoryError):
                validate_history_answer(answer, pack)

    def test_no_table_source_is_returned_as_an_exclusion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            batch = self._batch(root)
            stem = "capture_revision_no_tables"
            file_name = "VP+CD GMI DT line NG image report.xlsx"
            source = {
                "revisionUid": stem,
                "contentSha256": "sha-no-tables",
                "fileName": file_name,
                "sourcePath": str(root / file_name),
            }
            request = {
                "schemaVersion": "table-first-request-v1",
                "requestId": "request_no_tables",
                "source": source,
                "codeOwnedTermDictionary": {"status": "NOT_FOUND", "aliasGroups": []},
                "tables": [],
                "textBlocks": [],
            }
            analysis = {
                "schemaVersion": "table-first-analysis-v1",
                "requestId": "request_no_tables",
                "revisionUid": stem,
                "status": "NO_TABLES",
                "workbookSummary": "No captured tabular evidence; images are out of scope.",
                "tables": [],
            }
            projection = {
                "schemaVersion": "table-first-projection-v1",
                "requestId": "request_no_tables",
                "source": source,
                "analysisStatus": "NO_TABLES",
                "verificationStatus": "EXCLUDED",
                "queryEligibility": "NOT_ELIGIBLE",
                "studies": [],
                "textBlocks": [],
            }
            for kind, value in (
                ("requests", request),
                ("analyses", analysis),
                ("projections", projection),
            ):
                _write(batch / kind / f"{stem}.json", value)
            report_path = batch / "batch-report.json"
            report = json.loads(report_path.read_text(encoding="utf-8"))
            report["items"].append(
                {
                    "index": 3,
                    "fileName": file_name,
                    "request": str(batch / "requests" / f"{stem}.json"),
                    "analysis": str(batch / "analyses" / f"{stem}.json"),
                    "projection": str(batch / "projections" / f"{stem}.json"),
                }
            )
            _write(report_path, report)
            database = root / "history.sqlite"
            index = build_history_index(batch, database)
            self.assertEqual(index["workbookCount"], 3)

            pack = build_history_pack(
                database,
                "VP+CD GMI DT line NG 이미지 자료",
                limit=10,
            )
            self.assertEqual(len(pack["sourceExclusions"]), 1)
            self.assertEqual(
                pack["sourceExclusions"][0]["analysisStatus"],
                "NO_TABLES",
            )
            answer = build_history_answer(pack)
            self.assertIn("표 기반 분석에서 제외", answer["markdown"])
            self.assertEqual(answer["coverage"]["eligibleEffectCount"], 0)

    def test_incomplete_batch_is_rejected_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with self.assertRaises(TableFirstHistoryError):
                build_history_index(
                    self._batch(root, status="running"),
                    root / "history.sqlite",
                )

    def test_acceptance_requires_primary_source_retrieval(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = root / "history.sqlite"
            build_history_index(self._batch(root), database)
            manifest = root / "acceptance.json"
            _write(
                manifest,
                {
                    "workbooks": [
                        {
                            "id": "P01",
                            "relativePath": (
                                "009.MSU Report bonding VP+CD 2025.06.05.xlsx"
                            ),
                        }
                    ],
                    "goldenQuestions": [
                        {
                            "id": "GQ01",
                            "question": "VP+CD bonding Function NG rate",
                            "primaryPilotIds": ["P01"],
                        }
                    ],
                },
            )
            report = run_history_acceptance(database, manifest, query_limit=10)
            self.assertEqual(report["status"], "PASS")
            self.assertEqual(report["summary"]["passed"], 1)
            self.assertEqual(
                report["questions"][0]["retrievedPrimaryFiles"],
                ["009.MSU Report bonding VP+CD 2025.06.05.xlsx"],
            )


if __name__ == "__main__":
    unittest.main()
