from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from inference_data_ai_cli import ensure_universal_schema
from inference_data_ai_schema import ensure_knowledge_schema
from inference_data_ai_source_ingest import (
    bridge_capture_to_canonical_source,
    ensure_capture_v2_schema,
    extract_workbook,
    import_capture,
)
from inference_data_ai_structure_canonical_import import (
    build_recipe_manifest,
    run_structure_canonical_import,
)


class StructureCanonicalImportTests(unittest.TestCase):
    def _replay_item(self, revision_uid: str) -> dict:
        return {
            "passed": True,
            "requestFile": f"{revision_uid}.json",
            "tableId": "table-one",
            "sheet": "Data",
            "range": "A1:B3",
            "extraction": {
                "tableId": "table-one",
                "tableStructureId": "structure-one",
                "recipeId": "recipe-one",
                "sheet": "Data",
                "range": "A1:B3",
                "semantic": {
                    "title": "Measured tension",
                    "tableType": "DESCRIPTIVE",
                },
                "deterministicNumericFacts": [
                    {
                        "columnId": "table-one_col_B",
                        "columnRole": "MEASURE_VALUE",
                        "name": "Tension",
                        "unit": "N",
                        "numericCount": 2,
                        "min": 10.0,
                        "max": 20.0,
                        "average": 15.0,
                        "sourceRange": "B2:B3",
                        "calculationAuthority": (
                            "CODE_FROM_CAPTURED_RAW_VALUES"
                        ),
                    }
                ],
                "deterministicCellFacts": [
                    {
                        "columnId": "table-one_col_B",
                        "coordinate": "B2",
                        "displayRole": "NUMBER",
                    },
                    {
                        "columnId": "table-one_col_B",
                        "coordinate": "B3",
                        "displayRole": "NUMBER",
                    },
                ],
            },
        }

    def test_recipe_manifest_is_valid_and_code_owned(self) -> None:
        manifest = build_recipe_manifest(
            source={
                "dataset": "Fixture",
                "sourcePath": "D:/fixture.xlsx",
                "revisionUid": "capture_revision_fixture",
                "contentSha256": "a" * 64,
                "fileName": "fixture.xlsx",
            },
            replay_items=[
                self._replay_item("capture_revision_fixture")
            ],
        )

        study = manifest["studies"][0]
        observation = study["outcomes"][0]["observations"][0]
        self.assertEqual("NEEDS_REVIEW", study["verificationStatus"])
        self.assertEqual(15.0, observation["average"])
        self.assertEqual(2, observation["sampleSize"])
        self.assertEqual(
            "CODE_FROM_CAPTURED_RAW_VALUES",
            observation["details"]["calculationAuthority"],
        )
        self.assertEqual([], study["comparisons"])

    def test_apply_is_idempotent_and_creates_queryable_observation(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "fixture.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Data"
            sheet.append(["Sample", "Tension"])
            sheet.append(["A", 10.0])
            sheet.append(["B", 20.0])
            workbook.save(source)
            workbook.close()

            database = root / "fixture.sqlite"
            connection = sqlite3.connect(database)
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys=ON")
            try:
                ensure_universal_schema(connection)
                ensure_capture_v2_schema(connection)
                ensure_knowledge_schema(
                    connection,
                    lambda: "2026-07-27T00:00:00+00:00",
                )
                payload = extract_workbook(source)
                capture = import_capture(
                    connection,
                    payload,
                    captured_at="2026-07-27T00:00:00+00:00",
                )
                bridge = bridge_capture_to_canonical_source(
                    connection,
                    dataset="Fixture",
                    payload=payload,
                    capture_result=capture,
                    captured_at="2026-07-27T00:00:00+00:00",
                )
                connection.commit()
            finally:
                connection.close()

            batch = root / "batch"
            batch.mkdir()
            replay = batch / "replay.json"
            replay.write_text(
                json.dumps(
                    {
                        "summary": {"passed": 1, "failed": 0},
                        "items": [
                            self._replay_item(bridge["revisionUid"])
                        ],
                    }
                ),
                encoding="utf-8",
            )
            (batch / "report.json").write_text(
                json.dumps(
                    {
                        "summary": {"eligibleWorkbookCount": 1},
                        "workbooks": [
                            {
                                "captureRevisionId": capture["revisionId"],
                            }
                        ],
                        "failures": [],
                    }
                ),
                encoding="utf-8",
            )
            (batch / "recipe-registry.json").write_text(
                json.dumps(
                    {
                        "summary": {"registeredTableCount": 1},
                        "recipes": [{"replayFile": str(replay)}],
                    }
                ),
                encoding="utf-8",
            )

            for index in (1, 2):
                result = run_structure_canonical_import(
                    database_path=database,
                    batch_root=batch,
                    artifact_root=root / "artifacts",
                    output_path=root / f"result-{index}.json",
                    apply=True,
                )
                self.assertEqual(
                    1,
                    result["summary"]["databaseCoverageAfter"][
                        "workbooksWithActiveCanonicalAnalysis"
                    ],
                )
                self.assertEqual(
                    0,
                    result["summary"]["databaseCoverageAfter"][
                        "workbooksWithMultipleActiveCanonicalAnalyses"
                    ],
                )
                self.assertEqual(0, result["aiUsage"]["aiCallCount"])
                self.assertEqual(
                    {
                        "analysisCount": 1,
                        "studyCount": 1,
                        "armCount": 1,
                        "outcomeCount": 1,
                        "observationCount": 1,
                        "entityEvidenceLinkCount": 4,
                        "distinctEvidenceCount": 4,
                        "analysisStatusCounts": [
                            {
                                "analysisStatus": "NEEDS_REVIEW",
                                "verificationStatus": "NEEDS_REVIEW",
                                "count": 1,
                            }
                        ],
                    },
                    result["summary"]["importerCoverageAfter"],
                )

            connection = sqlite3.connect(database)
            counts = {
                "analyses": connection.execute(
                    "SELECT COUNT(*) FROM workbook_analyses"
                ).fetchone()[0],
                "studies": connection.execute(
                    "SELECT COUNT(*) FROM knowledge_studies"
                ).fetchone()[0],
                "observations": connection.execute(
                    "SELECT COUNT(*) FROM knowledge_observations"
                ).fetchone()[0],
                "evidence": connection.execute(
                    "SELECT COUNT(*) FROM evidence_items"
                ).fetchone()[0],
            }
            average = connection.execute(
                "SELECT average_value FROM knowledge_observations"
            ).fetchone()[0]
            connection.close()

        self.assertEqual(
            {
                "analyses": 1,
                "studies": 1,
                "observations": 1,
                "evidence": 4,
            },
            counts,
        )
        self.assertEqual(15.0, average)


if __name__ == "__main__":
    unittest.main()
