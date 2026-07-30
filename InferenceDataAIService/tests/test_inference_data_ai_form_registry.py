from __future__ import annotations

import json
import sqlite3
import subprocess
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

from inference_data_ai_form_registry import (
    CONTRACT_SCHEMA_VERSION,
    analyze_form_family,
    build_form_group_review,
    decide_form_family,
    reclassify_form_preflight_report,
)
from inference_data_ai_schema import ensure_knowledge_schema
from inference_data_ai_source_ingest import (
    ensure_capture_v2_schema,
    extract_workbook,
    import_capture,
)


class FormRegistryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.database = self.root / "canonical.sqlite"
        self.archive = self.root / "archive"
        self.archive.mkdir()
        self.report_path = self.root / "form-preflight" / "latest.json"
        self.report_path.parent.mkdir()
        self.revision_ids = [
            self._capture_workbook("sample-a.xlsx", 5),
            self._capture_workbook("sample-b.xlsx", 12),
        ]
        self._write_report()

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def _capture_workbook(self, name: str, row_count: int) -> int:
        source = self.archive / name
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Report"
        sheet.append(["TEST", "MIN", "MAX", "RESULT"])
        for index in range(row_count):
            sheet.append([f"sample-{index + 1}", 1, 10, index + 2])
        workbook.save(source)
        workbook.close()
        payload = extract_workbook(source)
        with closing(sqlite3.connect(self.database)) as connection:
            connection.row_factory = sqlite3.Row
            ensure_knowledge_schema(
                connection,
                lambda: "2026-07-25T00:00:00Z",
            )
            ensure_capture_v2_schema(connection)
            capture = import_capture(
                connection,
                payload,
                captured_at="2026-07-25T00:00:00Z",
            )
            connection.commit()
        return int(capture["revisionId"])

    def _write_report(self) -> None:
        items = []
        for source, revision_id in zip(
            sorted(self.archive.glob("*.xlsx")),
            self.revision_ids,
            strict=True,
        ):
            items.append(
                {
                    "sourcePath": str(source.resolve()),
                    "relativePath": source.name,
                    "fileName": source.name,
                    "contentSha256": extract_workbook(source)["source"][
                        "contentSha256"
                    ],
                    "sizeBytes": source.stat().st_size,
                    "captureAction": "REUSED_CAPTURE",
                    "captureRevisionId": revision_id,
                    "status": "NEW_FORM",
                    "similarity": 0.0,
                    "nearestKnownSource": "",
                    "nearestKnownFormSignatureId": "",
                    "reason": "검토 필요",
                    "formSignatureId": f"strict-{revision_id}",
                }
            )
        report = {
            "schemaVersion": "excel-form-preflight-v1",
            "classifierVersion": "fixture",
            "status": "COMPLETED",
            "generatedAt": "2026-07-25T00:00:00Z",
            "databasePath": str(self.database),
            "sourceRoot": str(self.archive),
            "knownCatalogCount": 0,
            "summary": {
                "total": len(items),
                "knownForms": 0,
                "similarReview": 0,
                "newForms": len(items),
                "captureFailed": 0,
                "fullProcessingAllowed": False,
            },
            "knownFormManifestPath": str(
                self.report_path.with_name(
                    "latest.known-forms.manifest.json"
                )
            ),
            "items": items,
        }
        self.report_path.write_text(
            json.dumps(report, ensure_ascii=False),
            encoding="utf-8",
        )

    def _review(self) -> dict:
        report = json.loads(
            self.report_path.read_text(encoding="utf-8")
        )
        with closing(sqlite3.connect(self.database)) as connection:
            connection.row_factory = sqlite3.Row
            return build_form_group_review(
                connection=connection,
                report=report,
            )

    def test_row_growth_is_grouped_into_one_review_family(self) -> None:
        review = self._review()

        self.assertEqual(1, review["summary"]["groupCount"])
        self.assertEqual(2, review["summary"]["workbookCount"])
        group = review["groups"][0]
        self.assertEqual(2, group["memberCount"])
        self.assertEqual(2, len(group["sampleSources"]))
        self.assertEqual("PENDING", group["decisionStatus"])

    def test_ai_validation_and_human_approval_reclassify_manifest(
        self,
    ) -> None:
        group = self._review()["groups"][0]
        family_id = group["familyId"]

        def fake_run(
            arguments: list[str],
            **_: object,
        ) -> subprocess.CompletedProcess[str]:
            self.assertEqual("-", arguments[-1])
            self.assertLess(arguments.index("-c"), len(arguments) - 1)
            schema_path = Path(
                arguments[arguments.index("--output-schema") + 1]
            )
            schema = json.loads(schema_path.read_text(encoding="utf-8"))
            self.assertEqual(
                {
                    "type": "string",
                    "const": CONTRACT_SCHEMA_VERSION,
                },
                schema["properties"]["schemaVersion"],
            )
            self.assertEqual(
                "string",
                schema["properties"]["recommendation"]["type"],
            )
            response_path = Path(
                arguments[
                    arguments.index("--output-last-message") + 1
                ]
            )
            response_path.write_text(
                json.dumps(
                    {
                        "schemaVersion": CONTRACT_SCHEMA_VERSION,
                        "familyId": family_id,
                        "familyName": "Fixture report",
                        "documentType": "test report",
                        "extractionContract": {
                            "targetSheets": ["Report"],
                            "headerPatterns": [
                                "TEST | MIN | MAX | RESULT"
                            ],
                            "tableRules": ["header row then data rows"],
                            "requiredFields": ["TEST", "RESULT"],
                            "cautions": ["preserve source coordinates"],
                        },
                        "sampleValidation": [
                            {
                                "sourcePath": source,
                                "compatible": True,
                                "reason": "same layout",
                            }
                            for source in group["sampleSources"]
                        ],
                        "confidence": 0.95,
                        "recommendation": "REGISTER_NEW",
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(arguments, 0, "", "")

        contract_path = self.root / "contracts" / f"{family_id}.json"
        with mock.patch(
            "inference_data_ai_form_registry.subprocess.run",
            side_effect=fake_run,
        ):
            analyzed = analyze_form_family(
                database_path=self.database,
                report_path=self.report_path,
                family_id=family_id,
                output_path=contract_path,
                codex_executable="codex",
            )
        self.assertEqual("PASSED", analyzed["validationStatus"])

        decision = decide_form_family(
            database_path=self.database,
            report_path=self.report_path,
            family_id=family_id,
            decision="REGISTER_NEW",
            reviewer="tester",
        )
        self.assertEqual("APPROVED_NEW", decision["status"])

        report = reclassify_form_preflight_report(
            database_path=self.database,
            report_path=self.report_path,
        )
        self.assertEqual(2, report["summary"]["knownForms"])
        self.assertEqual(
            {"KNOWN_FORM"},
            {item["status"] for item in report["items"]},
        )
        manifest = json.loads(
            Path(report["knownFormManifestPath"]).read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(2, len(manifest["workbooks"]))
        self.assertTrue(
            all(
                item["registryDecision"] == "APPROVED_NEW"
                and "extractionContract" in item
                for item in manifest["workbooks"]
            )
        )

    def test_exclusion_needs_reviewer_and_reclassifies_family(
        self,
    ) -> None:
        family_id = self._review()["groups"][0]["familyId"]
        with self.assertRaisesRegex(ValueError, "reviewer"):
            decide_form_family(
                database_path=self.database,
                report_path=self.report_path,
                family_id=family_id,
                decision="EXCLUDE",
                reviewer="",
            )

        decide_form_family(
            database_path=self.database,
            report_path=self.report_path,
            family_id=family_id,
            decision="EXCLUDE",
            reviewer="tester",
            notes="not an analysis report",
        )
        report = reclassify_form_preflight_report(
            database_path=self.database,
            report_path=self.report_path,
        )
        self.assertEqual(2, report["summary"]["excludedForms"])
        self.assertEqual(0, report["summary"]["knownForms"])
        self.assertEqual(
            {"EXCLUDED_FORM"},
            {item["status"] for item in report["items"]},
        )


if __name__ == "__main__":
    unittest.main()
