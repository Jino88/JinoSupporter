from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import inference_data_ai_form_pipeline as pipeline


class FormPipelineTests(unittest.TestCase):
    def test_manifest_preserves_leading_space_in_workbook_name(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / " (2).xlsx"
            source.write_bytes(b"fixture")
            manifest = root / "manifest.json"
            manifest.write_text(
                json.dumps(
                    {
                        "workbooks": [
                            {
                                "relativePath": source.name,
                                "contentSha256": hashlib.sha256(
                                    b"fixture"
                                ).hexdigest(),
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )

            relative_paths = pipeline._manifest_relative_paths(
                manifest,
                root,
            )

        self.assertEqual([" (2).xlsx"], relative_paths)

    def test_auto_decision_is_fail_closed(self) -> None:
        self.assertEqual(
            "REGISTER_NEW",
            pipeline.choose_auto_form_decision(
                recommendation="REGISTER_NEW",
                validation_status="PASSED",
                nearest_known_form_signature_id="",
            ),
        )
        self.assertEqual(
            "LINK_EXISTING",
            pipeline.choose_auto_form_decision(
                recommendation="LINK_EXISTING",
                validation_status="PASSED",
                nearest_known_form_signature_id="form-known",
            ),
        )
        self.assertEqual(
            "REGISTER_NEW",
            pipeline.choose_auto_form_decision(
                recommendation="LINK_EXISTING",
                validation_status="PASSED",
                nearest_known_form_signature_id="",
            ),
        )
        self.assertEqual(
            "EXCLUDE",
            pipeline.choose_auto_form_decision(
                recommendation="REGISTER_NEW",
                validation_status="FAILED",
                nearest_known_form_signature_id="",
            ),
        )

    def test_corpus_retries_rejected_drafts_until_no_failures(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = root / "canonical.sqlite"
            database.write_bytes(b"fixture")
            source_root = root / "archive"
            source_root.mkdir()
            source = source_root / "report.xlsx"
            source.write_bytes(b"fixture")
            manifest = root / "manifest.json"
            manifest.write_text(
                json.dumps(
                    {
                        "workbooks": [
                            {
                                "relativePath": source.name,
                                "contentSha256": hashlib.sha256(
                                    b"fixture"
                                ).hexdigest(),
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )
            failed = {
                "status": "COMPLETED_WITH_ERRORS",
                "summary": {
                    "attempted": 1,
                    "completedThisRun": 0,
                    "failedThisRun": 1,
                    "currentStatusCounts": {"FAILED": 1},
                },
            }
            completed = {
                "status": "COMPLETED",
                "summary": {
                    "attempted": 1,
                    "completedThisRun": 1,
                    "failedThisRun": 0,
                    "currentStatusCounts": {"COMPLETED": 1},
                },
            }
            with mock.patch.object(
                pipeline,
                "run_corpus_ingest",
                side_effect=[failed, failed, completed],
            ) as ingest:
                result = pipeline._run_corpus(
                    database_path=database,
                    source_root=source_root,
                    output_root=root / "output",
                    manifest_path=manifest,
                    dataset="Fixture",
                    draft_monolithic_max_bytes=80_000,
                    progress_callback=None,
                )

        self.assertEqual(3, ingest.call_count)
        self.assertEqual(
            80_000,
            ingest.call_args.kwargs["ingest_options"][
                "draft_monolithic_max_bytes"
            ],
        )
        self.assertEqual("COMPLETED", result["status"])
        self.assertEqual(
            [1, 1, 0],
            [
                item["remainingFailed"]
                for item in result["retryPasses"]
            ],
        )

    def test_complete_pipeline_analyzes_and_decides_all_groups(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = root / "canonical.sqlite"
            database.write_bytes(b"fixture")
            source = root / "archive"
            source.mkdir()
            output = root / "output"
            manifest = output / "form-preflight" / "manifest.json"
            groups = [
                {
                    "familyId": "family-a",
                    "displayName": "A",
                    "decisionStatus": "PENDING",
                    "memberCount": 2,
                    "representativeSource": "a.xlsx",
                    "nearestKnownFormSignatureId": "",
                },
                {
                    "familyId": "family-b",
                    "displayName": "B",
                    "decisionStatus": "PENDING",
                    "memberCount": 1,
                    "representativeSource": "b.xlsx",
                    "nearestKnownFormSignatureId": "",
                },
            ]
            first_review = {
                "summary": {
                    "groupCount": 2,
                    "pendingCount": 2,
                    "approvedCount": 0,
                    "excludedCount": 0,
                    "workbookCount": 3,
                },
                "groups": groups,
            }
            final_review = {
                "summary": {
                    "groupCount": 2,
                    "pendingCount": 0,
                    "approvedCount": 1,
                    "excludedCount": 1,
                    "workbookCount": 3,
                },
                "groups": [],
            }

            def analyze(**arguments: object) -> dict[str, object]:
                family_id = str(arguments["family_id"])
                expected_group = next(
                    group
                    for group in groups
                    if group["familyId"] == family_id
                )
                self.assertIs(
                    expected_group,
                    arguments["group_snapshot"],
                )
                return {
                    "familyId": family_id,
                    "recommendation": (
                        "REGISTER_NEW"
                        if family_id == "family-a"
                        else "EXCLUDE"
                    ),
                    "validationStatus": (
                        "PASSED"
                        if family_id == "family-a"
                        else "FAILED"
                    ),
                }

            decisions: list[tuple[str, str]] = []

            def decide(**arguments: object) -> dict[str, object]:
                family_id = str(arguments["family_id"])
                decision = str(arguments["decision"])
                expected_group = next(
                    group
                    for group in groups
                    if group["familyId"] == family_id
                )
                self.assertIs(
                    expected_group,
                    arguments["group_snapshot"],
                )
                decisions.append((family_id, decision))
                return {
                    "status": {
                        "REGISTER_NEW": "APPROVED_NEW",
                        "LINK_EXISTING": "LINKED_EXISTING",
                        "EXCLUDE": "EXCLUDED",
                    }[decision],
                    "familyId": family_id,
                    "memberCount": 1,
                    "linkedFormSignatureId": "",
                }

            with (
                mock.patch.object(
                    pipeline,
                    "run_form_preflight",
                    return_value={
                        "status": "COMPLETED",
                        "summary": {"knownForms": 0},
                    },
                ),
                mock.patch.object(
                    pipeline,
                    "write_form_group_review",
                    side_effect=[first_review, final_review],
                ),
                mock.patch.object(
                    pipeline,
                    "analyze_form_family",
                    side_effect=analyze,
                ),
                mock.patch.object(
                    pipeline,
                    "decide_form_family",
                    side_effect=decide,
                ),
                mock.patch.object(
                    pipeline,
                    "reclassify_form_preflight_report",
                    return_value={
                        "knownFormManifestPath": str(manifest),
                        "summary": {
                            "knownForms": 2,
                            "captureFailed": 0,
                        },
                    },
                ),
            ):
                result = pipeline.run_form_pipeline_complete(
                    database_path=database,
                    source_root=source,
                    output_root=output,
                    reviewer="cli-test",
                    analysis_workers=2,
                    run_corpus=False,
                )

        self.assertEqual("COMPLETED", result["status"])
        self.assertCountEqual(
            [
                ("family-a", "REGISTER_NEW"),
                ("family-b", "EXCLUDE"),
            ],
            decisions,
        )
        self.assertEqual(2, len(result["outcomes"]))
        self.assertEqual([], result["errors"])


if __name__ == "__main__":
    unittest.main()
