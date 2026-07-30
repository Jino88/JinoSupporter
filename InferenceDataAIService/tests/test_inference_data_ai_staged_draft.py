from __future__ import annotations

import copy
import tempfile
import unittest
from pathlib import Path

import inference_data_ai_staged_draft as staged


class StagedDraftTests(unittest.TestCase):
    def packet(self, chunk_count: int = 32) -> dict:
        chunks = []
        for index in range(chunk_count):
            chunk_number = index + 1
            chunks.append(
                {
                    "chunkId": f"chunk-{chunk_number:02d}",
                    "sheet": {
                        "sheetIndex": 1,
                        "title": "Data",
                    },
                    "sectionIndex": 1,
                    "chunkIndex": chunk_number,
                    "primaryRange": f"A{chunk_number}",
                    "cells": [
                        {
                            "coordinate": f"A{chunk_number}",
                            "sourceCellKey": (
                                f"revision-1:1:A{chunk_number}"
                            ),
                            "rawValue": f"value-{chunk_number}",
                        }
                    ],
                    "contextCells": [],
                }
            )
        return {
            "chunks": chunks,
            "inventory": {
                "sourceRevision": {
                    "revisionUid": "revision-1",
                    "contentSha256": "a" * 64,
                }
            },
        }

    def locators(
        self,
        packet: dict,
        candidate_ids: set[str],
    ) -> list[dict]:
        results = []
        for chunk in packet["chunks"]:
            chunk_id = chunk["chunkId"]
            is_candidate = chunk_id in candidate_ids
            results.append(
                {
                    "chunkId": chunk_id,
                    "status": (
                        "CANDIDATES" if is_candidate else "NO_CANDIDATE"
                    ),
                    "candidates": (
                        [{"key": f"candidate-{chunk_id}"}]
                        if is_candidate
                        else []
                    ),
                }
            )
        return results

    def test_32_chunk_section_is_bounded_and_includes_continuations(
        self,
    ) -> None:
        packet = self.packet()
        locators = self.locators(packet, {"chunk-01"})
        kwargs = {
            "packet_set": packet,
            "locator_results": locators,
            "prompt_version": "prompt-v1",
            "max_chunks": 4,
            "max_cells": 4,
            "max_serialized_bytes": 10_000,
        }

        first = staged.plan_study_draft_parts(**kwargs)
        second = staged.plan_study_draft_parts(**kwargs)

        self.assertEqual(first, second)
        self.assertEqual(8, len(first["parts"]))
        self.assertEqual(["chunk-01"], first["candidateChunkIds"])
        self.assertEqual(
            [f"chunk-{value:02d}" for value in range(2, 33)],
            first["continuationChunkIds"],
        )
        self.assertEqual(
            [f"chunk-{value:02d}" for value in range(1, 33)],
            [
                chunk_id
                for part in first["parts"]
                for chunk_id in part["chunkIds"]
            ],
        )
        self.assertEqual(32, first["ownedSourceCellCount"])
        self.assertTrue(
            all(part["chunkCount"] <= 4 for part in first["parts"])
        )
        self.assertTrue(
            all(part["cellCount"] <= 4 for part in first["parts"])
        )
        self.assertTrue(
            all(
                part["serializedBytes"] <= 10_000
                for part in first["parts"]
            )
        )

    def fragment_manifest(
        self,
        *,
        study_key: str,
        study_title: str,
        evidence_range: str,
        limitation: str,
    ) -> dict:
        evidence = [
            {
                "sheet": "Data",
                "range": evidence_range,
                "role": "SOURCE",
                "sourceText": study_title,
                "note": "",
            }
        ]
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "dataset": "Fixture",
                "sourcePath": "fixture.xlsx",
                "revisionUid": "revision-1",
                "contentSha256": "a" * 64,
                "contentComplete": False,
            },
            "workbookAnalysis": {
                "key": "fragment-analysis",
                "title": "Fixture",
                "summary": study_title,
                "status": "NEEDS_REVIEW",
                "verificationStatus": "NEEDS_REVIEW",
                "limitations": [limitation],
                "evidence": evidence,
            },
            "studies": [
                {
                    "key": study_key,
                    "title": study_title,
                    "verificationStatus": "NEEDS_REVIEW",
                    "comparabilityStatus": "UNASSESSED",
                    "confoundingStatus": "UNASSESSED",
                    "evidence": evidence,
                    "contexts": [],
                    "factors": [],
                    "arms": [],
                    "outcomes": [],
                    "measurementSeries": [],
                    "comparisons": [],
                    "conclusions": [],
                    "limitations": [limitation],
                }
            ],
        }

    def test_consolidation_is_deterministic_and_rejects_key_collision(
        self,
    ) -> None:
        packet = self.packet(chunk_count=2)
        plan = staged.plan_study_draft_parts(
            packet_set=packet,
            locator_results=self.locators(packet, {"chunk-01"}),
            prompt_version="prompt-v1",
            max_chunks=1,
            max_cells=1,
            max_serialized_bytes=10_000,
        )
        first_manifest = self.fragment_manifest(
            study_key="same-key",
            study_title="First",
            evidence_range="A1",
            limitation="Shared limitation",
        )
        second_manifest = self.fragment_manifest(
            study_key="same-key",
            study_title="Second",
            evidence_range="A2",
            limitation="Second limitation",
        )
        source = {
            "dataset": "Fixture",
            "sourcePath": "fixture.xlsx",
            "revisionUid": "revision-1",
            "contentSha256": "a" * 64,
        }
        parts = [
            (plan["parts"][0], first_manifest),
            (plan["parts"][1], second_manifest),
        ]

        first = staged.consolidate_study_draft_parts(
            plan=plan,
            part_manifests=parts,
            source=source,
            content_complete=True,
        )
        second = staged.consolidate_study_draft_parts(
            plan=plan,
            part_manifests=parts,
            source=source,
            content_complete=True,
        )

        self.assertEqual(first, second)
        self.assertIs(True, first["source"]["contentComplete"])
        self.assertEqual(
            ["First", "Second"],
            [study["title"] for study in first["studies"]],
        )
        self.assertEqual(
            2,
            len({study["key"] for study in first["studies"]}),
        )
        self.assertEqual(
            ["Shared limitation", "Second limitation"],
            first["workbookAnalysis"]["limitations"],
        )
        self.assertEqual(
            ["A1", "A2"],
            [
                item["range"]
                for item in first["workbookAnalysis"]["evidence"]
            ],
        )
        self.assertTrue(
            all(
                not study["comparisons"]
                for study in first["studies"]
            )
        )

        collided = copy.deepcopy(plan)
        collided["parts"][1]["partId"] = collided["parts"][0]["partId"]
        with self.assertRaisesRegex(
            staged.StagedDraftError,
            "Study key collision",
        ):
            staged.consolidate_study_draft_parts(
                plan=collided,
                part_manifests=[
                    (collided["parts"][0], first_manifest),
                    (collided["parts"][1], second_manifest),
                ],
                source=source,
                content_complete=True,
            )

    def test_part_paths_and_provenance_are_stable_and_exact(
        self,
    ) -> None:
        packet = self.packet(chunk_count=2)
        plan = staged.plan_study_draft_parts(
            packet_set=packet,
            locator_results=self.locators(packet, {"chunk-01"}),
            prompt_version="prompt-v1",
            max_chunks=1,
            max_cells=1,
            max_serialized_bytes=10_000,
        )
        part = plan["parts"][0]
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest_path, provenance_path = staged.part_artifact_paths(
                Path(temp_dir),
                part,
            )
            self.assertEqual(
                Path(temp_dir) / "draft-parts",
                manifest_path.parent,
            )
            self.assertEqual(manifest_path.parent, provenance_path.parent)
            provenance = staged.part_provenance_value(
                plan=plan,
                part=part,
                manifest_path=manifest_path,
                manifest_sha256="manifest-sha",
                generated_at="2026-07-18T00:00:00Z",
            )
            self.assertTrue(
                staged.part_provenance_matches(
                    provenance=provenance,
                    plan=plan,
                    part=part,
                    manifest_sha256="manifest-sha",
                )
            )
            changed = copy.deepcopy(provenance)
            changed["chunkIds"] = ["other"]
            self.assertFalse(
                staged.part_provenance_matches(
                    provenance=changed,
                    plan=plan,
                    part=part,
                    manifest_sha256="manifest-sha",
                )
            )


if __name__ == "__main__":
    unittest.main()
