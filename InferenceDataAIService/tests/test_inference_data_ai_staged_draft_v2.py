from __future__ import annotations

import copy
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import inference_data_ai_staged_draft_v2 as staged
import inference_data_ai_staged_runner_v2 as runner


class StagedDraftV2Tests(unittest.TestCase):
    source = {
        "dataset": "Fixture",
        "sourcePath": "fixture.xlsx",
        "revisionUid": "revision-1",
        "contentSha256": "a" * 64,
    }
    workbook = {"fileName": "fixture.xlsx"}

    def cell(
        self,
        coordinate: str,
        value: object,
        *,
        formula: str = "",
        cached_value: object = None,
    ) -> dict:
        column = ord(coordinate[0].upper()) - ord("A") + 1
        row = int(coordinate[1:])
        return {
            "sourceCellKey": f"revision-1:1:{coordinate}",
            "row": row,
            "column": column,
            "coordinate": coordinate,
            "rawValue": None if formula else value,
            "formula": formula,
            "cachedValue": cached_value,
            "displayValue": cached_value if formula else value,
            "dataType": "n" if isinstance(value, (int, float)) else "s",
            "cachedDataType": (
                "n" if isinstance(cached_value, (int, float)) else None
            ),
            "numberFormat": "General",
        }

    def chunk(
        self,
        chunk_id: str,
        cells: list[dict],
        *,
        section: int = 1,
        context: list[dict] | None = None,
    ) -> dict:
        coordinates = [str(cell["coordinate"]) for cell in cells]
        return {
            "chunkId": chunk_id,
            "sheet": {"sheetIndex": 1, "title": "Data"},
            "sectionIndex": section,
            "primaryRange": (
                coordinates[0]
                if len(coordinates) == 1
                else f"{coordinates[0]}:{coordinates[-1]}"
            ),
            "cells": cells,
            "contextCells": context or [],
        }

    def locator(
        self,
        chunk: dict,
        *,
        candidate: bool = True,
        evidence_range: str | None = None,
    ) -> dict:
        return {
            "chunkId": chunk["chunkId"],
            "status": "CANDIDATES" if candidate else "NO_CANDIDATE",
            "candidates": (
                [
                    {
                        "key": f"candidate-{chunk['chunkId']}",
                        "title": "Source review",
                        "evidence": [
                            {
                                "sheet": "Data",
                                "range": (
                                    evidence_range
                                    or chunk["primaryRange"]
                                ),
                                "role": "CANDIDATE_REGION",
                            }
                        ],
                    }
                ]
                if candidate
                else []
            ),
        }

    def universe(
        self,
        chunks: list[dict],
        locators: list[dict],
    ) -> dict:
        return staged.select_draft_universe(
            packet_set={"chunks": chunks},
            locator_results=locators,
        )

    def planned(
        self,
        chunks: list[dict],
        locators: list[dict],
        *,
        max_chunks: int = 1,
    ) -> tuple[dict, dict, dict]:
        universe = self.universe(chunks, locators)
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=universe,
        )
        plan = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=max_chunks,
            max_cells=100,
            max_serialized_bytes=100_000,
        )
        return universe, registry, plan

    def envelope(
        self,
        universe: dict,
        registry: dict,
        plan: dict,
        part_index: int = 0,
    ) -> tuple[dict, dict]:
        part = plan["parts"][part_index]
        return part, staged.finalize_fragment_envelope(
            staged.build_fragment_envelope(
                source=self.source,
                workbook=self.workbook,
                plan=plan,
                part=part,
                focused_chunks=staged.chunks_for_part_v2(
                    universe,
                    part,
                ),
                locator_results=staged.locators_for_part_v2(
                    universe,
                    part,
                ),
                registry_slice=staged.registry_for_part(
                    registry,
                    part,
                ),
            )
        )

    def record(
        self,
        *,
        envelope: dict,
        logical_id: str,
        record_type: str,
        identity_key: str,
        label: str,
        payload: dict,
        evidence_range: str,
    ) -> dict:
        subtype_record = {
            "recordType": record_type,
            "payload": payload,
        }
        subtype = staged._record_semantic_subtype(subtype_record)
        return {
            "recordType": record_type,
            "recordId": staged.stable_record_id(
                revision_uid=self.source["revisionUid"],
                logical_study_id=logical_id,
                record_type=record_type,
                identity_cell_keys=[identity_key],
                exact_source_label=label,
                semantic_subtype=subtype,
            ),
            "logicalStudyId": logical_id,
            "identityCellKeys": [identity_key],
            "exactSourceLabel": label,
            "payload": payload,
            "evidence": [
                {
                    "sheet": "Data",
                    "range": evidence_range,
                    "role": "SOURCE",
                }
            ],
        }

    def fragment(
        self,
        envelope: dict,
        records: list[dict],
        dispositions: list[dict] | None = None,
    ) -> dict:
        if dispositions is None:
            dispositions = [
                {
                    "sourceCellKey": key,
                    "disposition": "CONTEXT_ONLY",
                    "recordIds": [],
                    "reason": "Fixture context.",
                }
                for key in envelope["ownedSourceCellKeys"]
            ]
        return {
            "schemaVersion": "study-draft-fragment-v2",
            "source": {
                **self.source,
                "contentComplete": False,
            },
            "planId": envelope["planId"],
            "partId": envelope["partId"],
            "inputEnvelopeSha256": envelope[
                "inputEnvelopeSha256"
            ],
            "records": records,
            "coverageDispositions": dispositions,
        }

    def merged_value(
        self,
        records: list[dict],
        chunks: list[dict],
        *,
        records_sha256: str,
    ) -> dict:
        dispositions = []
        for chunk in chunks:
            for cell in chunk.get("cells", []):
                key = cell["sourceCellKey"]
                record_ids = [
                    record["recordId"]
                    for record in records
                    if key
                    in staged.evidence_cell_keys(
                        record.get("evidence", []),
                        chunks=chunks,
                    )
                ]
                dispositions.append(
                    {
                        "sourceCellKey": key,
                        "disposition": (
                            "RECORD_EVIDENCE"
                            if record_ids
                            else "CONTEXT_ONLY"
                        ),
                        "recordIds": record_ids,
                        "reason": "Fixture final coverage.",
                    }
                )
        return {
            "recordsSha256": records_sha256,
            "records": records,
            "coverageDispositions": dispositions,
        }

    def test_exact_request_budget_boundary_and_runner_prompt_hash(
        self,
    ) -> None:
        request = staged.build_monolithic_request(
            source=self.source,
            workbook=self.workbook,
            universe={
                "selectedLocatorResults": [],
                "selectedChunks": [],
            },
            content_complete=True,
            prompt_text="가" * 17,
        )
        exact = request["promptBytes"]
        self.assertEqual(
            "MONOLITHIC",
            staged.assess_one_call_budget(
                request=request,
                max_prompt_bytes=exact,
            )["mode"],
        )
        self.assertEqual(
            "STAGED_V2",
            staged.assess_one_call_budget(
                request=request,
                max_prompt_bytes=exact - 1,
            )["mode"],
        )
        cell_limited_request = staged.build_monolithic_request(
            source=self.source,
            workbook=self.workbook,
            universe={
                "selectedLocatorResults": [],
                "selectedChunks": [
                    self.chunk(
                        "cell-limit",
                        [
                            self.cell("A1", "one"),
                            self.cell("A2", "two"),
                        ],
                    )
                ],
            },
            content_complete=True,
            prompt_text="small",
        )
        self.assertEqual(
            "MONOLITHIC",
            staged.assess_one_call_budget(
                request=cell_limited_request,
                max_prompt_bytes=10_000,
                max_source_cells=2,
            )["mode"],
        )
        cell_budget = staged.assess_one_call_budget(
            request=cell_limited_request,
            max_prompt_bytes=10_000,
            max_source_cells=1,
        )
        self.assertEqual("STAGED_V2", cell_budget["mode"])
        self.assertEqual(2, cell_budget["sourceCellCount"])
        self.assertEqual(1, cell_budget["maxSourceCells"])

    def test_planner_splits_on_exact_finalized_prompt_and_is_stable(
        self,
    ) -> None:
        first = self.chunk(
            "c1",
            [
                self.cell("A1", "A" * 1_000),
                self.cell("B1", "first"),
            ],
        )
        second = self.chunk(
            "c2",
            [
                self.cell("A2", "B" * 1_000),
                self.cell("B2", "second"),
            ],
        )
        locators = [
            self.locator(first),
            self.locator(second),
        ]
        universe = self.universe([first, second], locators)
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=universe,
        )
        combined = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=2,
            max_cells=100,
            max_serialized_bytes=1_000_000,
        )
        self.assertEqual(1, len(combined["parts"]))
        combined_part = combined["parts"][0]
        separate = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=1,
            max_cells=100,
            max_serialized_bytes=1_000_000,
        )
        prompt_limit = max(
            part["promptBytes"] for part in separate["parts"]
        ) + 256
        self.assertGreater(
            combined_part["promptBytes"],
            prompt_limit,
        )

        bounded = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=2,
            max_cells=100,
            max_serialized_bytes=prompt_limit,
        )
        repeated = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=2,
            max_cells=100,
            max_serialized_bytes=prompt_limit,
        )
        self.assertEqual(bounded, repeated)
        self.assertEqual(2, len(bounded["parts"]))
        self.assertEqual(
            universe["ownedSourceCellKeys"],
            [
                key
                for part in bounded["parts"]
                for key in part["ownedSourceCellKeys"]
            ],
        )
        self.assertTrue(
            all(
                part["promptBytes"] <= prompt_limit
                for part in bounded["parts"]
            )
        )
        for index, part in enumerate(bounded["parts"]):
            _part, envelope = self.envelope(
                universe,
                registry,
                bounded,
                index,
            )
            self.assertEqual(
                part["promptBytes"],
                envelope["promptBytes"],
            )

    def test_planner_segments_one_oversized_chunk_source_losslessly(
        self,
    ) -> None:
        header = self.cell("A1", "Shared section header")
        merged_anchor = self.cell("B2", "Merged arm label")
        merged_anchor["mergeRange"] = "B2:B4"
        merged_anchor["mergeRole"] = "anchor"
        cells = [
            merged_anchor,
            *[
                self.cell(
                    f"B{row}",
                    f"value-{row}-" + ("가" * 600),
                )
                for row in range(5, 17)
            ],
        ]
        chunk = self.chunk(
            "oversized",
            cells,
            context=[header],
        )
        locator = self.locator(
            chunk,
            evidence_range="B2:B16",
        )
        universe = self.universe([chunk], [locator])
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=universe,
        )
        unbounded = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=8,
            max_cells=100,
            max_serialized_bytes=1_000_000,
        )
        exact_limit = unbounded["parts"][0]["promptBytes"] - 1
        plan = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=8,
            max_cells=100,
            max_serialized_bytes=exact_limit,
        )

        segmented_parts = [
            part for part in plan["parts"] if part["sourceSegments"]
        ]
        self.assertGreater(len(segmented_parts), 1)
        self.assertEqual(
            ["oversized"],
            plan["sourceSegmentation"]["segmentedChunkIds"],
        )
        expected_keys = [
            cell["sourceCellKey"] for cell in chunk["cells"]
        ]
        actual_keys = [
            key
            for part in plan["parts"]
            for key in part["ownedSourceCellKeys"]
        ]
        self.assertEqual(expected_keys, actual_keys)
        self.assertEqual(len(actual_keys), len(set(actual_keys)))
        self.assertTrue(
            all(
                part["promptBytes"] <= exact_limit
                for part in plan["parts"]
            )
        )
        self.assertTrue(
            all(
                part["logicalStudyIds"]
                == segmented_parts[0]["logicalStudyIds"]
                for part in segmented_parts
            )
        )

        merged_key = merged_anchor["sourceCellKey"]
        saw_shared_merged_anchor = False
        for part in segmented_parts:
            descriptor = part["sourceSegments"][0]
            start = expected_keys.index(
                descriptor["firstSourceCellKey"]
            )
            end = start + descriptor["sourceCellCount"]
            self.assertEqual(
                expected_keys[start:end],
                part["ownedSourceCellKeys"],
            )
            self.assertEqual(
                part["ownedSourceCellKeys"][-1],
                descriptor["lastSourceCellKey"],
            )
            focused = staged.chunks_for_part_v2(universe, part)
            self.assertEqual(1, len(focused))
            self.assertEqual("oversized", focused[0]["chunkId"])
            self.assertEqual(
                part["ownedSourceCellKeys"],
                [
                    cell["sourceCellKey"]
                    for cell in focused[0]["cells"]
                ],
            )
            self.assertEqual(
                [locator],
                staged.locators_for_part_v2(universe, part),
            )
            context_by_key = {
                cell["sourceCellKey"]: cell
                for cell in focused[0]["contextCells"]
            }
            self.assertIn(header["sourceCellKey"], context_by_key)
            self.assertTrue(
                context_by_key[header["sourceCellKey"]]["contextOnly"]
            )
            self.assertFalse(
                context_by_key[header["sourceCellKey"]]["primary"]
            )
            self.assertTrue(
                set(context_by_key).issubset(
                    set(part["sharedAnchorCellKeys"])
                )
            )
            if merged_key not in part["ownedSourceCellKeys"]:
                self.assertIn(merged_key, context_by_key)
                saw_shared_merged_anchor = True
            envelope = self.envelope(
                universe,
                registry,
                plan,
                part["partIndex"] - 1,
            )[1]
            self.assertEqual(
                part["promptBytes"],
                envelope["promptBytes"],
            )
            self.assertEqual(
                part["sourceSegments"],
                envelope["sourceSegments"],
            )
        self.assertTrue(saw_shared_merged_anchor)

        first_part = segmented_parts[0]
        first_envelope = self.envelope(
            universe,
            registry,
            plan,
            first_part["partIndex"] - 1,
        )[1]
        provenance = staged.part_provenance_v2(
            plan=plan,
            part=first_part,
            envelope=first_envelope,
            output_path=Path("segment.json"),
            output_sha256="segment-sha",
            generated_at="now",
        )
        self.assertTrue(
            staged.part_provenance_v2_matches(
                provenance=provenance,
                plan=plan,
                part=first_part,
                envelope=first_envelope,
                output_sha256="segment-sha",
                output_path=Path("segment.json"),
            )
        )
        stale_provenance = copy.deepcopy(provenance)
        stale_provenance["sourceSegments"][0][
            "lastSourceCellKey"
        ] = "stale"
        self.assertFalse(
            staged.part_provenance_v2_matches(
                provenance=stale_provenance,
                plan=plan,
                part=first_part,
                envelope=first_envelope,
                output_sha256="segment-sha",
                output_path=Path("segment.json"),
            )
        )
        stale_universe = copy.deepcopy(universe)
        stale_universe["selectedChunks"][0]["cells"][0][
            "displayValue"
        ] = "changed-after-plan"
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "no longer matches its chunk",
        ):
            staged.chunks_for_part_v2(
                stale_universe,
                first_part,
            )

    def test_planner_rejects_oversized_atomic_cell_context_envelope(
        self,
    ) -> None:
        context = self.cell("A1", "context-" + ("X" * 20_000))
        cell = self.cell("B2", "value-" + ("Y" * 20_000))
        chunk = self.chunk("atomic", [cell], context=[context])
        locator = self.locator(chunk, evidence_range="B2")
        universe = self.universe([chunk], [locator])
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=universe,
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "atomic source cell.*context envelope",
        ):
            staged.plan_study_draft_v2(
                source=self.source,
                workbook=self.workbook,
                universe=universe,
                registry=registry,
                prompt_version="draft-v1",
                max_chunks=8,
                max_cells=100,
                max_serialized_bytes=2_000,
            )

    def test_segment_registry_preserves_linked_logical_study_anchor(
        self,
    ) -> None:
        header = self.cell("A1", "Shared Study header")
        first = self.chunk(
            "anchor-chunk",
            [header, self.cell("B1", "Baseline")],
        )
        oversized = self.chunk(
            "segmented-chunk",
            [
                self.cell(f"B{row}", f"result-{row}")
                for row in range(2, 8)
            ],
            context=[copy.deepcopy(header)],
        )
        first_locator = self.locator(
            first,
            evidence_range="A1:B1",
        )
        oversized_locator = self.locator(
            oversized,
            evidence_range="B2:B7",
        )
        universe = self.universe(
            [first, oversized],
            [first_locator, oversized_locator],
        )
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=universe,
        )
        self.assertEqual(1, len(registry["studies"]))
        logical_id = registry["studies"][0]["logicalStudyId"]
        plan = staged.plan_study_draft_v2(
            source=self.source,
            workbook=self.workbook,
            universe=universe,
            registry=registry,
            prompt_version="draft-v1",
            max_chunks=8,
            max_cells=2,
            max_serialized_bytes=100_000,
        )
        segmented_parts = [
            part
            for part in plan["parts"]
            if part["sourceSegments"]
            and part["sourceSegments"][0]["sourceChunkId"]
            == "segmented-chunk"
        ]
        self.assertGreater(len(segmented_parts), 1)
        for part in segmented_parts:
            self.assertEqual([logical_id], part["logicalStudyIds"])
            self.assertIn(
                header["sourceCellKey"],
                part["sharedAnchorCellKeys"],
            )
            scoped = staged.registry_for_part(registry, part)
            self.assertEqual(
                "SOURCE_SEGMENT",
                scoped["scope"]["mode"],
            )
            self.assertEqual(
                registry["registrySha256"],
                scoped["scope"]["fullRegistrySha256"],
            )
            self.assertEqual([logical_id], scoped["scope"][
                "logicalStudyIds"
            ])
            self.assertEqual(1, len(scoped["candidateAnchors"]))
            self.assertEqual(
                "segmented-chunk",
                scoped["candidateAnchors"][0]["chunkId"],
            )
            scoped_study = scoped["studies"][0]
            self.assertEqual(
                2,
                len(scoped_study["memberCandidateAnchorIds"]),
            )
            self.assertEqual(
                1,
                len(scoped_study["focusedCandidateAnchorIds"]),
            )
            self.assertTrue(
                scoped_study["fullAnchorEvidenceSha256"]
            )
            self.assertEqual(
                [oversized_locator],
                staged.locators_for_part_v2(universe, part),
            )

    def test_plan_and_part_ids_bind_every_fragment_contract_version(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "value")])
        universe = self.universe([chunk], [self.locator(chunk)])
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=universe,
        )

        def plan_value() -> dict:
            return staged.plan_study_draft_v2(
                source=self.source,
                workbook=self.workbook,
                universe=universe,
                registry=registry,
                prompt_version="draft-v1",
                max_chunks=1,
                max_cells=100,
                max_serialized_bytes=100_000,
            )

        baseline = plan_value()
        baseline_part = baseline["parts"][0]
        baseline_envelope = self.envelope(
            universe,
            registry,
            baseline,
        )[1]
        baseline_part_provenance = staged.part_provenance_v2(
            plan=baseline,
            part=baseline_part,
            envelope=baseline_envelope,
            output_path=Path("part.json"),
            output_sha256="part-sha",
            generated_at="now",
        )
        baseline_final_provenance = staged.final_provenance_v2(
            plan=baseline,
            registry=registry,
            ordered_part_hashes=[
                {
                    "partId": baseline_part["partId"],
                    "outputSha256": "part-sha",
                }
            ],
            merged_path=Path("merged.json"),
            merged_sha256="merged-sha",
            final_path=Path("final.json"),
            final_sha256="final-sha",
            generated_at="now",
        )
        self.assertEqual(
            staged.json_sha256(baseline["fragmentIdentity"]),
            baseline["fragmentIdentitySha256"],
        )
        self.assertEqual(
            baseline["fragmentIdentitySha256"],
            baseline_part["fragmentIdentitySha256"],
        )

        version_fields = (
            ("FRAGMENT_PROMPT_VERSION", "promptVersion"),
            (
                "FRAGMENT_CONTRACT_VERSION",
                "fragmentContractVersion",
            ),
            (
                "FRAGMENT_VALIDATOR_VERSION",
                "validatorContractVersion",
            ),
            (
                "CONSOLIDATOR_CONTRACT_VERSION",
                "consolidatorContractVersion",
            ),
        )
        for attribute, identity_field in version_fields:
            with self.subTest(attribute=attribute):
                changed_version = getattr(staged, attribute) + "-future"
                with mock.patch.object(
                    staged,
                    attribute,
                    changed_version,
                ):
                    changed = plan_value()
                    changed_part = changed["parts"][0]
                    self.assertEqual(
                        changed_version,
                        changed["fragmentIdentity"][identity_field],
                    )
                    self.assertEqual(
                        staged.json_sha256(
                            changed["fragmentIdentity"]
                        ),
                        changed["fragmentIdentitySha256"],
                    )
                    self.assertNotEqual(
                        baseline_part["partId"],
                        changed_part["partId"],
                    )
                    self.assertNotEqual(
                        baseline["planId"],
                        changed["planId"],
                    )
                    _part, changed_envelope = self.envelope(
                        universe,
                        registry,
                        changed,
                    )
                    self.assertEqual(
                        changed["fragmentIdentitySha256"],
                        changed_envelope[
                            "fragmentIdentitySha256"
                        ],
                    )
                    with self.assertRaisesRegex(
                        staged.StagedDraftV2Error,
                        "identity does not match",
                    ):
                        self.envelope(
                            universe,
                            registry,
                            baseline,
                        )
                    self.assertFalse(
                        staged.part_provenance_v2_matches(
                            provenance=baseline_part_provenance,
                            plan=baseline,
                            part=baseline_part,
                            envelope=baseline_envelope,
                            output_sha256="part-sha",
                            output_path=Path("part.json"),
                        )
                    )
                    self.assertFalse(
                        staged.final_provenance_v2_matches(
                            provenance=baseline_final_provenance,
                            plan=baseline,
                            registry=registry,
                            final_sha256="final-sha",
                        )
                    )

    def test_runner_uses_stable_unique_transport_and_observes_ai_call(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "label")])
        locator = self.locator(chunk)
        universe, registry, plan = self.planned(
            [chunk],
            [locator],
        )
        _part, envelope = self.envelope(
            universe,
            registry,
            plan,
        )
        observed: dict[str, object] = {}

        def fake_run(
            command: list[str],
            **kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            observed["input"] = kwargs["input"]
            output_index = command.index("--output-last-message") + 1
            output_path = Path(command[output_index])
            output_path.write_text(
                json.dumps(self.fragment(envelope, [])),
                encoding="utf-8",
            )
            observed["outputPath"] = output_path
            schema_index = command.index("--output-schema") + 1
            observed["schemaPath"] = Path(command[schema_index])
            return subprocess.CompletedProcess(command, 0, "", "")

        ai_calls: list[str] = []
        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "fragment.json"
            result = runner.run_codex_study_fragment_v2(
                envelope=envelope,
                all_selected_chunks=[chunk],
                output_path=target,
                codex_command=["codex"],
                run_command=fake_run,
                ai_call_observer=lambda: ai_calls.append("called"),
            )
            output_transport = observed["outputPath"]
            schema_transport = observed["schemaPath"]
            self.assertIsInstance(output_transport, Path)
            self.assertIsInstance(schema_transport, Path)
            self.assertEqual(target.parent, output_transport.parent)
            self.assertEqual(target.parent, schema_transport.parent)
            self.assertNotEqual(target, output_transport)
            self.assertNotEqual(target, schema_transport)
            self.assertFalse(output_transport.exists())
            self.assertFalse(schema_transport.exists())
            self.assertTrue(target.is_file())
        self.assertEqual(envelope["promptText"], observed["input"])
        self.assertEqual(
            envelope["inputHashes"]["promptSha256"],
            staged.bytes_sha256(
                str(observed["input"]).encode("utf-8")
            ),
        )
        self.assertEqual([], result["records"])
        self.assertEqual(["called"], ai_calls)

        invalid_transport = self.fragment(envelope, [])
        invalid_transport["partId"] = "wrong-part"

        def invalid_run(
            command: list[str],
            **_kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            output_path = Path(
                command[
                    command.index("--output-last-message") + 1
                ]
            )
            output_path.write_text(
                json.dumps(invalid_transport),
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(command, 0, "", "")

        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "fragment.json"
            with self.assertRaisesRegex(
                staged.StagedDraftV2Error,
                "partId",
            ):
                runner.run_codex_study_fragment_v2(
                    envelope=envelope,
                    all_selected_chunks=[chunk],
                    output_path=target,
                    codex_command=["codex"],
                    run_command=invalid_run,
                )
            rejected_path = (
                target.parent / "fragment.rejected.json"
            )
            self.assertFalse(target.exists())
            self.assertTrue(rejected_path.is_file())
            rejected = json.loads(
                rejected_path.read_text(encoding="utf-8")
            )
            self.assertEqual("wrong-part", rejected["partId"])

        changed = copy.deepcopy(envelope)
        changed["promptText"] += "changed"
        blocked_calls: list[str] = []
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "provenance hash",
        ):
            runner.run_codex_study_fragment_v2(
                envelope=changed,
                all_selected_chunks=[chunk],
                output_path="unused.json",
                codex_command=["codex"],
                run_command=lambda *_args, **_kwargs: self.fail(
                    "stale prompt must not execute"
                ),
                ai_call_observer=lambda: blocked_calls.append(
                    "called"
                ),
            )
        self.assertEqual([], blocked_calls)

        failed_paths: list[Path] = []
        failed_calls: list[str] = []

        def failed_run(
            command: list[str],
            **_kwargs: object,
        ) -> subprocess.CompletedProcess[str]:
            failed_paths.extend(
                [
                    Path(
                        command[
                            command.index("--output-last-message") + 1
                        ]
                    ),
                    Path(
                        command[command.index("--output-schema") + 1]
                    ),
                ]
            )
            return subprocess.CompletedProcess(
                command,
                2,
                "",
                "fixture transport failure",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(
                staged.StagedDraftV2Error,
                "exit code 2",
            ):
                runner.run_codex_study_fragment_v2(
                    envelope=envelope,
                    all_selected_chunks=[chunk],
                    output_path=Path(temp_dir) / "failed.json",
                    codex_command=["codex"],
                    run_command=failed_run,
                    ai_call_observer=lambda: failed_calls.append(
                        "called"
                    ),
                )
            self.assertEqual(["called"], failed_calls)
            self.assertTrue(failed_paths)
            self.assertTrue(
                all(not path.exists() for path in failed_paths)
            )

    def test_runner_promotes_prior_contract_fragment_without_ai_call(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "label")])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        prior = self.fragment(envelope, [])
        prior["planId"] = "old-plan"
        prior["partId"] = "study-draft-part-v2_old"
        prior["inputEnvelopeSha256"] = "old-envelope"

        with tempfile.TemporaryDirectory() as temp_dir:
            parent = Path(temp_dir)
            prior_path = parent / (
                "study-draft-part-v2_old.fragment.rejected.json"
            )
            prior_path.write_text(
                json.dumps(prior),
                encoding="utf-8",
            )
            target = parent / (
                "study-draft-part-v2_current.fragment.json"
            )
            ai_calls: list[str] = []
            result = runner.run_codex_study_fragment_v2(
                envelope=envelope,
                all_selected_chunks=[chunk],
                output_path=target,
                codex_command=["codex"],
                run_command=mock.Mock(
                    side_effect=AssertionError(
                        "Validated prior fragment must skip the model"
                    )
                ),
                ai_call_observer=lambda: ai_calls.append("called"),
            )

            self.assertTrue(target.is_file())
            self.assertEqual(envelope["planId"], result["planId"])
            self.assertEqual(envelope["partId"], result["partId"])
            self.assertEqual([], ai_calls)

    def test_fragment_transport_schema_is_recursively_strict(
        self,
    ) -> None:
        schema = runner.fragment_output_schema_v2()

        def assert_strict_objects(
            value: object,
            path: str,
        ) -> None:
            if isinstance(value, dict):
                if value.get("type") == "object":
                    properties = value.get("properties")
                    self.assertIsInstance(
                        properties,
                        dict,
                        msg=path,
                    )
                    self.assertIs(
                        False,
                        value.get("additionalProperties"),
                        msg=path,
                    )
                    self.assertEqual(
                        set(properties),
                        set(value.get("required", [])),
                        msg=path,
                    )
                for key, child in value.items():
                    assert_strict_objects(child, f"{path}.{key}")
            elif isinstance(value, list):
                for index, child in enumerate(value):
                    assert_strict_objects(
                        child,
                        f"{path}[{index}]",
                    )

        assert_strict_objects(schema, "$")
        record_properties = schema["properties"]["records"]["items"][
            "properties"
        ]
        self.assertNotIn("payload", record_properties)
        self.assertEqual(
            {"type": "string"},
            record_properties["payloadJson"],
        )

    def test_runner_decodes_free_form_payload_json_fail_closed(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "Study")])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(
            universe,
            registry,
            plan,
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        payload = {
            "title": "Free-form study",
            "openDomain": "VP+CD",
            "nested": {
                "items": [
                    {
                        "meaning": "preserved",
                        "flags": [True, None],
                    }
                ]
            },
        }
        record = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="STUDY_PATCH",
            identity_key=chunk["cells"][0]["sourceCellKey"],
            label="Study",
            payload=payload,
            evidence_range="A1",
        )
        record_id = record["recordId"]
        transport = self.fragment(
            envelope,
            [record],
            dispositions=[
                {
                    "sourceCellKey": chunk["cells"][0]["sourceCellKey"],
                    "disposition": "RECORD_EVIDENCE",
                    "recordIds": [record_id],
                    "reason": "Fixture evidence.",
                }
            ],
        )
        transport["source"] = {
            "revisionUid": self.source["revisionUid"],
            "contentSha256": self.source["contentSha256"],
            "contentComplete": False,
        }
        transport_record = transport["records"][0]
        transport_record["payloadJson"] = json.dumps(
            transport_record.pop("payload"),
            ensure_ascii=False,
            separators=(",", ":"),
        )
        transport_record["evidence"][0].update(
            {"sourceText": "Study", "note": ""}
        )

        def result_runner(
            value: dict,
        ):
            def fake_run(
                command: list[str],
                **_kwargs: object,
            ) -> subprocess.CompletedProcess[str]:
                output_path = Path(
                    command[
                        command.index("--output-last-message") + 1
                    ]
                )
                output_path.write_text(
                    json.dumps(value, ensure_ascii=False),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess(
                    command,
                    0,
                    "",
                    "",
                )

            return fake_run

        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "fragment.json"
            result = runner.run_codex_study_fragment_v2(
                envelope=envelope,
                all_selected_chunks=[chunk],
                output_path=target,
                codex_command=["codex"],
                run_command=result_runner(transport),
            )
            self.assertEqual(payload, result["records"][0]["payload"])
            self.assertNotIn(
                "payloadJson",
                result["records"][0],
            )
            self.assertTrue(target.is_file())

        self.assertIn("payloadJson", envelope["promptText"])
        self.assertEqual(
            envelope["inputHashes"]["promptSha256"],
            staged.bytes_sha256(
                envelope["promptText"].encode("utf-8")
            ),
        )

        invalid_transport = copy.deepcopy(transport)
        invalid_transport["records"][0]["payloadJson"] = "[]"
        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "invalid.json"
            with self.assertRaisesRegex(
                staged.StagedDraftV2Error,
                "must encode an object",
            ):
                runner.run_codex_study_fragment_v2(
                    envelope=envelope,
                    all_selected_chunks=[chunk],
                    output_path=target,
                    codex_command=["codex"],
                    run_command=result_runner(invalid_transport),
                )
            self.assertFalse(target.exists())

    def test_registry_links_continuations_only_by_exact_structure(
        self,
    ) -> None:
        header = self.cell("A1", "Shared header")
        first = self.chunk(
            "c1",
            [header, self.cell("B1", "Arm A")],
        )
        second = self.chunk(
            "c2",
            [self.cell("A2", "Arm B")],
            context=[copy.deepcopy(header)],
        )
        locators = [
            self.locator(first),
            self.locator(second),
        ]
        universe, registry, plan = self.planned(
            [first, second],
            locators,
        )

        self.assertEqual(1, len(registry["studies"]))
        self.assertEqual(1, len(registry["linkProposals"]))
        self.assertEqual(
            ["c1", "c2"],
            universe["selectedChunkIds"],
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        self.assertTrue(
            all(
                logical_id in part["logicalStudyIds"]
                for part in plan["parts"]
            )
        )
        self.assertIn(
            header["sourceCellKey"],
            plan["parts"][1]["sharedAnchorCellKeys"],
        )

        disjoint_second = self.chunk(
            "c2",
            [self.cell("A2", "Unrelated")],
        )
        disjoint_universe = self.universe(
            [first, disjoint_second],
            [
                self.locator(first),
                self.locator(disjoint_second),
            ],
        )
        disjoint_registry = staged.build_study_registry_v2(
            source=self.source,
            universe=disjoint_universe,
        )
        self.assertEqual(2, len(disjoint_registry["studies"]))

        anchors = disjoint_registry["candidateAnchors"]
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "does not touch candidate",
        ):
            staged.build_study_registry_v2(
                source=self.source,
                universe=disjoint_universe,
                link_proposals=[
                    {
                        "memberCandidateAnchorIds": [
                            anchor["candidateAnchorId"]
                            for anchor in anchors
                        ],
                        "linkEvidence": [
                            {
                                "sheet": "Data",
                                "range": "A1",
                            }
                        ],
                    }
                ],
            )

    def test_registry_uses_spanning_links_for_dense_candidate_overlap(
        self,
    ) -> None:
        chunk = self.chunk(
            "dense",
            [
                self.cell("A1", "Header"),
                self.cell("A2", "Arm A"),
                self.cell("A3", "Arm B"),
            ],
        )
        locator = self.locator(chunk)
        locator["candidates"] = [
            {
                "key": f"candidate-{index}",
                "title": f"Candidate {index}",
                "evidence": [
                    {
                        "sheet": "Data",
                        "range": "A1:A3",
                        "role": "CANDIDATE_REGION",
                    }
                ],
            }
            for index in range(3)
        ]
        registry = staged.build_study_registry_v2(
            source=self.source,
            universe=self.universe([chunk], [locator]),
        )

        self.assertEqual(1, len(registry["studies"]))
        self.assertEqual(2, len(registry["linkProposals"]))
        self.assertTrue(
            all(
                len(proposal["linkEvidence"]) == 1
                and len(proposal["linkEvidenceCellKeys"]) == 1
                for proposal in registry["linkProposals"]
            )
        )

    def test_no_candidate_and_unselected_inventory_fail_closed(
        self,
    ) -> None:
        singleton = self.chunk(
            "single",
            [self.cell("A1", 19000)],
            section=1,
        )
        audit = staged.audit_no_candidate_source_inventory(
            packet_set={"chunks": [singleton]},
            locator_results=[self.locator(singleton, candidate=False)],
        )
        self.assertEqual(1, audit["requiredCellCount"])

        detached = self.chunk(
            "tail",
            [
                self.cell("A1", 1),
                self.cell("B1", 2),
                self.cell("A20", 31001),
                self.cell("A21", 41002),
            ],
        )
        tail_audit = staged.audit_no_candidate_source_inventory(
            packet_set={"chunks": [detached]},
            locator_results=[self.locator(detached, candidate=False)],
        )
        excluded_coordinates = {
            item["coordinate"]
            for item in tail_audit["excludedCells"]
        }
        self.assertEqual({"A20", "A21"}, excluded_coordinates)

        formula = self.chunk(
            "formula",
            [
                self.cell(
                    "A1",
                    None,
                    formula="=SUM(B1:B2)",
                    cached_value=None,
                )
            ],
        )
        formula_audit = staged.audit_no_candidate_source_inventory(
            packet_set={"chunks": [formula]},
            locator_results=[self.locator(formula, candidate=False)],
        )
        self.assertEqual(1, formula_audit["requiredCellCount"])

        categorical = self.chunk(
            "categorical",
            [
                self.cell("A1", "Result"),
                self.cell("B2", "PASSED"),
                self.cell("B3", "FAILED"),
            ],
        )
        categorical_audit = (
            staged.audit_no_candidate_source_inventory(
                packet_set={"chunks": [categorical]},
                locator_results=[
                    self.locator(categorical, candidate=False)
                ],
            )
        )
        self.assertEqual(
            {"B2", "B3"},
            {
                item["coordinate"]
                for item in categorical_audit["requiredCells"]
                if item.get("contentClass")
                == "CATEGORICAL_RESULT"
            },
        )

        semantic = self.chunk(
            "semantic",
            [self.cell("C40", "Assembly method")],
            section=4,
        )
        semantic_audit = staged.audit_no_candidate_source_inventory(
            packet_set={"chunks": [semantic]},
            locator_results=[
                self.locator(semantic, candidate=False)
            ],
        )
        self.assertEqual(
            ["C40"],
            [
                item["coordinate"]
                for item in semantic_audit["requiredCells"]
                if item.get("contentClass") == "SEMANTIC_LABEL"
            ],
        )

        selected = self.chunk(
            "selected",
            [self.cell("A1", "candidate")],
            section=1,
        )
        missed = self.chunk(
            "missed",
            [self.cell("B20", 7)],
            section=2,
        )
        partial = staged.audit_unselected_source_inventory(
            packet_set={"chunks": [selected, missed]},
            locator_results=[
                self.locator(selected),
                self.locator(missed, candidate=False),
            ],
            selected_source_cell_keys=[
                selected["cells"][0]["sourceCellKey"]
            ],
        )
        self.assertEqual(
            [missed["cells"][0]["sourceCellKey"]],
            [
                item["sourceCellKey"]
                for item in partial["requiredCells"]
            ],
        )

        missed_status = self.chunk(
            "missed-status",
            [self.cell("B30", "NG")],
            section=3,
        )
        categorical_partial = (
            staged.audit_unselected_source_inventory(
                packet_set={
                    "chunks": [selected, missed_status]
                },
                locator_results=[
                    self.locator(selected),
                    self.locator(
                        missed_status,
                        candidate=False,
                    ),
                ],
                selected_source_cell_keys=[
                    selected["cells"][0]["sourceCellKey"]
                ],
            )
        )
        self.assertEqual(
            [missed_status["cells"][0]["sourceCellKey"]],
            [
                item["sourceCellKey"]
                for item in categorical_partial["requiredCells"]
            ],
        )

        semantic_partial = (
            staged.audit_unselected_source_inventory(
                packet_set={"chunks": [selected, semantic]},
                locator_results=[
                    self.locator(selected),
                    self.locator(semantic, candidate=False),
                ],
                selected_source_cell_keys=[
                    selected["cells"][0]["sourceCellKey"]
                ],
            )
        )
        self.assertEqual(
            [semantic["cells"][0]["sourceCellKey"]],
            [
                item["sourceCellKey"]
                for item in semantic_partial["requiredCells"]
                if item.get("contentClass") == "SEMANTIC_LABEL"
            ],
        )

    def test_required_source_promotion_selects_missed_section(
        self,
    ) -> None:
        selected = self.chunk(
            "selected",
            [self.cell("A1", "candidate")],
            section=1,
        )
        missed = self.chunk(
            "missed",
            [self.cell("M73", 7)],
            section=2,
        )
        locator_results = [
            self.locator(selected),
            self.locator(missed, candidate=False),
        ]
        partial = staged.audit_unselected_source_inventory(
            packet_set={"chunks": [selected, missed]},
            locator_results=locator_results,
            selected_source_cell_keys=[
                selected["cells"][0]["sourceCellKey"]
            ],
        )

        promoted = staged.promote_required_source_locator_sections(
            locator_results=locator_results,
            required_cells=partial["requiredCells"],
        )
        promoted_locator = promoted[1]
        universe = staged.select_draft_universe(
            packet_set={"chunks": [selected, missed]},
            locator_results=promoted,
        )

        self.assertEqual("NEEDS_REVIEW", promoted_locator["status"])
        self.assertEqual(
            "M73",
            promoted_locator["candidates"][0]["evidence"][0]["range"],
        )
        self.assertEqual(
            [missed["cells"][0]["sourceCellKey"]],
            promoted_locator["deterministicCoveragePromotion"][
                "requiredSourceCellKeys"
            ],
        )
        self.assertEqual(
            ["selected", "missed"],
            universe["selectedChunkIds"],
        )

    def test_entity_ids_include_subtype_and_fragment_scope_is_exact(
        self,
    ) -> None:
        owned_cell = self.cell("A1", 10)
        shared_cell = self.cell("B1", 20)
        outside_cell = self.cell("C1", 30)
        first = self.chunk("c1", [owned_cell])
        second = self.chunk("c2", [shared_cell, outside_cell])
        universe, registry, plan = self.planned(
            [first, second],
            [
                self.locator(first),
                self.locator(second, candidate=False),
            ],
        )
        _part, envelope = self.envelope(
            universe,
            registry,
            plan,
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        envelope["sharedAnchorCellKeys"] = [
            shared_cell["sourceCellKey"]
        ]
        arm = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=owned_cell["sourceCellKey"],
            label="Result",
            payload={
                "entityType": "ARM",
                "key": "arm",
                "label": "Result",
            },
            evidence_range="A1",
        )
        outcome = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=owned_cell["sourceCellKey"],
            label="Result",
            payload={
                "entityType": "OUTCOME",
                "key": "outcome",
                "originalLabel": "Result",
                "metricType": "numeric_measurement",
            },
            evidence_range="A1",
        )
        self.assertNotEqual(arm["recordId"], outcome["recordId"])
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "without a canonical observation/series binding",
        ):
            staged.validate_fragment_v2(
                fragment=self.fragment(envelope, [arm]),
                envelope=envelope,
                all_selected_chunks=[first, second],
            )

        shared_observation = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=shared_cell["sourceCellKey"],
            label="20",
            payload={
                "outcome": "outcome",
                "arm": "arm",
                "valueNumber": 20,
            },
            evidence_range="B1",
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "shared-only numeric",
        ):
            staged.validate_fragment_v2(
                fragment=self.fragment(
                    envelope,
                    [shared_observation],
                ),
                envelope=envelope,
                all_selected_chunks=[first, second],
            )

        series = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=owned_cell["sourceCellKey"],
            label="Series",
            payload={
                "outcome": "outcome",
                "arm": "arm",
                "sheet": "Data",
                "headerRange": "A1",
                "valueRange": "B1",
                "rowIdentityRange": "A1",
                "axisSource": "ROW_IDENTITY",
            },
            evidence_range="A1:B1",
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "valueRange must include an owned",
        ):
            staged.validate_fragment_v2(
                fragment=self.fragment(envelope, [series]),
                envelope=envelope,
                all_selected_chunks=[first, second],
            )

        outside = copy.deepcopy(arm)
        outside["payload"]["sheet"] = "Data"
        outside["payload"]["sourceRange"] = "C1"
        outside["recordId"] = staged.stable_record_id(
            revision_uid=self.source["revisionUid"],
            logical_study_id=logical_id,
            record_type=outside["recordType"],
            identity_cell_keys=outside["identityCellKeys"],
            exact_source_label=outside["exactSourceLabel"],
            semantic_subtype="ARM",
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "exceeds owned/shared scope",
        ):
            staged.validate_fragment_v2(
                fragment=self.fragment(envelope, [outside]),
                envelope=envelope,
                all_selected_chunks=[first, second],
            )

        owned_observation = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=owned_cell["sourceCellKey"],
            label="10",
            payload={
                "outcome": "outcome",
                "arm": "arm",
                "valueNumber": 10,
            },
            evidence_range="A1",
        )
        unknown_disposition = self.fragment(
            envelope,
            [arm, outcome, owned_observation],
            dispositions=[
                {
                    "sourceCellKey": owned_cell["sourceCellKey"],
                    "disposition": "RECORD_EVIDENCE",
                    "recordIds": ["unknown-record"],
                    "reason": "Fixture.",
                }
            ],
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "disposition records",
        ):
            staged.validate_fragment_v2(
                fragment=unknown_disposition,
                envelope=envelope,
                all_selected_chunks=[first, second],
            )

    def test_normalize_fragment_evidence_dispositions_materializes_links(
        self,
    ) -> None:
        value_cell = self.cell("A1", 10)
        context_cell = self.cell("B1", "Lot A")
        chunk = self.chunk("c1", [value_cell, context_cell])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        observation = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=value_cell["sourceCellKey"],
            label="10",
            payload={
                "outcome": "height",
                "arm": "lot-a",
                "valueNumber": 10,
            },
            evidence_range="A1",
        )
        fragment = self.fragment(
            envelope,
            [observation],
            dispositions=[
                {
                    "sourceCellKey": value_cell["sourceCellKey"],
                    "disposition": "CONTEXT_ONLY",
                    "recordIds": [],
                    "reason": "Incorrect model classification.",
                },
                {
                    "sourceCellKey": context_cell["sourceCellKey"],
                    "disposition": "RECORD_EVIDENCE",
                    "recordIds": [observation["recordId"]],
                    "reason": "Row identity.",
                },
            ],
        )

        normalized = staged.normalize_fragment_evidence_dispositions(
            fragment=fragment,
            all_selected_chunks=[chunk],
        )

        value_disposition, context_disposition = normalized[
            "coverageDispositions"
        ]
        self.assertEqual("RECORD_EVIDENCE", value_disposition["disposition"])
        self.assertEqual(
            [observation["recordId"]],
            value_disposition["recordIds"],
        )
        self.assertEqual(
            "RECORD_EVIDENCE",
            context_disposition["disposition"],
        )
        self.assertEqual(
            [observation["recordId"]],
            context_disposition["recordIds"],
        )
        self.assertEqual(
            "B1",
            normalized["records"][0]["evidence"][-1]["range"],
        )
        self.assertEqual(
            "DECLARED_SOURCE",
            normalized["records"][0]["evidence"][-1]["role"],
        )

    def test_normalize_required_outcome_and_exact_series_header(
        self,
    ) -> None:
        header = self.cell("H13", "Input")
        value = self.cell("H15", 798)
        chunk = self.chunk("c1", [header, value])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        outcome = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=header["sourceCellKey"],
            label="Input",
            payload={
                "entityType": "OUTCOME",
                "key": "input",
            },
            evidence_range="H13",
        )
        series = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=header["sourceCellKey"],
            label="Input",
            payload={
                "key": "input-series",
                "sheet": "Data",
                "headerRange": "H14",
                "valueRange": "H15",
            },
            evidence_range="H13:H15",
        )
        series["evidence"] = [
            {
                "sheet": "Data",
                "range": "H13:H15",
                "role": "HEADER_AND_VALUES",
                "sourceText": "Input quantities for the reported rows",
            },
        ]

        normalized = (
            staged.normalize_fragment_required_fields_and_series_headers(
                fragment={"records": [outcome, series]},
                all_selected_chunks=universe["selectedChunks"],
            )
        )

        self.assertEqual(
            "Input",
            normalized["records"][0]["payload"]["originalLabel"],
        )
        self.assertEqual(
            "source_labeled_result",
            normalized["records"][0]["payload"]["metricType"],
        )
        self.assertEqual(
            "H13",
            normalized["records"][1]["payload"]["headerRange"],
        )

    def test_normalize_series_header_placeholder_to_merged_anchor(
        self,
    ) -> None:
        header = self.cell("F3", "100~750 Hz")
        header["mergeRange"] = "F3:G4"
        value = self.cell("G8", 108.4)
        chunk = self.chunk("c1", [header, value])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        series = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=value["sourceCellKey"],
            label="100~750 Hz difference",
            payload={
                "key": "difference-series",
                "sheet": "Data",
                "headerRange": "G3",
                "valueRange": "G8",
            },
            evidence_range="G8",
        )

        normalized = (
            staged.normalize_fragment_required_fields_and_series_headers(
                fragment={"records": [series]},
                all_selected_chunks=universe["selectedChunks"],
            )
        )

        self.assertEqual(
            "F3:G4",
            normalized["records"][0]["payload"]["headerRange"],
        )

    def test_normalize_derived_series_header_from_aggregate_source(
        self,
    ) -> None:
        aggregate_header = self.cell("B50", "STD_AVG")
        identity = self.cell("G53", 1060)
        aggregate_value = self.cell("H53", 115.58)
        derived_value = self.cell("N53", 113.58)
        chunk = self.chunk(
            "c1",
            [
                aggregate_header,
                identity,
                aggregate_value,
                derived_value,
            ],
        )
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        aggregate = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=aggregate_header["sourceCellKey"],
            label="STD_AVG",
            payload={
                "key": "std-average",
                "sheet": "Data",
                "headerRange": "B50",
                "valueRange": "H53",
                "rowIdentityRange": "G53",
                "outcome": "spl",
                "arm": "std",
                "axisSource": "ROW_IDENTITY",
            },
            evidence_range="B50",
        )
        derived = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=derived_value["sourceCellKey"],
            label="H minus applicable tolerance",
            payload={
                "key": "std-lower-limit",
                "sheet": "Data",
                "headerRange": "N8",
                "valueRange": "N53",
                "rowIdentityRange": "G53",
                "outcome": "spl",
                "arm": "std",
                "axisSource": "ROW_IDENTITY",
                "aggregateOfSeries": "std-average",
            },
            evidence_range="N53",
        )

        normalized = (
            staged.normalize_fragment_required_fields_and_series_headers(
                fragment={"records": [aggregate, derived]},
                all_selected_chunks=universe["selectedChunks"],
            )
        )

        self.assertEqual(
            "B50",
            normalized["records"][1]["payload"]["headerRange"],
        )

    def test_normalize_wide_series_header_from_complete_context_row(
        self,
    ) -> None:
        context_headers = [
            self.cell("R4", 107.9),
            self.cell("S4", 107.8),
            self.cell("T4", 107.7),
        ]
        for cell in context_headers:
            cell["contextOnly"] = True
            cell["primary"] = False
        identity = self.cell("G50", 1060)
        values = [
            self.cell("R50", 115.5),
            self.cell("S50", 115.6),
            self.cell("T50", 115.7),
        ]
        chunk = self.chunk(
            "c1",
            [*context_headers, identity, *values],
        )
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        series = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=values[0]["sourceCellKey"],
            label="Normal raw segment",
            payload={
                "key": "normal-raw",
                "sheet": "Data",
                "headerRange": "R8:T8",
                "valueRange": "R50:T50",
                "rowIdentityRange": "G50",
                "outcome": "spl",
                "arm": "normal",
                "axisSource": "ROW_IDENTITY",
            },
            evidence_range="R50:T50",
        )

        normalized = (
            staged.normalize_fragment_required_fields_and_series_headers(
                fragment={"records": [series]},
                all_selected_chunks=universe["selectedChunks"],
            )
        )

        self.assertEqual(
            "R4:T4",
            normalized["records"][0]["payload"]["headerRange"],
        )

    def test_normalize_multi_arm_row_series_to_exact_arm_segments(
        self,
    ) -> None:
        header = self.cell("F3", "100~750 Hz")
        identities = [
            self.cell("C8", "STD"),
            self.cell("C9", "Normal"),
            self.cell("C10", "Frame V4"),
        ]
        values = [
            self.cell("F8", 108.12),
            self.cell("F9", 107.61),
            self.cell("F10", 107.52),
        ]
        chunk = self.chunk("c1", [header, *identities, *values])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        arm_records = []
        for identity, arm_key in zip(
            identities,
            ("std", "normal", "frame-v4"),
            strict=True,
        ):
            arm_records.append(
                self.record(
                    envelope=envelope,
                    logical_id=logical_id,
                    record_type="ENTITY_DECLARATION",
                    identity_key=identity["sourceCellKey"],
                    label=str(identity["displayValue"]),
                    payload={
                        "entityType": "ARM",
                        "key": arm_key,
                        "label": str(identity["displayValue"]),
                        "role": "OTHER",
                    },
                    evidence_range=identity["coordinate"],
                )
            )
        series = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="SERIES_SEGMENT_APPEND",
            identity_key=values[0]["sourceCellKey"],
            label="100~750 Hz measured SPL by configuration",
            payload={
                "key": "spl-series",
                "seriesRole": "RAW",
                "aggregationFunction": None,
                "aggregateOfSeries": None,
                "outcome": "spl",
                "arm": None,
                "sheet": "Data",
                "headerRange": "F3",
                "valueRange": "F8:F10",
                "rowIdentityRange": "C8:C10",
                "aggregateReplicateRanges": [],
                "axisSource": "ROW_IDENTITY",
            },
            evidence_range="C8:F10",
        )

        normalized = staged.normalize_fragment_multi_arm_series_rows(
            fragment={"records": [*arm_records, series]},
            all_selected_chunks=universe["selectedChunks"],
        )
        split_series = [
            item
            for item in normalized["records"]
            if item["recordType"] == "SERIES_SEGMENT_APPEND"
        ]

        self.assertEqual(3, len(split_series))
        self.assertEqual(
            ["std", "normal", "frame-v4"],
            [item["payload"]["arm"] for item in split_series],
        )
        self.assertEqual(
            ["F8", "F9", "F10"],
            [item["payload"]["valueRange"] for item in split_series],
        )
        self.assertEqual(
            ["C8", "C9", "C10"],
            [
                item["payload"]["rowIdentityRange"]
                for item in split_series
            ],
        )

    def test_normalize_complete_dispositions_adds_every_owned_cell(
        self,
    ) -> None:
        value_cell = self.cell("A1", 10)
        context_cell = self.cell("B1", "context")
        chunk = self.chunk("c1", [value_cell, context_cell])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        observation = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=value_cell["sourceCellKey"],
            label="10",
            payload={
                "outcome": "height",
                "arm": "lot-a",
                "valueNumber": 10,
            },
            evidence_range="A1",
        )

        normalized = staged.normalize_fragment_complete_dispositions(
            fragment={
                "records": [observation],
                "coverageDispositions": [],
            },
            envelope=envelope,
            all_selected_chunks=universe["selectedChunks"],
        )

        self.assertEqual(
            envelope["ownedSourceCellKeys"],
            [
                item["sourceCellKey"]
                for item in normalized["coverageDispositions"]
            ],
        )
        self.assertEqual(
            "RECORD_EVIDENCE",
            normalized["coverageDispositions"][0]["disposition"],
        )
        self.assertEqual(
            "CONTEXT_ONLY",
            normalized["coverageDispositions"][1]["disposition"],
        )

    def test_arm_sample_size_accepts_exact_owned_unit_text(
        self,
    ) -> None:
        source_cell = self.cell("A1", "semi VP+CD: 10pcs")
        chunk = self.chunk("c1", [source_cell])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        arm = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=source_cell["sourceCellKey"],
            label="semi VP+CD: 10pcs",
            payload={
                "entityType": "ARM",
                "key": "semi-vp-cd",
                "label": "semi VP+CD: 10pcs",
                "sampleSize": 10,
            },
            evidence_range="A1",
        )
        disposition = [
            {
                "sourceCellKey": source_cell["sourceCellKey"],
                "disposition": "RECORD_EVIDENCE",
                "recordIds": [arm["recordId"]],
                "reason": "",
            }
        ]

        validated = staged.validate_fragment_v2(
            fragment=self.fragment(
                envelope,
                [arm],
                dispositions=disposition,
            ),
            envelope=envelope,
            all_selected_chunks=[chunk],
        )
        self.assertEqual(10, validated["records"][0]["payload"]["sampleSize"])

        wrong = copy.deepcopy(arm)
        wrong["payload"]["sampleSize"] = 11
        wrong["recordId"] = staged.stable_record_id(
            revision_uid=self.source["revisionUid"],
            logical_study_id=logical_id,
            record_type=wrong["recordType"],
            identity_cell_keys=wrong["identityCellKeys"],
            exact_source_label=wrong["exactSourceLabel"],
            semantic_subtype="ARM",
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "numeric claim lacks an owned value cell",
        ):
            staged.validate_fragment_v2(
                fragment=self.fragment(
                    envelope,
                    [wrong],
                    dispositions=[
                        {
                            "sourceCellKey": source_cell["sourceCellKey"],
                            "disposition": "RECORD_EVIDENCE",
                            "recordIds": [wrong["recordId"]],
                            "reason": "",
                        }
                    ],
                ),
                envelope=envelope,
                all_selected_chunks=[chunk],
            )

    def test_embedded_ratio_keeps_text_but_not_numeric_claims(
        self,
    ) -> None:
        narrative_cell = self.cell(
            "A1",
            "Bond 201 around enclosure assy ok 0/20",
        )
        exact_ratio_cell = self.cell("A2", "1/8 pcs")
        chunk = self.chunk("c1", [narrative_cell, exact_ratio_cell])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        arm = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=narrative_cell["sourceCellKey"],
            label="Bond 201",
            payload={
                "entityType": "ARM",
                "key": "bond-201",
                "label": "Bond 201",
                "sampleSize": 20,
                "sampleBasis": "Embedded ratio",
            },
            evidence_range="A1",
        )
        embedded = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=narrative_cell["sourceCellKey"],
            label="ok 0/20",
            payload={
                "outcome": "reported-result",
                "arm": "bond-201",
                "valueNumber": None,
                "valueText": "ok 0/20",
                "numerator": 0,
                "denominator": 20,
                "sampleSize": 20,
            },
            evidence_range="A1",
        )
        exact = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=exact_ratio_cell["sourceCellKey"],
            label="1/8 pcs",
            payload={
                "outcome": "exact-result",
                "arm": "bond-201",
                "valueNumber": None,
                "valueText": "1/8 pcs",
                "numerator": 1,
                "denominator": 8,
                "sampleSize": 8,
            },
            evidence_range="A2",
        )

        normalized = (
            staged.normalize_fragment_unsupported_text_numeric_claims(
                fragment={"records": [arm, embedded, exact]},
                envelope=envelope,
                all_selected_chunks=universe["selectedChunks"],
            )
        )

        self.assertIsNone(
            normalized["records"][0]["payload"]["sampleSize"]
        )
        self.assertEqual(
            "",
            normalized["records"][0]["payload"]["sampleBasis"],
        )
        self.assertEqual(
            "ok 0/20",
            normalized["records"][1]["payload"]["valueText"],
        )
        self.assertIsNone(
            normalized["records"][1]["payload"]["numerator"]
        )
        self.assertIsNone(
            normalized["records"][1]["payload"]["denominator"]
        )
        self.assertEqual(
            1,
            normalized["records"][2]["payload"]["numerator"],
        )
        self.assertEqual(
            8,
            normalized["records"][2]["payload"]["denominator"],
        )

    def test_percent_source_binds_exact_scaled_rate_claims(
        self,
    ) -> None:
        source_cell = {"numberFormat": "0.0%"}
        self.assertTrue(
            staged._observation_claim_matches_numeric_source(
                payload={"ratePpm": 562_500},
                source_number=0.5625,
                source_cell=source_cell,
            )
        )
        self.assertTrue(
            staged._observation_claim_matches_numeric_source(
                payload={"valueNumber": 56.25},
                source_number=0.5625,
                source_cell=source_cell,
            )
        )
        self.assertFalse(
            staged._observation_claim_matches_numeric_source(
                payload={"valueNumber": 56.3},
                source_number=0.5625,
                source_cell=source_cell,
            )
        )

    def test_observation_replicate_key_adds_same_row_identity_evidence(
        self,
    ) -> None:
        identity_cell = self.cell("B5", 1.2)
        label_cell = self.cell("E5", "Hearing 8V")
        value_cell = self.cell("K5", "Pass")
        chunk = self.chunk(
            "replicate-row",
            [identity_cell, label_cell, value_cell],
        )
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        observation = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=value_cell["sourceCellKey"],
            label="Hearing result",
            payload={
                "outcome": "hearing-result",
                "replicateKey": "1.2-Hearing 8V",
                "valueText": "Pass",
            },
            evidence_range="E5:K5",
        )

        normalized = (
            staged.normalize_fragment_observation_replicate_evidence(
                fragment={"records": [observation]},
                envelope=envelope,
                all_selected_chunks=universe["selectedChunks"],
            )
        )

        self.assertTrue(
            staged._replicate_key_preserves_source_identity(
                "1.2-Hearing 8V",
                "1.2",
            )
        )
        self.assertFalse(
            staged._replicate_key_preserves_source_identity(
                "11.2-Hearing 8V",
                "1.2",
            )
        )
        identity_evidence = [
            item
            for item in normalized["records"][0]["evidence"]
            if item.get("role") == "REPLICATE_IDENTITY"
        ]
        self.assertEqual(1, len(identity_evidence))
        self.assertEqual("B5", identity_evidence[0]["range"])
        self.assertEqual("1.2", identity_evidence[0]["sourceText"])

    def test_numeric_arm_entity_preserves_exact_row_identity(
        self,
    ) -> None:
        identity_cell = self.cell("B6", 1.2)
        chunk = self.chunk("numeric-arm", [identity_cell])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        arm = self.record(
            envelope=envelope,
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=identity_cell["sourceCellKey"],
            label="1.2",
            payload={
                "entityType": "ARM",
                "key": "row-1-2",
                "label": "1.2",
                "role": "OTHER",
            },
            evidence_range="B6",
        )
        fragment = self.fragment(
            envelope,
            [arm],
            dispositions=[
                {
                    "sourceCellKey": identity_cell["sourceCellKey"],
                    "disposition": "RECORD_EVIDENCE",
                    "recordIds": [arm["recordId"]],
                    "reason": "",
                }
            ],
        )

        validated = staged.validate_fragment_v2(
            fragment=fragment,
            envelope=envelope,
            all_selected_chunks=universe["selectedChunks"],
        )

        self.assertEqual(1, len(validated["records"]))

    def test_arm_less_observations_use_one_exact_descriptive_label(
        self,
    ) -> None:
        label_cell = self.cell("C20", "2. Particle")
        value_cell = self.cell("G20", 1)
        chunk = self.chunk("c1", [label_cell, value_cell])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        logical_id = registry["studies"][0]["logicalStudyId"]
        observations = [
            self.record(
                envelope=envelope,
                logical_id=logical_id,
                record_type="OBSERVATION_APPEND",
                identity_key=value_cell["sourceCellKey"],
                label=str(index),
                payload={
                    "outcome": f"particle-{index}",
                    "arm": None,
                    "valueNumber": 1,
                },
                evidence_range="C20:G20",
            )
            for index in (1, 2)
        ]

        normalized = staged.normalize_fragment_missing_observation_arms(
            fragment={"records": observations},
            envelope=envelope,
            all_selected_chunks=universe["selectedChunks"],
        )

        self.assertEqual(3, len(normalized["records"]))
        arm = normalized["records"][2]
        self.assertEqual("ENTITY_DECLARATION", arm["recordType"])
        self.assertEqual("ARM", arm["payload"]["entityType"])
        self.assertEqual("2. Particle", arm["payload"]["label"])
        self.assertEqual(
            arm["payload"]["key"],
            normalized["records"][0]["payload"]["arm"],
        )
        self.assertEqual(
            arm["payload"]["key"],
            normalized["records"][1]["payload"]["arm"],
        )

    def test_result_table_merged_header_outcomes_have_unique_identity(
        self,
    ) -> None:
        merged_header = self.cell("B2", "Peak")
        merged_header["mergeRange"] = "B2:C2"
        chunk = self.chunk(
            "result-table",
            [
                self.cell("A1", "RESULT TENSION"),
                self.cell("A2", "Sample"),
                merged_header,
                self.cell("A3", "Normal"),
                self.cell("B3", 1.1),
                self.cell("C3", 1.2),
            ],
        )
        locator = self.locator(chunk)
        universe, registry, plan = self.planned(
            [chunk],
            [locator],
        )
        _part, envelope = self.envelope(
            universe,
            registry,
            plan,
        )

        fragment = staged.build_deterministic_result_table_fragment_v2(
            envelope=envelope,
            all_selected_chunks=[chunk],
        )

        self.assertIsNotNone(fragment)
        assert fragment is not None
        outcomes = [
            record
            for record in fragment["records"]
            if record["recordType"] == "ENTITY_DECLARATION"
            and record["payload"]["entityType"] == "OUTCOME"
            and record["payload"]["key"]
            in {"result_c2_value", "result_c3_value"}
        ]
        self.assertEqual(2, len(outcomes))
        self.assertEqual(
            2,
            len({record["recordId"] for record in outcomes}),
        )
        self.assertEqual(
            {
                ("revision-1:1:B2", "revision-1:1:B3"),
                ("revision-1:1:B2", "revision-1:1:C3"),
            },
            {
                tuple(record["identityCellKeys"])
                for record in outcomes
            },
        )

    def test_result_table_continuation_uses_shared_result_title(
        self,
    ) -> None:
        chunk = self.chunk(
            "result-table-continuation",
            [
                self.cell("D53", "#6"),
                self.cell("E53", 1.1),
                self.cell("F53", 1.2),
            ],
            context=[
                self.cell("B2", "RELIABILITY TEST RESULT"),
            ],
        )
        locator = self.locator(chunk)
        universe, registry, plan = self.planned(
            [chunk],
            [locator],
        )
        _part, envelope = self.envelope(
            universe,
            registry,
            plan,
        )

        fragment = staged.build_deterministic_result_table_fragment_v2(
            envelope=envelope,
            all_selected_chunks=[chunk],
        )

        self.assertIsNotNone(fragment)
        assert fragment is not None
        observations = [
            record
            for record in fragment["records"]
            if record["recordType"] == "OBSERVATION_APPEND"
        ]
        self.assertEqual(2, len(observations))
        self.assertEqual(
            {
                "revision-1:1:E53",
                "revision-1:1:F53",
            },
            {
                record["identityCellKeys"][0]
                for record in observations
            },
        )
        outcomes = [
            record
            for record in fragment["records"]
            if record["recordType"] == "ENTITY_DECLARATION"
            and record["payload"]["entityType"] == "OUTCOME"
        ]
        self.assertEqual(
            {
                "RELIABILITY TEST RESULT | Column E",
                "RELIABILITY TEST RESULT | Column F",
            },
            {
                record["payload"]["originalLabel"]
                for record in outcomes
            },
        )

    def test_result_table_coalesces_chunks_and_prefers_specific_bounds(
        self,
    ) -> None:
        first = self.chunk(
            "result-first",
            [
                self.cell("A1", "RESULT CHECK"),
                self.cell("A2", "Item"),
                self.cell("B2", "Input"),
                self.cell("A3", "Specific"),
                self.cell("B3", 240),
            ],
        )
        second = self.chunk(
            "result-second",
            [
                self.cell("A4", "Overall"),
                self.cell("B4", 520),
            ],
            section=2,
            context=[self.cell("A1", "RESULT CHECK")],
        )
        chunks = [first, second]
        owned = [
            cell["sourceCellKey"]
            for chunk in chunks
            for cell in chunk["cells"]
        ]
        envelope = {
            "focusedChunks": chunks,
            "locatorResults": [
                self.locator(first),
                self.locator(second),
            ],
            "ownedSourceCellKeys": owned,
            "sharedAnchorCellKeys": [],
            "registry": {
                "studies": [
                    {
                        "logicalStudyId": "logical-overall",
                        "anchorEvidenceCellKeys": [
                            "revision-1:1:A1",
                            "revision-1:1:B4",
                        ],
                    },
                    {
                        "logicalStudyId": "logical-specific",
                        "anchorEvidenceCellKeys": [
                            "revision-1:1:A2",
                            "revision-1:1:B3",
                        ],
                    },
                ]
            },
            "source": self.source,
            "planId": "plan-multi-result",
            "partId": "part-multi-result",
            "inputEnvelopeSha256": "fixture-envelope",
        }

        fragment = staged.build_deterministic_result_table_fragment_v2(
            envelope=envelope,
            all_selected_chunks=chunks,
        )

        self.assertIsNotNone(fragment)
        assert fragment is not None
        observations = {
            record["identityCellKeys"][0]: record
            for record in fragment["records"]
            if record["recordType"] == "OBSERVATION_APPEND"
        }
        self.assertEqual(
            "logical-specific",
            observations["revision-1:1:B3"]["logicalStudyId"],
        )
        self.assertEqual(
            "logical-overall",
            observations["revision-1:1:B4"]["logicalStudyId"],
        )
        self.assertEqual(
            owned,
            [
                item["sourceCellKey"]
                for item in fragment["coverageDispositions"]
            ],
        )

    def test_exact_mask_profile_builds_deterministic_series_fragment(
        self,
    ) -> None:
        def mask_cell(coordinate: str, value: object) -> dict:
            column = 31 if coordinate.startswith("AE") else 32
            row = int(coordinate[2:])
            return {
                "sourceCellKey": f"revision-1:1:{coordinate}",
                "row": row,
                "column": column,
                "coordinate": coordinate,
                "rawValue": value,
                "formula": "",
                "cachedValue": None,
                "displayValue": value,
                "dataType": (
                    "n" if isinstance(value, (int, float)) else "s"
                ),
                "cachedDataType": None,
                "numberFormat": "General",
            }

        cells = [
            mask_cell("AE8", "MASK"),
            mask_cell("AE9", "HZ"),
            mask_cell("AF9", "%"),
        ]
        for row, hz, percentage in (
            (10, 100, 90),
            (11, 240, 42),
            (12, 312, 35),
            (13, 500, 25),
            (14, 1000, 10),
            (15, 2000, 8),
            (16, 4500, 6),
            (17, 7000, 10),
            (18, 14000, 8),
        ):
            cells.extend(
                [
                    mask_cell(f"AE{row}", hz),
                    mask_cell(f"AF{row}", percentage),
                ]
            )
        chunk = self.chunk("mask", cells)
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        fragment = staged.build_deterministic_mask_fragment_v2(
            envelope=envelope,
            all_selected_chunks=universe["selectedChunks"],
        )
        self.assertIsNotNone(fragment)
        assert fragment is not None
        series = next(
            record
            for record in fragment["records"]
            if record["recordType"] == "SERIES_SEGMENT_APPEND"
        )
        self.assertEqual("Data", series["payload"]["sheet"])
        self.assertEqual("AF10:AF18", series["payload"]["valueRange"])
        self.assertEqual(
            "AE10:AE18",
            series["payload"]["rowIdentityRange"],
        )
        self.assertTrue(
            all(
                disposition["disposition"] == "RECORD_EVIDENCE"
                for disposition in fragment["coverageDispositions"]
            )
        )

        changed = copy.deepcopy(envelope)
        changed["focusedChunks"][0]["cells"][0]["rawValue"] = "MASS"
        self.assertIsNone(
            staged.build_deterministic_mask_fragment_v2(
                envelope=changed,
                all_selected_chunks=universe["selectedChunks"],
            )
        )

    def test_exact_fo_table_builds_cell_bound_observations(
        self,
    ) -> None:
        def grid_cell(coordinate: str, value: object) -> dict:
            column = ord(coordinate[0]) - ord("A") + 1
            row = int(coordinate[1:])
            return {
                "sourceCellKey": f"revision-1:1:{coordinate}",
                "row": row,
                "column": column,
                "coordinate": coordinate,
                "rawValue": value,
                "formula": "",
                "cachedValue": None,
                "displayValue": value,
                "dataType": (
                    "n" if isinstance(value, (int, float)) else "s"
                ),
                "cachedDataType": None,
                "numberFormat": "General",
            }

        cells = [grid_cell("B3", "RESULT CHECKING FO")]
        for row in range(5, 15):
            index = row - 4
            cells.extend(
                [
                    grid_cell(f"B{row}", f"Test 1 #{index}"),
                    grid_cell(f"C{row}", 640 + index),
                    grid_cell(f"D{row}", f"Test 2 #{index}"),
                    grid_cell(f"E{row}", 650 + index),
                    grid_cell(f"F{row}", f"Normal #{index}"),
                    grid_cell(f"G{row}", 660 + index),
                ]
            )
            if row == 5:
                cells.extend(
                    [
                        grid_cell("H5", "ST"),
                        grid_cell("I5", 670),
                    ]
                )
        cells.extend(
            [
                grid_cell("B15", "AVG test 1"),
                grid_cell("C15", 645.5),
                grid_cell("D15", "AVG test 2"),
                grid_cell("E15", 655.5),
                grid_cell("F15", "AVG normal"),
                grid_cell("G15", 665.5),
                grid_cell("H15", "AVG ST"),
                grid_cell("I15", 670),
            ]
        )
        chunk = self.chunk("fo", cells)
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        _part, envelope = self.envelope(universe, registry, plan)
        fragment = staged.build_deterministic_fo_fragment_v2(
            envelope=envelope,
            all_selected_chunks=universe["selectedChunks"],
        )
        self.assertIsNotNone(fragment)
        assert fragment is not None
        record_types = [
            record["recordType"] for record in fragment["records"]
        ]
        self.assertEqual(35, record_types.count("OBSERVATION_APPEND"))
        self.assertEqual(5, record_types.count("ENTITY_DECLARATION"))
        normal = next(
            record
            for record in fragment["records"]
            if record["recordType"] == "ENTITY_DECLARATION"
            and record["payload"].get("key") == "fo_normal"
        )
        self.assertEqual("REFERENCE", normal["payload"]["role"])
        self.assertEqual("Normal group", normal["payload"]["label"])
        self.assertTrue(
            all(
                disposition["disposition"] == "RECORD_EVIDENCE"
                for disposition in fragment["coverageDispositions"]
            )
        )

        changed = copy.deepcopy(envelope)
        changed["focusedChunks"][0]["cells"][0]["rawValue"] = "OTHER"
        self.assertIsNone(
            staged.build_deterministic_fo_fragment_v2(
                envelope=changed,
                all_selected_chunks=universe["selectedChunks"],
            )
        )

    def test_merge_is_order_independent_and_conflicts_fail(
        self,
    ) -> None:
        first = self.chunk("c1", [self.cell("A1", "one")])
        second = self.chunk("c2", [self.cell("A2", "two")])
        universe, registry, plan = self.planned(
            [first, second],
            [self.locator(first), self.locator(second, candidate=False)],
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        shared_identity = first["cells"][0]["sourceCellKey"]
        records = []
        fragments = []
        for index, part in enumerate(plan["parts"]):
            evidence_range = "A1" if index == 0 else "A2"
            payload = (
                {"title": "Merged study"}
                if index == 0
                else {"summary": "Merged summary"}
            )
            record = self.record(
                envelope={},
                logical_id=logical_id,
                record_type="STUDY_PATCH",
                identity_key=shared_identity,
                label="Study",
                payload=payload,
                evidence_range=evidence_range,
            )
            records.append(record)
            fragments.append(
                (
                    part,
                    {
                        "records": [record],
                        "coverageDispositions": [
                            {
                                "sourceCellKey": key,
                                "disposition": "RECORD_EVIDENCE",
                                "recordIds": [record["recordId"]],
                                "reason": "Fixture evidence.",
                            }
                            for key in part[
                                "ownedSourceCellKeys"
                            ]
                        ],
                    },
                )
            )
        shared_limitation = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="LIMITATION_APPEND",
            identity_key=shared_identity,
            label="Shared source limitation",
            payload={
                "text": "The second fragment also cites the shared anchor.",
                "scope": "STUDY",
            },
            evidence_range="A1:A2",
        )
        fragments[1][1]["records"].append(shared_limitation)
        fragments[1][1]["coverageDispositions"][0]["recordIds"].append(
            shared_limitation["recordId"]
        )
        forward = staged.merge_fragment_records(
            plan=plan,
            fragments=fragments,
            selected_chunks=universe["selectedChunks"],
        )
        reverse = staged.merge_fragment_records(
            plan=plan,
            fragments=list(reversed(fragments)),
            selected_chunks=universe["selectedChunks"],
        )
        self.assertEqual(forward, reverse)
        self.assertEqual(
            {
                "title": "Merged study",
                "summary": "Merged summary",
            },
            next(
                record["payload"]
                for record in forward["records"]
                if record["recordType"] == "STUDY_PATCH"
            ),
        )
        first_disposition = forward["coverageDispositions"][0]
        self.assertEqual(
            {
                records[0]["recordId"],
                shared_limitation["recordId"],
            },
            set(first_disposition["recordIds"]),
        )
        forward_manifest = staged.project_canonical_manifest(
            merged=forward,
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=universe["selectedChunks"],
        )
        reverse_manifest = staged.project_canonical_manifest(
            merged=reverse,
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=universe["selectedChunks"],
        )
        self.assertEqual(
            staged.json_sha256(forward_manifest),
            staged.json_sha256(reverse_manifest),
        )

        conflicted = copy.deepcopy(fragments)
        conflicted[1][1]["records"][0]["payload"] = {
            "title": "Conflicting title"
        }
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "Conflicting nonempty",
        ):
            staged.merge_fragment_records(
                plan=plan,
                fragments=conflicted,
                selected_chunks=universe["selectedChunks"],
            )

    def test_entity_merge_ignores_source_bound_entity_ids_only(
        self,
    ) -> None:
        logical_id = "logical-study-v2_fixture"
        first = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key="cell-a1",
            label="Peak",
            payload={
                "entityType": "OUTCOME",
                "key": "result_peak",
                "originalLabel": "Peak",
                "metricType": "measurement",
                "unit": "N",
            },
            evidence_range="A1",
        )
        second = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key="cell-b1",
            label="Peak",
            payload={
                "entityType": "OUTCOME",
                "key": "result_peak",
                "originalLabel": "Peak",
                "metricType": "measurement",
                "unit": "N",
            },
            evidence_range="B1",
        )
        first["payload"]["entityId"] = first["recordId"]
        second["payload"]["entityId"] = second["recordId"]
        second["payload"]["originalLabel"] = "Peak result"

        entities, _cross_index = staged._entity_maps([first, second])
        entity = entities[
            (logical_id, "OUTCOME", "result_peak")
        ]
        self.assertNotIn("entityId", entity["payload"])
        self.assertEqual("Peak", entity["payload"]["originalLabel"])
        self.assertEqual(2, len(entity["evidence"]))

        conflicted = copy.deepcopy(second)
        conflicted["payload"]["unit"] = "kgf"
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "Conflicting nonempty",
        ):
            staged._entity_maps([first, conflicted])

    def test_merge_coalesces_same_entity_key_across_source_sections(
        self,
    ) -> None:
        first_chunk = self.chunk(
            "c1",
            [self.cell("A1", "Result")],
        )
        second_chunk = self.chunk(
            "c2",
            [self.cell("A2", "Result")],
        )
        universe, registry, plan = self.planned(
            [first_chunk, second_chunk],
            [
                self.locator(first_chunk),
                self.locator(second_chunk, candidate=False),
            ],
            max_chunks=1,
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        records = [
            self.record(
                envelope={},
                logical_id=logical_id,
                record_type="ENTITY_DECLARATION",
                identity_key=chunk["cells"][0]["sourceCellKey"],
                label="Result",
                payload={
                    "entityType": "OUTCOME",
                    "key": "result_c5_value",
                    "originalLabel": "Result",
                    "metricType": "measurement",
                    "unit": "N",
                    "favorableDirection": "UNKNOWN",
                },
                evidence_range=f"A{index}",
            )
            for index, chunk in enumerate(
                (first_chunk, second_chunk),
                start=1,
            )
        ]
        fragments = [
            (
                part,
                {
                    "records": [record],
                    "coverageDispositions": [
                        {
                            "sourceCellKey": source_key,
                            "disposition": "RECORD_EVIDENCE",
                            "recordIds": [record["recordId"]],
                            "reason": "",
                        }
                        for source_key in part["ownedSourceCellKeys"]
                    ],
                },
            )
            for part, record in zip(plan["parts"], records)
        ]

        merged = staged.merge_fragment_records(
            plan=plan,
            fragments=fragments,
            selected_chunks=universe["selectedChunks"],
        )

        declarations = [
            record
            for record in merged["records"]
            if record["recordType"] == "ENTITY_DECLARATION"
        ]
        self.assertEqual(1, len(declarations))
        self.assertEqual(
            {"A1", "A2"},
            {
                evidence["range"]
                for evidence in declarations[0]["evidence"]
            },
        )

    def test_merge_canonicalizes_cross_part_entity_key_aliases(
        self,
    ) -> None:
        first_chunk = self.chunk("c1", [self.cell("A1", "Gap B-A")])
        second_chunk = self.chunk("c2", [self.cell("A2", "42")])
        universe, registry, plan = self.planned(
            [first_chunk, second_chunk],
            [
                self.locator(first_chunk),
                self.locator(second_chunk, candidate=False),
            ],
            max_chunks=1,
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        identity_key = first_chunk["cells"][0]["sourceCellKey"]
        value_key = second_chunk["cells"][0]["sourceCellKey"]

        canonical_outcome = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=identity_key,
            label="Gap B-A",
            payload={
                "entityType": "OUTCOME",
                "key": "gauss_gap",
                "originalLabel": "Gap B-A",
                "metricType": "measurement",
                "unit": "G",
                "favorableDirection": "UNKNOWN",
            },
            evidence_range="A1",
        )
        canonical_arm = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=identity_key,
            label="Gap B-A",
            payload={
                "entityType": "ARM",
                "key": "gap_before_minus_after",
                "role": "DERIVED",
                "label": "Gap B-A",
                "condition": "Before minus after",
                "sampleSize": None,
                "sampleBasis": "Paired samples",
                "matchingBasis": "Same sample number",
                "factorValues": [],
            },
            evidence_range="A1",
        )
        aliased_outcome = copy.deepcopy(canonical_outcome)
        aliased_outcome["payload"]["key"] = "gap_value"
        aliased_outcome["payload"]["originalLabel"] = "Gap value"
        aliased_arm = copy.deepcopy(canonical_arm)
        aliased_arm["payload"]["key"] = "gap_b_minus_a"
        aliased_arm["payload"]["condition"] = (
            "Worksheet-computed difference"
        )
        observation = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=value_key,
            label="Gap B-A",
            payload={
                "outcome": "gap_value",
                "arm": "gap_b_minus_a",
                "valueNumber": 42,
                "valueText": "42",
            },
            evidence_range="A2",
        )
        old_observation_id = observation["recordId"]

        fragments = [
            (
                plan["parts"][0],
                {
                    "records": [canonical_outcome, canonical_arm],
                    "coverageDispositions": [
                        {
                            "sourceCellKey": key,
                            "disposition": "RECORD_EVIDENCE",
                            "recordIds": [
                                canonical_outcome["recordId"],
                                canonical_arm["recordId"],
                            ],
                            "reason": "Canonical declarations.",
                        }
                        for key in plan["parts"][0][
                            "ownedSourceCellKeys"
                        ]
                    ],
                },
            ),
            (
                plan["parts"][1],
                {
                    "records": [
                        aliased_outcome,
                        aliased_arm,
                        observation,
                    ],
                    "coverageDispositions": [
                        {
                            "sourceCellKey": key,
                            "disposition": "RECORD_EVIDENCE",
                            "recordIds": [observation["recordId"]],
                            "reason": "Aliased observation.",
                        }
                        for key in plan["parts"][1][
                            "ownedSourceCellKeys"
                        ]
                    ],
                },
            ),
        ]
        forward = staged.merge_fragment_records(
            plan=plan,
            fragments=fragments,
            selected_chunks=universe["selectedChunks"],
        )
        reverse = staged.merge_fragment_records(
            plan=plan,
            fragments=list(reversed(fragments)),
            selected_chunks=universe["selectedChunks"],
        )
        self.assertEqual(forward, reverse)
        entities = [
            record
            for record in forward["records"]
            if record["recordType"] == "ENTITY_DECLARATION"
        ]
        self.assertEqual(
            {"gauss_gap", "gap_before_minus_after"},
            {record["payload"]["key"] for record in entities},
        )
        merged_observation = next(
            record
            for record in forward["records"]
            if record["recordType"] == "OBSERVATION_APPEND"
        )
        self.assertEqual("gauss_gap", merged_observation["payload"]["outcome"])
        self.assertEqual(
            "gap_before_minus_after",
            merged_observation["payload"]["arm"],
        )
        self.assertNotEqual(old_observation_id, merged_observation["recordId"])

        conflicted = copy.deepcopy(fragments)
        conflicted[1][1]["records"][0]["payload"]["unit"] = "kgf"
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "Conflicting nonempty.*unit",
        ):
            staged.merge_fragment_records(
                plan=plan,
                fragments=conflicted,
                selected_chunks=universe["selectedChunks"],
            )

    def test_merge_reconciles_compatible_entities_on_same_source_cell(
        self,
    ) -> None:
        title = "X526TOP SAMPLE NG OLD JIG"
        first_chunk = self.chunk("c1", [self.cell("A1", title)])
        second_chunk = self.chunk(
            "c2",
            [self.cell("A2", 68)],
            context=[self.cell("A1", title)],
        )
        universe, registry, plan = self.planned(
            [first_chunk, second_chunk],
            [
                self.locator(first_chunk),
                self.locator(second_chunk, candidate=False),
            ],
            max_chunks=1,
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        identity_key = first_chunk["cells"][0]["sourceCellKey"]
        value_key = second_chunk["cells"][0]["sourceCellKey"]
        canonical_arm = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=identity_key,
            label="OLD JIG",
            payload={
                "entityType": "ARM",
                "key": "old_jig",
                "role": "CONDITION",
                "label": "OLD JIG",
                "condition": "Old jig",
                "factorValues": [],
            },
            evidence_range="A1",
        )
        expanded_arm = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="ENTITY_DECLARATION",
            identity_key=identity_key,
            label=title,
            payload={
                "entityType": "ARM",
                "key": "sample_ng_old_jig",
                "role": "DESCRIPTIVE",
                "label": "Sample NG old jig",
                "condition": title,
                "factorValues": [],
            },
            evidence_range="A1",
        )
        observation = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="OBSERVATION_APPEND",
            identity_key=value_key,
            label="Gauss",
            payload={
                "outcome": "gauss",
                "arm": "sample_ng_old_jig",
                "valueNumber": 68,
            },
            evidence_range="A2",
        )
        fragments = [
            (
                plan["parts"][0],
                {
                    "records": [canonical_arm],
                    "coverageDispositions": [
                        {
                            "sourceCellKey": identity_key,
                            "disposition": "RECORD_EVIDENCE",
                            "recordIds": [canonical_arm["recordId"]],
                            "reason": "",
                        }
                    ],
                },
            ),
            (
                plan["parts"][1],
                {
                    "records": [expanded_arm, observation],
                    "coverageDispositions": [
                        {
                            "sourceCellKey": value_key,
                            "disposition": "RECORD_EVIDENCE",
                            "recordIds": [observation["recordId"]],
                            "reason": "",
                        }
                    ],
                },
            ),
        ]

        merged = staged.merge_fragment_records(
            plan=plan,
            fragments=fragments,
            selected_chunks=universe["selectedChunks"],
        )

        arms = [
            record
            for record in merged["records"]
            if record["recordType"] == "ENTITY_DECLARATION"
            and record["payload"]["entityType"] == "ARM"
        ]
        self.assertEqual(1, len(arms))
        self.assertEqual("old_jig", arms[0]["payload"]["key"])
        merged_observation = next(
            record
            for record in merged["records"]
            if record["recordType"] == "OBSERVATION_APPEND"
        )
        self.assertEqual(
            "old_jig",
            merged_observation["payload"]["arm"],
        )

    def projection_records(
        self,
        logical_id: str,
        cell_key: str,
    ) -> list[dict]:
        values = [
            (
                "ENTITY_DECLARATION",
                "Control",
                {"entityType": "ARM", "key": "control", "role": "CONTROL"},
            ),
            (
                "ENTITY_DECLARATION",
                "Test",
                {"entityType": "ARM", "key": "test", "role": "TEST"},
            ),
            (
                "ENTITY_DECLARATION",
                "Result",
                {
                    "entityType": "OUTCOME",
                    "key": "result",
                    "originalLabel": "Result",
                    "outcomeType": "CATEGORICAL",
                    "unit": "",
                    "betterDirection": "UNKNOWN",
                },
            ),
            (
                "OBSERVATION_APPEND",
                "Control pass",
                {
                    "outcome": "result",
                    "arm": "control",
                    "valueText": "Pass",
                },
            ),
            (
                "OBSERVATION_APPEND",
                "Test pass",
                {
                    "outcome": "result",
                    "arm": "test",
                    "valueText": "Pass",
                },
            ),
        ]
        return [
            self.record(
                envelope={},
                logical_id=logical_id,
                record_type=record_type,
                identity_key=cell_key,
                label=label,
                payload=payload,
                evidence_range="A1",
            )
            for record_type, label, payload in values
        ]

    def test_projection_clears_only_incomplete_observation_rate_pairs(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", 10)])
        universe, registry, _plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        records = self.projection_records(
            logical_id,
            chunk["cells"][0]["sourceCellKey"],
        )
        observations = [
            record
            for record in records
            if record["recordType"] == "OBSERVATION_APPEND"
        ]
        observations[0]["payload"].update(
            {
                "valueNumber": 0.1,
                "valueText": "10%",
                "numerator": 1,
                "denominator": None,
                "ratePpm": 100_000,
            }
        )
        observations[1]["payload"].update(
            {
                "valueNumber": 0.2,
                "valueText": "20%",
                "numerator": 2,
                "denominator": 10,
                "ratePpm": 200_000,
            }
        )

        manifest = staged.project_canonical_manifest(
            merged=self.merged_value(
                records,
                [chunk],
                records_sha256="incomplete-rate-pair",
            ),
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=universe["selectedChunks"],
        )

        projected = manifest["studies"][0]["outcomes"][0]["observations"]
        by_arm = {observation["arm"]: observation for observation in projected}
        self.assertIsNone(by_arm["control"]["numerator"])
        self.assertIsNone(by_arm["control"]["denominator"])
        self.assertEqual(100_000, by_arm["control"]["ratePpm"])
        self.assertEqual(2, by_arm["test"]["numerator"])
        self.assertEqual(10, by_arm["test"]["denominator"])

    def test_projection_maps_literal_normal_arm_to_reference(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "Normal")])
        universe, registry, _plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        records = self.projection_records(
            logical_id,
            chunk["cells"][0]["sourceCellKey"],
        )
        normal = next(
            record
            for record in records
            if record["recordType"] == "ENTITY_DECLARATION"
            and record["payload"].get("entityType") == "ARM"
            and record["payload"].get("key") == "control"
        )
        normal["payload"].update(
            {
                "role": "CONTROL",
                "label": "Normal",
                "condition": "Normal",
            }
        )

        manifest = staged.project_canonical_manifest(
            merged=self.merged_value(
                records,
                [chunk],
                records_sha256="literal-normal-reference",
            ),
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=universe["selectedChunks"],
        )

        projected = next(
            arm
            for arm in manifest["studies"][0]["arms"]
            if arm["key"] == "control"
        )
        self.assertEqual("REFERENCE", projected["role"])

    def test_projection_recovers_direct_formula_pair_and_percent_scale(
        self,
    ) -> None:
        formula_cell = self.cell(
            "C1",
            None,
            formula="=A1/B1",
            cached_value=0.1,
        )
        formula_cell["numberFormat"] = "0.0%"
        formula_cell["displayValue"] = "10.0%"
        chunk = self.chunk(
            "c1",
            [
                self.cell("A1", 1),
                self.cell("B1", 10),
                formula_cell,
            ],
        )
        universe, registry, _plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        records = self.projection_records(
            logical_id,
            chunk["cells"][0]["sourceCellKey"],
        )
        outcome = next(
            record
            for record in records
            if record["recordType"] == "ENTITY_DECLARATION"
            and record["payload"]["entityType"] == "OUTCOME"
        )
        outcome["payload"]["unit"] = "percent"
        observation = next(
            record
            for record in records
            if record["recordType"] == "OBSERVATION_APPEND"
            and record["payload"]["arm"] == "control"
        )
        observation["payload"].update(
            {
                "valueNumber": 0.1,
                "valueText": None,
                "numerator": None,
                "denominator": None,
                "ratePpm": 100_000,
            }
        )
        observation["evidence"] = [
            {
                "sheet": "Data",
                "range": "C1",
                "role": "RATE",
            }
        ]

        manifest = staged.project_canonical_manifest(
            merged=self.merged_value(
                records,
                [chunk],
                records_sha256="direct-formula-rate-pair",
            ),
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=universe["selectedChunks"],
        )

        projected = next(
            value
            for value in manifest["studies"][0]["outcomes"][0][
                "observations"
            ]
            if value["arm"] == "control"
        )
        self.assertEqual(1, projected["numerator"])
        self.assertEqual(10, projected["denominator"])
        self.assertEqual(10, projected["sampleSize"])
        self.assertEqual(10.0, projected["valueNumber"])
        self.assertEqual("10.0%", projected["valueText"])
        self.assertEqual(
            {"A1", "B1", "C1"},
            {item["range"] for item in projected["evidence"]},
        )

    def test_projection_clears_redundant_ppm_from_labeled_percent_text(
        self,
    ) -> None:
        cell = self.cell(
            "A1",
            "NG R&B and hearing 10.5% (Voltage 130% NG: 2.0%)",
        )
        chunk = self.chunk("c1", [cell])
        result = staged._normalize_projected_percent_observation(
            payload={
                "valueNumber": 10.5,
                "valueText": "10.5%",
                "numerator": None,
                "denominator": None,
                "ratePpm": 105_000,
            },
            evidence=[
                {
                    "sheet": "Data",
                    "range": "A1",
                    "role": "OBSERVATION",
                }
            ],
            outcome_unit="%",
            outcome_label="NG R&B and hearing percentage",
            selected_chunks=[chunk],
            by_key={
                cell["sourceCellKey"]: ("Data", "A1", cell),
            },
        )

        self.assertEqual(10.5, result["valueNumber"])
        self.assertEqual("10.5%", result["valueText"])
        self.assertIsNone(result["ratePpm"])

    def test_projection_augments_exact_categorical_status_rows(
        self,
    ) -> None:
        studies = [
            {
                "key": "study",
                "arms": [
                    {
                        "key": "result_row_38",
                        "label": "#1",
                        "evidence": [
                            {
                                "sheet": "Data",
                                "range": "D38",
                                "sourceText": "#1",
                            }
                        ],
                    }
                ],
                "outcomes": [],
            }
        ]
        chunk = self.chunk(
            "status",
            [
                self.cell("Q34", "THD"),
                self.cell("Q37", "After"),
                self.cell("D38", "#1"),
                self.cell("Q38", "OK"),
            ],
        )

        augmented = (
            staged._augment_projected_categorical_status_observations(
                studies=studies,
                selected_chunks=[chunk],
                revision_uid="revision-1",
            )
        )

        outcome = augmented[0]["outcomes"][0]
        self.assertEqual("source_status_c17", outcome["key"])
        self.assertEqual("THD | After", outcome["originalLabel"])
        self.assertEqual(
            {"Q34", "Q37"},
            {item["range"] for item in outcome["evidence"]},
        )
        self.assertEqual("result_row_38", outcome["observations"][0]["arm"])
        self.assertEqual("OK", outcome["observations"][0]["valueText"])
        self.assertEqual(
            "Q38",
            outcome["observations"][0]["evidence"][0]["range"],
        )

    def test_projection_adds_status_only_study_for_unmatched_rows(
        self,
    ) -> None:
        chunk = self.chunk(
            "status-only",
            [
                self.cell("B2", "RELIABILITY TEST RESULT"),
                self.cell("D14", "#2"),
                self.cell("S14", "PASS"),
            ],
        )

        augmented = (
            staged._augment_projected_categorical_status_observations(
                studies=[],
                selected_chunks=[chunk],
                revision_uid="revision-1",
            )
        )

        self.assertEqual(1, len(augmented))
        study = augmented[0]
        self.assertEqual("RELIABILITY TEST RESULT", study["title"])
        self.assertEqual("#2", study["arms"][0]["label"])
        self.assertEqual(
            "PASS",
            study["outcomes"][0]["observations"][0]["valueText"],
        )
        self.assertEqual(
            study["arms"][0]["key"],
            study["outcomes"][0]["observations"][0]["arm"],
        )

    def test_mixed_series_splits_numeric_runs_and_preserves_text(
        self,
    ) -> None:
        chunk = self.chunk(
            "mixed-series",
            [
                self.cell("B3", 1),
                self.cell("C3", 68.8),
                self.cell("B4", 2),
                self.cell("C4", "Press"),
                self.cell("B5", 3),
                self.cell("C5", 23.5),
                self.cell("B6", 4),
                self.cell("C6", 43.9),
            ],
        )
        by_coordinate, _by_key, _order = staged._source_cell_maps(
            [chunk]
        )
        record = {
            "recordType": "SERIES_SEGMENT_APPEND",
            "recordId": "series-record",
            "payload": {
                "key": "mixed",
                "outcome": "air_leak",
                "arm": "normal",
                "sheet": "Data",
                "headerRange": "C2",
                "valueRange": "C3:C6",
                "rowIdentityRange": "B3:B6",
            },
        }

        split, text_observations = (
            staged._split_mixed_numeric_series_record(
                record=record,
                by_coordinate=by_coordinate,
                revision_uid="revision-1",
            )
        )

        self.assertEqual(
            ["C3", "C5:C6"],
            [item["payload"]["valueRange"] for item in split],
        )
        self.assertEqual(
            ["B3", "B5:B6"],
            [item["payload"]["rowIdentityRange"] for item in split],
        )
        self.assertEqual(1, len(text_observations))
        self.assertEqual("Press", text_observations[0]["valueText"])
        self.assertEqual("2", text_observations[0]["replicateKey"])
        self.assertEqual(
            {"B4", "C4"},
            {
                item["range"]
                for item in text_observations[0]["evidence"]
            },
        )

    def test_comparison_intent_requires_local_entities_and_both_arms(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "Pass")])
        universe, registry, _plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        logical_id = registry["studies"][0]["logicalStudyId"]
        records = self.projection_records(
            logical_id,
            chunk["cells"][0]["sourceCellKey"],
        )
        base_merged = self.merged_value(
            records,
            [chunk],
            records_sha256="records",
        )
        without_intent = staged.project_canonical_manifest(
            merged=base_merged,
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=[chunk],
        )
        self.assertEqual(
            [],
            without_intent["studies"][0]["comparisons"],
        )

        intent = self.record(
            envelope={},
            logical_id=logical_id,
            record_type="COMPARISON_LINK_INTENT",
            identity_key=chunk["cells"][0]["sourceCellKey"],
            label="Control versus Test",
            payload={
                "comparedArm": "test",
                "controlArm": "control",
                "outcomes": ["result"],
                "designType": "CONTROL_VS_TEST",
                "matchingBasis": "Explicit source table.",
            },
            evidence_range="A1",
        )
        with_intent = staged.project_canonical_manifest(
            merged=self.merged_value(
                [*records, intent],
                [chunk],
                records_sha256="records-intent",
            ),
            registry=registry,
            source=self.source,
            workbook=self.workbook,
            selected_chunks=[chunk],
        )
        comparison = with_intent["studies"][0]["comparisons"][0]
        self.assertEqual("NEEDS_REVIEW", comparison["validityStatus"])
        self.assertFalse(comparison["aggregationEligible"])
        self.assertEqual([], comparison["effects"])

        dangling = [
            record
            for record in records
            if not (
                record["recordType"] == "OBSERVATION_APPEND"
                and record["payload"]["arm"] == "control"
            )
        ]
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "dangling",
        ):
            staged.project_canonical_manifest(
                merged=self.merged_value(
                    [*dangling, intent],
                    [chunk],
                    records_sha256="dangling",
                ),
                registry=registry,
                source=self.source,
                workbook=self.workbook,
                selected_chunks=[chunk],
            )

    def test_unknown_and_cross_study_entity_references_are_rejected(
        self,
    ) -> None:
        first = self.chunk("c1", [self.cell("A1", "Arm")])
        second = self.chunk("c2", [self.cell("A2", "Factor")])
        universe, registry, _plan = self.planned(
            [first, second],
            [self.locator(first), self.locator(second)],
        )
        self.assertEqual(2, len(registry["studies"]))
        first_id, second_id = [
            study["logicalStudyId"]
            for study in registry["studies"]
        ]
        arm = self.record(
            envelope={},
            logical_id=first_id,
            record_type="ENTITY_DECLARATION",
            identity_key=first["cells"][0]["sourceCellKey"],
            label="Arm",
            payload={
                "entityType": "ARM",
                "key": "arm",
                "factorValues": [
                    {
                        "factor": "foreign-factor",
                        "value": "x",
                        "valueNumber": None,
                        "unit": "",
                        "isBaseline": False,
                        "heldConstant": False,
                    }
                ],
            },
            evidence_range="A1",
        )
        foreign_factor = self.record(
            envelope={},
            logical_id=second_id,
            record_type="ENTITY_DECLARATION",
            identity_key=second["cells"][0]["sourceCellKey"],
            label="Factor",
            payload={
                "entityType": "FACTOR",
                "key": "foreign-factor",
                "originalLabel": "Factor",
            },
            evidence_range="A2",
        )
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "cross-Study reference",
        ):
            staged.project_canonical_manifest(
                merged=self.merged_value(
                    [arm, foreign_factor],
                    universe["selectedChunks"],
                    records_sha256="cross",
                ),
                registry=registry,
                source=self.source,
                workbook=self.workbook,
                selected_chunks=universe["selectedChunks"],
            )

        unknown = copy.deepcopy(arm)
        unknown["payload"]["factorValues"][0][
            "factor"
        ] = "unknown-factor"
        with self.assertRaisesRegex(
            staged.StagedDraftV2Error,
            "references unknown FACTOR",
        ):
            staged.project_canonical_manifest(
                merged=self.merged_value(
                    [unknown],
                    universe["selectedChunks"],
                    records_sha256="unknown",
                ),
                registry=registry,
                source=self.source,
                workbook=self.workbook,
                selected_chunks=universe["selectedChunks"],
            )

    def test_series_merge_requires_exact_adjacency(self) -> None:
        def series(
            record_id: str,
            value_range: str,
            identity_range: str,
        ) -> dict:
            return {
                "recordId": record_id,
                "logicalStudyId": "study",
                "payload": {
                    "outcome": "result",
                    "arm": "test",
                    "axisSource": "ROW",
                    "sheet": "Data",
                    "valueUnit": "dB",
                    "headerRange": "B1",
                    "valueRange": value_range,
                    "rowIdentityRange": identity_range,
                },
                "evidence": [],
            }

        adjacent = staged.merge_adjacent_series_segments(
            [
                series("one", "B2:B3", "A2:A3"),
                series("two", "B4:B5", "A4:A5"),
            ]
        )
        self.assertEqual(1, len(adjacent))
        self.assertEqual(
            "B2:B5",
            adjacent[0]["payload"]["valueRange"],
        )
        gap = staged.merge_adjacent_series_segments(
            [
                series("one", "B2:B3", "A2:A3"),
                series("two", "B5:B6", "A5:A6"),
            ]
        )
        self.assertEqual(2, len(gap))

    def test_part_and_final_provenance_invalidate_every_contract_hash(
        self,
    ) -> None:
        chunk = self.chunk("c1", [self.cell("A1", "value")])
        universe, registry, plan = self.planned(
            [chunk],
            [self.locator(chunk)],
        )
        part, envelope = self.envelope(universe, registry, plan)
        provenance = staged.part_provenance_v2(
            plan=plan,
            part=part,
            envelope=envelope,
            output_path=Path("part.json"),
            output_sha256="part-sha",
            generated_at="now",
        )
        self.assertTrue(
            staged.part_provenance_v2_matches(
                provenance=provenance,
                plan=plan,
                part=part,
                envelope=envelope,
                output_sha256="part-sha",
                output_path=Path("part.json"),
            )
        )
        for field in envelope["inputHashes"]:
            changed = copy.deepcopy(provenance)
            changed["inputHashes"][field] = "changed"
            self.assertFalse(
                staged.part_provenance_v2_matches(
                    provenance=changed,
                    plan=plan,
                    part=part,
                    envelope=envelope,
                    output_sha256="part-sha",
                    output_path=Path("part.json"),
                ),
                field,
            )
        for field in (
            "fragmentIdentity",
            "fragmentIdentitySha256",
            "outputPath",
            "imagesAnalyzed",
        ):
            changed = copy.deepcopy(provenance)
            changed[field] = "changed"
            self.assertFalse(
                staged.part_provenance_v2_matches(
                    provenance=changed,
                    plan=plan,
                    part=part,
                    envelope=envelope,
                    output_sha256="part-sha",
                    output_path=Path("part.json"),
                ),
                field,
            )

        ordered_hashes = [
            {"partId": part["partId"], "outputSha256": "part-sha"}
        ]
        final = staged.final_provenance_v2(
            plan=plan,
            registry=registry,
            ordered_part_hashes=ordered_hashes,
            merged_path=Path("merged.json"),
            merged_sha256="merged-sha",
            final_path=Path("final.json"),
            final_sha256="final-sha",
            generated_at="now",
        )
        self.assertTrue(
            staged.final_provenance_v2_matches(
                provenance=final,
                plan=plan,
                registry=registry,
                final_sha256="final-sha",
                ordered_part_hashes=ordered_hashes,
                merged_path=Path("merged.json"),
                merged_sha256="merged-sha",
                final_path=Path("final.json"),
            )
        )
        for field in (
            "planId",
            "fragmentIdentity",
            "fragmentIdentitySha256",
            "source",
            "registrySha256",
            "fragmentContractVersion",
            "validatorContractVersion",
            "consolidatorContractVersion",
            "finalManifestSha256",
            "orderedPartOutputHashes",
            "mergedArtifactPath",
            "mergedArtifactSha256",
            "finalManifestPath",
            "imagesAnalyzed",
        ):
            changed = copy.deepcopy(final)
            changed[field] = "changed"
            self.assertFalse(
                staged.final_provenance_v2_matches(
                    provenance=changed,
                    plan=plan,
                    registry=registry,
                    final_sha256="final-sha",
                    ordered_part_hashes=ordered_hashes,
                    merged_path=Path("merged.json"),
                    merged_sha256="merged-sha",
                    final_path=Path("final.json"),
                ),
                field,
            )


if __name__ == "__main__":
    unittest.main()
