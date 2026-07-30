from __future__ import annotations

import copy
import unittest

import inference_data_ai_b04_b08_repair as repair


class B04B08RepairTests(unittest.TestCase):
    def evidence(
        self,
        address: str,
        source_text: str,
        *,
        note: str = "",
    ) -> dict:
        return {
            "sheet": "Sheet1",
            "range": address,
            "role": "SOURCE",
            "sourceText": source_text,
            "note": note,
        }

    def cell(
        self,
        revision_uid: str,
        coordinate: str,
        raw_value: object,
        data_type: str,
        *,
        merge_range: str | None = None,
    ) -> dict:
        return {
            "coordinate": coordinate,
            "dataType": data_type,
            "rawValue": raw_value,
            "displayValue": raw_value,
            "mergeRange": merge_range,
            "mergeRole": "anchor" if merge_range else "none",
            "primary": True,
            "contextOnly": False,
            "sourceCellKey": f"{revision_uid}:1:{coordinate}",
        }

    def chunk(
        self,
        revision_uid: str,
        content_sha256: str,
        cells: list[dict],
    ) -> dict:
        return {
            "sheet": {"title": "Sheet1", "sheetIndex": 1},
            "sourceRevision": {
                "revisionUid": revision_uid,
                "contentSha256": content_sha256,
            },
            "cells": cells,
        }

    def b04_baseline(self) -> dict:
        sample_text = "Check height of sample  ( 50pcs ) "
        conclusion_text = (
            "6 Posistion of jig not same when setting sensor zero "
        )
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "dataset": "InputDataFinish",
                "sourcePath": "D:\\source\\" + repair.B04_FILE_NAME,
                "revisionUid": repair.B04_REVISION_UID,
                "contentSha256": repair.B04_CONTENT_SHA256,
                "contentComplete": True,
            },
            "workbookAnalysis": {"key": "unchanged"},
            "studies": [
                {
                    "key": "sensor-zero-set-by-jig-position",
                    "conclusions": [
                        {
                            "key": "six-jig-positions-not-same",
                            "text": conclusion_text,
                            "claimType": "SOURCE_CONCLUSION",
                            "causalStrength": "DESCRIPTIVE",
                            "evidence": [
                                self.evidence(
                                    "I6:K6",
                                    conclusion_text,
                                )
                            ],
                        }
                    ],
                },
                {
                    "key": "sample-height-ng-summary",
                    "summary": "must remain unchanged",
                },
                {
                    "key": "sample-height-50pcs-measurements",
                    "arms": [
                        {
                            "key": "height-measurement-50pcs",
                            "sampleSize": 50,
                            "evidence": [
                                self.evidence(
                                    "D23:M23",
                                    sample_text,
                                )
                            ],
                        }
                    ],
                    "outcomes": [
                        {
                            "key": "height-measurement-sample-size",
                            "originalLabel": sample_text,
                            "metricType": "sample_size",
                            "unit": "pcs",
                            "evidence": [
                                self.evidence(
                                    "D23:M23",
                                    sample_text,
                                )
                            ],
                            "observations": [
                                {
                                    "key": (
                                        "height-measurement-"
                                        "sample-size-50"
                                    ),
                                    "arm": "height-measurement-50pcs",
                                    "valueNumber": 50,
                                    "valueText": "50pcs",
                                    "sampleSize": 50,
                                    "evidence": [
                                        self.evidence(
                                            "D23:M23",
                                            sample_text,
                                        )
                                    ],
                                }
                            ],
                        }
                    ],
                    "measurementSeries": [
                        {
                            "key": "sample-height-raw-matrix",
                            "seriesRole": "RAW",
                            "outcome": "sample-height",
                            "arm": "height-measurement-50pcs",
                            "sheet": "Sheet1",
                            "headerRange": "D24:M24",
                            "valueRange": "D25:M29",
                            "rowIdentityRange": "C25:C29",
                            "axisSource": "HEADER",
                        }
                    ],
                },
            ],
        }

    def b04_chunks(self) -> list[dict]:
        uid = repair.B04_REVISION_UID
        cells = [
            self.cell(
                uid,
                "B3",
                "CHECK SENSOR ZERO SET ON THE JIG ( 6 POSITION )",
                "s",
            ),
            self.cell(uid, "B4", "Date", "s", merge_range="B4:B5"),
            self.cell(
                uid,
                "C4",
                "Posistion on the jig ",
                "s",
                merge_range="C4:H4",
            ),
            self.cell(
                uid,
                "I4",
                "Note",
                "s",
                merge_range="I4:K5",
            ),
            self.cell(
                uid,
                "I6",
                "6 Posistion of jig not same when setting sensor zero ",
                "s",
                merge_range="I6:K6",
            ),
            self.cell(
                uid,
                "D23",
                "Check height of sample  ( 50pcs ) ",
                "s",
                merge_range="D23:M23",
            ),
        ]
        for value, column in enumerate("DEFGHIJKLM", start=1):
            cells.append(self.cell(uid, f"{column}24", value, "n"))
        for row_identity, row in enumerate(range(25, 30), start=1):
            cells.append(self.cell(uid, f"C{row}", row_identity, "n"))
            for offset, column in enumerate("DEFGHIJKLM"):
                cells.append(
                    self.cell(
                        uid,
                        f"{column}{row}",
                        1.9 + row_identity / 100 + offset / 1000,
                        "n",
                    )
                )
        return [
            self.chunk(
                uid,
                repair.B04_CONTENT_SHA256,
                cells,
            )
        ]

    def b08_baseline(self) -> dict:
        line_1 = (
            "- Lot test frame new mold check SPK OK -> Continue move "
            "modul  test "
        )
        line_2 = " => Can use"
        return {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                "dataset": "InputDataFinish",
                "sourcePath": "D:\\source\\" + repair.B08_FILE_NAME,
                "revisionUid": repair.B08_REVISION_UID,
                "contentSha256": repair.B08_CONTENT_SHA256,
                "contentComplete": True,
            },
            "workbookAnalysis": {"key": "unchanged"},
            "studies": [
                {
                    "key": "new-mold-frame-lot-performance",
                    "summary": "must remain unchanged",
                    "conclusions": [
                        {
                            "key": "new-mold-lot-spk-disposition",
                            "text": line_1,
                            "claimType": "SOURCE_CONCLUSION",
                            "causalStrength": "DESCRIPTIVE",
                            "evidence": [
                                self.evidence("B25", line_1)
                            ],
                        },
                        {
                            "key": "new-mold-can-use-disposition",
                            "text": line_2,
                            "claimType": "SOURCE_CONCLUSION",
                            "causalStrength": "DESCRIPTIVE",
                            "evidence": [
                                self.evidence("B26", line_2)
                            ],
                        },
                    ],
                }
            ],
        }

    def b08_chunks(self) -> list[dict]:
        uid = repair.B08_REVISION_UID
        return [
            self.chunk(
                uid,
                repair.B08_CONTENT_SHA256,
                [
                    self.cell(
                        uid,
                        "B25",
                        (
                            "- Lot test frame new mold check SPK OK -> "
                            "Continue move modul  test "
                        ),
                        "s",
                        merge_range="B25:T25",
                    ),
                    self.cell(uid, "B26", " => Can use", "s"),
                ],
            )
        ]

    def target(
        self,
        baseline: dict,
        chunks: list[dict],
        error: str,
    ) -> dict:
        target = repair.b04_b08_repair_target(
            error,
            baseline,
            chunks,
        )
        self.assertIsNotNone(target)
        assert target is not None
        return target

    def test_b04_exact_projection_preserves_queryable_context(self) -> None:
        baseline = self.b04_baseline()
        original = copy.deepcopy(baseline)
        target = self.target(
            baseline,
            self.b04_chunks(),
            repair.B04_NUMERIC_VALIDATION_ERROR,
        )

        repaired = repair.apply_b04_b08_repair(baseline, target)
        repair.validate_b04_b08_repair(
            baseline,
            repaired,
            target,
        )

        self.assertEqual(original, baseline)
        observation = repaired["studies"][2]["outcomes"][0][
            "observations"
        ][0]
        self.assertIsNone(observation["valueNumber"])
        self.assertEqual("50pcs", observation["valueText"])
        self.assertEqual(50, observation["sampleSize"])
        conclusion = repaired["studies"][0]["conclusions"][0]
        contexts = conclusion["evidence"][1:]
        self.assertEqual(
            [
                ("B3", "CHECK SENSOR ZERO SET ON THE JIG ( 6 POSITION )"),
                ("B4", "Date"),
                ("C4", "Posistion on the jig "),
                ("I4", "Note"),
            ],
            [
                (item["range"], item["sourceText"])
                for item in contexts
            ],
        )
        self.assertTrue(
            all(
                "not a separate conclusion" in item["note"]
                for item in contexts
            )
        )

        stripped = copy.deepcopy(repaired)
        stripped["studies"][2]["outcomes"][0]["observations"][0][
            "valueNumber"
        ] = 50
        del stripped["studies"][0]["conclusions"][0]["evidence"][1:]
        self.assertEqual(baseline, stripped)

    def test_b08_changes_only_adjacent_evidence_range_and_text(self) -> None:
        baseline = self.b08_baseline()
        original = copy.deepcopy(baseline)
        target = self.target(
            baseline,
            self.b08_chunks(),
            repair.B08_CONCLUSION_VALIDATION_ERROR,
        )

        repaired = repair.apply_b04_b08_repair(baseline, target)
        repair.validate_b04_b08_repair(
            baseline,
            repaired,
            target,
        )

        self.assertEqual(original, baseline)
        evidence = repaired["studies"][0]["conclusions"][1][
            "evidence"
        ][0]
        self.assertEqual("B25:B26", evidence["range"])
        self.assertEqual(
            (
                "- Lot test frame new mold check SPK OK -> Continue "
                "move modul  test; => Can use"
            ),
            evidence["sourceText"],
        )
        stripped = copy.deepcopy(repaired)
        stripped_evidence = stripped["studies"][0]["conclusions"][1][
            "evidence"
        ][0]
        original_evidence = baseline["studies"][0]["conclusions"][1][
            "evidence"
        ][0]
        stripped_evidence["range"] = original_evidence["range"]
        stripped_evidence["sourceText"] = original_evidence["sourceText"]
        self.assertEqual(baseline, stripped)

    def test_both_repairs_are_idempotent(self) -> None:
        cases = (
            (
                self.b04_baseline(),
                self.b04_chunks(),
                repair.B04_NUMERIC_VALIDATION_ERROR,
            ),
            (
                self.b08_baseline(),
                self.b08_chunks(),
                repair.B08_CONCLUSION_VALIDATION_ERROR,
            ),
        )
        for baseline, chunks, error in cases:
            with self.subTest(source=baseline["source"]["revisionUid"]):
                target = self.target(baseline, chunks, error)
                once = repair.apply_b04_b08_repair(
                    baseline,
                    target,
                )
                twice = repair.apply_b04_b08_repair(once, target)
                self.assertEqual(once, twice)

                repeated_target = self.target(once, chunks, error)
                self.assertEqual(
                    once,
                    repair.apply_b04_b08_repair(
                        once,
                        repeated_target,
                    ),
                )

    def test_source_identity_and_error_fail_closed(self) -> None:
        cases = (
            (
                self.b04_baseline(),
                self.b04_chunks(),
                repair.B04_NUMERIC_VALIDATION_ERROR,
            ),
            (
                self.b08_baseline(),
                self.b08_chunks(),
                repair.B08_CONCLUSION_VALIDATION_ERROR,
            ),
        )
        for baseline, chunks, error in cases:
            for name, mutation in (
                (
                    "revision",
                    lambda value: value["source"].__setitem__(
                        "revisionUid",
                        "capture_revision_wrong",
                    ),
                ),
                (
                    "sha",
                    lambda value: value["source"].__setitem__(
                        "contentSha256",
                        "0" * 64,
                    ),
                ),
                (
                    "file",
                    lambda value: value["source"].__setitem__(
                        "sourcePath",
                        "D:\\source\\wrong.xlsx",
                    ),
                ),
            ):
                with self.subTest(error=error, mutation=name):
                    changed = copy.deepcopy(baseline)
                    mutation(changed)
                    self.assertIsNone(
                        repair.b04_b08_repair_target(
                            error,
                            changed,
                            chunks,
                        )
                    )
            self.assertIsNone(
                repair.b04_b08_repair_target(
                    error + " changed",
                    baseline,
                    chunks,
                )
            )

    def test_b04_source_geometry_fail_closed(self) -> None:
        mutations: list[tuple[str, list[dict]]] = []

        wrong_prose = self.b04_chunks()
        self.find_cell(wrong_prose, "D23")["rawValue"] = "50"
        mutations.append(("prose", wrong_prose))

        wrong_merge = self.b04_chunks()
        self.find_cell(wrong_merge, "B4")["mergeRange"] = None
        mutations.append(("context merge", wrong_merge))

        missing_point = self.b04_chunks()
        missing_point[0]["cells"] = [
            cell
            for cell in missing_point[0]["cells"]
            if cell["coordinate"] != "M29"
        ]
        mutations.append(("matrix shape", missing_point))

        text_point = self.b04_chunks()
        self.find_cell(text_point, "D25")["dataType"] = "s"
        mutations.append(("matrix type", text_point))

        wrong_key = self.b04_chunks()
        self.find_cell(wrong_key, "I4")["sourceCellKey"] = "wrong"
        mutations.append(("source key", wrong_key))

        for name, chunks in mutations:
            with self.subTest(name=name):
                self.assertIsNone(
                    repair.b04_b08_repair_target(
                        repair.B04_NUMERIC_VALIDATION_ERROR,
                        self.b04_baseline(),
                        chunks,
                    )
                )

    def test_b08_source_geometry_fail_closed(self) -> None:
        mutations: list[tuple[str, list[dict]]] = []

        wrong_context = self.b08_chunks()
        self.find_cell(wrong_context, "B25")["rawValue"] += "changed"
        mutations.append(("context text", wrong_context))

        wrong_merge = self.b08_chunks()
        self.find_cell(wrong_merge, "B25")["mergeRange"] = "B25:S25"
        mutations.append(("context merge", wrong_merge))

        missing_claim = self.b08_chunks()
        missing_claim[0]["cells"] = [
            cell
            for cell in missing_claim[0]["cells"]
            if cell["coordinate"] != "B26"
        ]
        mutations.append(("claim missing", missing_claim))

        duplicate = self.b08_chunks()
        duplicate[0]["cells"].append(
            copy.deepcopy(self.find_cell(duplicate, "B26"))
        )
        mutations.append(("duplicate", duplicate))

        for name, chunks in mutations:
            with self.subTest(name=name):
                self.assertIsNone(
                    repair.b04_b08_repair_target(
                        repair.B08_CONCLUSION_VALIDATION_ERROR,
                        self.b08_baseline(),
                        chunks,
                    )
                )

    def test_manifest_geometry_and_partial_states_fail_closed(self) -> None:
        b04_cases = []
        wrong_value = self.b04_baseline()
        wrong_value["studies"][2]["outcomes"][0]["observations"][0][
            "valueNumber"
        ] = 49
        b04_cases.append(wrong_value)

        wrong_series = self.b04_baseline()
        wrong_series["studies"][2]["measurementSeries"][0][
            "valueRange"
        ] = "D25:M28"
        b04_cases.append(wrong_series)

        partial_b04 = self.b04_baseline()
        partial_b04["studies"][2]["outcomes"][0]["observations"][0][
            "valueNumber"
        ] = None
        b04_cases.append(partial_b04)

        for baseline in b04_cases:
            self.assertIsNone(
                repair.b04_b08_repair_target(
                    repair.B04_NUMERIC_VALIDATION_ERROR,
                    baseline,
                    self.b04_chunks(),
                )
            )

        b08_cases = []
        wrong_claim = self.b08_baseline()
        wrong_claim["studies"][0]["conclusions"][1]["text"] = "Can use"
        b08_cases.append(wrong_claim)

        partial_b08 = self.b08_baseline()
        partial_b08["studies"][0]["conclusions"][1]["evidence"][0][
            "range"
        ] = "B25:B26"
        b08_cases.append(partial_b08)

        for baseline in b08_cases:
            self.assertIsNone(
                repair.b04_b08_repair_target(
                    repair.B08_CONCLUSION_VALIDATION_ERROR,
                    baseline,
                    self.b08_chunks(),
                )
            )

    def test_projection_and_tampered_target_fail_closed(self) -> None:
        baseline = self.b08_baseline()
        target = self.target(
            baseline,
            self.b08_chunks(),
            repair.B08_CONCLUSION_VALIDATION_ERROR,
        )
        changed = copy.deepcopy(baseline)
        changed["workbookAnalysis"]["key"] = "changed"
        with self.assertRaisesRegex(
            repair.B04B08RepairError,
            "protected repair projection",
        ):
            repair.apply_b04_b08_repair(changed, target)

        repaired = repair.apply_b04_b08_repair(baseline, target)
        altered = copy.deepcopy(repaired)
        altered["studies"][0]["summary"] = "unrelated mutation"
        with self.assertRaisesRegex(
            repair.B04B08RepairError,
            "outside the exact",
        ):
            repair.validate_b04_b08_repair(
                baseline,
                altered,
                target,
            )

        tampered = copy.deepcopy(target)
        tampered["repairKind"] = "OTHER"
        with self.assertRaisesRegex(
            repair.B04B08RepairError,
            "target is not exact",
        ):
            repair.apply_b04_b08_repair(baseline, tampered)

    def find_cell(
        self,
        chunks: list[dict],
        coordinate: str,
    ) -> dict:
        return next(
            cell
            for chunk in chunks
            for cell in chunk["cells"]
            if cell["coordinate"] == coordinate
        )


if __name__ == "__main__":
    unittest.main()
