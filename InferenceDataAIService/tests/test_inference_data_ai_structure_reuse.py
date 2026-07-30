from __future__ import annotations

import copy
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from inference_data_ai_extraction_recipe import (
    EXTRACTION_RECIPE_SCHEMA_VERSION,
    create_form_template,
    validate_extraction_recipe,
)
from inference_data_ai_recipe_executor import (
    RecipeExecutionError,
    execute_recipe,
)
from inference_data_ai_incremental_match_audit import rank_catalog_structures
from inference_data_ai_incremental_table_match import match_table_request
from inference_data_ai_recipe_matcher import decide_template_match
from inference_data_ai_recipe_validation import validate_extraction
from inference_data_ai_source_ingest import import_capture
from inference_data_ai_structure_fingerprint import (
    build_structure_fingerprint,
    build_structure_fingerprint_from_database,
)
from inference_data_ai_template_bootstrap import (
    build_template_bootstrap_catalog,
)
from inference_data_ai_table_structure_catalog import (
    table_structure_fingerprint,
)
from inference_data_ai_table_recipe_proposal import (
    PROPOSAL_DECISION_SCHEMA_VERSION,
    adapt_decision_to_priority_item,
    build_priority_extension_report,
    build_table_recipe_decision_prompt,
    build_table_recipe_priority_report,
    compile_structure_recipe,
    execute_structure_recipe,
    redact_representative_table,
    replay_structure_recipe,
    run_codex_table_recipe_decision,
    semantic_header_signature,
    table_structure_similarity,
    validate_table_recipe_decision,
    _semantic_header_sha256,
    _safety_status,
)
from inference_data_ai_structure_batch_control import (
    build_batch_control,
    build_recipe_registry,
    match_registered_recipe,
)
from inference_data_ai_historical_semantic_bootstrap import (
    build_historical_semantic_contract_catalog,
)
from inference_data_ai_incremental_coverage import (
    build_incremental_coverage_report,
)
from inference_data_ai_structure_completion import (
    COMPLETION_STATE_SCHEMA_VERSION,
    _record_registered_outcomes,
    _summarize,
    complete_structure_queue,
)
from inference_data_ai_single_measure_bootstrap import (
    SOURCE_OWNED_DECISION_SOURCE,
    build_source_owned_single_measure_decision,
)


def _cell(
    row: int,
    column: int,
    coordinate: str,
    value: object,
    *,
    number_format: str = "General",
    merge_range: str | None = None,
    merge_role: str = "none",
) -> dict:
    return {
        "row": row,
        "column": column,
        "coordinate": coordinate,
        "rawValue": value,
        "displayValue": value,
        "cachedValue": None,
        "formula": None,
        "dataType": "n" if isinstance(value, (int, float)) else "s",
        "numberFormat": number_format,
        "styleId": 0,
        "style": {},
        "mergeRange": merge_range,
        "mergeRole": merge_role,
    }


def capture_payload(
    *,
    first_value: float = 12.34,
    second_value: float = 8.1,
    digest_character: str = "a",
    judgement_header: str = "판정",
) -> dict:
    cells = [
        _cell(
            1,
            1,
            "A1",
            "시험결과",
            merge_range="A1:D1",
            merge_role="anchor",
        ),
        _cell(3, 1, "A3", "항목"),
        _cell(3, 2, "B3", "단위"),
        _cell(3, 3, "C3", "측정값"),
        _cell(3, 4, "D3", judgement_header),
        _cell(4, 1, "A4", "압축강도"),
        _cell(4, 2, "B4", "MPa"),
        _cell(4, 3, "C4", first_value, number_format="0.00"),
        _cell(4, 4, "D4", "적합"),
        _cell(5, 1, "A5", "인장강도"),
        _cell(5, 2, "B5", "MPa"),
        _cell(5, 3, "C5", second_value, number_format="0.00"),
        _cell(5, 4, "D5", "적합"),
    ]
    return {
        "schemaVersion": "input-data-openxml-capture-v2",
        "captureContract": "openpyxl-sparse-source-capture-v2.0",
        "source": {
            "sourcePath": "fixture.xlsx",
            "contentSha256": digest_character * 64,
        },
        "workbook": {
            "status": "CAPTURED",
            "sheetCount": 1,
            "visibleSheetCount": 1,
            "tabularSheetCount": 1,
            "sheets": [
                {
                    "sheetIndex": 1,
                    "title": "시험결과",
                    "sheetState": "visible",
                    "status": "CAPTURED",
                    "hasTabularEvidence": True,
                    "usedBounds": {
                        "minRow": 1,
                        "minColumn": 1,
                        "maxRow": 5,
                        "maxColumn": 4,
                        "rowCount": 5,
                        "columnCount": 4,
                        "address": "A1:D5",
                    },
                    "contentBounds": {
                        "minRow": 1,
                        "minColumn": 1,
                        "maxRow": 5,
                        "maxColumn": 4,
                        "rowCount": 5,
                        "columnCount": 4,
                        "address": "A1:D5",
                    },
                    "formulaCellCount": 0,
                    "mergeCount": 1,
                    "mergedRanges": [
                        {
                            "minRow": 1,
                            "minColumn": 1,
                            "maxRow": 1,
                            "maxColumn": 4,
                            "rowCount": 1,
                            "columnCount": 4,
                            "address": "A1:D1",
                            "anchor": "A1",
                        }
                    ],
                    "cells": cells,
                }
            ],
        },
    }


def extraction_recipe() -> dict:
    return {
        "schemaVersion": EXTRACTION_RECIPE_SCHEMA_VERSION,
        "recipeId": "recipe-test-result",
        "recipeVersion": 1,
        "templateId": "tpl-test-result",
        "sheetSelectors": [
            {
                "id": "resultSheet",
                "titleAliases": ["시험결과", "결과"],
                "requiredAnchors": ["판정"],
                "fallbackRole": "tabular-result",
                "cardinality": "exactly-one",
            }
        ],
        "anchors": [
            {
                "id": "resultHeader",
                "sheet": "resultSheet",
                "textRegex": "^시험결과$",
                "normalized": True,
                "uniqueness": "one",
            }
        ],
        "regions": [
            {
                "id": "resultTable",
                "sheet": "resultSheet",
                "start": {"below": "resultHeader", "rows": 1},
                "stop": {"firstBlankKeyColumnRun": 1},
                "headerDepth": 2,
                "repeatMode": "rows",
            }
        ],
        "axes": {
            "rowKey": {
                "region": "resultTable",
                "columnRole": "item-label",
            },
            "columns": [
                {"role": "item-label", "headerAliases": ["항목"]},
                {"role": "unit", "headerAliases": ["단위"]},
                {"role": "value", "headerAliases": ["측정값", "결과값"]},
                {"role": "judgement", "headerAliases": ["판정"]},
            ],
        },
        "fields": [
            {
                "parameter": "measurement",
                "source": {
                    "region": "resultTable",
                    "row": "each",
                    "valueColumnRole": "value",
                    "labelColumnRole": "item-label",
                    "unitColumnRole": "unit",
                },
                "valueType": "decimal",
                "evidence": "exact-source-cell",
                "required": True,
            }
        ],
        "validationRules": [
            {"rule": "required-field-coverage", "minimum": 1.0},
            {"rule": "evidence-cell-exists"},
        ],
    }


def database_capture_payload(
    source_path: Path,
    *,
    first_value: float,
    digest_character: str,
) -> dict:
    result = capture_payload(
        first_value=first_value,
        digest_character=digest_character,
    )
    result["extractor"] = {"name": "test", "version": "1"}
    result["source"].update(
        {
            "sourcePath": str(source_path),
            "fileName": source_path.name,
            "extension": ".xlsx",
            "sizeBytes": 1,
            "mtimeNs": 1,
        }
    )
    result["workbook"].update(
        {
            "isTrulyEmpty": False,
            "nonEmptySheetCount": 1,
            "metadata": {},
        }
    )
    sheet = result["workbook"]["sheets"][0]
    sheet.update(
        {
            "isTrulyEmpty": False,
            "nonEmptyCellCount": len(sheet["cells"]),
            "structuralCellCount": 0,
            "capturedCellCount": len(sheet["cells"]),
            "freezePanes": None,
            "autoFilter": None,
            "sheetMetadata": {},
            "rowDimensions": [],
            "columnDimensions": [],
        }
    )
    return result


class StructureFingerprintTests(unittest.TestCase):
    def test_numeric_values_and_source_sha_do_not_change_structure_digest(self) -> None:
        first = build_structure_fingerprint(capture_payload())
        second = build_structure_fingerprint(
            capture_payload(
                first_value=999.25,
                second_value=-3.5,
                digest_character="b",
            )
        )

        self.assertNotEqual(first["sourceSha256"], second["sourceSha256"])
        self.assertEqual(first["fingerprintSha256"], second["fingerprintSha256"])

    def test_header_change_changes_structure_digest(self) -> None:
        first = build_structure_fingerprint(capture_payload())
        second = build_structure_fingerprint(
            capture_payload(judgement_header="상태")
        )

        self.assertNotEqual(first["fingerprintSha256"], second["fingerprintSha256"])

    def test_database_fingerprint_matches_payload_fingerprint(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            payload = database_capture_payload(
                root / "fixture.xlsx",
                first_value=12.34,
                digest_character="d",
            )
            connection = sqlite3.connect(root / "capture.sqlite")
            imported = import_capture(
                connection,
                payload,
                captured_at="2026-07-27T00:00:00Z",
            )
            connection.commit()

            expected = build_structure_fingerprint(payload)
            actual = build_structure_fingerprint_from_database(
                connection,
                imported["revisionId"],
            )
            connection.close()

        self.assertEqual(expected, actual)


class RecipeReuseVerticalSliceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.capture = capture_payload()
        self.recipe = extraction_recipe()
        validate_extraction_recipe(self.recipe)
        fingerprint = build_structure_fingerprint(self.capture)
        self.template = create_form_template(
            template_id="tpl-test-result",
            family_id="family-test-result",
            fingerprint=fingerprint,
            recipe_ref="recipe-test-result@1",
            required_anchors=["시험결과", "판정"],
            status="APPROVED",
        )

    def test_exact_structure_uses_no_ai_and_extracts_exact_cell_evidence(self) -> None:
        candidate = capture_payload(
            first_value=44.5,
            second_value=19.25,
            digest_character="c",
        )
        fingerprint = build_structure_fingerprint(candidate)

        decision = decide_template_match(fingerprint, [self.template])
        execution = execute_recipe(
            candidate,
            self.recipe,
            match_decision=decision,
        )
        report = validate_extraction(candidate, self.recipe, execution)

        self.assertEqual("EXACT_REUSE", decision["decision"])
        self.assertFalse(decision["ai"]["used"])
        self.assertEqual("VERIFIED", report["status"])
        parameters = execution["result"]["parameters"]
        self.assertEqual([44.5, 19.25], [item["value"] for item in parameters])
        self.assertEqual(["C4", "C5"], [
            next(
                evidence["cell"]
                for evidence in item["evidence"]
                if evidence["role"] == "value"
            )
            for item in parameters
        ])
        self.assertEqual(["MPa", "MPa"], [item["unit"] for item in parameters])

    def test_missing_required_anchor_fails_closed_before_recipe_execution(self) -> None:
        candidate = capture_payload(judgement_header="상태")
        fingerprint = build_structure_fingerprint(candidate)

        decision = decide_template_match(fingerprint, [self.template])

        self.assertEqual("NEW_TEMPLATE_REQUIRED", decision["decision"])
        self.assertEqual(
            ["REQUIRED_ANCHOR_MISSING"],
            decision["topCandidates"][0]["hardGateFailures"],
        )
        with self.assertRaisesRegex(
            RecipeExecutionError,
            "does not authorize recipe execution",
        ):
            execute_recipe(candidate, self.recipe, match_decision=decision)

    def test_tampered_evidence_is_rejected(self) -> None:
        decision = decide_template_match(
            build_structure_fingerprint(self.capture),
            [self.template],
        )
        execution = execute_recipe(
            self.capture,
            self.recipe,
            match_decision=decision,
        )
        tampered = copy.deepcopy(execution)
        tampered["result"]["parameters"][0]["evidence"][1]["rawValue"] = 0

        report = validate_extraction(self.capture, self.recipe, tampered)

        self.assertEqual("FAILED", report["status"])
        self.assertIn("EVIDENCE_VALUE_MISMATCH", report["failureCodes"])


class TemplateBootstrapTests(unittest.TestCase):
    def test_completed_requests_build_one_reusable_exact_structure_without_ai(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = root / "capture.sqlite"
            batch = root / "batch"
            for name in ("requests", "analyses", "projections"):
                (batch / name).mkdir(parents=True)
            connection = sqlite3.connect(database)
            imported: list[tuple[dict, dict]] = []
            for index, (value, character) in enumerate(
                ((12.34, "e"), (88.0, "f")),
                start=1,
            ):
                payload = database_capture_payload(
                    root / f"fixture-{index}.xlsx",
                    first_value=value,
                    digest_character=character,
                )
                result = import_capture(
                    connection,
                    payload,
                    captured_at=f"2026-07-27T00:00:0{index}Z",
                )
                revision_uid = connection.execute(
                    """
                    SELECT revision_uid
                    FROM capture_v2_revisions
                    WHERE revision_id=?
                    """,
                    (result["revisionId"],),
                ).fetchone()[0]
                imported.append((payload, {**result, "revisionUid": revision_uid}))
            connection.commit()
            connection.close()

            for index, (payload, imported_result) in enumerate(imported, start=1):
                name = f"capture_revision_test_{index}.json"
                request = {
                    "schemaVersion": "table-first-analysis-request-v1",
                    "source": {
                        "contentSha256": payload["source"]["contentSha256"],
                        "fileName": payload["source"]["fileName"],
                        "sourcePath": payload["source"]["sourcePath"],
                        "revisionUid": imported_result["revisionUid"],
                    },
                }
                (batch / "requests" / name).write_text(
                    json.dumps(request),
                    encoding="utf-8",
                )
                (batch / "analyses" / name).write_text("{}", encoding="utf-8")
                (batch / "projections" / name).write_text("{}", encoding="utf-8")

            catalog = build_template_bootstrap_catalog(
                database_path=database,
                table_first_batch_root=batch,
                generated_at="2026-07-27T00:00:10+00:00",
            )

        self.assertEqual(0, catalog["inputs"]["aiCalls"])
        self.assertEqual(2, catalog["summary"]["fingerprintedFileCount"])
        self.assertEqual(1, catalog["summary"]["exactStructureCount"])
        self.assertEqual(2, catalog["summary"]["exactReusableFileCount"])
        self.assertEqual(1, catalog["summary"]["candidateFamilyCount"])
        self.assertEqual(2, catalog["families"][0]["fileCount"])

    def test_incremental_audit_ranks_same_structure_with_changed_values_first(
        self,
    ) -> None:
        old_same = build_structure_fingerprint(capture_payload())
        old_other = build_structure_fingerprint(
            capture_payload(judgement_header="상태")
        )
        candidate = build_structure_fingerprint(
            capture_payload(
                first_value=771.5,
                second_value=-22.0,
                digest_character="9",
            )
        )
        structures = [
            {
                "structureId": "other",
                "fingerprintSha256": old_other["fingerprintSha256"],
                "fileCount": 3,
                "fingerprint": old_other,
                "members": [{"fileName": "other.xlsx", "revisionId": 2}],
            },
            {
                "structureId": "same",
                "fingerprintSha256": old_same["fingerprintSha256"],
                "fileCount": 1,
                "fingerprint": old_same,
                "members": [{"fileName": "same.xlsx", "revisionId": 1}],
            },
        ]

        ranked = rank_catalog_structures(candidate, structures, top_k=2)

        self.assertEqual("same", ranked[0]["structureId"])
        self.assertTrue(ranked[0]["fingerprintExact"])
        self.assertEqual(1.0, ranked[0]["score"])

    def test_table_structure_fingerprint_excludes_numeric_samples(self) -> None:
        table = {
            "bounds": {
                "minRow": 3,
                "minColumn": 1,
                "maxRow": 5,
                "maxColumn": 3,
            },
            "sourceCellCount": 9,
            "numericCellCount": 2,
            "numericColumnCount": 1,
            "previewRows": [
                {
                    "row": 3,
                    "omittedCellCount": 0,
                    "cells": [
                        {"coordinate": "A3", "kind": "TEXT", "value": "항목"},
                        {"coordinate": "B3", "kind": "TEXT", "value": "단위"},
                        {"coordinate": "C3", "kind": "TEXT", "value": "측정값"},
                    ],
                },
                {
                    "row": 4,
                    "omittedCellCount": 0,
                    "cells": [
                        {"coordinate": "A4", "kind": "TEXT", "value": "압축강도"},
                        {"coordinate": "B4", "kind": "TEXT", "value": "MPa"},
                        {"coordinate": "C4", "kind": "NUMBER", "value": 12.34},
                    ],
                },
                {
                    "row": 5,
                    "omittedCellCount": 0,
                    "cells": [
                        {"coordinate": "A5", "kind": "TEXT", "value": "인장강도"},
                        {"coordinate": "B5", "kind": "TEXT", "value": "MPa"},
                        {"coordinate": "C5", "kind": "NUMBER", "value": 8.1},
                    ],
                },
            ],
            "rowLabels": [
                {"row": 3, "labels": ["항목", "단위", "측정값"]},
                {"row": 4, "labels": ["압축강도", "MPa"]},
                {"row": 5, "labels": ["인장강도", "MPa"]},
            ],
            "numericColumns": [
                {
                    "column": "C",
                    "columnRole": "MEASURE_VALUE",
                    "numberFormats": ["0.00"],
                    "numericCount": 2,
                    "headerTexts": ["측정값"],
                }
            ],
            "numericSeries": [],
            "aggregateChecks": [],
        }
        changed = copy.deepcopy(table)
        changed["previewRows"][1]["cells"][2]["value"] = 999
        changed["previewRows"][2]["cells"][2]["value"] = -5

        first = table_structure_fingerprint(table)
        second = table_structure_fingerprint(changed)

        self.assertEqual(first["fingerprintSha256"], second["fingerprintSha256"])

    def test_incremental_table_request_exactly_matches_historical_block(self) -> None:
        table = {
            "tableId": "new-table",
            "sheetIndex": 1,
            "sheet": "Result",
            "range": "A3:C5",
            "bounds": {
                "minRow": 3,
                "minColumn": 1,
                "maxRow": 5,
                "maxColumn": 3,
            },
            "sourceCellCount": 9,
            "numericCellCount": 2,
            "numericColumnCount": 1,
            "previewRows": [
                {
                    "row": 3,
                    "omittedCellCount": 0,
                    "cells": [
                        {"coordinate": "A3", "kind": "TEXT", "value": "항목"},
                        {"coordinate": "B3", "kind": "TEXT", "value": "단위"},
                        {"coordinate": "C3", "kind": "TEXT", "value": "값"},
                    ],
                },
                {
                    "row": 4,
                    "omittedCellCount": 0,
                    "cells": [
                        {"coordinate": "A4", "kind": "TEXT", "value": "A"},
                        {"coordinate": "B4", "kind": "TEXT", "value": "MPa"},
                        {"coordinate": "C4", "kind": "NUMBER", "value": 5},
                    ],
                },
            ],
            "rowLabels": [
                {"row": 3, "labels": ["항목", "단위", "값"]},
                {"row": 4, "labels": ["A", "MPa"]},
            ],
            "numericColumns": [
                {
                    "column": "C",
                    "columnRole": "MEASURE_VALUE",
                    "numberFormats": ["0.00"],
                    "numericCount": 2,
                    "headerTexts": ["값"],
                }
            ],
            "numericSeries": [],
            "aggregateChecks": [],
        }
        fingerprint = table_structure_fingerprint(table)
        structure = {
            "tableStructureId": "historical-structure",
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "tableCount": 12,
            "workbookCount": 9,
            "dominantSemanticType": "DESCRIPTIVE",
            "semanticConsistency": 1.0,
        }

        result = match_table_request(
            {"tables": [table]},
            {fingerprint["fingerprintSha256"]: structure},
            verified_recipes={
                "historical-structure": {
                    "recipeId": "verified-recipe",
                    "recipeVersion": 1,
                }
            },
        )

        self.assertEqual(1, result["exactMatchedTableCount"])
        self.assertEqual(1, result["exactMatchedQuantitativeTableCount"])
        self.assertEqual(1, result["verifiedRecipeMatchCount"])
        self.assertEqual(
            "VERIFIED_RECIPE_MATCH",
            result["tables"][0]["status"],
        )


class TableRecipeProposalTests(unittest.TestCase):
    @staticmethod
    def _table(
        table_id: str,
        *,
        first_value: float,
        second_value: float,
    ) -> dict:
        return {
            "tableId": table_id,
            "sheetIndex": 1,
            "sheet": "Result",
            "range": "A1:C3",
            "bounds": {
                "minRow": 1,
                "minColumn": 1,
                "maxRow": 3,
                "maxColumn": 3,
            },
            "sourceCellCount": 9,
            "numericCellCount": 4,
            "numericColumnCount": 2,
            "titleCandidates": ["Strength result"],
            "previewRows": [
                {
                    "row": 1,
                    "omittedCellCount": 0,
                    "cells": [
                        {
                            "coordinate": "A1",
                            "kind": "TEXT",
                            "value": "Item",
                        },
                        {
                            "coordinate": "B1",
                            "kind": "TEXT",
                            "value": "Compression",
                        },
                        {
                            "coordinate": "C1",
                            "kind": "TEXT",
                            "value": "Tension",
                        },
                    ],
                },
                {
                    "row": 2,
                    "omittedCellCount": 0,
                    "cells": [
                        {
                            "coordinate": "A2",
                            "kind": "TEXT",
                            "value": "Sample A",
                        },
                        {
                            "coordinate": "B2",
                            "kind": "NUMBER",
                            "value": first_value,
                        },
                        {
                            "coordinate": "C2",
                            "kind": "NUMBER",
                            "value": second_value,
                        },
                    ],
                },
                {
                    "row": 3,
                    "omittedCellCount": 0,
                    "cells": [
                        {
                            "coordinate": "A3",
                            "kind": "TEXT",
                            "value": "Sample B",
                        },
                        {
                            "coordinate": "B3",
                            "kind": "NUMBER",
                            "value": first_value + 1,
                            "numberFormat": "0.00",
                        },
                        {
                            "coordinate": "C3",
                            "kind": "NUMBER",
                            "value": second_value + 1,
                            "numberFormat": "0.00",
                        },
                    ],
                },
            ],
            "rowLabels": [
                {"row": 1, "labels": ["Item", "Compression", "Tension"]},
                {"row": 2, "labels": ["Sample A"]},
                {"row": 3, "labels": ["Sample B"]},
            ],
            "numericColumns": [
                {
                    "column": "B",
                    "columnId": f"{table_id}_col_B",
                    "columnRole": "MEASURE_VALUE",
                    "numberFormats": ["0.00"],
                    "numericCount": 2,
                    "headerTexts": ["Compression"],
                    "sourceRange": "B2:B3",
                    "min": first_value,
                    "max": first_value + 1,
                    "average": first_value + 0.5,
                },
                {
                    "column": "C",
                    "columnId": f"{table_id}_col_C",
                    "columnRole": "MEASURE_VALUE",
                    "numberFormats": ["0.00"],
                    "numericCount": 2,
                    "headerTexts": ["Tension"],
                    "sourceRange": "C2:C3",
                    "min": second_value,
                    "max": second_value + 1,
                    "average": second_value + 0.5,
                },
            ],
            "numericSeries": [],
            "aggregateChecks": [],
        }

    @staticmethod
    def _decision(
        *,
        structure_id: str,
        fingerprint_sha256: str,
    ) -> dict:
        return {
            "schemaVersion": PROPOSAL_DECISION_SCHEMA_VERSION,
            "targetTableStructureId": structure_id,
            "targetFingerprintSha256": fingerprint_sha256,
            "decision": "NEW_RECIPE",
            "historicalSourceTableStructureId": "",
            "confidence": "HIGH",
            "rationale": "Headers identify two measured strength outcomes.",
            "semanticContract": {
                "title": "Strength result",
                "tableType": "DESCRIPTIVE",
                "studyGroup": "strength-result",
                "groups": [],
                "metricColumns": [
                    {
                        "relativeColumn": 1,
                        "canonicalName": "Compression",
                        "unit": "MPa",
                    },
                    {
                        "relativeColumn": 2,
                        "canonicalName": "Tension",
                        "unit": "MPa",
                    },
                ],
                "comparisonRelations": [],
                "limitations": [],
            },
        }

    def test_redacted_representative_and_similarity_exclude_values(self) -> None:
        first_table = self._table(
            "table-one",
            first_value=12.34,
            second_value=8.1,
        )
        second_table = self._table(
            "table-two",
            first_value=999.0,
            second_value=-7.5,
        )
        first = table_structure_fingerprint(first_table)
        second = table_structure_fingerprint(second_table)

        similarity = table_structure_similarity(first, second)
        redacted = redact_representative_table(first_table)
        serialized = json.dumps(redacted, ensure_ascii=False)

        self.assertEqual(1.0, similarity["score"])
        self.assertTrue(similarity["fingerprintExact"])
        self.assertNotIn("12.34", serialized)
        self.assertNotIn('"average"', serialized)
        self.assertFalse(redacted["valuePolicy"]["rawValuesIncluded"])

    def test_single_measure_structure_requires_explicit_expansion(self) -> None:
        structure = {
            "fingerprint": {
                "numericColumnCount": 1,
                "numericColumns": [
                    {
                        "relativeColumn": 1,
                        "columnRole": "MEASURE_VALUE",
                    }
                ],
                "headerTokens": ["RESULT"],
            }
        }

        default_status, default_reasons = _safety_status(structure)
        expanded_status, expanded_reasons = _safety_status(
            structure,
            minimum_measure_columns=1,
        )

        self.assertEqual("NOT_PARAMETER_TABLE", default_status)
        self.assertEqual(["FEWER_THAN_2_MEASURE_COLUMNS"], default_reasons)
        self.assertEqual("PROPOSAL_READY", expanded_status)
        self.assertEqual([], expanded_reasons)

    def test_extension_report_selects_only_new_structure_ids(self) -> None:
        baseline = {
            "queue": [
                {
                    "tableStructureId": "table-structure-existing",
                    "status": "PROPOSAL_READY",
                    "tableCount": 2,
                    "workbookCount": 2,
                }
            ]
        }
        expanded = {
            "inputs": {
                "newCatalog": "new-catalog.json",
                "newBatchRoot": "new-batch",
                "minimumMeasureColumns": 1,
            },
            "queue": [
                baseline["queue"][0],
                {
                    "tableStructureId": "table-structure-new",
                    "status": "PROPOSAL_READY",
                    "tableCount": 3,
                    "workbookCount": 3,
                },
            ],
        }

        extension = build_priority_extension_report(
            expanded_priority_report=expanded,
            baseline_priority_report=baseline,
            generated_at="2026-07-27T00:00:00+00:00",
        )

        self.assertEqual(
            ["table-structure-new"],
            [item["tableStructureId"] for item in extension["queue"]],
        )
        self.assertEqual(3, extension["summary"]["coveredTableCount"])
        self.assertEqual(1, extension["inputs"]["baselineQueueStructureCount"])

    def test_single_measure_source_header_is_mapped_or_quarantined(self) -> None:
        table = self._table(
            "table-single-measure",
            first_value=12.34,
            second_value=8.1,
        )
        fingerprint = table_structure_fingerprint(table)
        base_item = {
            "tableStructureId": "table-structure-single-measure",
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "semanticHeaderSha256": "e" * 64,
            "representativeTable": {
                "titleCandidates": [],
                "numericColumns": [
                    {
                        "relativeColumn": 1,
                        "columnRole": "MEASURE_VALUE",
                        "headerTexts": ["Q'ty (pcs)"],
                    }
                ],
            },
            "historicalTopK": [],
        }

        mapped = build_source_owned_single_measure_decision(base_item)
        generic_item = copy.deepcopy(base_item)
        generic_item["representativeTable"]["numericColumns"][0][
            "headerTexts"
        ] = ["ME"]
        quarantined = build_source_owned_single_measure_decision(generic_item)

        self.assertEqual("NEW_RECIPE", mapped["decision"])
        self.assertEqual(
            {
                "relativeColumn": 1,
                "canonicalName": "Q'ty (pcs)",
                "unit": "pcs",
            },
            mapped["semanticContract"]["metricColumns"][0],
        )
        self.assertEqual("QUARANTINE", quarantined["decision"])
        self.assertEqual(
            [],
            quarantined["semanticContract"]["metricColumns"],
        )
        self.assertEqual(
            "SOURCE_OWNED_SINGLE_MEASURE_HEADER",
            SOURCE_OWNED_DECISION_SOURCE,
        )

    def test_priority_report_and_recipe_replay_one_structure_without_ai(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            batch = root / "new-batch"
            (batch / "requests").mkdir(parents=True)
            first_table = self._table(
                "table-one",
                first_value=12.34,
                second_value=8.1,
            )
            second_table = self._table(
                "table-two",
                first_value=77.0,
                second_value=55.0,
            )
            fingerprint = table_structure_fingerprint(first_table)
            structure_id = (
                "table-structure-" + fingerprint["fingerprintSha256"][:20]
            )
            members = []
            for index, table in enumerate(
                (first_table, second_table),
                start=1,
            ):
                request_file = f"request-{index}.json"
                (batch / "requests" / request_file).write_text(
                    json.dumps({"tables": [table]}),
                    encoding="utf-8",
                )
                members.append(
                    {
                        "fileName": f"sample-{index}.xlsx",
                        "requestFile": request_file,
                        "tableId": table["tableId"],
                        "sheet": table["sheet"],
                        "range": table["range"],
                        "contentSha256": str(index) * 64,
                    }
                )
            new_catalog = {
                "structures": [
                    {
                        "tableStructureId": structure_id,
                        "fingerprintSha256": fingerprint[
                            "fingerprintSha256"
                        ],
                        "fingerprint": fingerprint,
                        "tableCount": 2,
                        "workbookCount": 2,
                        "quantitative": True,
                        "members": members,
                    }
                ]
            }
            old_catalog = {
                "structures": [
                    {
                        "tableStructureId": "old-structure",
                        "fingerprintSha256": fingerprint[
                            "fingerprintSha256"
                        ],
                        "fingerprint": fingerprint,
                        "tableCount": 4,
                        "workbookCount": 4,
                        "dominantSemanticType": "DESCRIPTIVE",
                        "semanticConsistency": 1.0,
                        "members": [],
                    }
                ]
            }
            new_catalog_path = root / "new-catalog.json"
            old_catalog_path = root / "old-catalog.json"
            new_catalog_path.write_text(
                json.dumps(new_catalog),
                encoding="utf-8",
            )
            old_catalog_path.write_text(
                json.dumps(old_catalog),
                encoding="utf-8",
            )

            report = build_table_recipe_priority_report(
                new_catalog_path=new_catalog_path,
                historical_catalog_path=old_catalog_path,
                new_batch_root=batch,
                generated_at="2026-07-27T00:00:00+00:00",
            )
            item = report["queue"][0]
            decision = self._decision(
                structure_id=structure_id,
                fingerprint_sha256=fingerprint["fingerprintSha256"],
            )
            recipe = compile_structure_recipe(
                decision,
                priority_item=item,
                generated_at="2026-07-27T00:00:01+00:00",
            )
            replay = replay_structure_recipe(
                recipe=recipe,
                priority_report=report,
                generated_at="2026-07-27T00:00:02+00:00",
            )

        self.assertEqual(0, report["summary"]["aiCalls"])
        self.assertEqual("PROPOSAL_READY", item["status"])
        self.assertEqual(1.0, item["historicalTopK"][0]["score"])
        self.assertEqual(2, replay["summary"]["passed"])
        self.assertEqual(0, replay["summary"]["failed"])
        self.assertEqual(4, replay["summary"]["deterministicFactCount"])
        self.assertEqual(
            8,
            replay["summary"]["deterministicCellFactCount"],
        )
        self.assertEqual(
            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_NEEDS_CANONICAL_REVIEW",
            recipe["status"],
        )
        first_min = replay["items"][0]["extraction"][
            "deterministicNumericFacts"
        ][0]["min"]
        second_min = replay["items"][1]["extraction"][
            "deterministicNumericFacts"
        ][0]["min"]
        self.assertEqual((12.34, 77.0), (first_min, second_min))
        self.assertEqual(
            "B2",
            replay["items"][0]["extraction"][
                "deterministicCellFacts"
            ][0]["coordinate"],
        )

    def test_recipe_decision_rejects_unknown_relative_column(self) -> None:
        table = self._table(
            "table-one",
            first_value=12.34,
            second_value=8.1,
        )
        fingerprint = table_structure_fingerprint(table)
        structure_id = "table-structure-test"
        item = {
            "tableStructureId": structure_id,
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "representativeTable": redact_representative_table(table),
            "historicalTopK": [],
        }
        decision = self._decision(
            structure_id=structure_id,
            fingerprint_sha256=fingerprint["fingerprintSha256"],
        )
        decision["semanticContract"]["metricColumns"][0][
            "relativeColumn"
        ] = 99

        with self.assertRaisesRegex(
            RuntimeError,
            "Invalid metric column mapping",
        ):
            validate_table_recipe_decision(
                decision,
                priority_item=item,
            )

    def test_comparison_relation_references_are_canonicalized(self) -> None:
        table = self._table(
            "table-one",
            first_value=12.34,
            second_value=8.1,
        )
        fingerprint = table_structure_fingerprint(table)
        structure_id = "table-structure-relation-normalization"
        item = {
            "tableStructureId": structure_id,
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "representativeTable": redact_representative_table(table),
            "historicalTopK": [],
        }
        decision = self._decision(
            structure_id=structure_id,
            fingerprint_sha256=fingerprint["fingerprintSha256"],
        )
        decision["semanticContract"].update(
            {
                "tableType": "COMPARISON",
                "groups": [
                    {
                        "label": "Normal",
                        "role": "REFERENCE",
                        "basis": "Source-authored baseline.",
                    },
                    {
                        "label": "Test bond",
                        "role": "TEST",
                        "basis": "Source-authored test condition.",
                    },
                ],
                "comparisonRelations": [
                    {
                        "leftGroup": " normal ",
                        "rightGroup": "TEST BOND",
                        "basis": "Shared source metrics.",
                    }
                ],
            }
        )

        validated = validate_table_recipe_decision(
            decision,
            priority_item=item,
        )

        self.assertEqual(
            {
                "leftGroup": "Normal",
                "rightGroup": "Test bond",
                "basis": "Shared source metrics.",
            },
            validated["semanticContract"]["comparisonRelations"][0],
        )

    def test_recipe_prompt_requires_exact_group_relation_references(self) -> None:
        table = self._table(
            "table-one",
            first_value=12.34,
            second_value=8.1,
        )
        fingerprint = table_structure_fingerprint(table)
        structure_id = "table-structure-prompt-relations"
        report = {
            "queue": [
                {
                    "tableStructureId": structure_id,
                    "fingerprintSha256": fingerprint["fingerprintSha256"],
                    "tableCount": 1,
                    "workbookCount": 1,
                    "representativeTable": redact_representative_table(table),
                    "historicalTopK": [],
                }
            ]
        }

        prompt = build_table_recipe_decision_prompt(
            report,
            table_structure_id=structure_id,
        )

        self.assertIn(
            "must exactly copy two distinct label values declared in groups",
            prompt,
        )

    def test_rejected_ai_decision_is_preserved_in_telemetry(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            table = self._table(
                "table-one",
                first_value=12.34,
                second_value=8.1,
            )
            fingerprint = table_structure_fingerprint(table)
            structure_id = "table-structure-rejected-decision"
            item = {
                "tableStructureId": structure_id,
                "fingerprintSha256": fingerprint["fingerprintSha256"],
                "tableCount": 1,
                "workbookCount": 1,
                "status": "PROPOSAL_READY",
                "representativeTable": redact_representative_table(table),
                "historicalTopK": [],
            }
            decision = self._decision(
                structure_id=structure_id,
                fingerprint_sha256=fingerprint["fingerprintSha256"],
            )
            decision["semanticContract"].update(
                {
                    "tableType": "COMPARISON",
                    "groups": [
                        {
                            "label": "Normal",
                            "role": "REFERENCE",
                            "basis": "Source-authored baseline.",
                        },
                        {
                            "label": "Test",
                            "role": "TEST",
                            "basis": "Source-authored condition.",
                        },
                    ],
                    "comparisonRelations": [
                        {
                            "leftGroup": "Unknown",
                            "rightGroup": "Normal",
                            "basis": "Invalid cross-reference.",
                        }
                    ],
                }
            )
            output = root / "decision.json"
            telemetry = root / "telemetry.json"

            def fake_run(command: list[str], **_: object) -> mock.Mock:
                message_path = Path(
                    command[command.index("--output-last-message") + 1]
                )
                message_path.write_text(
                    json.dumps(decision),
                    encoding="utf-8",
                )
                return mock.Mock(returncode=0, stdout="", stderr="")

            with mock.patch(
                "inference_data_ai_table_recipe_proposal.subprocess.run",
                side_effect=fake_run,
            ):
                with self.assertRaisesRegex(
                    RuntimeError,
                    "Invalid comparison relation",
                ):
                    run_codex_table_recipe_decision(
                        priority_report={"queue": [item]},
                        table_structure_id=structure_id,
                        output_path=output,
                        telemetry_path=telemetry,
                        codex_command=["codex"],
                    )

            telemetry_value = json.loads(
                telemetry.read_text(encoding="utf-8")
            )
            self.assertEqual("FAILED", telemetry_value["status"])
            self.assertEqual(
                decision,
                telemetry_value["rejectedDecision"],
            )
            self.assertFalse(output.exists())

    def test_recipe_extracts_group_labels_from_each_source_table(self) -> None:
        representative = self._table(
            "table-one",
            first_value=12.34,
            second_value=8.1,
        )
        fingerprint = table_structure_fingerprint(representative)
        structure_id = "table-structure-dynamic-groups"
        item = {
            "tableStructureId": structure_id,
            "baseTableStructureId": structure_id,
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "semanticHeaderSha256": "f" * 64,
            "semanticHeaderSignature": [
                {
                    "relativeColumn": 1,
                    "columnRole": "MEASURE_VALUE",
                    "headerTexts": ["compression"],
                },
                {
                    "relativeColumn": 2,
                    "columnRole": "MEASURE_VALUE",
                    "headerTexts": ["tension"],
                },
            ],
            "representativeTable": redact_representative_table(
                representative
            ),
            "historicalTopK": [],
        }
        decision = self._decision(
            structure_id=structure_id,
            fingerprint_sha256=fingerprint["fingerprintSha256"],
        )
        decision["semanticContract"].update(
            {
                "tableType": "COMPARISON",
                "groups": [
                    {
                        "label": "Sample A",
                        "role": "TEST",
                        "basis": "First source condition.",
                    },
                    {
                        "label": "Sample B",
                        "role": "REFERENCE",
                        "basis": "Second source condition.",
                    },
                ],
                "comparisonRelations": [
                    {
                        "leftGroup": "Sample A",
                        "rightGroup": "Sample B",
                        "basis": "Two source conditions share metrics.",
                    }
                ],
            }
        )
        recipe = compile_structure_recipe(
            decision,
            priority_item=item,
        )
        candidate = self._table(
            "table-two",
            first_value=77.0,
            second_value=55.0,
        )
        candidate["previewRows"][1]["cells"][0]["value"] = "After"
        candidate["previewRows"][2]["cells"][0]["value"] = "Before"
        target_item = {
            **item,
            "tableStructureId": "table-structure-dynamic-groups-target",
            "baseTableStructureId": (
                "table-structure-dynamic-groups-target"
            ),
            "representativeTable": redact_representative_table(candidate),
        }
        adapted = adapt_decision_to_priority_item(
            decision,
            source_item=item,
            target_item=target_item,
        )

        extraction = execute_structure_recipe(recipe, candidate)

        self.assertEqual(
            ["After", "Before"],
            [
                group["label"]
                for group in extraction["semantic"]["groups"]
            ],
        )
        self.assertEqual(
            {
                "leftGroup": "After",
                "rightGroup": "Before",
                "basis": "Two source conditions share metrics.",
            },
            extraction["semantic"]["comparisonRelations"][0],
        )
        self.assertTrue(
            all(
                group["calculationAuthority"]
                == "CODE_FROM_CAPTURED_TABLE_PREVIEW"
                for group in extraction["semantic"]["groups"]
            )
        )
        self.assertEqual(
            ["After", "Before"],
            [
                group["label"]
                for group in adapted["semanticContract"]["groups"]
            ],
        )

    def test_recipe_resolves_repeated_and_composite_groups_from_capture_v2(
        self,
    ) -> None:
        table = self._table(
            "table-captured-groups",
            first_value=12.34,
            second_value=8.1,
        )
        table["bounds"]["maxRow"] = 4
        table["range"] = "A1:C4"
        table["previewRows"][2]["cells"][0]["value"] = "Sample A"
        fingerprint = table_structure_fingerprint(table)
        structure_id = "table-structure-captured-groups"
        item = {
            "tableStructureId": structure_id,
            "baseTableStructureId": structure_id,
            "fingerprintSha256": fingerprint["fingerprintSha256"],
            "semanticHeaderSha256": "f" * 64,
            "semanticHeaderSignature": semantic_header_signature(table),
            "representativeMember": {
                "range": table["range"],
            },
            "representativeTable": redact_representative_table(table),
            "historicalTopK": [],
        }
        decision = self._decision(
            structure_id=structure_id,
            fingerprint_sha256=fingerprint["fingerprintSha256"],
        )
        decision["semanticContract"].update(
            {
                "tableType": "COMPARISON",
                "groups": [
                    {
                        "label": "Sample A",
                        "role": "REFERENCE",
                        "basis": "Repeated source-authored condition.",
                    },
                    {
                        "label": "Normal voltage",
                        "role": "TEST",
                        "basis": "Source-authored composite header.",
                    },
                ],
                "comparisonRelations": [
                    {
                        "leftGroup": "Sample A",
                        "rightGroup": "Normal voltage",
                        "basis": "Two source-authored conditions.",
                    }
                ],
            }
        )
        captured_cells = [
            {
                "row": 2,
                "column": 1,
                "coordinate": "A2",
                "rawValue": "Sample A",
                "formula": None,
                "cachedValue": None,
                "displayValue": "Sample A",
                "numberFormat": "",
            },
            {
                "row": 3,
                "column": 1,
                "coordinate": "A3",
                "rawValue": "Sample A",
                "formula": None,
                "cachedValue": None,
                "displayValue": "Sample A",
                "numberFormat": "",
            },
            {
                "row": 4,
                "column": 1,
                "coordinate": "A4",
                "rawValue": "Total NG voltage normal",
                "formula": None,
                "cachedValue": None,
                "displayValue": "Total NG voltage normal",
                "numberFormat": "",
            },
        ]
        for row, first, second in (
            (2, 12.34, 8.1),
            (3, 13.34, 9.1),
        ):
            captured_cells.extend(
                [
                    {
                        "row": row,
                        "column": 2,
                        "coordinate": f"B{row}",
                        "rawValue": first,
                        "formula": None,
                        "cachedValue": None,
                        "displayValue": first,
                        "numberFormat": "0.00",
                    },
                    {
                        "row": row,
                        "column": 3,
                        "coordinate": f"C{row}",
                        "rawValue": second,
                        "formula": None,
                        "cachedValue": None,
                        "displayValue": second,
                        "numberFormat": "0.00",
                    },
                ]
            )

        recipe = compile_structure_recipe(
            decision,
            priority_item=item,
            representative_captured_cells=captured_cells,
        )
        extraction = execute_structure_recipe(
            recipe,
            table,
            captured_cells=captured_cells,
        )

        selectors = [
            group["sourceSelector"]
            for group in recipe["semanticContract"]["groups"]
        ]
        self.assertEqual(
            [(1, 0), (3, 0)],
            [
                (
                    selector["relativeRow"],
                    selector["relativeColumn"],
                )
                for selector in selectors
            ],
        )
        self.assertEqual(
            ["EXACT_SOURCE_TEXT", "SOURCE_TOKEN_SUBSET"],
            [selector["matchMode"] for selector in selectors],
        )
        self.assertEqual(
            ["Sample A", "Total NG voltage normal"],
            [
                group["label"]
                for group in extraction["semantic"]["groups"]
            ],
        )
        self.assertTrue(
            all(
                group["calculationAuthority"]
                == "CODE_FROM_CAPTURE_V2_DATABASE"
                for group in extraction["semantic"]["groups"]
            )
        )


class StructureBatchControlTests(unittest.TestCase):
    def test_registry_and_run_budget_prevent_second_structure_ai_call(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            recipes = root / "recipes"
            replays = root / "replays"
            telemetry = root / "telemetry"
            for path in (recipes, replays, telemetry):
                path.mkdir()
            structure_id = "table-structure-ready"
            recipe_id = "structure-recipe-ready"
            fingerprint = "a" * 64
            report = {
                "queue": [
                    {
                        "rank": 1,
                        "tableStructureId": structure_id,
                        "baseTableStructureId": structure_id,
                        "fingerprintSha256": fingerprint,
                        "semanticHeaderSha256": "d" * 64,
                        "semanticHeaderSignature": [
                            {
                                "relativeColumn": 1,
                                "columnRole": "MEASURE_VALUE",
                                "headerTexts": ["result"],
                            }
                        ],
                        "status": "PROPOSAL_READY",
                        "safetyReasons": [],
                        "tableCount": 11,
                        "workbookCount": 10,
                    },
                    {
                        "rank": 2,
                        "tableStructureId": "table-structure-next",
                        "fingerprintSha256": "b" * 64,
                        "status": "PROPOSAL_READY",
                        "safetyReasons": [],
                        "tableCount": 9,
                        "workbookCount": 9,
                    },
                    {
                        "rank": 3,
                        "tableStructureId": "table-structure-review",
                        "fingerprintSha256": "c" * 64,
                        "status": "REVIEW_BEFORE_PROPOSAL",
                        "safetyReasons": ["HIGH_DIMENSIONAL_NUMERIC_MATRIX"],
                        "tableCount": 3,
                        "workbookCount": 3,
                    },
                ]
            }
            (recipes / "recipe.json").write_text(
                json.dumps(
                    {
                        "recipeId": recipe_id,
                        "recipeVersion": 1,
                        "status": (
                            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_"
                            "NEEDS_CANONICAL_REVIEW"
                        ),
                        "decision": {
                            "source": "BOUNDED_AI_STRUCTURE_DECISION",
                            "aiCallCount": 1,
                        },
                        "match": {
                            "tableStructureId": structure_id,
                            "baseTableStructureId": structure_id,
                            "fingerprintSha256": fingerprint,
                            "semanticHeaderSha256": "d" * 64,
                            "semanticHeaderSignature": [
                                {
                                    "relativeColumn": 1,
                                    "columnRole": "MEASURE_VALUE",
                                    "headerTexts": ["result"],
                                }
                            ],
                        },
                    }
                ),
                encoding="utf-8",
            )
            (replays / "replay.json").write_text(
                json.dumps(
                    {
                        "recipeId": recipe_id,
                        "status": (
                            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_"
                            "NEEDS_CANONICAL_REVIEW"
                        ),
                        "summary": {
                            "memberCount": 11,
                            "passed": 11,
                            "failed": 0,
                            "deterministicFactCount": 77,
                            "deterministicCellFactCount": 264,
                        },
                    }
                ),
                encoding="utf-8",
            )
            (telemetry / "ai.json").write_text(
                json.dumps(
                    {
                        "tableStructureId": structure_id,
                        "status": "SUCCEEDED",
                        "aiCallsAttempted": 1,
                        "aiCallsSucceeded": 1,
                        "retryCount": 0,
                    }
                ),
                encoding="utf-8",
            )

            registry = build_recipe_registry(
                priority_report=report,
                recipe_root=recipes,
                replay_root=replays,
                telemetry_root=telemetry,
                generated_at="2026-07-27T00:00:00+00:00",
            )
            control = build_batch_control(
                priority_report=report,
                registry=registry,
                telemetry_root=telemetry,
                max_ai_calls=1,
                generated_at="2026-07-27T00:00:01+00:00",
            )

        self.assertEqual(1, registry["summary"]["registeredRecipeCount"])
        self.assertEqual(0, registry["summary"]["rejectedRecipeCount"])
        self.assertEqual(1, control["budget"]["consumedAiCalls"])
        self.assertEqual(0, control["budget"]["remainingAiCalls"])
        self.assertEqual(
            [
                "EXACT_RECIPE_REPLAY_READY",
                "AI_BUDGET_WAIT",
                "MANUAL_REVIEW_REQUIRED",
            ],
            [item["action"] for item in control["actions"]],
        )
        self.assertFalse(control["policy"]["fileLevelAiEnabled"])
        self.assertEqual(0, control["policy"]["retryCount"])

    def test_explicit_source_quarantine_is_not_ai_authorized(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            telemetry = root / "telemetry"
            decisions = root / "decisions"
            telemetry.mkdir()
            decisions.mkdir()
            (decisions / "source-owned.json").write_text(
                json.dumps(
                    {
                        "targetTableStructureId": "structure-generic",
                        "decision": "QUARANTINE",
                    }
                ),
                encoding="utf-8",
            )
            control = build_batch_control(
                priority_report={
                    "queue": [
                        {
                            "rank": 1,
                            "tableStructureId": "structure-generic",
                            "tableCount": 3,
                            "workbookCount": 3,
                            "status": "PROPOSAL_READY",
                        }
                    ]
                },
                registry={"recipes": []},
                telemetry_root=telemetry,
                decision_root=decisions,
                max_ai_calls=1,
            )

        self.assertEqual(
            "SOURCE_PRECHECK_QUARANTINED",
            control["actions"][0]["action"],
        )
        self.assertEqual(0, control["budget"]["newlyAuthorizedAiCalls"])

    def test_registered_recipe_match_is_exact_and_ai_free(self) -> None:
        table = TableRecipeProposalTests._table(
            "table-one",
            first_value=12.34,
            second_value=8.1,
        )
        fingerprint = table_structure_fingerprint(table)
        registry = {
            "recipes": [
                {
                    "recipeId": "recipe-one",
                    "fingerprintSha256": fingerprint[
                        "fingerprintSha256"
                    ],
                    "semanticHeaderSignature": [
                        {
                            "relativeColumn": 1,
                            "columnRole": "MEASURE_VALUE",
                            "headerTexts": ["compression"],
                        },
                        {
                            "relativeColumn": 2,
                            "columnRole": "MEASURE_VALUE",
                            "headerTexts": ["tension"],
                        },
                    ],
                }
            ]
        }

        result = match_registered_recipe(table, registry)
        changed = copy.deepcopy(table)
        changed["previewRows"][0]["cells"][1]["value"] = "Other metric"
        changed["numericColumns"][0]["headerTexts"] = ["Other metric"]
        no_match = match_registered_recipe(changed, registry)

        self.assertEqual("EXACT_RECIPE_MATCH", result["status"])
        self.assertEqual(0, result["aiCalls"])
        self.assertEqual("NO_REGISTERED_RECIPE", no_match["status"])

    def test_registry_accepts_verified_zero_ai_historical_recipe(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for name in ("recipes", "replays", "telemetry"):
                (root / name).mkdir()
            structure_id = "table-structure-historical"
            recipe_id = "structure-recipe-historical"
            signature = [
                {
                    "relativeColumn": 1,
                    "columnRole": "MEASURE_VALUE",
                    "headerTexts": ["compression"],
                }
            ]
            priority = {
                "queue": [
                    {
                        "tableStructureId": structure_id,
                        "fingerprintSha256": "a" * 64,
                        "semanticHeaderSignature": signature,
                        "workbookCount": 2,
                    }
                ]
            }
            (root / "recipes" / "recipe.json").write_text(
                json.dumps(
                    {
                        "recipeId": recipe_id,
                        "recipeVersion": 1,
                        "status": (
                            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_"
                            "NEEDS_CANONICAL_REVIEW"
                        ),
                        "match": {
                            "tableStructureId": structure_id,
                            "fingerprintSha256": "a" * 64,
                            "semanticHeaderSha256": "b" * 64,
                            "semanticHeaderSignature": signature,
                        },
                        "decision": {
                            "source": "HISTORICAL_989_CONSENSUS",
                            "aiCallCount": 0,
                        },
                    }
                ),
                encoding="utf-8",
            )
            (root / "replays" / "replay.json").write_text(
                json.dumps(
                    {
                        "recipeId": recipe_id,
                        "status": (
                            "VERIFIED_DETERMINISTIC_STRUCTURE_REPLAY_"
                            "NEEDS_CANONICAL_REVIEW"
                        ),
                        "summary": {
                            "memberCount": 2,
                            "passed": 2,
                            "failed": 0,
                        },
                    }
                ),
                encoding="utf-8",
            )

            registry = build_recipe_registry(
                priority_report=priority,
                recipe_root=root / "recipes",
                replay_root=root / "replays",
                telemetry_root=root / "telemetry",
            )

        self.assertEqual(1, registry["summary"]["registeredRecipeCount"])
        self.assertEqual(
            "HISTORICAL_989_CONSENSUS",
            registry["recipes"][0]["decisionSource"],
        )
        self.assertEqual(0, registry["recipes"][0]["decisionAiCalls"])


class HistoricalSemanticBootstrapTests(unittest.TestCase):
    def test_consistent_descriptive_history_builds_zero_ai_contract(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "requests").mkdir()
            (root / "analyses").mkdir()
            table = TableRecipeProposalTests._table(
                "table-one",
                first_value=12.34,
                second_value=8.1,
            )
            signature = semantic_header_signature(table)
            signature_sha256 = _semantic_header_sha256(signature)
            for index in (1, 2):
                table_copy = copy.deepcopy(table)
                table_copy["tableId"] = f"table-{index}"
                for column in table_copy["numericColumns"]:
                    column["columnId"] = (
                        f"table-{index}_col_{column['column']}"
                    )
                request_name = f"request-{index}.json"
                (root / "requests" / request_name).write_text(
                    json.dumps({"tables": [table_copy]}),
                    encoding="utf-8",
                )
                (root / "analyses" / request_name).write_text(
                    json.dumps(
                        {
                            "tables": [
                                {
                                    "tableId": table_copy["tableId"],
                                    "type": "DESCRIPTIVE",
                                    "confidence": "HIGH",
                                    "metrics": [
                                        {
                                            "name": "Compression",
                                            "unit": "MPa",
                                            "axisRefs": [
                                                table_copy[
                                                    "numericColumns"
                                                ][0]["columnId"]
                                            ],
                                        },
                                        {
                                            "name": "Tension",
                                            "unit": "MPa",
                                            "axisRefs": [
                                                table_copy[
                                                    "numericColumns"
                                                ][1]["columnId"]
                                            ],
                                        },
                                    ],
                                }
                            ]
                        }
                    ),
                    encoding="utf-8",
                )
            priority = {
                "queue": [
                    {
                        "semanticHeaderSha256": signature_sha256,
                        "semanticHeaderSignature": signature,
                    }
                ]
            }

            catalog = build_historical_semantic_contract_catalog(
                priority_report=priority,
                historical_batch_root=root,
                generated_at="2026-07-27T00:00:00+00:00",
            )

        self.assertEqual(0, catalog["summary"]["aiCalls"])
        self.assertEqual(1, catalog["summary"]["readyContractCount"])
        contract = catalog["contracts"][0]
        self.assertEqual(2, contract["fullMappingSupport"])
        self.assertEqual(
            ["Compression", "Tension"],
            [
                metric["canonicalName"]
                for metric in contract["metricColumns"]
            ],
        )


class StructureCompletionRecoveryTests(unittest.TestCase):
    def test_verified_registry_replaces_prior_contract_quarantine(self) -> None:
        state = {
            "outcomes": {
                "table-structure-recovered": {
                    "status": "QUARANTINED_RECIPE_CONTRACT_FAILURE",
                    "error": "old selector failure",
                }
            }
        }
        registry = {
            "recipes": [
                {
                    "tableStructureId": "table-structure-recovered",
                    "decisionSource": "BOUNDED_AI_STRUCTURE_DECISION",
                    "decisionAiCalls": 1,
                    "tableCount": 2,
                    "workbookCount": 2,
                    "recipeId": "structure-recipe-recovered",
                }
            ]
        }

        _record_registered_outcomes(state, registry)

        self.assertEqual(
            {
                "status": "REGISTERED_AI_REPLAY",
                "decisionSource": "BOUNDED_AI_STRUCTURE_DECISION",
                "decisionAiCalls": 1,
                "tableCount": 2,
                "workbookCount": 2,
                "recipeId": "structure-recipe-recovered",
            },
            state["outcomes"]["table-structure-recovered"],
        )

    def test_summary_ignores_outcomes_outside_current_priority_queue(
        self,
    ) -> None:
        state = {
            "inputs": {},
            "outcomes": {
                "table-structure-current": {"status": "REGISTERED"},
                "table-structure-stale": {
                    "status": "QUARANTINED_AI_DECISION"
                },
            },
        }

        _summarize(
            state,
            priority_report={
                "queue": [
                    {"tableStructureId": "table-structure-current"}
                ]
            },
            telemetry={},
            registry={"recipes": [], "summary": {}},
        )

        self.assertEqual(1, state["summary"]["queueStructureCount"])
        self.assertEqual(1, state["summary"]["completedStructureCount"])
        self.assertEqual(0, state["summary"]["unresolvedStructureCount"])
        self.assertEqual(
            {"REGISTERED": 1},
            state["summary"]["outcomeCounts"],
        )

    def test_resume_updates_priority_report_identity(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            priority = root / "expanded-priority.json"
            state_path = root / "completion-state.json"
            priority.write_text(
                json.dumps({"queue": []}),
                encoding="utf-8",
            )
            state_path.write_text(
                json.dumps(
                    {
                        "schemaVersion": COMPLETION_STATE_SCHEMA_VERSION,
                        "engineVersion": "test",
                        "startedAt": "2026-07-28T00:00:00+00:00",
                        "updatedAt": "2026-07-28T00:00:00+00:00",
                        "status": "COMPLETED",
                        "inputs": {
                            "priorityReport": "stale-priority.json",
                        },
                        "outcomes": {},
                        "events": [],
                        "summary": {},
                    }
                ),
                encoding="utf-8",
            )

            completed = complete_structure_queue(
                priority_report_path=priority,
                recipe_root=root / "recipes",
                replay_root=root / "replays",
                decision_root=root / "decisions",
                telemetry_root=root / "telemetry",
                quarantine_root=root / "quarantine",
                registry_path=root / "registry.json",
                state_path=state_path,
                max_total_ai_calls=0,
            )

        self.assertEqual(
            str(priority.resolve()),
            completed["inputs"]["priorityReport"],
        )
        self.assertEqual(0, completed["summary"]["unresolvedStructureCount"])


class IncrementalCoverageTests(unittest.TestCase):
    def test_report_closes_workbook_table_queue_and_ai_accounting(
        self,
    ) -> None:
        report = build_incremental_coverage_report(
            table_match_report={
                "summary": {
                    "eligibleWorkbookCount": 3,
                    "completedWorkbookCount": 2,
                    "failedWorkbookCount": 1,
                    "tableCount": 5,
                },
                "failures": [
                    {
                        "relativePath": "broken.xlsx",
                        "errorType": "SemanticPacketError",
                        "message": "cell count mismatch",
                    }
                ],
            },
            table_structure_catalog={
                "summary": {
                    "tableCount": 5,
                    "quantitativeTablesInReusableStructures": 3,
                },
                "structures": [
                    {
                        "tableStructureId": "base-quantitative",
                        "tableCount": 3,
                        "workbookCount": 2,
                        "quantitative": True,
                    },
                    {
                        "tableStructureId": "base-text",
                        "tableCount": 2,
                        "workbookCount": 2,
                        "quantitative": False,
                    },
                ],
            },
            priority_report={
                "summary": {
                    "coveredTableCount": 2,
                    "coveredWorkbookReferences": 2,
                },
                "queue": [
                    {
                        "tableStructureId": "queue-one",
                        "baseTableStructureId": "base-quantitative",
                        "tableCount": 2,
                        "workbookCount": 2,
                    }
                ],
            },
            completion_state={
                "status": "COMPLETED",
                "outcomes": {
                    "queue-one": {
                        "status": "REGISTERED_AI_REPLAY",
                        "tableCount": 2,
                        "workbookCount": 2,
                    }
                },
                "summary": {
                    "unresolvedStructureCount": 0,
                    "aiCallsAttempted": 1,
                    "aiCallsSucceeded": 1,
                    "aiCallsFailed": 0,
                    "retryCount": 0,
                    "fileLevelAiCalls": 0,
                },
            },
            recipe_registry={
                "summary": {
                    "registeredRecipeCount": 1,
                    "registeredTableCount": 2,
                }
            },
            telemetry_values=[
                {
                    "tableStructureId": "queue-one",
                    "status": "SUCCEEDED",
                    "aiCallBudget": 1,
                    "aiCallsAttempted": 1,
                    "aiCallsSucceeded": 1,
                    "retryCount": 0,
                    "durationMs": 2_000,
                    "promptBytes": 100,
                    "outputBytes": 20,
                }
            ],
            generated_at="2026-07-27T00:00:00+00:00",
        )

        self.assertEqual(
            "ACCOUNTED_WITH_EXPLICIT_TABLE_REQUEST_FAILURES",
            report["status"],
        )
        self.assertEqual(
            {
                "eligible": 3,
                "sourceAndCapturePresent": 3,
                "tableRequestCompleted": 2,
                "tableRequestFailed": 1,
                "tableRequestFailures": [
                    {
                        "relativePath": "broken.xlsx",
                        "errorType": "SemanticPacketError",
                        "message": "cell count mismatch",
                    }
                ],
            },
            report["workbookCoverage"],
        )
        self.assertEqual(
            5,
            report["tableCoverage"]["catalogedTableCount"],
        )
        self.assertEqual(2, report["tableCoverage"][
            "registeredProgramExtractionTableCount"
        ])
        self.assertEqual(3, report["tableCoverage"][
            "outsideRepeatedQueueTableCount"
        ])
        self.assertEqual(
            1,
            report["tableCoverage"]["outsideRepeatedQueue"][
                "quantitativeTableCount"
            ],
        )
        self.assertEqual(0, report["queueCoverage"][
            "unresolvedStructureCount"
        ])
        self.assertEqual(0, report["aiUsage"]["retryCount"])


if __name__ == "__main__":
    unittest.main()
