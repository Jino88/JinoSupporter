from __future__ import annotations

import importlib.util
import sqlite3
import tempfile
import unittest
from pathlib import Path


SERVICE_DIR = Path(__file__).parents[1]
NOW = "2026-07-17T15:00:00Z"


def load_module(name: str, file_name: str):
    specification = importlib.util.spec_from_file_location(
        name,
        SERVICE_DIR / file_name,
    )
    assert specification is not None and specification.loader is not None
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


schema = load_module("related_fixture_schema", "inference_data_ai_schema.py")
related = load_module("inference_data_ai_related", "inference_data_ai_related.py")


class RelatedStudyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.database = self.root / "knowledge.sqlite"
        self.connection = sqlite3.connect(self.database)
        self._create_legacy_parent_stubs()
        schema.ensure_knowledge_schema(self.connection, lambda: NOW)
        self._insert_fixture()
        self.connection.commit()

    def tearDown(self) -> None:
        self.connection.close()
        self.temp.cleanup()

    def _create_legacy_parent_stubs(self) -> None:
        for table, primary_key in [
            ("workbooks", "workbook_id"),
            ("analysis_reports", "analysis_report_id"),
            ("analysis_review_items", "review_item_id"),
            ("analysis_cohorts", "cohort_id"),
            ("analysis_metrics", "metric_id"),
            ("analysis_metric_values", "metric_value_id"),
            ("analysis_comparisons", "comparison_id"),
            ("analysis_evidence", "evidence_id"),
            ("analysis_conclusions", "conclusion_id"),
        ]:
            self.connection.execute(
                f"CREATE TABLE {table}({primary_key} INTEGER PRIMARY KEY)"
            )

    def _concept(self, uid: str, kind: str, name: str) -> int:
        cursor = self.connection.execute(
            """
            INSERT INTO knowledge_concepts(
                concept_uid, concept_kind, canonical_name, normalized_name,
                description, created_at, updated_at
            ) VALUES (?, ?, ?, ?, '', ?, ?)
            """,
            (uid, kind, name, related.normalize_text(name), NOW, NOW),
        )
        return int(cursor.lastrowid)

    def _source(
        self,
        suffix: str,
        *,
        content_hash: str,
        path_name: str,
        current: int = 1,
    ) -> tuple[int, int]:
        document = self.connection.execute(
            """
            INSERT INTO source_documents(
                document_uid, dataset, source_path, original_file_name,
                source_kind, lifecycle_status, created_at, updated_at
            ) VALUES (?, 'fixture', ?, ?, 'XLSX', 'ACTIVE', ?, ?)
            """,
            (
                f"document_{suffix}",
                str(self.root / path_name),
                path_name,
                NOW,
                NOW,
            ),
        )
        document_id = int(document.lastrowid)
        revision = self.connection.execute(
            """
            INSERT INTO source_revisions(
                revision_uid, document_id, source_fingerprint,
                fingerprint_kind, content_sha256, extractor_name,
                extractor_version, capture_contract, capture_status,
                is_current, captured_at
            ) VALUES (?, ?, ?, 'SHA256', ?, 'capture-v2', '2',
                      'openxml-v2', 'CAPTURED', ?, ?)
            """,
            (
                f"revision_{suffix}",
                document_id,
                f"fingerprint_{suffix}",
                content_hash,
                current,
                NOW,
            ),
        )
        return document_id, int(revision.lastrowid)

    def _study(
        self,
        suffix: str,
        document_id: int,
        revision_id: int,
        *,
        data_id: str,
        title: str,
        factor_original: str = "",
        factor_concept: int | None = None,
        outcome_original: str = "",
        outcome_concept: int | None = None,
        context_original: str = "",
        context_concept: int | None = None,
    ) -> int:
        analysis = self.connection.execute(
            """
            INSERT INTO workbook_analyses(
                analysis_uid, public_analysis_id, document_id, revision_id,
                analysis_key, title, analysis_status, verification_status,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'COMPLETE', 'VERIFIED', ?, ?)
            """,
            (
                f"analysis_{suffix}",
                f"ANALYSIS-{suffix.upper()}",
                document_id,
                revision_id,
                f"analysis-{suffix}",
                title,
                NOW,
                NOW,
            ),
        )
        study = self.connection.execute(
            """
            INSERT INTO knowledge_studies(
                study_uid, public_data_id, workbook_analysis_id, study_key,
                title, verification_status, comparability_status,
                confounding_status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, 'VERIFIED', 'VALID', 'NONE', ?, ?)
            """,
            (
                f"study_{suffix}",
                data_id,
                int(analysis.lastrowid),
                f"study-{suffix}",
                title,
                NOW,
                NOW,
            ),
        )
        study_id = int(study.lastrowid)
        if factor_original or factor_concept is not None:
            self.connection.execute(
                """
                INSERT INTO knowledge_factors(
                    factor_uid, study_id, concept_id, factor_key,
                    original_label, verification_status
                ) VALUES (?, ?, ?, 'factor', ?, 'VERIFIED')
                """,
                (
                    f"factor_{suffix}",
                    study_id,
                    factor_concept,
                    factor_original,
                ),
            )
        if outcome_original or outcome_concept is not None:
            self.connection.execute(
                """
                INSERT INTO knowledge_outcomes(
                    outcome_uid, study_id, outcome_key, concept_id,
                    original_label, verification_status
                ) VALUES (?, ?, 'outcome', ?, ?, 'VERIFIED')
                """,
                (
                    f"outcome_{suffix}",
                    study_id,
                    outcome_concept,
                    outcome_original,
                ),
            )
        if context_original or context_concept is not None:
            self.connection.execute(
                """
                INSERT INTO knowledge_study_contexts(
                    context_uid, study_id, context_kind, concept_id,
                    original_value, normalized_value, verification_status
                ) VALUES (?, ?, 'ARBITRARY_CONTEXT', ?, ?, ?, 'VERIFIED')
                """,
                (
                    f"context_{suffix}",
                    study_id,
                    context_concept,
                    context_original,
                    related.normalize_text(context_original),
                ),
            )
        return study_id

    def _insert_fixture(self) -> None:
        factor = self._concept(
            "concept_factor_zephyr",
            "ARBITRARY_FACTOR",
            "Zephyr flux",
        )
        self.connection.execute(
            """
            INSERT INTO knowledge_concept_aliases(
                alias_uid, concept_id, alias_text, normalized_alias,
                language, source, confidence, created_at
            ) VALUES (
                'alias_factor_whisper', ?, 'adhesive-whisper',
                'adhesive-whisper', '', 'HUMAN_APPROVED', 1, ?
            )
            """,
            (factor, NOW),
        )
        outcome = self._concept(
            "concept_outcome_quasar",
            "ARBITRARY_OUTCOME",
            "Quasar fracture",
        )
        context = self._concept(
            "concept_context_monsoon",
            "ARBITRARY_CONTEXT",
            "Monsoon lot",
        )
        same_hash = "a" * 64

        target_doc, target_rev = self._source(
            "target",
            content_hash=same_hash,
            path_name="target.xlsx",
        )
        self._study(
            "target",
            target_doc,
            target_rev,
            data_id="DATA-TARGET",
            title="Zephyr trial alpha",
            factor_original="제피르 유량",
            factor_concept=factor,
            outcome_original="퀘이사 파손",
            outcome_concept=outcome,
            context_original="우기 lot A",
            context_concept=context,
        )

        duplicate_doc, duplicate_rev = self._source(
            "duplicate",
            content_hash=same_hash,
            path_name="copied elsewhere.xlsx",
        )
        self._study(
            "duplicate",
            duplicate_doc,
            duplicate_rev,
            data_id="DATA-DUPLICATE",
            title="Zephyr trial alpha",
            factor_original="제피르 유량",
            factor_concept=factor,
            outcome_original="퀘이사 파손",
            outcome_concept=outcome,
        )

        close_doc, close_rev = self._source(
            "close",
            content_hash="b" * 64,
            path_name="close.xlsx",
        )
        self._study(
            "close",
            close_doc,
            close_rev,
            data_id="DATA-CLOSE",
            title="Zephyr trial beta",
            factor_original="다른 원문",
            factor_concept=factor,
            outcome_original="다른 결과",
            outcome_concept=outcome,
            context_original="건기 lot B",
            context_concept=context,
        )

        factor_doc, factor_rev = self._source(
            "factor_only",
            content_hash="c" * 64,
            path_name="factor-only.xlsx",
        )
        self._study(
            "factor_only",
            factor_doc,
            factor_rev,
            data_id="DATA-FACTOR-ONLY",
            title="Unrelated heading",
            factor_original="제피르 흐름",
            factor_concept=factor,
            outcome_original="독립 출력",
        )

        unrelated_doc, unrelated_rev = self._source(
            "unrelated",
            content_hash="d" * 64,
            path_name="unrelated.xlsx",
        )
        self._study(
            "unrelated",
            unrelated_doc,
            unrelated_rev,
            data_id="DATA-UNRELATED",
            title="Ocean salinity",
            factor_original="염도",
            outcome_original="전도도",
            context_original="해수",
        )

        stale_doc, stale_rev = self._source(
            "stale",
            content_hash="e" * 64,
            path_name="stale.xlsx",
            current=0,
        )
        self._study(
            "stale",
            stale_doc,
            stale_rev,
            data_id="DATA-STALE",
            title="Zephyr trial stale",
            factor_original="제피르 유량",
            factor_concept=factor,
            outcome_original="퀘이사 파손",
            outcome_concept=outcome,
        )

    def test_data_id_reports_exact_duplicates_and_ranked_current_studies(self) -> None:
        report = related.build_related_studies(
            self.connection,
            "data-target",
            limit=10,
        )

        self.assertEqual(
            related.RELATED_STUDIES_SCHEMA_VERSION,
            report["schemaVersion"],
        )
        self.assertEqual("PUBLIC_DATA_ID", report["targetIdentifierType"])
        self.assertEqual("DATA-TARGET", report["targetIdentifier"])
        self.assertEqual(
            ["DATA-DUPLICATE"],
            report["exactContentDuplicates"][0]["publicDataIds"],
        )
        self.assertTrue(
            report["exactContentDuplicates"][0]["source"]["sourcePath"].endswith(
                "copied elsewhere.xlsx"
            )
        )

        ids = [item["publicDataId"] for item in report["relatedStudies"]]
        self.assertEqual(["DATA-CLOSE", "DATA-FACTOR-ONLY"], ids)
        self.assertNotIn("DATA-DUPLICATE", ids)
        self.assertNotIn("DATA-UNRELATED", ids)
        self.assertNotIn("DATA-STALE", ids)
        self.assertGreater(
            report["relatedStudies"][0]["similarityScore"],
            report["relatedStudies"][1]["similarityScore"],
        )
        self.assertIn(
            "adhesive-whisper",
            report["target"]["termProfile"]["factor"],
        )
        self.assertIn(
            "conceptAliases",
            report["scoring"]["fieldScope"]["factor"],
        )
        first = report["relatedStudies"][0]
        self.assertEqual(
            ["context", "factor", "outcome", "studyTitle"],
            sorted(reason["category"] for reason in first["sharedTermReasons"]),
        )
        self.assertFalse(first["similarityIsRelationshipEvidence"])
        self.assertFalse(first["similarityIsCausalEvidence"])
        self.assertFalse(report["safety"]["imagesAnalyzed"])

    def test_revision_uid_uses_all_revision_studies_and_is_deterministic(self) -> None:
        first = related.build_related_studies(
            self.connection,
            "REVISION_TARGET",
            limit=1,
        )
        second = related.build_related_studies(
            self.connection,
            "revision_target",
            limit=1,
        )

        self.assertEqual("REVISION_UID", first["targetIdentifierType"])
        self.assertEqual("revision_target", first["targetIdentifier"])
        self.assertEqual(1, len(first["relatedStudies"]))
        self.assertEqual(
            related.related_studies_json_bytes(first),
            related.related_studies_json_bytes(second),
        )

    def test_readonly_path_api_does_not_mutate_database(self) -> None:
        before = self.database.read_bytes()
        report = related.build_related_studies_from_db(
            self.database,
            "DATA-TARGET",
        )
        after = self.database.read_bytes()

        self.assertEqual("DATA-TARGET", report["targetIdentifier"])
        self.assertEqual(before, after)

    def test_unknown_target_and_invalid_limit_are_rejected(self) -> None:
        with self.assertRaisesRegex(related.RelatedStudyError, "Unknown"):
            related.build_related_studies(self.connection, "DATA-MISSING")
        with self.assertRaisesRegex(ValueError, "positive integer"):
            related.build_related_studies(
                self.connection,
                "DATA-TARGET",
                limit=0,
            )


if __name__ == "__main__":
    unittest.main()
