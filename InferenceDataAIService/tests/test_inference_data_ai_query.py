from __future__ import annotations

import importlib.util
import sqlite3
import tempfile
import unittest
from pathlib import Path


SERVICE_DIR = Path(__file__).parents[1]


def load_module(name: str, file_name: str):
    specification = importlib.util.spec_from_file_location(name, SERVICE_DIR / file_name)
    assert specification is not None and specification.loader is not None
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


schema = load_module("query_fixture_schema", "inference_data_ai_schema.py")
query = load_module("inference_data_ai_query", "inference_data_ai_query.py")


NOW = "2026-07-17T04:00:00Z"


class CanonicalEvidenceQueryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.database = self.root / "knowledge.sqlite"
        self.connection = sqlite3.connect(self.database)
        self.connection.execute(
            """
            CREATE TABLE schema_migrations(
                migration_name TEXT PRIMARY KEY,
                applied_at TEXT NOT NULL
            )
            """
        )
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

    def _insert_fixture(self) -> None:
        source_path = str(self.root / "arbitrary-novel-study.xlsx")
        self.source_path = source_path
        self.connection.execute(
            """
            INSERT INTO source_documents(
                document_uid, dataset, source_path, original_file_name,
                source_kind, created_at, updated_at
            ) VALUES ('document_novel', 'fixture', ?, 'arbitrary-novel-study.xlsx',
                      'XLSX', ?, ?)
            """,
            (source_path, NOW, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO source_revisions(
                revision_uid, document_id, source_fingerprint,
                fingerprint_kind, content_sha256, extractor_name,
                extractor_version, capture_contract, capture_status,
                is_current, captured_at
            ) VALUES (
                'revision_novel', 1, 'fingerprint-novel', 'SHA256',
                'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
                'capture-v2', '2', 'openxml-v2', 'CAPTURED', 1, ?
            )
            """,
            (NOW,),
        )
        self.connection.execute(
            """
            INSERT INTO workbook_analyses(
                analysis_uid, public_analysis_id, document_id, revision_id,
                analysis_key, title, purpose, scope_text, analysis_status,
                verification_status, decision_text, consolidated_summary,
                created_at, updated_at
            ) VALUES (
                'analysis_novel', 'ANALYSIS-NOVEL', 1, 1, 'novel-analysis',
                'Arbitrary cryoflux evaluation', 'Compare novel conditions',
                'No predefined domain vocabulary', 'COMPLETE', 'VERIFIED',
                'Use verified comparisons only', 'Four candidate studies', ?, ?
            )
            """,
            (NOW, NOW),
        )

        factor_concept_id = self._insert_concept(
            "concept_cryoflux",
            "PROCESS_FACTOR",
            "CryoFlux Ω cadence",
            "cryoflux ω cadence",
            "극저온플럭스",
        )
        outcome_concept_id = self._insert_concept(
            "concept_nebula",
            "OUTCOME",
            "Nebula fracture rate",
            "nebula fracture rate",
            "성운파손",
        )

        self._insert_study_case(
            suffix="good",
            public_data_id="DATA-NOVEL-GOOD",
            public_comparison_id="CMP-NOVEL-GOOD",
            public_effect_id="EFF-NOVEL-GOOD",
            public_evidence_id="EVD-NOVEL-GOOD",
            factor_concept_id=factor_concept_id,
            outcome_concept_id=outcome_concept_id,
            comparison_validity="VALID",
            comparison_confounding="NONE",
            comparison_verification="VERIFIED",
            comparison_eligible=1,
            effect_verification="VERIFIED",
            effect_eligible=1,
            evidence_sheet="Raw Ω Results",
            evidence_range="C4:H12",
        )
        self._insert_study_case(
            suffix="confounded",
            public_data_id="DATA-NOVEL-CONFOUNDED",
            public_comparison_id="CMP-NOVEL-CONFOUNDED",
            public_effect_id="EFF-NOVEL-CONFOUNDED",
            public_evidence_id="EVD-NOVEL-CONFOUNDED",
            factor_concept_id=factor_concept_id,
            outcome_concept_id=outcome_concept_id,
            comparison_validity="VALID",
            comparison_confounding="CONFOUNDED",
            comparison_verification="VERIFIED",
            comparison_eligible=0,
            effect_verification="VERIFIED",
            effect_eligible=0,
            evidence_sheet="Confounded",
            evidence_range="A2:F8",
        )
        self._insert_study_case(
            suffix="review",
            public_data_id="DATA-NOVEL-REVIEW",
            public_comparison_id="CMP-NOVEL-REVIEW",
            public_effect_id="EFF-NOVEL-REVIEW",
            public_evidence_id="EVD-NOVEL-REVIEW",
            factor_concept_id=factor_concept_id,
            outcome_concept_id=outcome_concept_id,
            comparison_validity="NEEDS_REVIEW",
            comparison_confounding="UNASSESSED",
            comparison_verification="NEEDS_REVIEW",
            comparison_eligible=0,
            effect_verification="NEEDS_REVIEW",
            effect_eligible=0,
            evidence_sheet="Review",
            evidence_range="B3:E7",
        )
        self._insert_study_case(
            suffix="invalid",
            public_data_id="DATA-NOVEL-INVALID",
            public_comparison_id="CMP-NOVEL-INVALID",
            public_effect_id="EFF-NOVEL-INVALID",
            public_evidence_id="EVD-NOVEL-INVALID",
            factor_concept_id=factor_concept_id,
            outcome_concept_id=outcome_concept_id,
            comparison_validity="INVALID",
            comparison_confounding="NONE",
            comparison_verification="VERIFIED",
            comparison_eligible=0,
            effect_verification="INVALID",
            effect_eligible=0,
            evidence_sheet="Invalid",
            evidence_range="D5:G9",
        )

        # A fully unrelated study proves that retrieval is based on actual
        # question terms rather than returning every database row.
        self.connection.execute(
            """
            INSERT INTO knowledge_studies(
                study_uid, public_data_id, workbook_analysis_id, study_key,
                title, summary_text, verification_status,
                comparability_status, confounding_status, created_at, updated_at
            ) VALUES (
                'study_unrelated', 'DATA-UNRELATED', 1, 'unrelated',
                'Ocean salinity baseline', 'Marine conductivity only',
                'VERIFIED', 'VALID', 'NONE', ?, ?
            )
            """,
            (NOW, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO source_documents(
                document_uid, dataset, source_path, original_file_name,
                source_kind, created_at, updated_at
            ) VALUES (
                'document_terminal', 'fixture', ?,
                'Zephyr XRAY image review.xlsx', 'XLSX', ?, ?
            )
            """,
            (
                str(self.root / "Zephyr XRAY image review.xlsx"),
                NOW,
                NOW,
            ),
        )
        self.connection.execute(
            """
            INSERT INTO source_revisions(
                revision_uid, document_id, source_fingerprint,
                fingerprint_kind, content_sha256, extractor_name,
                extractor_version, capture_contract, capture_status,
                source_content_status, is_current, captured_at
            ) VALUES (
                'revision_terminal', 2, 'fingerprint-terminal', 'SHA256',
                ?, 'capture-v2', '2', 'openxml-v2', 'CAPTURED',
                'CAPTURED', 1, ?
            )
            """,
            ("f" * 64, NOW),
        )
        self.connection.execute(
            """
            INSERT INTO workbook_analyses(
                analysis_uid, public_analysis_id, document_id, revision_id,
                analysis_key, title, analysis_status, verification_status,
                consolidated_summary, limitations_json, created_at, updated_at
            ) VALUES (
                'analysis_terminal', 'ANALYSIS-TERMINAL', 2, 2,
                'terminal-analysis', 'Zephyr XRAY image review',
                'NO_TABULAR_EVIDENCE', 'EXCLUDED',
                'No reviewable Zephyr table was captured.',
                '["Images are outside scope."]', ?, ?
            )
            """,
            (NOW, NOW),
        )

    def _insert_concept(
        self,
        uid: str,
        kind: str,
        canonical_name: str,
        normalized_name: str,
        alias: str,
    ) -> int:
        cursor = self.connection.execute(
            """
            INSERT INTO knowledge_concepts(
                concept_uid, concept_kind, canonical_name, normalized_name,
                description, created_at, updated_at
            ) VALUES (?, ?, ?, ?, 'fixture-only novel term', ?, ?)
            """,
            (uid, kind, canonical_name, normalized_name, NOW, NOW),
        )
        concept_id = int(cursor.lastrowid)
        self.connection.execute(
            """
            INSERT INTO knowledge_concept_aliases(
                alias_uid, concept_id, alias_text, normalized_alias,
                language, source, confidence, created_at
            ) VALUES (?, ?, ?, ?, 'ko', 'FIXTURE', 1, ?)
            """,
            (
                f"alias_{uid}",
                concept_id,
                alias,
                query.normalize_text(alias),
                NOW,
            ),
        )
        return concept_id

    def _insert_study_case(
        self,
        *,
        suffix: str,
        public_data_id: str,
        public_comparison_id: str,
        public_effect_id: str,
        public_evidence_id: str,
        factor_concept_id: int,
        outcome_concept_id: int,
        comparison_validity: str,
        comparison_confounding: str,
        comparison_verification: str,
        comparison_eligible: int,
        effect_verification: str,
        effect_eligible: int,
        evidence_sheet: str,
        evidence_range: str,
    ) -> None:
        cursor = self.connection.execute(
            """
            INSERT INTO knowledge_studies(
                study_uid, public_data_id, workbook_analysis_id, study_key,
                title, purpose, hypothesis, objective, design_type,
                comparison_basis, analysis_status, verification_status,
                comparability_status, confounding_status, decision_text,
                summary_text, created_at, updated_at
            ) VALUES (?, ?, 1, ?, ?, 'Novel factor evaluation',
                      'Cryoflux may alter nebula fracture',
                      'Compare matched arms', 'CONTROL_COMPARISON',
                      'Matched lot and sample basis', 'COMPLETE', 'VERIFIED',
                      'VALID', ?, 'Inspect exact evidence',
                      'ZQ-17 plasma cadence versus Nebula fracture rate', ?, ?)
            """,
            (
                f"study_{suffix}",
                public_data_id,
                suffix,
                f"CryoFlux Ω study {suffix}",
                (
                    "CONFOUNDED"
                    if comparison_confounding == "CONFOUNDED"
                    else "NONE"
                ),
                NOW,
                NOW,
            ),
        )
        study_id = int(cursor.lastrowid)
        self.connection.execute(
            """
            INSERT INTO knowledge_study_contexts(
                context_uid, study_id, context_kind, original_value,
                normalized_value, verification_status
            ) VALUES (?, ?, 'EQUIPMENT', 'AuroraRig-Ξ',
                      'aurorarig-ξ', 'VERIFIED')
            """,
            (f"context_{suffix}", study_id),
        )
        factor_cursor = self.connection.execute(
            """
            INSERT INTO knowledge_factors(
                factor_uid, study_id, concept_id, factor_key, factor_domain,
                original_label, baseline_condition, changed_condition,
                change_direction, isolation_status, verification_status
            ) VALUES (?, ?, ?, 'zq17-cadence', 'Novel unmapped process',
                      'ZQ-17 plasma cadence', '2 pulses', '5 pulses',
                      'INCREASE', ?, 'VERIFIED')
            """,
            (
                f"factor_{suffix}",
                study_id,
                factor_concept_id,
                (
                    "CONFOUNDED"
                    if comparison_confounding == "CONFOUNDED"
                    else "ISOLATED"
                ),
            ),
        )
        factor_id = int(factor_cursor.lastrowid)
        control_cursor = self.connection.execute(
            """
            INSERT INTO knowledge_arms(
                arm_uid, study_id, arm_key, arm_role, label, condition_text,
                sample_size, sample_basis, matching_basis, verification_status
            ) VALUES (?, ?, 'control', 'CONTROL', 'Baseline arm',
                      '2 pulses', 100, 'units', 'same lot', 'VERIFIED')
            """,
            (f"arm_{suffix}_control", study_id),
        )
        control_id = int(control_cursor.lastrowid)
        compared_cursor = self.connection.execute(
            """
            INSERT INTO knowledge_arms(
                arm_uid, study_id, arm_key, arm_role, label, condition_text,
                sample_size, sample_basis, matching_basis, verification_status
            ) VALUES (?, ?, 'changed', 'TREATMENT', 'Changed arm',
                      '5 pulses', 100, 'units', 'same lot', 'VERIFIED')
            """,
            (f"arm_{suffix}_changed", study_id),
        )
        compared_id = int(compared_cursor.lastrowid)
        self.connection.executemany(
            """
            INSERT INTO knowledge_arm_factor_values(
                arm_id, factor_id, original_value, value_number,
                is_baseline, held_constant
            ) VALUES (?, ?, ?, ?, ?, 0)
            """,
            [
                (control_id, factor_id, "2 pulses", 2, 1),
                (compared_id, factor_id, "5 pulses", 5, 0),
            ],
        )
        outcome_cursor = self.connection.execute(
            """
            INSERT INTO knowledge_outcomes(
                outcome_uid, study_id, outcome_key, concept_id,
                original_label, outcome_domain, metric_type, original_unit,
                denominator_basis, favorable_direction, definition_text,
                verification_status
            ) VALUES (?, ?, 'nebula-fracture', ?,
                      'Nebula fracture rate', 'Novel unmapped outcome',
                      'RATE', '%', '100 tested units', 'LOWER',
                      'Fractured nebula units divided by tested units',
                      'VERIFIED')
            """,
            (f"outcome_{suffix}", study_id, outcome_concept_id),
        )
        outcome_id = int(outcome_cursor.lastrowid)
        self.connection.executemany(
            """
            INSERT INTO knowledge_observations(
                observation_uid, outcome_id, arm_id, observation_key,
                value_number, value_text, numerator, denominator,
                sample_size, result_status, verification_status
            ) VALUES (?, ?, ?, 'overall', ?, ?, ?, 100, 100, 'OBSERVED', 'VERIFIED')
            """,
            [
                (f"observation_{suffix}_control", outcome_id, control_id, 10, "10%", 10),
                (f"observation_{suffix}_changed", outcome_id, compared_id, 15, "15%", 15),
            ],
        )
        comparison_cursor = self.connection.execute(
            """
            INSERT INTO knowledge_comparisons(
                comparison_uid, public_comparison_id, study_id,
                comparison_key, compared_arm_id, control_arm_id,
                design_type, matching_basis, validity_status,
                confounding_status, exclusion_reason, direction,
                summary_text, aggregation_eligible, verification_status
            ) VALUES (?, ?, ?, 'changed-vs-control', ?, ?,
                      'CONTROL_COMPARISON', 'same lot and sample basis',
                      ?, ?, ?, 'INCREASE',
                      'Changed arm is five percentage points higher',
                      ?, ?)
            """,
            (
                f"comparison_{suffix}",
                public_comparison_id,
                study_id,
                compared_id,
                control_id,
                comparison_validity,
                comparison_confounding,
                (
                    "Multiple factors changed"
                    if comparison_confounding == "CONFOUNDED"
                    else ""
                ),
                comparison_eligible,
                comparison_verification,
            ),
        )
        comparison_id = int(comparison_cursor.lastrowid)
        self.connection.execute(
            """
            INSERT INTO knowledge_effects(
                effect_uid, public_effect_id, comparison_id, outcome_id,
                effect_type, estimate, original_unit, formula_version,
                calculation_text, direction, aggregation_eligible,
                verification_status
            ) VALUES (?, ?, ?, ?, 'ABSOLUTE_DIFFERENCE', 5, '%p',
                      'fixture-v1', '15 - 10 = 5 %p', 'INCREASE', ?, ?)
            """,
            (
                f"effect_{suffix}",
                public_effect_id,
                comparison_id,
                outcome_id,
                effect_eligible,
                effect_verification,
            ),
        )
        start_col = ord(evidence_range.split(":")[0][0]) - ord("A") + 1
        start_row = int(evidence_range.split(":")[0][1:])
        end_col = ord(evidence_range.split(":")[1][0]) - ord("A") + 1
        end_row = int(evidence_range.split(":")[1][1:])
        evidence_cursor = self.connection.execute(
            """
            INSERT INTO evidence_items(
                evidence_uid, public_evidence_id, revision_id, evidence_kind,
                sheet_name, start_row, start_col, end_row, end_col,
                range_address, evidence_role, source_text, note,
                verification_status, created_at
            ) VALUES (?, ?, 1, 'CELL_RANGE', ?, ?, ?, ?, ?, ?,
                      'RESULT', 'Exact control and changed observations',
                      'Fixture exact citation', 'VERIFIED', ?)
            """,
            (
                f"evidence_{suffix}",
                public_evidence_id,
                evidence_sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                evidence_range,
                NOW,
            ),
        )
        evidence_id = int(evidence_cursor.lastrowid)
        self.connection.execute(
            """
            INSERT INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id, evidence_role, claim_scope
            ) VALUES ('EFFECT', ?, ?, 'RESULT', 'effect estimate')
            """,
            (f"effect_{suffix}", evidence_id),
        )

    def _insert_measurement_series(self, suffix: str) -> None:
        study_id = int(
            self.connection.execute(
                "SELECT study_id FROM knowledge_studies WHERE study_uid=?",
                (f"study_{suffix}",),
            ).fetchone()[0]
        )
        outcome_id = int(
            self.connection.execute(
                "SELECT outcome_id FROM knowledge_outcomes WHERE outcome_uid=?",
                (f"outcome_{suffix}",),
            ).fetchone()[0]
        )
        arm_id = int(
            self.connection.execute(
                "SELECT arm_id FROM knowledge_arms WHERE arm_uid=?",
                (f"arm_{suffix}_changed",),
            ).fetchone()[0]
        )
        series_cursor = self.connection.execute(
            """
            INSERT INTO knowledge_measurement_series(
                series_uid, public_series_id, study_id, outcome_id, arm_id,
                series_key, sheet_name, header_range, value_range,
                row_identity_range, axis_name, axis_source,
                original_axis_unit,
                original_value_unit, stratum_key, verification_status
            ) VALUES (
                ?, 'SER-SPECTRAL-REVIEW', ?, ?, ?, 'spectral-sweep',
                'Spectrum', 'B1:C1', 'B2:C3', 'A2:A3',
                'Frequency band', 'ROW_IDENTITY', 'Hz', 'dBSPL',
                'chamber-alpha',
                'NEEDS_REVIEW'
            )
            """,
            (f"series_{suffix}", study_id, outcome_id, arm_id),
        )
        series_id = int(series_cursor.lastrowid)
        self.connection.executemany(
            """
            INSERT INTO knowledge_measurement_points(
                point_uid, public_point_id, series_id, row_ordinal,
                column_ordinal, axis_label, axis_value,
                original_axis_unit, replicate_key, stratum_key,
                value_number, original_value_unit, source_revision_id,
                source_sheet_name, source_row_index, source_column_index,
                source_coordinate, axis_source_coordinate,
                replicate_source_coordinate, source_value_json,
                verification_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'Hz', ?, 'chamber-alpha',
                      ?, 'dBSPL', 1, 'Spectrum', ?, ?, ?, ?, ?, ?,
                      'NEEDS_REVIEW')
            """,
            [
                (
                    "point_review_1",
                    "MPT-REVIEW-1",
                    series_id,
                    1,
                    1,
                    "Low band",
                    100,
                    "replicate-alpha",
                    10,
                    2,
                    2,
                    "B2",
                    "A2",
                    "B1",
                    "10",
                ),
                (
                    "point_review_2",
                    "MPT-REVIEW-2",
                    series_id,
                    1,
                    2,
                    "Low band",
                    100,
                    "replicate-beta",
                    20,
                    2,
                    3,
                    "C2",
                    "A2",
                    "C1",
                    "20",
                ),
                (
                    "point_review_3",
                    "MPT-REVIEW-3",
                    series_id,
                    2,
                    1,
                    "High band",
                    200,
                    "replicate-alpha",
                    30,
                    3,
                    2,
                    "B3",
                    "A3",
                    "B1",
                    "30",
                ),
                (
                    "point_review_4",
                    "MPT-REVIEW-4",
                    series_id,
                    2,
                    2,
                    "High band",
                    200,
                    "replicate-beta",
                    40,
                    3,
                    3,
                    "C3",
                    "A3",
                    "C1",
                    "40",
                ),
            ],
        )
        evidence_cursor = self.connection.execute(
            """
            INSERT INTO evidence_items(
                evidence_uid, public_evidence_id, revision_id, evidence_kind,
                sheet_name, start_row, start_col, end_row, end_col,
                range_address, evidence_role, verification_status, created_at
            ) VALUES (
                'evidence_series_review', 'EVD-SERIES-REVIEW', 1,
                'CELL_RANGE', 'Spectrum', 2, 2, 3, 3, 'B2:C3',
                'MEASUREMENT_VALUES', 'VERIFIED', ?
            )
            """,
            (NOW,),
        )
        self.connection.execute(
            """
            INSERT INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id, evidence_role,
                claim_scope
            ) VALUES (
                'MEASUREMENT_SERIES', ?, ?, 'MEASUREMENT_VALUES',
                'descriptive point summary'
            )
            """,
            (f"series_{suffix}", int(evidence_cursor.lastrowid)),
        )
        self.connection.commit()

    def test_unicode_tokenizer_is_not_domain_limited(self) -> None:
        self.assertEqual(
            ["극저온플럭스", "zq-17", "α/β", "성운파손"],
            query.unicode_tokens("극저온플럭스 ZQ-17 α/β 성운파손"),
        )

    def test_korean_request_words_and_particles_do_not_pollute_search(
        self,
    ) -> None:
        tokens = query._unique_tokens(
            "AWF cooling time 중 Function NG가 가장 낮은 조건은 "
            "무엇이며 각 세부 NG를 함께 비교해줘"
        )
        self.assertEqual(
            ["awf", "cooling", "time", "function", "ng"],
            query._search_tokens(tokens),
        )

    def test_retrieval_returns_all_relevant_studies_and_exact_eligible_evidence(self) -> None:
        result = query.build_evidence_pack(
            self.connection,
            "극저온플럭스 성운파손",
        )

        self.assertEqual("canonical-evidence-pack-v1", result["schemaVersion"])
        self.assertEqual(4, result["summary"]["relevantStudyCount"])
        self.assertEqual(
            {
                "DATA-NOVEL-GOOD",
                "DATA-NOVEL-CONFOUNDED",
                "DATA-NOVEL-REVIEW",
                "DATA-NOVEL-INVALID",
            },
            {
                candidate["publicDataId"]
                for candidate in result["studyCandidates"]
            },
        )
        self.assertNotIn(
            "DATA-UNRELATED",
            {
                candidate["publicDataId"]
                for candidate in result["studyCandidates"]
            },
        )

        self.assertEqual(1, len(result["answerEligibleEffects"]))
        eligible = result["answerEligibleEffects"][0]
        self.assertEqual("DATA-NOVEL-GOOD", eligible["publicDataId"])
        self.assertEqual("CMP-NOVEL-GOOD", eligible["publicComparisonId"])
        self.assertEqual("EFF-NOVEL-GOOD", eligible["publicEffectId"])
        self.assertEqual(["EVD-NOVEL-GOOD"], eligible["publicEvidenceIds"])
        self.assertEqual(self.source_path, eligible["sourcePath"])

        self.assertEqual("Changed arm", eligible["comparison"]["comparedArm"]["label"])
        self.assertEqual("Baseline arm", eligible["comparison"]["controlArm"]["label"])
        self.assertEqual(
            15,
            eligible["observations"]["comparedArm"][0]["valueNumber"],
        )
        self.assertEqual(
            10,
            eligible["observations"]["controlArm"][0]["valueNumber"],
        )
        self.assertEqual(5, eligible["effect"]["estimate"])
        self.assertEqual("%p", eligible["effect"]["unit"])

        citation = eligible["evidence"][0]
        self.assertEqual("EVD-NOVEL-GOOD", citation["publicEvidenceId"])
        self.assertEqual(self.source_path, citation["sourcePath"])
        self.assertEqual("Raw Ω Results", citation["sheet"])
        self.assertEqual("C4:H12", citation["range"])
        self.assertEqual({"row": 4, "column": 3}, citation["start"])
        self.assertEqual({"row": 12, "column": 8}, citation["end"])
        self.assertEqual("VERIFIED", citation["verificationStatus"])

    def test_measurement_series_is_compact_searchable_and_source_backed(
        self,
    ) -> None:
        self._insert_measurement_series("review")

        result = query.build_evidence_pack(
            self.connection,
            "spectral-sweep Frequency dBSPL replicate-alpha",
        )

        self.assertEqual(1, result["summary"]["relevantStudyCount"])
        candidate = result["studyCandidates"][0]
        self.assertEqual("DATA-NOVEL-REVIEW", candidate["publicDataId"])
        self.assertEqual(1, len(candidate["measurementSeries"]))
        series = candidate["measurementSeries"][0]
        self.assertEqual("SER-SPECTRAL-REVIEW", series["publicSeriesId"])
        self.assertEqual("Frequency band", series["axisLabel"])
        self.assertEqual("ROW_IDENTITY", series["axisSource"])
        self.assertEqual("dBSPL", series["valueUnit"])
        self.assertEqual(
            ["replicate-alpha", "replicate-beta"],
            series["replicateKeys"],
        )
        self.assertEqual(
            {
                "pointCount": 4,
                "rawPointCount": 4,
                "aggregatePointCount": 0,
                "minimum": 10.0,
                "maximum": 40.0,
                "average": 25.0,
                "distinctAxisCount": 2,
                "distinctReplicateCount": 2,
                "aggregateReplicateCount": 0,
            },
            series["pointSummary"],
        )
        self.assertNotIn("points", series)
        matched_fields = {
            field["field"]
            for field in candidate["relevance"]["matchedFields"]
        }
        self.assertIn(
            "measurementSeries[0].replicateKey",
            matched_fields,
        )
        excluded = result["excludedCandidates"][0]
        self.assertEqual(
            "SER-SPECTRAL-REVIEW",
            excluded["descriptiveMeasurementSeries"][0][
                "publicSeriesId"
            ],
        )
        series_evidence = next(
            item
            for item in excluded["evidence"]
            if item["publicEvidenceId"] == "EVD-SERIES-REVIEW"
        )
        self.assertEqual("Spectrum", series_evidence["sheet"])
        self.assertEqual("B2:C3", series_evidence["range"])
        self.assertIn(
            {
                "entityType": "MEASUREMENT_SERIES",
                "entityUid": "series_review",
            },
            series_evidence["linkedEntities"],
        )
        numeric_value_query = query.build_evidence_pack(
            self.connection,
            "40",
        )
        self.assertEqual(
            0,
            numeric_value_query["summary"]["relevantStudyCount"],
        )

    def test_measurement_series_geometry_counts_source_coordinates_not_labels(
        self,
    ) -> None:
        self._insert_measurement_series("review")
        self.connection.execute(
            """
            UPDATE knowledge_measurement_points
            SET axis_label='repeated-axis',
                replicate_key='repeated-replicate'
            WHERE point_uid LIKE 'point_review_%'
            """
        )
        self.connection.commit()

        result = query.build_evidence_pack(
            self.connection,
            "spectral-sweep Frequency dBSPL repeated-replicate",
        )

        series = result["studyCandidates"][0]["measurementSeries"][0]
        self.assertEqual(
            ["repeated-replicate"],
            series["replicateKeys"],
        )
        self.assertEqual(2, series["pointSummary"]["distinctAxisCount"])
        self.assertEqual(
            2,
            series["pointSummary"]["distinctReplicateCount"],
        )

    def test_measurement_series_aggregate_replicate_is_not_raw_statistics(
        self,
    ) -> None:
        self._insert_measurement_series("review")
        self.connection.execute(
            """
            UPDATE knowledge_measurement_points
            SET replicate_role='AGGREGATE'
            WHERE replicate_source_coordinate='C1'
            """
        )
        self.connection.commit()

        result = query.build_evidence_pack(
            self.connection,
            "spectral-sweep Frequency dBSPL replicate-alpha",
        )
        series = result["studyCandidates"][0]["measurementSeries"][0]
        self.assertEqual(["replicate-alpha"], series["replicateKeys"])
        self.assertEqual(
            ["replicate-beta"],
            series["aggregateReplicateKeys"],
        )
        self.assertEqual(
            {
                "pointCount": 4,
                "rawPointCount": 2,
                "aggregatePointCount": 2,
                "minimum": 10.0,
                "maximum": 30.0,
                "average": 20.0,
                "distinctAxisCount": 2,
                "distinctReplicateCount": 1,
                "aggregateReplicateCount": 1,
            },
            series["pointSummary"],
        )

    def test_confounded_needs_review_and_invalid_candidates_are_excluded_with_reasons(self) -> None:
        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Ω cadence Nebula fracture rate",
        )
        excluded = {
            candidate["publicEffectId"]: {
                reason["code"]
                for reason in candidate["exclusionReasons"]
            }
            for candidate in result["excludedCandidates"]
        }

        self.assertEqual(3, len(excluded))
        self.assertIn(
            "COMPARISON_CONFOUNDED",
            excluded["EFF-NOVEL-CONFOUNDED"],
        )
        self.assertIn(
            "EFFECT_NOT_AGGREGATION_ELIGIBLE",
            excluded["EFF-NOVEL-CONFOUNDED"],
        )
        self.assertIn(
            "COMPARISON_NEEDS_REVIEW",
            excluded["EFF-NOVEL-REVIEW"],
        )
        self.assertIn("EFFECT_NEEDS_REVIEW", excluded["EFF-NOVEL-REVIEW"])
        self.assertIn(
            "COMPARISON_INVALID",
            excluded["EFF-NOVEL-INVALID"],
        )
        self.assertIn("EFFECT_INVALID", excluded["EFF-NOVEL-INVALID"])
        self.assertTrue(
            all(candidate["publicEvidenceIds"] for candidate in result["excludedCandidates"])
        )

    def test_effectless_comparison_preserves_comparison_gates_and_factor_difference(
        self,
    ) -> None:
        self.connection.execute(
            "DELETE FROM knowledge_effects WHERE effect_uid='effect_confounded'"
        )
        evidence_id = int(
            self.connection.execute(
                """
                SELECT evidence_id
                FROM evidence_items
                WHERE evidence_uid='evidence_confounded'
                """
            ).fetchone()[0]
        )
        self.connection.execute(
            """
            INSERT INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id, evidence_role,
                claim_scope
            ) VALUES (
                'COMPARISON', 'comparison_confounded', ?, 'RESULT',
                'confounded comparison'
            )
            """,
            (evidence_id,),
        )
        self.connection.commit()

        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Ω cadence Nebula fracture rate",
        )
        excluded = next(
            item
            for item in result["excludedCandidates"]
            if item["publicDataId"] == "DATA-NOVEL-CONFOUNDED"
        )

        self.assertIsNone(excluded["publicEffectId"])
        self.assertTrue(
            {
                "COMPARISON_CONFOUNDED",
                "COMPARISON_NOT_AGGREGATION_ELIGIBLE",
                "NO_EFFECT_RECORD",
            }.issubset(
                {
                reason["code"]
                for reason in excluded["exclusionReasons"]
                }
            )
        )
        comparison = excluded["comparison"]
        self.assertEqual(
            "comparison_confounded",
            comparison["comparisonUid"],
        )
        self.assertEqual(
            "same lot and sample basis",
            comparison["matchingBasis"],
        )
        self.assertEqual(
            "Multiple factors changed",
            comparison["exclusionReason"],
        )
        self.assertEqual(1, len(comparison["factorDifferences"]))
        difference = comparison["factorDifferences"][0]
        self.assertEqual("ZQ-17 plasma cadence", difference["factorLabel"])
        self.assertEqual("5 pulses", difference["comparedValue"])
        self.assertEqual("2 pulses", difference["controlValue"])
        self.assertEqual(
            ["EVD-NOVEL-CONFOUNDED"],
            excluded["publicEvidenceIds"],
        )
        self.assertEqual("STUDY", excluded["descriptiveScope"])
        self.assertEqual(
            "Nebula fracture rate",
            excluded["descriptiveOutcomes"][0]["outcome"]["originalLabel"],
        )
        self.assertEqual(
            {"Changed arm", "Baseline arm"},
            {
                item["arm"]["label"]
                for item in excluded["descriptiveOutcomes"][0][
                    "armObservations"
                ]
            },
        )

    def test_factor_metadata_flags_do_not_create_false_difference(
        self,
    ) -> None:
        control = {
            "factorValues": [
                {
                    "factorUid": "factor-same",
                    "factorLabel": "Same process",
                    "originalValue": "10%",
                    "valueNumber": 10,
                    "unit": "%",
                    "isBaseline": True,
                    "heldConstant": False,
                }
            ]
        }
        compared = {
            "factorValues": [
                {
                    "factorUid": "factor-same",
                    "factorLabel": "Same process",
                    "originalValue": "10%",
                    "valueNumber": 10,
                    "unit": "%",
                    "isBaseline": False,
                    "heldConstant": True,
                }
            ]
        }

        self.assertEqual(
            [],
            query._arm_factor_differences(compared, control),
        )

    def test_multi_factor_code_requires_two_real_value_differences(
        self,
    ) -> None:
        control = {
            "armId": 1,
            "factorValues": [
                {
                    "factorUid": "factor-a",
                    "factorLabel": "Factor A",
                    "originalValue": "A0",
                },
                {
                    "factorUid": "factor-b",
                    "factorLabel": "Factor B",
                    "originalValue": "B0",
                },
            ],
        }
        compared = {
            "armId": 2,
            "factorValues": [
                {
                    "factorUid": "factor-a",
                    "factorLabel": "Factor A",
                    "originalValue": "A1",
                },
                {
                    "factorUid": "factor-b",
                    "factorLabel": "Factor B",
                    "originalValue": "B1",
                },
            ],
        }
        candidate = {"arms": [control, compared]}
        comparison = {
            "control_arm_id": 1,
            "compared_arm_id": 2,
            "confounding_status": "CONFOUNDED",
        }

        self.assertTrue(
            query._is_multi_factor_confounding(candidate, comparison)
        )
        compared["factorValues"][1]["originalValue"] = "B0"
        self.assertFalse(
            query._is_multi_factor_confounding(candidate, comparison)
        )

    def test_selected_series_citations_do_not_reinclude_candidate_series(
        self,
    ) -> None:
        evidence_index = {
            ("MEASUREMENT_SERIES", "selected"): [
                {
                    "evidenceId": 1,
                    "publicEvidenceId": "EVD-SELECTED",
                }
            ],
            ("MEASUREMENT_SERIES", "unrelated"): [
                {
                    "evidenceId": 2,
                    "publicEvidenceId": "EVD-UNRELATED",
                }
            ],
        }

        citations = query._series_citations(
            [{"seriesUid": "selected"}],
            evidence_index,
        )

        self.assertEqual(
            ["EVD-SELECTED"],
            [item["publicEvidenceId"] for item in citations],
        )

    def test_raw_unknown_terms_are_searchable_without_concept_rule(self) -> None:
        result = query.build_evidence_pack(
            self.connection,
            "ZQ-17 plasma cadence AuroraRig-Ξ",
        )
        self.assertEqual(4, result["summary"]["relevantStudyCount"])
        self.assertTrue(
            all(
                {"zq-17", "plasma", "cadence"}.issubset(
                    set(candidate["relevance"]["matchedTerms"])
                )
                for candidate in result["studyCandidates"]
            )
        )
        agglutinative = query.build_evidence_pack(
            self.connection,
            "극저온플럭스가 성운파손에 미치는가",
        )
        self.assertEqual(4, agglutinative["summary"]["relevantStudyCount"])

    def test_one_unknown_concept_plus_question_words_still_retrieves(self) -> None:
        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux 알려줘",
        )
        self.assertEqual(4, result["summary"]["relevantStudyCount"])
        self.assertTrue(
            all(
                "cryoflux" in candidate["relevance"]["matchedTerms"]
                for candidate in result["studyCandidates"]
            )
        )

    def test_relationship_query_requires_context_or_factor_and_outcome_terms(self) -> None:
        outcome_concept_id = self.connection.execute(
            """
            SELECT concept_id
            FROM knowledge_concepts
            WHERE concept_uid='concept_nebula'
            """
        ).fetchone()[0]
        study_id = self.connection.execute(
            """
            INSERT INTO knowledge_studies(
                study_uid, public_data_id, workbook_analysis_id, study_key,
                title, summary_text, verification_status,
                comparability_status, confounding_status, created_at, updated_at
            ) VALUES (
                'study_outcome_only', 'DATA-OUTCOME-ONLY', 1, 'outcome-only',
                'Nebula fracture reference', 'Outcome reference only',
                'VERIFIED', 'VALID', 'NONE', ?, ?
            )
            """,
            (NOW, NOW),
        ).lastrowid
        self.connection.execute(
            """
            INSERT INTO knowledge_outcomes(
                outcome_uid, study_id, outcome_key, concept_id,
                original_label, outcome_domain, metric_type,
                favorable_direction, verification_status
            ) VALUES (
                'outcome_only_nebula', ?, 'nebula-only', ?,
                'Nebula fracture rate', 'Reference', 'RATE',
                'LOWER', 'VERIFIED'
            )
            """,
            (study_id, outcome_concept_id),
        )
        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Nebula fracture",
        )
        self.assertTrue(result["queryRoleHints"]["relationGateApplied"])
        self.assertEqual(
            ["cryoflux"],
            result["queryRoleHints"]["contextOrFactorTerms"],
        )
        self.assertNotIn(
            "DATA-OUTCOME-ONLY",
            {
                candidate["publicDataId"]
                for candidate in result["studyCandidates"]
            },
        )

    def test_study_title_can_proxy_broad_outcome_with_detailed_submetrics(
        self,
    ) -> None:
        self.connection.execute(
            """
            UPDATE knowledge_studies
            SET title='CryoFlux Nebula archived review'
            WHERE study_uid='study_good'
            """
        )
        self.connection.execute(
            """
            UPDATE knowledge_outcomes
            SET concept_id=NULL, outcome_key='hearing-profile',
                original_label='Hearing profile',
                outcome_domain='Acoustic detail', metric_type='PROFILE'
            WHERE outcome_uid='outcome_good'
            """
        )

        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Nebula",
        )

        self.assertTrue(result["queryRoleHints"]["relationGateApplied"])
        self.assertIn(
            "DATA-NOVEL-GOOD",
            {
                candidate["publicDataId"]
                for candidate in result["studyCandidates"]
            },
        )

    def test_stable_effect_and_evidence_ids_are_directly_queryable(self) -> None:
        for public_id in ("EFF-NOVEL-GOOD", "EVD-NOVEL-GOOD"):
            result = query.build_evidence_pack(self.connection, public_id)
            self.assertEqual(1, result["summary"]["relevantStudyCount"])
            candidate = result["studyCandidates"][0]
            self.assertEqual("DATA-NOVEL-GOOD", candidate["publicDataId"])
            self.assertEqual(
                [public_id],
                candidate["relevance"]["directIdentifierMatches"],
            )

    def test_stale_or_unverified_parent_cannot_be_answer_eligible(self) -> None:
        self.connection.execute(
            "UPDATE source_revisions SET is_current=0 WHERE revision_uid='revision_novel'"
        )
        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Nebula",
        )
        self.assertEqual(0, result["summary"]["answerEligibleEffectCount"])
        good = next(
            item
            for item in result["excludedCandidates"]
            if item["publicEffectId"] == "EFF-NOVEL-GOOD"
        )
        self.assertIn(
            "SOURCE_NOT_CURRENT",
            {reason["code"] for reason in good["exclusionReasons"]},
        )
        self.assertEqual("COMPARISON", good["descriptiveScope"])
        self.assertEqual(
            [15.0, 10.0],
            [
                item["observations"][0]["valueNumber"]
                for item in good["descriptiveOutcomes"][0][
                    "armObservations"
                ]
            ],
        )

    def test_superseded_analysis_is_not_in_default_retrieval(
        self,
    ) -> None:
        self.connection.execute(
            """
            UPDATE workbook_analyses
            SET verification_status='STALE',
                analysis_status='STALE'
            WHERE analysis_uid='analysis_novel'
            """
        )
        result = query.build_evidence_pack(
            self.connection,
            "DATA-NOVEL-GOOD",
        )
        self.assertEqual(0, result["summary"]["relevantStudyCount"])
        self.assertEqual([], result["studyCandidates"])
        self.assertEqual([], result["excludedCandidates"])

    def test_study_and_evidence_hash_gates_block_quantitative_answer(self) -> None:
        self.connection.execute(
            """
            UPDATE knowledge_studies
            SET comparability_status='UNASSESSED'
            WHERE study_uid='study_good'
            """
        )
        self.connection.execute(
            """
            UPDATE evidence_items
            SET content_sha256=?
            WHERE evidence_uid='evidence_good'
            """,
            ("b" * 64,),
        )
        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Nebula",
        )
        good = next(
            item
            for item in result["excludedCandidates"]
            if item["publicEffectId"] == "EFF-NOVEL-GOOD"
        )
        reasons = {reason["code"] for reason in good["exclusionReasons"]}
        self.assertIn("STUDY_COMPARABILITY_UNASSESSED", reasons)
        self.assertIn("EFFECT_EVIDENCE_CONTENT_MISMATCH", reasons)

    def test_missing_effect_estimate_is_never_answer_eligible(self) -> None:
        self.connection.execute(
            """
            UPDATE knowledge_effects
            SET estimate=NULL
            WHERE effect_uid='effect_good'
            """
        )
        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux Nebula",
        )
        self.assertEqual(0, result["summary"]["answerEligibleEffectCount"])
        good = next(
            item
            for item in result["excludedCandidates"]
            if item["publicEffectId"] == "EFF-NOVEL-GOOD"
        )
        self.assertIn(
            "EFFECT_ESTIMATE_MISSING",
            {reason["code"] for reason in good["exclusionReasons"]},
        )

    def test_terminal_source_is_retrieved_as_source_level_exclusion(self) -> None:
        result = query.build_evidence_pack(
            self.connection,
            "Zephyr XRAY image",
        )

        self.assertEqual(1, result["summary"]["relevantSourceExclusionCount"])
        terminal = result["sourceExclusions"][0]
        self.assertEqual("ANALYSIS-TERMINAL", terminal["publicAnalysisId"])
        self.assertEqual(
            "NO_TABULAR_EVIDENCE",
            terminal["sourceContentStatus"],
        )
        self.assertFalse(terminal["imagesAnalyzed"])
        self.assertEqual(
            {"IMAGES_NOT_ANALYZED", "NO_TABULAR_EVIDENCE"},
            {
                reason["code"]
                for reason in terminal["exclusionReasons"]
            },
        )

    def test_terminal_source_allows_one_distinctive_exact_filename_term(
        self,
    ) -> None:
        result = query.build_evidence_pack(
            self.connection,
            "Zephyr paired result",
        )

        self.assertEqual(1, result["summary"]["relevantSourceExclusionCount"])
        self.assertEqual(
            "ANALYSIS-TERMINAL",
            result["sourceExclusions"][0]["publicAnalysisId"],
        )
        self.assertEqual(
            ["zephyr"],
            result["sourceExclusions"][0]["relevance"]["matchedTerms"],
        )

    def test_relationship_query_excludes_terminal_source_matching_only_context(
        self,
    ) -> None:
        self.connection.execute(
            """
            UPDATE source_documents
            SET original_file_name='CryoFlux Zephyr XRAY image review.xlsx'
            WHERE document_uid='document_terminal'
            """
        )
        self.connection.execute(
            """
            UPDATE workbook_analyses
            SET title='CryoFlux Zephyr XRAY image review'
            WHERE analysis_uid='analysis_terminal'
            """
        )
        self.connection.commit()

        result = query.build_evidence_pack(
            self.connection,
            "CryoFlux cadence Nebula fracture rate",
        )

        self.assertTrue(result["queryRoleHints"]["relationGateApplied"])
        self.assertEqual(0, result["summary"]["relevantSourceExclusionCount"])

    def test_database_path_entry_point_enforces_read_only_access(self) -> None:
        result = query.build_evidence_pack_from_db(
            self.database,
            "극저온플럭스 성운파손",
        )
        self.assertEqual(1, result["summary"]["answerEligibleEffectCount"])
        with query.connect_knowledge_readonly(self.database) as connection:
            with self.assertRaises(sqlite3.OperationalError):
                connection.execute(
                    "UPDATE knowledge_studies SET title='changed'"
                )


if __name__ == "__main__":
    unittest.main()
