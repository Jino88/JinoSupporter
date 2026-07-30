from __future__ import annotations

import sqlite3
import unittest

import inference_data_ai_review as review
from inference_data_ai_schema import (
    ensure_knowledge_schema,
    public_id,
    stable_uid,
)


NOW = "2026-07-17T00:00:00Z"
SOURCE_SHA = "a" * 64


class CanonicalHumanReviewTests(unittest.TestCase):
    def setUp(self) -> None:
        self.connection = sqlite3.connect(":memory:")
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys=ON")
        self.connection.executescript(
            """
            CREATE TABLE workbooks(workbook_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_reports(analysis_report_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_review_items(review_item_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_cohorts(cohort_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_metrics(metric_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_metric_values(metric_value_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_comparisons(comparison_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_conclusions(conclusion_id INTEGER PRIMARY KEY);
            CREATE TABLE analysis_evidence(evidence_id INTEGER PRIMARY KEY);
            """
        )
        ensure_knowledge_schema(self.connection, lambda: NOW)
        self._insert_fixture()
        self.connection.commit()

    def tearDown(self) -> None:
        self.connection.close()

    def _insert_fixture(self) -> None:
        document_uid = stable_uid("document", "fixture", "generic.xlsx")
        revision_uid = stable_uid("revision", document_uid, SOURCE_SHA)
        analysis_uid = stable_uid("analysis", revision_uid, "generic")
        study_uid = stable_uid("study", analysis_uid, "generic")
        compared_uid = stable_uid("arm", study_uid, "changed")
        control_uid = stable_uid("arm", study_uid, "control")
        comparison_uid = stable_uid("comparison", study_uid, "changed-control")
        self.comparison_uid = comparison_uid
        self.public_comparison_id = public_id("CMP", comparison_uid)

        document_id = self.connection.execute(
            """
            INSERT INTO source_documents(
                document_uid, dataset, source_path, original_file_name,
                source_kind, lifecycle_status, created_at, updated_at
            ) VALUES (?, 'Fixture', 'generic.xlsx', 'generic.xlsx',
                      'XLSX', 'ACTIVE', ?, ?)
            """,
            (document_uid, NOW, NOW),
        ).lastrowid
        revision_id = self.connection.execute(
            """
            INSERT INTO source_revisions(
                revision_uid, document_id, source_fingerprint,
                fingerprint_kind, content_sha256, size_bytes, mtime_ns,
                extractor_name, extractor_version, capture_contract,
                capture_status, is_current, captured_at,
                source_content_status
            ) VALUES (?, ?, ?, 'SHA256', ?, 10, 1, 'fixture', '1',
                      'capture-v2', 'CAPTURED', 1, ?, 'CAPTURED')
            """,
            (revision_uid, document_id, SOURCE_SHA, SOURCE_SHA, NOW),
        ).lastrowid
        analysis_id = self.connection.execute(
            """
            INSERT INTO workbook_analyses(
                analysis_uid, public_analysis_id, document_id, revision_id,
                analysis_key, title, analysis_status, verification_status,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, 'generic', 'Generic review', 'DRAFT',
                      'NEEDS_REVIEW', ?, ?)
            """,
            (
                analysis_uid,
                public_id("ANALYSIS", analysis_uid),
                document_id,
                revision_id,
                NOW,
                NOW,
            ),
        ).lastrowid
        study_id = self.connection.execute(
            """
            INSERT INTO knowledge_studies(
                study_uid, public_data_id, workbook_analysis_id, study_key,
                title, design_type, comparison_basis, analysis_status,
                verification_status, comparability_status,
                confounding_status, created_at, updated_at
            ) VALUES (?, ?, ?, 'generic', 'Generic study', 'CONTROL_TEST',
                      'same unit and period', 'DRAFT', 'NEEDS_REVIEW',
                      'VALID', 'NONE', ?, ?)
            """,
            (
                study_uid,
                public_id("DATA", study_uid),
                analysis_id,
                NOW,
                NOW,
            ),
        ).lastrowid
        compared_id = self.connection.execute(
            """
            INSERT INTO knowledge_arms(
                arm_uid, study_id, arm_key, arm_role, label,
                matching_basis, verification_status
            ) VALUES (?, ?, 'changed', 'TEST', 'Changed',
                      'same unit and period', 'NEEDS_REVIEW')
            """,
            (compared_uid, study_id),
        ).lastrowid
        control_id = self.connection.execute(
            """
            INSERT INTO knowledge_arms(
                arm_uid, study_id, arm_key, arm_role, label,
                matching_basis, verification_status
            ) VALUES (?, ?, 'control', 'CONTROL', 'Control',
                      'same unit and period', 'NEEDS_REVIEW')
            """,
            (control_uid, study_id),
        ).lastrowid
        self.connection.execute(
            """
            INSERT INTO knowledge_comparisons(
                comparison_uid, public_comparison_id, study_id,
                comparison_key, compared_arm_id, control_arm_id,
                design_type, matching_basis, validity_status,
                confounding_status, aggregation_eligible,
                verification_status
            ) VALUES (?, ?, ?, 'changed-control', ?, ?, 'CONTROL_TEST',
                      'same unit, period and measurement method', 'VALID',
                      'NONE', 0, 'NEEDS_REVIEW')
            """,
            (
                comparison_uid,
                self.public_comparison_id,
                study_id,
                compared_id,
                control_id,
            ),
        )
        self._add_evidence(
            revision_id,
            "COMPARISON",
            comparison_uid,
            "A1:D5",
        )
        self._add_outcome(
            revision_id,
            study_id,
            compared_id,
            control_id,
            study_uid,
            "metric-one",
            changed=12,
            control=10,
            row=2,
        )
        self._add_outcome(
            revision_id,
            study_id,
            compared_id,
            control_id,
            study_uid,
            "metric-two",
            changed=30,
            control=20,
            row=4,
        )

    def _add_outcome(
        self,
        revision_id: int,
        study_id: int,
        compared_id: int,
        control_id: int,
        study_uid: str,
        key: str,
        *,
        changed: float,
        control: float,
        row: int,
    ) -> None:
        outcome_uid = stable_uid("outcome", study_uid, key)
        outcome_id = self.connection.execute(
            """
            INSERT INTO knowledge_outcomes(
                outcome_uid, study_id, outcome_key, original_label,
                metric_type, original_unit, favorable_direction,
                verification_status
            ) VALUES (?, ?, ?, ?, 'continuous', 'mg', 'LOWER',
                      'NEEDS_REVIEW')
            """,
            (outcome_uid, study_id, key, key),
        ).lastrowid
        for arm_id, arm_key, value, evidence_row in (
            (compared_id, "changed", changed, row),
            (control_id, "control", control, row + 1),
        ):
            observation_uid = stable_uid(
                "observation",
                outcome_uid,
                arm_key,
                "primary",
            )
            self.connection.execute(
                """
                INSERT INTO knowledge_observations(
                    observation_uid, outcome_id, arm_id, observation_key,
                    value_number, verification_status
                ) VALUES (?, ?, ?, 'primary', ?, 'NEEDS_REVIEW')
                """,
                (observation_uid, outcome_id, arm_id, value),
            )
            self._add_evidence(
                revision_id,
                "OBSERVATION",
                observation_uid,
                f"A{evidence_row}:D{evidence_row}",
            )

    def _add_evidence(
        self,
        revision_id: int,
        entity_type: str,
        entity_uid: str,
        address: str,
    ) -> None:
        evidence_uid = stable_uid("evidence", entity_type, entity_uid, address)
        evidence_id = self.connection.execute(
            """
            INSERT INTO evidence_items(
                evidence_uid, public_evidence_id, revision_id,
                evidence_kind, sheet_name, start_row, start_col,
                end_row, end_col, range_address, evidence_role,
                verification_status, created_at
            ) VALUES (?, ?, ?, 'CELL_RANGE', 'Data', 1, 1, 5, 4, ?,
                      'SOURCE', 'VERIFIED', ?)
            """,
            (
                evidence_uid,
                public_id("EVD", evidence_uid),
                revision_id,
                address,
                NOW,
            ),
        ).lastrowid
        self.connection.execute(
            """
            INSERT INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id, evidence_role
            ) VALUES (?, ?, ?, 'SOURCE')
            """,
            (entity_type, entity_uid, evidence_id),
        )

    def _approve(self) -> dict:
        return review.decide_comparison(
            self.connection,
            self.public_comparison_id,
            decision="APPROVE",
            reviewer="reviewer-1",
            reason="Direct source rows and matching basis were checked.",
            decided_at=NOW,
        )

    def test_queue_detail_and_approval_create_verified_effects(self) -> None:
        queue = review.list_review_queue(self.connection)
        detail = review.get_review_detail(
            self.connection,
            self.public_comparison_id,
        )
        result = self._approve()

        self.assertEqual(1, queue["count"])
        self.assertFalse(queue["imagesAnalyzed"])
        self.assertTrue(detail["approvalReadiness"]["ready"])
        self.assertEqual(4, detail["approvalReadiness"]["plannedEffectCount"])
        self.assertEqual("APPROVE", result["decision"])
        self.assertTrue(result["comparisonAggregationEligible"])
        self.assertEqual(4, len(result["effectPublicIds"]))
        self.assertIn("valueNumber", detail["pairedObservations"][0][
            "comparedObservation"
        ])
        self.assertIn("evidence", detail["comparison"])
        rows = self.connection.execute(
            """
            SELECT e.public_effect_id, e.verification_status,
                   e.aggregation_eligible, o.outcome_uid, e.effect_uid
            FROM knowledge_effects AS e
            JOIN knowledge_outcomes AS o ON o.outcome_id=e.outcome_id
            ORDER BY e.public_effect_id
            """
        ).fetchall()
        self.assertEqual(4, len(rows))
        for row in rows:
            self.assertEqual("VERIFIED", row["verification_status"])
            self.assertEqual(1, row["aggregation_eligible"])
            self.assertEqual(
                stable_uid(
                    "effect",
                    self.comparison_uid,
                    row["outcome_uid"],
                    self.connection.execute(
                        "SELECT effect_type FROM knowledge_effects WHERE effect_uid=?",
                        (row["effect_uid"],),
                    ).fetchone()[0],
                    self.connection.execute(
                        "SELECT formula_version FROM knowledge_effects WHERE effect_uid=?",
                        (row["effect_uid"],),
                    ).fetchone()[0],
                ),
                row["effect_uid"],
            )
            direct_links = self.connection.execute(
                """
                SELECT COUNT(*)
                FROM entity_evidence_links
                WHERE entity_type='EFFECT' AND entity_uid=?
                  AND evidence_role='REVIEW_APPROVAL'
                """,
                (row["effect_uid"],),
            ).fetchone()[0]
            self.assertEqual(3, direct_links)

    def test_raw_count_outcomes_block_mean_difference_approval(self) -> None:
        self.connection.execute(
            "UPDATE knowledge_outcomes SET metric_type='defect_count'"
        )
        detail = review.get_review_detail(
            self.connection,
            self.public_comparison_id,
        )

        self.assertFalse(detail["approvalReadiness"]["ready"])
        self.assertEqual(0, detail["approvalReadiness"]["plannedEffectCount"])
        self.assertTrue(
            any(
                blocker["code"] == "EFFECT_CALCULATION_FAILED"
                and "raw counts cannot be treated as continuous effects"
                in blocker["message"]
                for blocker in detail["approvalReadiness"]["blockers"]
            )
        )

    def test_sample_size_outcome_is_denominator_only(self) -> None:
        first_outcome_id = self.connection.execute(
            "SELECT MIN(outcome_id) FROM knowledge_outcomes"
        ).fetchone()[0]
        self.connection.execute(
            """
            UPDATE knowledge_outcomes
            SET metric_type='sample_size'
            WHERE outcome_id=?
            """,
            (first_outcome_id,),
        )

        detail = review.get_review_detail(
            self.connection,
            self.public_comparison_id,
        )
        result = self._approve()

        self.assertTrue(detail["approvalReadiness"]["ready"])
        self.assertEqual(2, detail["approvalReadiness"]["plannedEffectCount"])
        self.assertEqual(2, len(result["effectPublicIds"]))

    def test_explicit_rate_outcome_suppresses_duplicate_count_effects(
        self,
    ) -> None:
        context = self.connection.execute(
            """
            SELECT s.study_id, s.study_uid, c.compared_arm_id,
                   c.control_arm_id, r.revision_id
            FROM knowledge_comparisons c
            JOIN knowledge_studies s ON s.study_id=c.study_id
            JOIN workbook_analyses wa
              ON wa.workbook_analysis_id=s.workbook_analysis_id
            JOIN source_revisions r ON r.revision_id=wa.revision_id
            """
        ).fetchone()
        self.connection.execute(
            """
            UPDATE knowledge_outcomes
            SET metric_type='defect_count'
            WHERE outcome_key='metric-one'
            """
        )
        self.connection.execute(
            """
            UPDATE knowledge_observations
            SET numerator=value_number, denominator=100
            WHERE outcome_id=(
                SELECT outcome_id
                FROM knowledge_outcomes
                WHERE outcome_key='metric-one'
            )
            """
        )
        self._add_outcome(
            int(context["revision_id"]),
            int(context["study_id"]),
            int(context["compared_arm_id"]),
            int(context["control_arm_id"]),
            str(context["study_uid"]),
            "metric-one-rate",
            changed=12,
            control=10,
            row=6,
        )
        self.connection.execute(
            """
            UPDATE knowledge_outcomes
            SET metric_type='defect_rate', original_unit='%'
            WHERE outcome_key='metric-one-rate'
            """
        )
        self.connection.execute(
            """
            UPDATE knowledge_observations
            SET numerator=value_number, denominator=100
            WHERE outcome_id=(
                SELECT outcome_id
                FROM knowledge_outcomes
                WHERE outcome_key='metric-one-rate'
            )
            """
        )

        detail = review.get_review_detail(
            self.connection,
            self.public_comparison_id,
        )
        result = self._approve()

        self.assertTrue(detail["approvalReadiness"]["ready"])
        self.assertEqual(
            7,
            detail["approvalReadiness"]["plannedEffectCount"],
        )
        planned_outcomes = {
            row[0]
            for row in self.connection.execute(
                """
                SELECT DISTINCT o.outcome_key
                FROM knowledge_effects e
                JOIN knowledge_outcomes o ON o.outcome_id=e.outcome_id
                """
            )
        }
        self.assertEqual(7, len(result["effectPublicIds"]))
        self.assertNotIn("metric-one", planned_outcomes)
        self.assertIn("metric-one-rate", planned_outcomes)

    def test_explicit_human_assessment_can_unlock_approval_atomically(
        self,
    ) -> None:
        self.connection.execute(
            """
            UPDATE knowledge_studies
            SET comparability_status='UNASSESSED',
                confounding_status='UNASSESSED'
            """
        )
        self.connection.execute(
            """
            UPDATE knowledge_comparisons
            SET validity_status='NEEDS_REVIEW',
                confounding_status='UNASSESSED',
                matching_basis=''
            """
        )
        self.connection.commit()

        result = review.decide_comparison(
            self.connection,
            self.public_comparison_id,
            decision="APPROVE",
            reviewer="reviewer-2",
            reason="I checked the current source rows and pairing.",
            decided_at=NOW,
            study_comparability_status="VALID",
            study_confounding_status="NONE",
            comparison_validity_status="VALID",
            comparison_confounding_status="NONE",
            matching_basis="same unit, period and measurement method",
        )

        self.assertTrue(result["comparisonAggregationEligible"])
        self.assertEqual(
            {
                "studyComparabilityStatus": "VALID",
                "studyConfoundingStatus": "NONE",
                "comparisonValidityStatus": "VALID",
                "comparisonConfoundingStatus": "NONE",
                "matchingBasis": "same unit, period and measurement method",
            },
            result["assessment"],
        )

    def test_missing_observation_evidence_fails_and_rolls_back(self) -> None:
        observation_uid = self.connection.execute(
            """
            SELECT observation_uid
            FROM knowledge_observations
            ORDER BY observation_id DESC
            LIMIT 1
            """
        ).fetchone()[0]
        self.connection.execute(
            """
            DELETE FROM entity_evidence_links
            WHERE entity_type='OBSERVATION' AND entity_uid=?
            """,
            (observation_uid,),
        )
        self.connection.commit()

        with self.assertRaisesRegex(
            review.ReviewGateError,
            "OBSERVATION_EVIDENCE_REQUIRED",
        ):
            self._approve()

        comparison = self.connection.execute(
            """
            SELECT verification_status, aggregation_eligible
            FROM knowledge_comparisons
            """
        ).fetchone()
        self.assertEqual("NEEDS_REVIEW", comparison["verification_status"])
        self.assertEqual(0, comparison["aggregation_eligible"])
        self.assertEqual(
            0,
            self.connection.execute(
                "SELECT COUNT(*) FROM knowledge_effects"
            ).fetchone()[0],
        )
        self.assertEqual(
            0,
            self.connection.execute(
                "SELECT COUNT(*) FROM review_decisions"
            ).fetchone()[0],
        )

    def test_stale_source_is_rejected(self) -> None:
        self.connection.execute(
            """
            UPDATE source_revisions
            SET is_current=0, capture_status='STALE'
            """
        )
        self.connection.commit()
        with self.assertRaisesRegex(
            review.ReviewGateError,
            "SOURCE_NOT_CURRENT",
        ):
            self._approve()
        self.assertEqual(
            0,
            self.connection.execute(
                "SELECT COUNT(*) FROM review_decisions"
            ).fetchone()[0],
        )

    def test_superseded_analysis_is_hidden_and_cannot_be_approved(
        self,
    ) -> None:
        self.connection.execute(
            """
            UPDATE workbook_analyses
            SET verification_status='STALE',
                analysis_status='STALE'
            """
        )
        self.connection.commit()
        queue = review.list_review_queue(self.connection)
        self.assertEqual(0, queue["count"])
        detail = review.get_review_detail(
            self.connection,
            self.public_comparison_id,
        )
        self.assertIn(
            "ANALYSIS_SUPERSEDED",
            {
                blocker["code"]
                for blocker in detail["approvalReadiness"]["blockers"]
            },
        )
        with self.assertRaisesRegex(
            review.ReviewGateError,
            "ANALYSIS_SUPERSEDED",
        ):
            self._approve()

    def test_reject_disables_an_already_approved_comparison_and_effects(self) -> None:
        self._approve()
        result = review.decide_comparison(
            self.connection,
            self.public_comparison_id,
            decision="REJECT",
            reviewer="reviewer-2",
            reason="The control selection was not acceptable.",
            decided_at="2026-07-17T01:00:00Z",
        )
        self.assertEqual("INVALID", result["comparisonValidityStatus"])
        self.assertFalse(result["comparisonAggregationEligible"])
        effects = self.connection.execute(
            """
            SELECT DISTINCT verification_status, aggregation_eligible
            FROM knowledge_effects
            """
        ).fetchall()
        self.assertEqual([("INVALID", 0)], [tuple(row) for row in effects])
        decisions = self.connection.execute(
            """
            SELECT decision, supersedes_decision_uid
            FROM review_decisions
            ORDER BY review_decision_id
            """
        ).fetchall()
        self.assertEqual(["APPROVE", "REJECT"], [row[0] for row in decisions])
        self.assertTrue(decisions[1][1])

    def test_repeated_identical_approval_is_idempotent(self) -> None:
        first = self._approve()
        first_effects = list(first["effectPublicIds"])
        second = self._approve()
        self.assertTrue(second["idempotent"])
        self.assertEqual(first["decisionUid"], second["decisionUid"])
        self.assertEqual(first_effects, second["effectPublicIds"])
        self.assertEqual(
            1,
            self.connection.execute(
                "SELECT COUNT(*) FROM review_decisions"
            ).fetchone()[0],
        )
        self.assertEqual(
            4,
            self.connection.execute(
                "SELECT COUNT(*) FROM knowledge_effects"
            ).fetchone()[0],
        )

    def test_multi_outcome_effect_rows_do_not_collide(self) -> None:
        self._approve()
        rows = self.connection.execute(
            """
            SELECT outcome_id, effect_type, formula_version, COUNT(*)
            FROM knowledge_effects
            GROUP BY outcome_id, effect_type, formula_version
            ORDER BY outcome_id, effect_type
            """
        ).fetchall()
        self.assertEqual(4, len(rows))
        self.assertTrue(all(row[3] == 1 for row in rows))
        self.assertEqual(
            2,
            self.connection.execute(
                "SELECT COUNT(DISTINCT outcome_id) FROM knowledge_effects"
            ).fetchone()[0],
        )


if __name__ == "__main__":
    unittest.main()
