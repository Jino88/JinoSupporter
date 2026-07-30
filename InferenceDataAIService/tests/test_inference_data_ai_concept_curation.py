from __future__ import annotations

import sqlite3
import unittest

import inference_data_ai_concept_curation as curation
import inference_data_ai_schema as schema


NOW = "2026-07-18T12:00:00Z"


class ConceptCurationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.connection = sqlite3.connect(":memory:")
        self.connection.row_factory = sqlite3.Row
        schema.ensure_knowledge_schema(
            self.connection,
            lambda: NOW,
        )

    def tearDown(self) -> None:
        self.connection.close()

    def candidate(
        self,
        kind: str,
        original: str,
        suggested: str = "",
    ) -> str:
        return schema.record_schema_candidate(
            self.connection,
            candidate_kind=kind,
            original_value=original,
            suggested_canonical_name=suggested,
            now_iso=lambda: NOW,
        )

    def concept_uid(self, kind: str, canonical_name: str) -> str:
        row = self.connection.execute(
            """
            SELECT concept_uid
            FROM knowledge_concepts
            WHERE concept_kind=? AND normalized_name=?
            """,
            (kind, schema.normalized_term(canonical_name)),
        ).fetchone()
        self.assertIsNotNone(row)
        return str(row[0])

    def test_migration_and_filtered_json_lists_are_stable(self) -> None:
        concept_uid = self.candidate(
            "CONCEPT:ARBITRARY_FACTOR",
            "Zephyr knob",
            "Zephyr control",
        )
        self.candidate("UNIT", "widgets")
        schema.ensure_knowledge_schema(
            self.connection,
            lambda: "2026-07-18T12:00:01Z",
        )
        self.assertEqual(
            1,
            self.connection.execute(
                """
                SELECT COUNT(*) FROM schema_migrations
                WHERE migration_name=?
                """,
                (curation.CONCEPT_CURATION_MIGRATION,),
            ).fetchone()[0],
        )

        candidates = curation.list_schema_candidates(
            self.connection,
            status="OPEN",
            candidate_kind="concept:arbitrary_factor",
            query="zephyr",
            limit=10,
        )
        self.assertEqual(
            curation.CANDIDATE_LIST_SCHEMA_VERSION,
            candidates["schemaVersion"],
        )
        self.assertEqual(1, candidates["count"])
        self.assertEqual(
            concept_uid,
            candidates["candidates"][0]["candidateUid"],
        )

        concepts = curation.list_canonical_concepts(
            self.connection,
            concept_kind="changed_factor",
            lifecycle_status="ACTIVE",
            query="bond amount",
            limit=10,
        )
        self.assertGreaterEqual(concepts["count"], 1)
        bonding = next(
            value
            for value in concepts["concepts"]
            if value["canonicalName"] == "Bonding amount"
        )
        self.assertTrue(
            any(
                alias["aliasText"] == "bond amount"
                for alias in bonding["aliases"]
            )
        )

    def test_create_is_atomic_immutable_and_exactly_idempotent(
        self,
    ) -> None:
        candidate_uid = self.candidate(
            "CONCEPT:ARBITRARY_FACTOR",
            "Zephyr knob",
            "Zephyr control",
        )
        request = {
            "candidate_uid": candidate_uid,
            "action": "CREATE",
            "canonical_name": "Zephyr control",
            "alias": "Zephyr knob",
            "reviewer": "human-1",
            "note": "Checked the source terminology.",
        }
        created = curation.resolve_schema_candidate(
            self.connection,
            **request,
            now_iso=lambda: NOW,
        )
        self.assertEqual("CREATE", created["action"])
        self.assertEqual("APPROVED", created["candidate"]["status"])
        self.assertEqual(
            "ARBITRARY_FACTOR",
            created["concept"]["conceptKind"],
        )
        self.assertEqual("HUMAN_APPROVED", created["alias"]["source"])
        self.assertFalse(created["idempotentReplay"])

        replay = curation.resolve_schema_candidate(
            self.connection,
            **request,
            now_iso=lambda: "2099-01-01T00:00:00Z",
        )
        self.assertTrue(replay["idempotentReplay"])
        self.assertEqual(created["resolutionUid"], replay["resolutionUid"])
        self.assertEqual(created["resolvedAt"], replay["resolvedAt"])
        self.assertEqual(
            1,
            self.connection.execute(
                "SELECT COUNT(*) FROM knowledge_concept_resolution_history"
            ).fetchone()[0],
        )
        self.assertEqual(
            1,
            self.connection.execute(
                """
                SELECT COUNT(*)
                FROM knowledge_concept_aliases
                WHERE concept_id=? AND normalized_alias=?
                """,
                (
                    created["concept"]["conceptId"],
                    schema.normalized_term("Zephyr knob"),
                ),
            ).fetchone()[0],
        )

        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "conflicting repeated resolution",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                **{**request, "alias": "Different alias"},
                now_iso=lambda: NOW,
            )
        with self.assertRaisesRegex(sqlite3.IntegrityError, "immutable"):
            self.connection.execute(
                """
                UPDATE knowledge_concept_resolution_history
                SET note='changed'
                """
            )
        with self.assertRaisesRegex(sqlite3.IntegrityError, "immutable"):
            self.connection.execute(
                "DELETE FROM knowledge_concept_resolution_history"
            )

    def test_merge_rejects_units_kind_mismatch_alias_collision_and_empty(
        self,
    ) -> None:
        bonding_uid = self.concept_uid(
            "CHANGED_FACTOR",
            "Bonding amount",
        )
        press_uid = self.concept_uid(
            "CHANGED_FACTOR",
            "Press amount",
        )
        merge_uid = self.candidate(
            "CONCEPT:CHANGED_FACTOR",
            "Glue dose",
        )
        merged = curation.resolve_schema_candidate(
            self.connection,
            candidate_uid=merge_uid,
            action="MERGE",
            concept_uid=bonding_uid,
            alias="Glue dose",
            reviewer="human-1",
            note="Same source meaning.",
            now_iso=lambda: NOW,
        )
        self.assertEqual("MERGE", merged["action"])
        self.assertEqual("MERGED", merged["candidate"]["status"])
        self.assertEqual(bonding_uid, merged["concept"]["conceptUid"])

        mismatch_uid = self.candidate(
            "CONCEPT:OUTCOME",
            "Foreign outcome",
        )
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "kind mismatch",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                candidate_uid=mismatch_uid,
                action="MERGE",
                concept_uid=bonding_uid,
                alias="Foreign outcome",
                reviewer="human-1",
                note="Invalid fixture.",
                now_iso=lambda: NOW,
            )

        collision_uid = self.candidate(
            "CONCEPT:CHANGED_FACTOR",
            "Another pressure term",
        )
        concept_count = self.connection.execute(
            "SELECT COUNT(*) FROM knowledge_concepts"
        ).fetchone()[0]
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "already owned",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                candidate_uid=collision_uid,
                action="MERGE",
                concept_uid=bonding_uid,
                alias="Press amount",
                reviewer="human-1",
                note="Collision fixture.",
                now_iso=lambda: NOW,
            )
        self.assertEqual(
            "OPEN",
            self.connection.execute(
                """
                SELECT status FROM knowledge_schema_candidates
                WHERE candidate_uid=?
                """,
                (collision_uid,),
            ).fetchone()[0],
        )
        self.assertEqual(
            concept_count,
            self.connection.execute(
                "SELECT COUNT(*) FROM knowledge_concepts"
            ).fetchone()[0],
        )
        self.assertEqual(
            press_uid,
            self.concept_uid("CHANGED_FACTOR", "Press amount"),
        )

        unit_uid = self.candidate("UNIT", "widgets")
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "unit curation path",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                candidate_uid=unit_uid,
                action="CREATE",
                canonical_name="Widgets",
                alias="widgets",
                reviewer="human-1",
                note="Wrong path.",
                now_iso=lambda: NOW,
            )

        empty_uid = self.candidate(
            "CONCEPT:ARBITRARY_FACTOR",
            "Empty checks",
        )
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "canonical_name must not be empty",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                candidate_uid=empty_uid,
                action="CREATE",
                canonical_name="",
                alias="alias",
                reviewer="human-1",
                note="Invalid fixture.",
                now_iso=lambda: NOW,
            )
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "alias must not be empty",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                candidate_uid=empty_uid,
                action="CREATE",
                canonical_name="Valid canonical",
                alias="",
                reviewer="human-1",
                note="Invalid fixture.",
                now_iso=lambda: NOW,
            )

    def test_reject_is_bounded_and_non_open_without_history_fails(
        self,
    ) -> None:
        candidate_uid = self.candidate(
            "CONCEPT:ARBITRARY_CONTEXT",
            "Not a reusable concept",
        )
        request = {
            "candidate_uid": candidate_uid,
            "action": "REJECT",
            "reviewer": "human-2",
            "note": "Workbook-local noise.",
        }
        rejected = curation.resolve_schema_candidate(
            self.connection,
            **request,
            now_iso=lambda: NOW,
        )
        self.assertEqual("REJECT", rejected["action"])
        self.assertEqual("REJECTED", rejected["candidate"]["status"])
        self.assertIsNone(rejected["concept"])
        self.assertIsNone(rejected["alias"])
        replay = curation.resolve_schema_candidate(
            self.connection,
            **request,
            now_iso=lambda: "2099-01-01T00:00:00Z",
        )
        self.assertTrue(replay["idempotentReplay"])
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "must not specify",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                **request,
                alias="unexpected",
                now_iso=lambda: NOW,
            )

        orphan_uid = self.candidate(
            "CONCEPT:ARBITRARY_CONTEXT",
            "Externally closed",
        )
        self.connection.execute(
            """
            UPDATE knowledge_schema_candidates
            SET status='APPROVED'
            WHERE candidate_uid=?
            """,
            (orphan_uid,),
        )
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "status must be OPEN",
        ):
            curation.resolve_schema_candidate(
                self.connection,
                candidate_uid=orphan_uid,
                action="REJECT",
                reviewer="human-2",
                note="Cannot adopt external status.",
                now_iso=lambda: NOW,
            )

    def test_standalone_human_alias_is_stable_audited_and_owned_once(
        self,
    ) -> None:
        bonding_uid = self.concept_uid(
            "CHANGED_FACTOR",
            "Bonding amount",
        )
        request = {
            "concept_uid": bonding_uid,
            "alias": "Adhesive dispense",
            "reviewer": "human-3",
            "note": "Approved terminology.",
        }
        first = curation.upsert_human_concept_alias(
            self.connection,
            **request,
            now_iso=lambda: NOW,
        )
        replay = curation.upsert_human_concept_alias(
            self.connection,
            **request,
            now_iso=lambda: "2099-01-01T00:00:00Z",
        )
        self.assertEqual(first["approvalUid"], replay["approvalUid"])
        self.assertEqual(
            first["alias"]["aliasUid"],
            replay["alias"]["aliasUid"],
        )
        self.assertTrue(replay["idempotentReplay"])

        second = curation.upsert_human_concept_alias(
            self.connection,
            concept_uid=bonding_uid,
            alias="Adhesive dispense",
            reviewer="human-4",
            note="Independent re-check.",
            now_iso=lambda: "2026-07-18T13:00:00Z",
        )
        self.assertEqual(
            first["alias"]["aliasUid"],
            second["alias"]["aliasUid"],
        )
        self.assertEqual(
            1,
            self.connection.execute(
                """
                SELECT COUNT(*) FROM knowledge_concept_aliases
                WHERE normalized_alias=?
                """,
                (schema.normalized_term("Adhesive dispense"),),
            ).fetchone()[0],
        )
        self.assertEqual(
            2,
            self.connection.execute(
                """
                SELECT COUNT(*)
                FROM knowledge_concept_alias_approval_history
                WHERE normalized_alias=?
                """,
                (schema.normalized_term("Adhesive dispense"),),
            ).fetchone()[0],
        )

        press_uid = self.concept_uid(
            "CHANGED_FACTOR",
            "Press amount",
        )
        with self.assertRaisesRegex(
            curation.ConceptCurationError,
            "already owned",
        ):
            curation.upsert_human_concept_alias(
                self.connection,
                concept_uid=bonding_uid,
                alias="Press amount",
                reviewer="human-3",
                note="Collision fixture.",
                now_iso=lambda: NOW,
            )
        self.assertNotEqual(bonding_uid, press_uid)

        seeded_before = self.connection.execute(
            """
            SELECT alias_uid
            FROM knowledge_concept_aliases
            WHERE concept_id=(
                SELECT concept_id FROM knowledge_concepts
                WHERE concept_uid=?
            ) AND normalized_alias=?
            """,
            (bonding_uid, schema.normalized_term("bond amount")),
        ).fetchone()[0]
        approved_seed = curation.upsert_human_concept_alias(
            self.connection,
            concept_uid=bonding_uid,
            alias="Bond Amount",
            reviewer="human-3",
            note="Approved seed spelling.",
            now_iso=lambda: NOW,
        )
        schema.ensure_knowledge_schema(
            self.connection,
            lambda: "2026-07-18T14:00:00Z",
        )
        seeded_after = self.connection.execute(
            """
            SELECT alias_uid, alias_text, source
            FROM knowledge_concept_aliases
            WHERE normalized_alias=?
              AND concept_id=(
                  SELECT concept_id FROM knowledge_concepts
                  WHERE concept_uid=?
              )
            """,
            (schema.normalized_term("bond amount"), bonding_uid),
        ).fetchone()
        self.assertEqual(seeded_before, approved_seed["alias"]["aliasUid"])
        self.assertEqual(seeded_before, seeded_after["alias_uid"])
        self.assertEqual("Bond Amount", seeded_after["alias_text"])
        self.assertEqual("HUMAN_APPROVED", seeded_after["source"])


if __name__ == "__main__":
    unittest.main()
