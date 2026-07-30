from __future__ import annotations

import sqlite3
import unittest

import inference_data_ai_schema as schema


class CanonicalSchemaMigrationTests(unittest.TestCase):
    def test_measurement_series_schema_is_additive_and_migratable(
        self,
    ) -> None:
        connection = sqlite3.connect(":memory:")
        try:
            schema.ensure_knowledge_schema(
                connection,
                lambda: "2026-07-18T00:00:00Z",
            )
            tables = {
                row[0]
                for row in connection.execute(
                    """
                    SELECT name
                    FROM sqlite_master
                    WHERE type='table'
                    """
                )
            }
            self.assertIn("knowledge_measurement_series", tables)
            self.assertIn("knowledge_measurement_points", tables)
            series_columns = {
                row[1]
                for row in connection.execute(
                    "PRAGMA table_info(knowledge_measurement_series)"
                )
            }
            point_columns = {
                row[1]
                for row in connection.execute(
                    "PRAGMA table_info(knowledge_measurement_points)"
                )
            }
            self.assertTrue(
                {
                    "series_uid",
                    "public_series_id",
                    "study_id",
                    "outcome_id",
                    "arm_id",
                    "header_range",
                    "value_range",
                    "row_identity_range",
                    "axis_source",
                }.issubset(series_columns)
            )
            self.assertTrue(
                {
                    "point_uid",
                    "public_point_id",
                    "axis_label",
                    "axis_value",
                    "replicate_key",
                    "replicate_role",
                    "stratum_key",
                    "value_number",
                    "source_coordinate",
                    "axis_source_coordinate",
                    "replicate_source_coordinate",
                }.issubset(point_columns)
            )
            migration_count = connection.execute(
                """
                SELECT COUNT(*)
                FROM schema_migrations
                WHERE migration_name=?
                """,
                (schema.MEASUREMENT_SERIES_MIGRATION,),
            ).fetchone()[0]
            self.assertEqual(1, migration_count)
            axis_migration_count = connection.execute(
                """
                SELECT COUNT(*)
                FROM schema_migrations
                WHERE migration_name=?
                """,
                (schema.MEASUREMENT_SERIES_AXIS_MIGRATION,),
            ).fetchone()[0]
            self.assertEqual(1, axis_migration_count)
            replicate_role_migration_count = connection.execute(
                """
                SELECT COUNT(*)
                FROM schema_migrations
                WHERE migration_name=?
                """,
                (schema.MEASUREMENT_POINT_REPLICATE_ROLE_MIGRATION,),
            ).fetchone()[0]
            self.assertEqual(1, replicate_role_migration_count)

            schema.ensure_knowledge_schema(
                connection,
                lambda: "2026-07-18T00:00:01Z",
            )
            self.assertEqual(
                1,
                connection.execute(
                    """
                    SELECT COUNT(*)
                    FROM schema_migrations
                    WHERE migration_name=?
                    """,
                    (schema.MEASUREMENT_SERIES_MIGRATION,),
                ).fetchone()[0],
            )
        finally:
            connection.close()

    def test_measurement_series_v1_table_gets_axis_source_column(
        self,
    ) -> None:
        connection = sqlite3.connect(":memory:")
        try:
            connection.execute(
                """
                CREATE TABLE knowledge_measurement_series (
                    series_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    series_uid TEXT NOT NULL UNIQUE,
                    public_series_id TEXT NOT NULL UNIQUE,
                    study_id INTEGER NOT NULL,
                    outcome_id INTEGER NOT NULL,
                    arm_id INTEGER NOT NULL,
                    series_key TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    header_range TEXT NOT NULL,
                    value_range TEXT NOT NULL,
                    row_identity_range TEXT NOT NULL,
                    axis_name TEXT NOT NULL DEFAULT '',
                    axis_unit_id INTEGER,
                    original_axis_unit TEXT NOT NULL DEFAULT '',
                    value_unit_id INTEGER,
                    original_value_unit TEXT NOT NULL DEFAULT '',
                    stratum_key TEXT NOT NULL DEFAULT '',
                    verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
                    details_json TEXT NOT NULL DEFAULT '{}',
                    UNIQUE(study_id, series_key)
                )
                """
            )
            schema.ensure_knowledge_schema(
                connection,
                lambda: "2026-07-18T00:00:00Z",
            )
            columns = {
                row[1]
                for row in connection.execute(
                    "PRAGMA table_info(knowledge_measurement_series)"
                )
            }
            self.assertIn("axis_source", columns)
            connection.execute(
                """
                INSERT INTO knowledge_measurement_series(
                    series_uid, public_series_id, study_id, outcome_id,
                    arm_id, series_key, sheet_name, header_range,
                    value_range, row_identity_range
                ) VALUES (
                    'series-1', 'SER-1', 1, 1, 1, 'series-1',
                    'Data', 'B1:C1', 'B2:C3', 'A2:A3'
                )
                """
            )
            self.assertEqual(
                "ROW_IDENTITY",
                connection.execute(
                    """
                    SELECT axis_source
                    FROM knowledge_measurement_series
                    """
                ).fetchone()[0],
            )
        finally:
            connection.close()

    def test_v1_effect_constraint_is_rebuilt_with_outcome_in_unique_key(
        self,
    ) -> None:
        connection = sqlite3.connect(":memory:")
        try:
            connection.execute(
                """
                CREATE TABLE knowledge_effects (
                    effect_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    effect_uid TEXT NOT NULL UNIQUE,
                    public_effect_id TEXT NOT NULL UNIQUE,
                    comparison_id INTEGER NOT NULL,
                    outcome_id INTEGER NOT NULL,
                    effect_type TEXT NOT NULL,
                    estimate REAL,
                    unit_id INTEGER,
                    original_unit TEXT NOT NULL DEFAULT '',
                    ci_lower REAL,
                    ci_upper REAL,
                    formula_version TEXT NOT NULL DEFAULT 'legacy-v1',
                    calculation_text TEXT NOT NULL DEFAULT '',
                    direction TEXT NOT NULL DEFAULT '',
                    aggregation_eligible INTEGER NOT NULL DEFAULT 0,
                    verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
                    details_json TEXT NOT NULL DEFAULT '{}',
                    UNIQUE(
                        comparison_id, effect_type, formula_version
                    )
                )
                """
            )
            schema.ensure_knowledge_schema(
                connection,
                lambda: "2026-07-17T00:00:00Z",
            )
            connection.commit()
            table_sql = connection.execute(
                """
                SELECT sql FROM sqlite_master
                WHERE type='table' AND name='knowledge_effects'
                """
            ).fetchone()[0]
            self.assertIn(
                "comparison_id, outcome_id, effect_type, formula_version",
                " ".join(str(table_sql).split()),
            )

            connection.execute("PRAGMA foreign_keys=OFF")
            base = (
                "ABSOLUTE_DIFFERENCE",
                1.0,
                None,
                "",
                None,
                None,
                "canonical-v1",
                "",
                "HIGHER",
                0,
                "NEEDS_REVIEW",
                "{}",
            )
            connection.execute(
                """
                INSERT INTO knowledge_effects(
                    effect_uid, public_effect_id, comparison_id, outcome_id,
                    effect_type, estimate, unit_id, original_unit, ci_lower,
                    ci_upper, formula_version, calculation_text, direction,
                    aggregation_eligible, verification_status, details_json
                ) VALUES ('effect-1', 'EFF-1', 1, 10, ?, ?, ?, ?, ?, ?, ?,
                          ?, ?, ?, ?, ?)
                """,
                base,
            )
            connection.execute(
                """
                INSERT INTO knowledge_effects(
                    effect_uid, public_effect_id, comparison_id, outcome_id,
                    effect_type, estimate, unit_id, original_unit, ci_lower,
                    ci_upper, formula_version, calculation_text, direction,
                    aggregation_eligible, verification_status, details_json
                ) VALUES ('effect-2', 'EFF-2', 1, 11, ?, ?, ?, ?, ?, ?, ?,
                          ?, ?, ?, ?, ?)
                """,
                base,
            )
            self.assertEqual(
                2,
                connection.execute(
                    "SELECT COUNT(*) FROM knowledge_effects"
                ).fetchone()[0],
            )
            migration_count = connection.execute(
                """
                SELECT COUNT(*) FROM schema_migrations
                WHERE migration_name=?
                """,
                (schema.EFFECT_OUTCOME_UNIQUENESS_MIGRATION,),
            ).fetchone()[0]
            self.assertEqual(1, migration_count)

            schema.ensure_knowledge_schema(
                connection,
                lambda: "2026-07-17T00:00:01Z",
            )
            self.assertEqual(
                2,
                connection.execute(
                    "SELECT COUNT(*) FROM knowledge_effects"
                ).fetchone()[0],
            )
        finally:
            connection.close()


if __name__ == "__main__":
    unittest.main()
