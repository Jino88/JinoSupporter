from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import unicodedata
from collections.abc import Callable
from typing import Any


KNOWLEDGE_MIGRATION = "canonical-study-evidence-v1"
EFFECT_OUTCOME_UNIQUENESS_MIGRATION = (
    "canonical-effect-outcome-uniqueness-v2"
)
MEASUREMENT_SERIES_MIGRATION = "canonical-measurement-series-v1"
MEASUREMENT_SERIES_AXIS_MIGRATION = (
    "canonical-measurement-series-axis-source-v2"
)
MEASUREMENT_POINT_REPLICATE_ROLE_MIGRATION = (
    "canonical-measurement-point-replicate-role-v1"
)


def normalize_key_part(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).strip().casefold()
    return re.sub(r"\s+", " ", text)


def normalized_term(value: object) -> str:
    text = normalize_key_part(value)
    text = re.sub(r"[^\w%+./-]+", " ", text, flags=re.UNICODE)
    return re.sub(r"\s+", " ", text).strip()


def stable_uid(prefix: str, *parts: object) -> str:
    payload = "\x1f".join(normalize_key_part(part) for part in parts)
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return f"{prefix.lower()}_{digest[:24]}"


def public_id(prefix: str, uid: str) -> str:
    digest = hashlib.sha256(uid.encode("utf-8")).hexdigest()[:12].upper()
    return f"{prefix.upper()}-{digest}"


def _json(value: object, fallback: object) -> str:
    return json.dumps(fallback if value is None else value, ensure_ascii=False, separators=(",", ":"))


def _table_exists(conn: sqlite3.Connection, table: str) -> bool:
    return (
        conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
            (table,),
        ).fetchone()
        is not None
    )


def ensure_knowledge_schema(conn: sqlite3.Connection, now_iso: Callable[[], str]) -> None:
    """Install the additive canonical Study/Comparison/Evidence schema.

    Legacy ``analysis_*`` tables remain readable and are treated as a
    compatibility input.  Public identifiers in this schema never depend on
    SQLite row ids.
    """

    conn.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS schema_migrations (
            migration_name TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS source_documents (
            document_id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_uid TEXT NOT NULL UNIQUE,
            dataset TEXT NOT NULL,
            source_path TEXT NOT NULL,
            original_file_name TEXT NOT NULL DEFAULT '',
            source_kind TEXT NOT NULL DEFAULT 'XLSX',
            lifecycle_status TEXT NOT NULL DEFAULT 'ACTIVE'
                CHECK(lifecycle_status IN ('ACTIVE','MISSING','EXCLUDED')),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(dataset, source_path)
        );

        CREATE TABLE IF NOT EXISTS source_revisions (
            revision_id INTEGER PRIMARY KEY AUTOINCREMENT,
            revision_uid TEXT NOT NULL UNIQUE,
            document_id INTEGER NOT NULL,
            legacy_workbook_id INTEGER,
            source_fingerprint TEXT NOT NULL,
            fingerprint_kind TEXT NOT NULL DEFAULT 'LEGACY_METADATA',
            content_sha256 TEXT NOT NULL DEFAULT '',
            size_bytes INTEGER NOT NULL DEFAULT 0,
            mtime_ns INTEGER NOT NULL DEFAULT 0,
            extractor_name TEXT NOT NULL DEFAULT '',
            extractor_version TEXT NOT NULL DEFAULT '',
            capture_contract TEXT NOT NULL DEFAULT '',
            capture_status TEXT NOT NULL DEFAULT 'CAPTURED'
                CHECK(capture_status IN ('CAPTURED','PARTIAL','FAILED','STALE')),
            is_current INTEGER NOT NULL DEFAULT 1 CHECK(is_current IN (0,1)),
            captured_at TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES source_documents(document_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_workbook_id) REFERENCES workbooks(workbook_id) ON DELETE SET NULL,
            UNIQUE(document_id, source_fingerprint, capture_contract)
        );

        CREATE TABLE IF NOT EXISTS knowledge_concepts (
            concept_id INTEGER PRIMARY KEY AUTOINCREMENT,
            concept_uid TEXT NOT NULL UNIQUE,
            concept_kind TEXT NOT NULL,
            canonical_name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            description TEXT NOT NULL DEFAULT '',
            lifecycle_status TEXT NOT NULL DEFAULT 'ACTIVE'
                CHECK(lifecycle_status IN ('ACTIVE','DEPRECATED')),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(concept_kind, normalized_name)
        );

        CREATE TABLE IF NOT EXISTS knowledge_concept_aliases (
            alias_id INTEGER PRIMARY KEY AUTOINCREMENT,
            alias_uid TEXT NOT NULL UNIQUE,
            concept_id INTEGER NOT NULL,
            alias_text TEXT NOT NULL,
            normalized_alias TEXT NOT NULL,
            language TEXT NOT NULL DEFAULT '',
            source TEXT NOT NULL DEFAULT 'SEED',
            confidence REAL NOT NULL DEFAULT 1.0 CHECK(confidence >= 0 AND confidence <= 1),
            created_at TEXT NOT NULL,
            FOREIGN KEY(concept_id) REFERENCES knowledge_concepts(concept_id) ON DELETE CASCADE,
            UNIQUE(concept_id, normalized_alias)
        );

        CREATE TABLE IF NOT EXISTS knowledge_units (
            unit_id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_uid TEXT NOT NULL UNIQUE,
            canonical_symbol TEXT NOT NULL UNIQUE,
            quantity_kind TEXT NOT NULL,
            scale_to_base REAL NOT NULL DEFAULT 1,
            offset_to_base REAL NOT NULL DEFAULT 0,
            aliases_json TEXT NOT NULL DEFAULT '[]',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS knowledge_schema_candidates (
            schema_candidate_id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_uid TEXT NOT NULL UNIQUE,
            candidate_kind TEXT NOT NULL,
            normalized_value TEXT NOT NULL,
            original_value TEXT NOT NULL,
            suggested_canonical_name TEXT NOT NULL DEFAULT '',
            occurrence_count INTEGER NOT NULL DEFAULT 1,
            status TEXT NOT NULL DEFAULT 'OPEN'
                CHECK(status IN ('OPEN','APPROVED','REJECTED','MERGED')),
            first_seen_at TEXT NOT NULL,
            last_seen_at TEXT NOT NULL,
            UNIQUE(candidate_kind, normalized_value)
        );

        CREATE TABLE IF NOT EXISTS workbook_analyses (
            workbook_analysis_id INTEGER PRIMARY KEY AUTOINCREMENT,
            analysis_uid TEXT NOT NULL UNIQUE,
            public_analysis_id TEXT NOT NULL UNIQUE,
            document_id INTEGER NOT NULL,
            revision_id INTEGER NOT NULL,
            legacy_analysis_report_id INTEGER,
            analysis_key TEXT NOT NULL,
            title TEXT NOT NULL DEFAULT '',
            analysis_type TEXT NOT NULL DEFAULT '',
            purpose TEXT NOT NULL DEFAULT '',
            scope_text TEXT NOT NULL DEFAULT '',
            analysis_status TEXT NOT NULL DEFAULT 'DRAFT',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                CHECK(verification_status IN ('VERIFIED','NEEDS_REVIEW','EXCLUDED','FAILED','STALE')),
            decision_text TEXT NOT NULL DEFAULT '',
            consolidated_summary TEXT NOT NULL DEFAULT '',
            limitations_json TEXT NOT NULL DEFAULT '[]',
            analyzer_name TEXT NOT NULL DEFAULT 'legacy-analysis-v1',
            analyzer_version TEXT NOT NULL DEFAULT '1',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES source_documents(document_id) ON DELETE CASCADE,
            FOREIGN KEY(revision_id) REFERENCES source_revisions(revision_id),
            FOREIGN KEY(legacy_analysis_report_id) REFERENCES analysis_reports(analysis_report_id) ON DELETE SET NULL,
            UNIQUE(document_id, analysis_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_studies (
            study_id INTEGER PRIMARY KEY AUTOINCREMENT,
            study_uid TEXT NOT NULL UNIQUE,
            public_data_id TEXT NOT NULL UNIQUE,
            workbook_analysis_id INTEGER NOT NULL,
            legacy_review_item_id INTEGER,
            study_key TEXT NOT NULL,
            title TEXT NOT NULL DEFAULT '',
            purpose TEXT NOT NULL DEFAULT '',
            hypothesis TEXT NOT NULL DEFAULT '',
            objective TEXT NOT NULL DEFAULT '',
            design_type TEXT NOT NULL DEFAULT 'UNSPECIFIED',
            comparison_basis TEXT NOT NULL DEFAULT '',
            analysis_status TEXT NOT NULL DEFAULT 'DRAFT',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                CHECK(verification_status IN ('VERIFIED','NEEDS_REVIEW','EXCLUDED','FAILED','STALE')),
            comparability_status TEXT NOT NULL DEFAULT 'UNASSESSED'
                CHECK(comparability_status IN ('VALID','PARTIAL','INVALID','UNASSESSED')),
            confounding_status TEXT NOT NULL DEFAULT 'UNASSESSED'
                CHECK(confounding_status IN ('NONE','POSSIBLE','CONFOUNDED','UNASSESSED')),
            decision_text TEXT NOT NULL DEFAULT '',
            summary_text TEXT NOT NULL DEFAULT '',
            limitations_json TEXT NOT NULL DEFAULT '[]',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(workbook_analysis_id) REFERENCES workbook_analyses(workbook_analysis_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_review_item_id) REFERENCES analysis_review_items(review_item_id) ON DELETE SET NULL,
            UNIQUE(workbook_analysis_id, study_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_study_contexts (
            context_id INTEGER PRIMARY KEY AUTOINCREMENT,
            context_uid TEXT NOT NULL UNIQUE,
            study_id INTEGER NOT NULL,
            context_kind TEXT NOT NULL,
            concept_id INTEGER,
            original_value TEXT NOT NULL,
            normalized_value TEXT NOT NULL DEFAULT '',
            value_number REAL,
            unit_id INTEGER,
            start_value TEXT NOT NULL DEFAULT '',
            end_value TEXT NOT NULL DEFAULT '',
            attributes_json TEXT NOT NULL DEFAULT '{}',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            FOREIGN KEY(study_id) REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(concept_id) REFERENCES knowledge_concepts(concept_id),
            FOREIGN KEY(unit_id) REFERENCES knowledge_units(unit_id)
        );

        CREATE TABLE IF NOT EXISTS knowledge_factors (
            factor_id INTEGER PRIMARY KEY AUTOINCREMENT,
            factor_uid TEXT NOT NULL UNIQUE,
            study_id INTEGER NOT NULL,
            concept_id INTEGER,
            factor_key TEXT NOT NULL,
            factor_domain TEXT NOT NULL DEFAULT '',
            original_label TEXT NOT NULL DEFAULT '',
            baseline_condition TEXT NOT NULL DEFAULT '',
            changed_condition TEXT NOT NULL DEFAULT '',
            change_direction TEXT NOT NULL DEFAULT '',
            isolation_status TEXT NOT NULL DEFAULT 'UNASSESSED'
                CHECK(isolation_status IN ('ISOLATED','MULTI_FACTOR','CONFOUNDED','UNASSESSED')),
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            attributes_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(study_id) REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(concept_id) REFERENCES knowledge_concepts(concept_id),
            UNIQUE(study_id, factor_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_arms (
            arm_id INTEGER PRIMARY KEY AUTOINCREMENT,
            arm_uid TEXT NOT NULL UNIQUE,
            study_id INTEGER NOT NULL,
            legacy_cohort_id INTEGER,
            arm_key TEXT NOT NULL,
            arm_role TEXT NOT NULL DEFAULT 'OTHER'
                CHECK(arm_role IN ('CONTROL','COMPARATOR','TREATMENT','TEST','BEFORE','AFTER','REFERENCE','OTHER')),
            label TEXT NOT NULL DEFAULT '',
            condition_text TEXT NOT NULL DEFAULT '',
            sample_size REAL,
            sample_basis TEXT NOT NULL DEFAULT '',
            matching_basis TEXT NOT NULL DEFAULT '',
            attributes_json TEXT NOT NULL DEFAULT '{}',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            FOREIGN KEY(study_id) REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_cohort_id) REFERENCES analysis_cohorts(cohort_id) ON DELETE SET NULL,
            UNIQUE(study_id, arm_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_arm_factor_values (
            arm_factor_value_id INTEGER PRIMARY KEY AUTOINCREMENT,
            arm_id INTEGER NOT NULL,
            factor_id INTEGER NOT NULL,
            original_value TEXT NOT NULL DEFAULT '',
            value_number REAL,
            unit_id INTEGER,
            is_baseline INTEGER NOT NULL DEFAULT 0 CHECK(is_baseline IN (0,1)),
            held_constant INTEGER NOT NULL DEFAULT 0 CHECK(held_constant IN (0,1)),
            FOREIGN KEY(arm_id) REFERENCES knowledge_arms(arm_id) ON DELETE CASCADE,
            FOREIGN KEY(factor_id) REFERENCES knowledge_factors(factor_id) ON DELETE CASCADE,
            FOREIGN KEY(unit_id) REFERENCES knowledge_units(unit_id),
            UNIQUE(arm_id, factor_id)
        );

        CREATE TABLE IF NOT EXISTS knowledge_outcomes (
            outcome_id INTEGER PRIMARY KEY AUTOINCREMENT,
            outcome_uid TEXT NOT NULL UNIQUE,
            study_id INTEGER NOT NULL,
            legacy_metric_id INTEGER,
            outcome_key TEXT NOT NULL,
            concept_id INTEGER,
            original_label TEXT NOT NULL DEFAULT '',
            outcome_domain TEXT NOT NULL DEFAULT '',
            metric_type TEXT NOT NULL DEFAULT '',
            unit_id INTEGER,
            original_unit TEXT NOT NULL DEFAULT '',
            denominator_basis TEXT NOT NULL DEFAULT '',
            favorable_direction TEXT NOT NULL DEFAULT 'UNKNOWN'
                CHECK(favorable_direction IN ('LOWER','HIGHER','TARGET','NONE','UNKNOWN')),
            definition_text TEXT NOT NULL DEFAULT '',
            spec_text TEXT NOT NULL DEFAULT '',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            attributes_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(study_id) REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_metric_id) REFERENCES analysis_metrics(metric_id) ON DELETE SET NULL,
            FOREIGN KEY(concept_id) REFERENCES knowledge_concepts(concept_id),
            FOREIGN KEY(unit_id) REFERENCES knowledge_units(unit_id),
            UNIQUE(study_id, outcome_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_observations (
            observation_id INTEGER PRIMARY KEY AUTOINCREMENT,
            observation_uid TEXT NOT NULL UNIQUE,
            outcome_id INTEGER NOT NULL,
            arm_id INTEGER NOT NULL,
            legacy_metric_value_id INTEGER,
            observation_key TEXT NOT NULL,
            stratum_key TEXT NOT NULL DEFAULT '',
            replicate_key TEXT NOT NULL DEFAULT '',
            observed_at TEXT NOT NULL DEFAULT '',
            value_number REAL,
            value_text TEXT NOT NULL DEFAULT '',
            numerator REAL CHECK(numerator IS NULL OR numerator >= 0),
            denominator REAL CHECK(denominator IS NULL OR denominator > 0),
            rate_ppm REAL,
            min_value REAL,
            max_value REAL,
            average_value REAL,
            sample_size REAL,
            result_status TEXT NOT NULL DEFAULT '',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            details_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(outcome_id) REFERENCES knowledge_outcomes(outcome_id) ON DELETE CASCADE,
            FOREIGN KEY(arm_id) REFERENCES knowledge_arms(arm_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_metric_value_id) REFERENCES analysis_metric_values(metric_value_id) ON DELETE SET NULL,
            UNIQUE(outcome_id, arm_id, observation_key, stratum_key, replicate_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_measurement_series (
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
            axis_source TEXT NOT NULL
                CHECK(axis_source IN ('HEADER','ROW_IDENTITY')),
            axis_unit_id INTEGER,
            original_axis_unit TEXT NOT NULL DEFAULT '',
            value_unit_id INTEGER,
            original_value_unit TEXT NOT NULL DEFAULT '',
            stratum_key TEXT NOT NULL DEFAULT '',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                CHECK(verification_status IN (
                    'VERIFIED','NEEDS_REVIEW','EXCLUDED','FAILED','STALE'
                )),
            details_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(study_id)
                REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(outcome_id)
                REFERENCES knowledge_outcomes(outcome_id) ON DELETE CASCADE,
            FOREIGN KEY(arm_id)
                REFERENCES knowledge_arms(arm_id) ON DELETE CASCADE,
            FOREIGN KEY(axis_unit_id) REFERENCES knowledge_units(unit_id),
            FOREIGN KEY(value_unit_id) REFERENCES knowledge_units(unit_id),
            UNIQUE(study_id, series_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_measurement_points (
            point_id INTEGER PRIMARY KEY AUTOINCREMENT,
            point_uid TEXT NOT NULL UNIQUE,
            public_point_id TEXT NOT NULL UNIQUE,
            series_id INTEGER NOT NULL,
            row_ordinal INTEGER NOT NULL CHECK(row_ordinal > 0),
            column_ordinal INTEGER NOT NULL CHECK(column_ordinal > 0),
            axis_label TEXT NOT NULL,
            axis_value REAL,
            axis_unit_id INTEGER,
            original_axis_unit TEXT NOT NULL DEFAULT '',
            replicate_key TEXT NOT NULL,
            replicate_role TEXT NOT NULL DEFAULT 'RAW'
                CHECK(replicate_role IN ('RAW','AGGREGATE')),
            stratum_key TEXT NOT NULL DEFAULT '',
            value_number REAL NOT NULL,
            value_unit_id INTEGER,
            original_value_unit TEXT NOT NULL DEFAULT '',
            source_revision_id INTEGER NOT NULL,
            source_sheet_name TEXT NOT NULL,
            source_row_index INTEGER NOT NULL CHECK(source_row_index > 0),
            source_column_index INTEGER NOT NULL
                CHECK(source_column_index > 0),
            source_coordinate TEXT NOT NULL,
            axis_source_coordinate TEXT NOT NULL,
            replicate_source_coordinate TEXT NOT NULL,
            source_value_json TEXT NOT NULL,
            source_formula_text TEXT NOT NULL DEFAULT '',
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                CHECK(verification_status IN (
                    'VERIFIED','NEEDS_REVIEW','EXCLUDED','FAILED','STALE'
                )),
            FOREIGN KEY(series_id)
                REFERENCES knowledge_measurement_series(series_id)
                ON DELETE CASCADE,
            FOREIGN KEY(axis_unit_id) REFERENCES knowledge_units(unit_id),
            FOREIGN KEY(value_unit_id) REFERENCES knowledge_units(unit_id),
            FOREIGN KEY(source_revision_id)
                REFERENCES source_revisions(revision_id),
            UNIQUE(series_id, row_ordinal, column_ordinal)
        );

        CREATE TABLE IF NOT EXISTS knowledge_comparisons (
            comparison_id INTEGER PRIMARY KEY AUTOINCREMENT,
            comparison_uid TEXT NOT NULL UNIQUE,
            public_comparison_id TEXT NOT NULL UNIQUE,
            study_id INTEGER NOT NULL,
            legacy_comparison_id INTEGER,
            comparison_key TEXT NOT NULL,
            compared_arm_id INTEGER NOT NULL,
            control_arm_id INTEGER NOT NULL,
            design_type TEXT NOT NULL DEFAULT 'UNSPECIFIED',
            matching_basis TEXT NOT NULL DEFAULT '',
            validity_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                CHECK(validity_status IN ('VALID','NEEDS_REVIEW','INVALID','EXCLUDED')),
            confounding_status TEXT NOT NULL DEFAULT 'UNASSESSED'
                CHECK(confounding_status IN ('NONE','POSSIBLE','CONFOUNDED','UNASSESSED')),
            exclusion_reason TEXT NOT NULL DEFAULT '',
            direction TEXT NOT NULL DEFAULT '',
            summary_text TEXT NOT NULL DEFAULT '',
            aggregation_eligible INTEGER NOT NULL DEFAULT 0 CHECK(aggregation_eligible IN (0,1)),
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            details_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(study_id) REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_comparison_id) REFERENCES analysis_comparisons(comparison_id) ON DELETE SET NULL,
            FOREIGN KEY(compared_arm_id) REFERENCES knowledge_arms(arm_id),
            FOREIGN KEY(control_arm_id) REFERENCES knowledge_arms(arm_id),
            CHECK(compared_arm_id <> control_arm_id),
            UNIQUE(study_id, comparison_key)
        );

        CREATE TABLE IF NOT EXISTS knowledge_effects (
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
            aggregation_eligible INTEGER NOT NULL DEFAULT 0 CHECK(aggregation_eligible IN (0,1)),
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                CHECK(verification_status IN ('VERIFIED','NEEDS_REVIEW','INVALID','EXCLUDED','STALE')),
            details_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(comparison_id) REFERENCES knowledge_comparisons(comparison_id) ON DELETE CASCADE,
            FOREIGN KEY(outcome_id) REFERENCES knowledge_outcomes(outcome_id) ON DELETE CASCADE,
            FOREIGN KEY(unit_id) REFERENCES knowledge_units(unit_id),
            UNIQUE(comparison_id, outcome_id, effect_type, formula_version)
        );

        CREATE TABLE IF NOT EXISTS evidence_items (
            evidence_id INTEGER PRIMARY KEY AUTOINCREMENT,
            evidence_uid TEXT NOT NULL UNIQUE,
            public_evidence_id TEXT NOT NULL UNIQUE,
            revision_id INTEGER NOT NULL,
            legacy_evidence_id INTEGER,
            evidence_kind TEXT NOT NULL DEFAULT 'CELL_RANGE'
                CHECK(evidence_kind IN ('CELL','CELL_RANGE','TABLE','TEXT')),
            sheet_name TEXT NOT NULL,
            start_row INTEGER NOT NULL CHECK(start_row > 0),
            start_col INTEGER NOT NULL CHECK(start_col > 0),
            end_row INTEGER NOT NULL CHECK(end_row >= start_row),
            end_col INTEGER NOT NULL CHECK(end_col >= start_col),
            range_address TEXT NOT NULL,
            evidence_role TEXT NOT NULL DEFAULT 'SOURCE',
            source_text TEXT NOT NULL DEFAULT '',
            note TEXT NOT NULL DEFAULT '',
            content_sha256 TEXT NOT NULL DEFAULT '',
            verification_status TEXT NOT NULL DEFAULT 'VERIFIED'
                CHECK(verification_status IN ('VERIFIED','NEEDS_REVIEW','INVALID','STALE')),
            created_at TEXT NOT NULL,
            FOREIGN KEY(revision_id) REFERENCES source_revisions(revision_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_evidence_id) REFERENCES analysis_evidence(evidence_id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS entity_evidence_links (
            entity_evidence_link_id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_type TEXT NOT NULL,
            entity_uid TEXT NOT NULL,
            evidence_id INTEGER NOT NULL,
            evidence_role TEXT NOT NULL DEFAULT 'SOURCE',
            claim_scope TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(evidence_id) REFERENCES evidence_items(evidence_id) ON DELETE CASCADE,
            UNIQUE(entity_type, entity_uid, evidence_id, evidence_role)
        );

        CREATE TABLE IF NOT EXISTS knowledge_claims (
            claim_id INTEGER PRIMARY KEY AUTOINCREMENT,
            claim_uid TEXT NOT NULL UNIQUE,
            public_claim_id TEXT NOT NULL UNIQUE,
            workbook_analysis_id INTEGER NOT NULL,
            study_id INTEGER,
            legacy_conclusion_id INTEGER,
            claim_key TEXT NOT NULL,
            claim_type TEXT NOT NULL DEFAULT 'SOURCE_CONCLUSION',
            claim_text TEXT NOT NULL,
            verdict TEXT NOT NULL DEFAULT '',
            causal_strength TEXT NOT NULL DEFAULT 'UNSPECIFIED'
                CHECK(causal_strength IN ('CAUSAL','ASSOCIATION','DESCRIPTIVE','UNSPECIFIED')),
            verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW',
            limitations_json TEXT NOT NULL DEFAULT '[]',
            FOREIGN KEY(workbook_analysis_id) REFERENCES workbook_analyses(workbook_analysis_id) ON DELETE CASCADE,
            FOREIGN KEY(study_id) REFERENCES knowledge_studies(study_id) ON DELETE CASCADE,
            FOREIGN KEY(legacy_conclusion_id) REFERENCES analysis_conclusions(conclusion_id) ON DELETE SET NULL,
            UNIQUE(workbook_analysis_id, study_id, claim_key)
        );

        CREATE TABLE IF NOT EXISTS validation_issues (
            validation_issue_id INTEGER PRIMARY KEY AUTOINCREMENT,
            issue_uid TEXT NOT NULL UNIQUE,
            entity_type TEXT NOT NULL,
            entity_uid TEXT NOT NULL,
            issue_code TEXT NOT NULL,
            severity TEXT NOT NULL CHECK(severity IN ('INFO','WARNING','ERROR','BLOCKING')),
            message TEXT NOT NULL,
            details_json TEXT NOT NULL DEFAULT '{}',
            status TEXT NOT NULL DEFAULT 'OPEN' CHECK(status IN ('OPEN','RESOLVED','ACCEPTED')),
            validator_name TEXT NOT NULL DEFAULT '',
            validator_version TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            resolved_at TEXT NOT NULL DEFAULT '',
            UNIQUE(entity_type, entity_uid, issue_code, status)
        );

        CREATE TABLE IF NOT EXISTS review_decisions (
            review_decision_id INTEGER PRIMARY KEY AUTOINCREMENT,
            decision_uid TEXT NOT NULL UNIQUE,
            entity_type TEXT NOT NULL,
            entity_uid TEXT NOT NULL,
            decision TEXT NOT NULL CHECK(decision IN ('APPROVE','REJECT','EXCLUDE','RETURN_TO_REVIEW')),
            reason TEXT NOT NULL DEFAULT '',
            reviewer TEXT NOT NULL DEFAULT '',
            decided_at TEXT NOT NULL,
            supersedes_decision_uid TEXT NOT NULL DEFAULT ''
        );

        CREATE INDEX IF NOT EXISTS idx_source_documents_dataset ON source_documents(dataset, lifecycle_status);
        CREATE INDEX IF NOT EXISTS idx_source_revisions_current ON source_revisions(document_id, is_current);
        CREATE INDEX IF NOT EXISTS idx_concept_alias_lookup ON knowledge_concept_aliases(normalized_alias);
        CREATE INDEX IF NOT EXISTS idx_schema_candidates_open ON knowledge_schema_candidates(status, candidate_kind);
        CREATE INDEX IF NOT EXISTS idx_workbook_analyses_status ON workbook_analyses(verification_status, document_id);
        CREATE INDEX IF NOT EXISTS idx_studies_status ON knowledge_studies(verification_status, comparability_status, confounding_status);
        CREATE INDEX IF NOT EXISTS idx_context_lookup ON knowledge_study_contexts(context_kind, normalized_value);
        CREATE INDEX IF NOT EXISTS idx_factor_concept ON knowledge_factors(concept_id, isolation_status);
        CREATE INDEX IF NOT EXISTS idx_arm_study_role ON knowledge_arms(study_id, arm_role);
        CREATE INDEX IF NOT EXISTS idx_outcome_concept ON knowledge_outcomes(concept_id, metric_type);
        CREATE INDEX IF NOT EXISTS idx_observation_outcome_arm ON knowledge_observations(outcome_id, arm_id);
        CREATE INDEX IF NOT EXISTS idx_measurement_series_study
            ON knowledge_measurement_series(study_id, outcome_id, arm_id);
        CREATE INDEX IF NOT EXISTS idx_measurement_point_series_axis
            ON knowledge_measurement_points(
                series_id, row_ordinal, column_ordinal
            );
        CREATE INDEX IF NOT EXISTS idx_measurement_point_source
            ON knowledge_measurement_points(
                source_revision_id, source_sheet_name,
                source_row_index, source_column_index
            );
        CREATE INDEX IF NOT EXISTS idx_comparison_eligibility ON knowledge_comparisons(aggregation_eligible, validity_status, confounding_status);
        CREATE INDEX IF NOT EXISTS idx_effect_eligibility ON knowledge_effects(aggregation_eligible, effect_type, verification_status);
        CREATE INDEX IF NOT EXISTS idx_evidence_revision_range ON evidence_items(revision_id, sheet_name, start_row, start_col);
        CREATE INDEX IF NOT EXISTS idx_entity_evidence ON entity_evidence_links(entity_type, entity_uid);
        CREATE INDEX IF NOT EXISTS idx_validation_open ON validation_issues(status, severity, entity_type);

        CREATE TRIGGER IF NOT EXISTS trg_effect_aggregation_guard_insert
        BEFORE INSERT ON knowledge_effects
        WHEN NEW.aggregation_eligible = 1
        BEGIN
            SELECT CASE
                WHEN NEW.verification_status <> 'VERIFIED'
                THEN RAISE(ABORT, 'aggregation-eligible effect must be VERIFIED')
            END;
            SELECT CASE
                WHEN NOT EXISTS (
                    SELECT 1 FROM knowledge_comparisons c
                    WHERE c.comparison_id=NEW.comparison_id
                      AND c.validity_status='VALID'
                      AND c.confounding_status='NONE'
                      AND c.aggregation_eligible=1
                      AND c.verification_status='VERIFIED'
                )
                THEN RAISE(ABORT, 'aggregation-eligible effect requires a valid unconfounded verified comparison')
            END;
        END;

        CREATE TRIGGER IF NOT EXISTS trg_effect_aggregation_guard_update
        BEFORE UPDATE OF aggregation_eligible, verification_status, comparison_id ON knowledge_effects
        WHEN NEW.aggregation_eligible = 1
        BEGIN
            SELECT CASE
                WHEN NEW.verification_status <> 'VERIFIED'
                THEN RAISE(ABORT, 'aggregation-eligible effect must be VERIFIED')
            END;
            SELECT CASE
                WHEN NOT EXISTS (
                    SELECT 1 FROM knowledge_comparisons c
                    WHERE c.comparison_id=NEW.comparison_id
                      AND c.validity_status='VALID'
                      AND c.confounding_status='NONE'
                      AND c.aggregation_eligible=1
                      AND c.verification_status='VERIFIED'
                )
                THEN RAISE(ABORT, 'aggregation-eligible effect requires a valid unconfounded verified comparison')
            END;
        END;
        """
    )
    _seed_units(conn, now_iso)
    _seed_concepts(conn, now_iso)
    revision_columns = {
        str(row[1])
        for row in conn.execute("PRAGMA table_info(source_revisions)")
    }
    if "capture_v2_revision_id" not in revision_columns:
        conn.execute(
            "ALTER TABLE source_revisions ADD COLUMN capture_v2_revision_id INTEGER"
        )
    if "source_content_status" not in revision_columns:
        conn.execute(
            """
            ALTER TABLE source_revisions
            ADD COLUMN source_content_status TEXT NOT NULL DEFAULT 'CAPTURED'
            """
        )
    series_columns = {
        str(row[1])
        for row in conn.execute(
            "PRAGMA table_info(knowledge_measurement_series)"
        )
    }
    if "axis_source" not in series_columns:
        conn.execute(
            """
            ALTER TABLE knowledge_measurement_series
            ADD COLUMN axis_source TEXT NOT NULL DEFAULT 'ROW_IDENTITY'
                CHECK(axis_source IN ('HEADER','ROW_IDENTITY'))
            """
        )
        conn.execute(
            """
            UPDATE knowledge_measurement_series
            SET verification_status='NEEDS_REVIEW'
            """
        )
        if _table_exists(conn, "knowledge_measurement_points"):
            conn.execute(
                """
                UPDATE knowledge_measurement_points
                SET verification_status='NEEDS_REVIEW'
                """
            )
    point_columns = {
        str(row[1])
        for row in conn.execute(
            "PRAGMA table_info(knowledge_measurement_points)"
        )
    }
    if "replicate_role" not in point_columns:
        conn.execute(
            """
            ALTER TABLE knowledge_measurement_points
            ADD COLUMN replicate_role TEXT NOT NULL DEFAULT 'RAW'
                CHECK(replicate_role IN ('RAW','AGGREGATE'))
            """
        )
    if _table_exists(conn, "capture_v2_revisions"):
        # Backfill already-bridged revisions without recapturing source files.
        # ``capture_status`` is the canonical revision lifecycle, whereas
        # ``source_content_status`` preserves terminal workbook semantics such
        # as EMPTY_WORKBOOK and NO_TABULAR_EVIDENCE.
        conn.execute(
            """
            UPDATE source_revisions
            SET source_content_status = COALESCE(
                (
                    SELECT capture.capture_status
                    FROM capture_v2_revisions AS capture
                    WHERE capture.revision_id =
                          source_revisions.capture_v2_revision_id
                ),
                source_content_status
            )
            WHERE capture_v2_revision_id IS NOT NULL
            """
        )
    _migrate_effect_outcome_uniqueness(conn, now_iso)
    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_source_revisions_capture_v2
        ON source_revisions(capture_v2_revision_id)
        WHERE capture_v2_revision_id IS NOT NULL
        """
    )
    conn.execute(
        "INSERT OR IGNORE INTO schema_migrations(migration_name, applied_at) VALUES (?, ?)",
        (KNOWLEDGE_MIGRATION, now_iso()),
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO schema_migrations(
            migration_name, applied_at
        ) VALUES (?, ?)
        """,
        (MEASUREMENT_SERIES_MIGRATION, now_iso()),
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO schema_migrations(
            migration_name, applied_at
        ) VALUES (?, ?)
        """,
        (MEASUREMENT_SERIES_AXIS_MIGRATION, now_iso()),
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO schema_migrations(
            migration_name, applied_at
        ) VALUES (?, ?)
        """,
        (MEASUREMENT_POINT_REPLICATE_ROLE_MIGRATION, now_iso()),
    )
    from inference_data_ai_concept_curation import (
        ensure_concept_curation_schema,
    )

    ensure_concept_curation_schema(conn, now_iso)


def _migrate_effect_outcome_uniqueness(
    conn: sqlite3.Connection,
    now_iso: Callable[[], str],
) -> None:
    """Allow one deterministic effect type per outcome in a comparison.

    The v1 table accidentally omitted ``outcome_id`` from its natural unique
    key. A comparison with several outcomes could therefore store only one
    MEAN_DIFFERENCE (or one rate effect) even though effect UIDs were already
    outcome-specific. Rebuild the small derived table in place while retaining
    every numeric primary key and stable public identifier.
    """

    row = conn.execute(
        """
        SELECT sql
        FROM sqlite_master
        WHERE type='table' AND name='knowledge_effects'
        """
    ).fetchone()
    if row is None:
        return
    normalized_sql = re.sub(r"\s+", "", str(row[0] or "")).lower()
    desired_key = (
        "unique(comparison_id,outcome_id,effect_type,formula_version)"
    )
    if desired_key not in normalized_sql:
        conn.executescript(
            """
            DROP TRIGGER IF EXISTS trg_effect_aggregation_guard_insert;
            DROP TRIGGER IF EXISTS trg_effect_aggregation_guard_update;
            DROP INDEX IF EXISTS idx_effect_eligibility;

            ALTER TABLE knowledge_effects
            RENAME TO knowledge_effects_v1_constraint;

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
                aggregation_eligible INTEGER NOT NULL DEFAULT 0
                    CHECK(aggregation_eligible IN (0,1)),
                verification_status TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                    CHECK(verification_status IN (
                        'VERIFIED','NEEDS_REVIEW','INVALID','EXCLUDED','STALE'
                    )),
                details_json TEXT NOT NULL DEFAULT '{}',
                FOREIGN KEY(comparison_id)
                    REFERENCES knowledge_comparisons(comparison_id)
                    ON DELETE CASCADE,
                FOREIGN KEY(outcome_id)
                    REFERENCES knowledge_outcomes(outcome_id)
                    ON DELETE CASCADE,
                FOREIGN KEY(unit_id) REFERENCES knowledge_units(unit_id),
                UNIQUE(
                    comparison_id, outcome_id, effect_type, formula_version
                )
            );

            INSERT INTO knowledge_effects(
                effect_id, effect_uid, public_effect_id, comparison_id,
                outcome_id, effect_type, estimate, unit_id, original_unit,
                ci_lower, ci_upper, formula_version, calculation_text,
                direction, aggregation_eligible, verification_status,
                details_json
            )
            SELECT
                effect_id, effect_uid, public_effect_id, comparison_id,
                outcome_id, effect_type, estimate, unit_id, original_unit,
                ci_lower, ci_upper, formula_version, calculation_text,
                direction, aggregation_eligible, verification_status,
                details_json
            FROM knowledge_effects_v1_constraint
            ORDER BY effect_id;

            DROP TABLE knowledge_effects_v1_constraint;
            """
        )
    conn.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_effect_eligibility
        ON knowledge_effects(
            aggregation_eligible, effect_type, verification_status
        );

        CREATE TRIGGER IF NOT EXISTS trg_effect_aggregation_guard_insert
        BEFORE INSERT ON knowledge_effects
        WHEN NEW.aggregation_eligible = 1
        BEGIN
            SELECT CASE
                WHEN NEW.verification_status <> 'VERIFIED'
                THEN RAISE(
                    ABORT,
                    'aggregation-eligible effect must be VERIFIED'
                )
            END;
            SELECT CASE
                WHEN NOT EXISTS (
                    SELECT 1 FROM knowledge_comparisons c
                    WHERE c.comparison_id=NEW.comparison_id
                      AND c.validity_status='VALID'
                      AND c.confounding_status='NONE'
                      AND c.aggregation_eligible=1
                      AND c.verification_status='VERIFIED'
                )
                THEN RAISE(
                    ABORT,
                    'aggregation-eligible effect requires a valid unconfounded verified comparison'
                )
            END;
        END;

        CREATE TRIGGER IF NOT EXISTS trg_effect_aggregation_guard_update
        BEFORE UPDATE OF aggregation_eligible, verification_status,
                         comparison_id
        ON knowledge_effects
        WHEN NEW.aggregation_eligible = 1
        BEGIN
            SELECT CASE
                WHEN NEW.verification_status <> 'VERIFIED'
                THEN RAISE(
                    ABORT,
                    'aggregation-eligible effect must be VERIFIED'
                )
            END;
            SELECT CASE
                WHEN NOT EXISTS (
                    SELECT 1 FROM knowledge_comparisons c
                    WHERE c.comparison_id=NEW.comparison_id
                      AND c.validity_status='VALID'
                      AND c.confounding_status='NONE'
                      AND c.aggregation_eligible=1
                      AND c.verification_status='VERIFIED'
                )
                THEN RAISE(
                    ABORT,
                    'aggregation-eligible effect requires a valid unconfounded verified comparison'
                )
            END;
        END;
        """
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO schema_migrations(
            migration_name, applied_at
        ) VALUES (?, ?)
        """,
        (EFFECT_OUTCOME_UNIQUENESS_MIGRATION, now_iso()),
    )


def _seed_units(conn: sqlite3.Connection, now_iso: Callable[[], str]) -> None:
    rows = [
        ("ppm", "RATE", 1.0, ["PPM"]),
        ("%", "RATE", 1.0, ["percent", "percentage", "퍼센트"]),
        ("%p", "RATE_DIFFERENCE", 1.0, ["percentage point", "percentage-point", "퍼센트포인트"]),
        ("mg", "MASS", 1.0, ["milligram", "밀리그램"]),
        ("kPa", "PRESSURE", 1.0, ["kpa", "킬로파스칼"]),
        ("s", "TIME", 1.0, ["sec", "second", "seconds", "초"]),
        ("°C", "TEMPERATURE", 1.0, ["℃", "degC", "celsius"]),
    ]
    for symbol, quantity, scale, aliases in rows:
        uid = stable_uid("unit", symbol)
        conn.execute(
            """
            INSERT INTO knowledge_units(
                unit_uid, canonical_symbol, quantity_kind, scale_to_base, offset_to_base,
                aliases_json, created_at
            ) VALUES (?, ?, ?, ?, 0, ?, ?)
            ON CONFLICT(canonical_symbol) DO UPDATE SET
                quantity_kind=excluded.quantity_kind,
                aliases_json=excluded.aliases_json
            """,
            (uid, symbol, quantity, scale, _json(aliases, []), now_iso()),
        )


def _seed_concepts(conn: sqlite3.Connection, now_iso: Callable[[], str]) -> None:
    concepts = [
        (
            "COMPONENT_PROCESS",
            "VP+CD assembly",
            "VP와 CD의 조립/접합 공정",
            ["VP+CD assembly", "VP-CD assembly", "VP+CD 조립", "VP CD 조립"],
        ),
        (
            "CHANGED_FACTOR",
            "Bonding amount",
            "접착제 또는 본드 도포량",
            ["bonding amount", "bond amount", "본드량", "접착제 도포량", "VP+CD 본드량"],
        ),
        (
            "CHANGED_FACTOR",
            "Press amount",
            "누름량 또는 가압량",
            ["press amount", "pressing amount", "누름량", "가압량"],
        ),
        (
            "CHANGED_FACTOR",
            "Assembly method",
            "조립 방법 또는 순서",
            ["assembly method", "조립 방식", "조립방법"],
        ),
        (
            "OUTCOME",
            "Function NG rate",
            "기능 불량률",
            ["function ng rate", "function defect rate", "function 불량률", "function ng", "기능 불량률"],
        ),
        (
            "OUTCOME",
            "Process NG rate",
            "공정 불량률",
            ["process ng rate", "process defect rate", "공정 불량률", "공정 ng율"],
        ),
        (
            "OUTCOME",
            "Tension",
            "인장력 또는 장력 측정 결과",
            ["tension", "인장력", "장력"],
        ),
    ]
    for kind, name, description, aliases in concepts:
        normalized_name = normalized_term(name)
        concept_uid = stable_uid("concept", kind, normalized_name)
        conn.execute(
            """
            INSERT INTO knowledge_concepts(
                concept_uid, concept_kind, canonical_name, normalized_name,
                description, lifecycle_status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, 'ACTIVE', ?, ?)
            ON CONFLICT(concept_kind, normalized_name) DO UPDATE SET
                canonical_name=excluded.canonical_name,
                description=excluded.description,
                lifecycle_status='ACTIVE',
                updated_at=excluded.updated_at
            """,
            (concept_uid, kind, name, normalized_name, description, now_iso(), now_iso()),
        )
        concept_id = int(
            conn.execute(
                "SELECT concept_id FROM knowledge_concepts WHERE concept_kind=? AND normalized_name=?",
                (kind, normalized_name),
            ).fetchone()[0]
        )
        for alias in aliases:
            normalized_alias = normalized_term(alias)
            alias_uid = stable_uid("alias", concept_uid, normalized_alias)
            conn.execute(
                """
                INSERT INTO knowledge_concept_aliases(
                    alias_uid, concept_id, alias_text, normalized_alias,
                    language, source, confidence, created_at
                ) VALUES (?, ?, ?, ?, '', 'SEED', 1, ?)
                ON CONFLICT(concept_id, normalized_alias) DO UPDATE SET
                    alias_text=CASE
                        WHEN knowledge_concept_aliases.source='HUMAN_APPROVED'
                        THEN knowledge_concept_aliases.alias_text
                        ELSE excluded.alias_text
                    END
                """,
                (alias_uid, concept_id, alias, normalized_alias, now_iso()),
            )


def _verification_status(value: object) -> str:
    status = str(value or "").strip().upper()
    if status in {"VERIFIED", "EXCLUDED", "FAILED", "STALE"}:
        return status
    return "NEEDS_REVIEW"


def _arm_role(value: object) -> str:
    role = str(value or "").strip().upper()
    aliases = {
        "VARIANT": "TREATMENT",
        "NORMAL": "REFERENCE",
        "BASELINE": "CONTROL",
    }
    role = aliases.get(role, role)
    allowed = {"CONTROL", "COMPARATOR", "TREATMENT", "TEST", "BEFORE", "AFTER", "REFERENCE", "OTHER"}
    return role if role in allowed else "OTHER"


def _unit_id(conn: sqlite3.Connection, symbol: object) -> int | None:
    text = str(symbol or "").strip()
    if not text:
        return None
    row = conn.execute(
        "SELECT unit_id FROM knowledge_units WHERE LOWER(canonical_symbol)=LOWER(?) LIMIT 1",
        (text,),
    ).fetchone()
    if row:
        return int(row[0])
    for unit_id, aliases_json in conn.execute("SELECT unit_id, aliases_json FROM knowledge_units"):
        try:
            aliases = json.loads(aliases_json or "[]")
        except json.JSONDecodeError:
            aliases = []
        if normalized_term(text) in {normalized_term(item) for item in aliases}:
            return int(unit_id)
    return None


def resolve_unit_id(conn: sqlite3.Connection, symbol: object) -> int | None:
    return _unit_id(conn, symbol)


def record_schema_candidate(
    conn: sqlite3.Connection,
    *,
    candidate_kind: str,
    original_value: object,
    suggested_canonical_name: object,
    now_iso: Callable[[], str],
) -> str:
    normalized_value = normalized_term(suggested_canonical_name or original_value)
    if not normalized_value:
        raise ValueError("schema candidate requires a non-empty value")
    candidate_uid = stable_uid("schema-candidate", candidate_kind, normalized_value)
    conn.execute(
        """
        INSERT INTO knowledge_schema_candidates(
            candidate_uid, candidate_kind, normalized_value, original_value,
            suggested_canonical_name, occurrence_count, status,
            first_seen_at, last_seen_at
        ) VALUES (?, ?, ?, ?, ?, 1, 'OPEN', ?, ?)
        ON CONFLICT(candidate_kind, normalized_value) DO UPDATE SET
            original_value=excluded.original_value,
            suggested_canonical_name=CASE
                WHEN excluded.suggested_canonical_name<>'' THEN excluded.suggested_canonical_name
                ELSE knowledge_schema_candidates.suggested_canonical_name
            END,
            occurrence_count=knowledge_schema_candidates.occurrence_count+1,
            last_seen_at=excluded.last_seen_at
        """,
        (
            candidate_uid,
            candidate_kind,
            normalized_value,
            str(original_value or "").strip(),
            str(suggested_canonical_name or "").strip(),
            now_iso(),
            now_iso(),
        ),
    )
    return candidate_uid


def find_concept_id(conn: sqlite3.Connection, text: object, concept_kind: str | None = None) -> int | None:
    needle = normalized_term(text)
    if not needle:
        return None
    sql = """
        SELECT c.concept_id, a.normalized_alias
        FROM knowledge_concept_aliases a
        JOIN knowledge_concepts c ON c.concept_id=a.concept_id
        WHERE c.lifecycle_status='ACTIVE'
    """
    params: tuple[object, ...] = ()
    if concept_kind:
        sql += " AND c.concept_kind=?"
        params = (concept_kind,)
    candidates: list[tuple[int, str]] = []
    for row in conn.execute(sql, params):
        alias = str(row[1] or "")
        if alias and (alias == needle or alias in needle):
            candidates.append((int(row[0]), alias))
    if not candidates:
        return None
    candidates.sort(key=lambda item: len(item[1]), reverse=True)
    return candidates[0][0]


def _upsert_source_revision(
    conn: sqlite3.Connection,
    report: sqlite3.Row | dict[str, Any],
    now_iso: Callable[[], str],
) -> tuple[int, int, str]:
    dataset = str(report["dataset"])
    source_path = str(report["source_path"])
    fingerprint = str(report["workbook_fingerprint"])
    current_workbook_id = report["current_workbook_id"]
    is_current = 1 if current_workbook_id is not None and str(report["overall_status"] or "").upper() != "STALE" else 0
    capture_status = "CAPTURED" if current_workbook_id is not None else "STALE"
    file_name = str(report["file_name"] or "") or re.split(r"[\\/]", source_path)[-1]
    document_uid = stable_uid("document", dataset, source_path)
    conn.execute(
        """
        INSERT INTO source_documents(
            document_uid, dataset, source_path, original_file_name,
            source_kind, lifecycle_status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, 'XLSX', 'ACTIVE', ?, ?)
        ON CONFLICT(dataset, source_path) DO UPDATE SET
            original_file_name=excluded.original_file_name,
            lifecycle_status='ACTIVE',
            updated_at=excluded.updated_at
        """,
        (
            document_uid,
            dataset,
            source_path,
            file_name,
            now_iso(),
            now_iso(),
        ),
    )
    document_id = int(
        conn.execute(
            "SELECT document_id FROM source_documents WHERE dataset=? AND source_path=?",
            (dataset, source_path),
        ).fetchone()[0]
    )
    if is_current:
        conn.execute("UPDATE source_revisions SET is_current=0 WHERE document_id=?", (document_id,))
    revision_uid = stable_uid("revision", document_uid, fingerprint, "legacy-analysis-v1")
    conn.execute(
        """
        INSERT INTO source_revisions(
            revision_uid, document_id, legacy_workbook_id, source_fingerprint,
            fingerprint_kind, content_sha256, size_bytes, mtime_ns,
            extractor_name, extractor_version, capture_contract, capture_status,
            is_current, captured_at
        ) VALUES (?, ?, ?, ?, 'LEGACY_METADATA', '', ?, ?, ?, '1',
                  'universal-grid-v1', ?, ?, ?)
        ON CONFLICT(revision_uid) DO UPDATE SET
            legacy_workbook_id=excluded.legacy_workbook_id,
            size_bytes=excluded.size_bytes,
            mtime_ns=excluded.mtime_ns,
            extractor_name=excluded.extractor_name,
            capture_status=excluded.capture_status,
            is_current=excluded.is_current
        """,
        (
            revision_uid,
            document_id,
            int(current_workbook_id) if current_workbook_id is not None else None,
            fingerprint,
            int(report["size_bytes"] or 0),
            int(report["mtime_ns"] or 0),
            str(report["extractor"] or ""),
            capture_status,
            is_current,
            now_iso(),
        ),
    )
    revision_id = int(conn.execute("SELECT revision_id FROM source_revisions WHERE revision_uid=?", (revision_uid,)).fetchone()[0])
    return document_id, revision_id, revision_uid


def sync_legacy_analysis_report(
    conn: sqlite3.Connection,
    analysis_report_id: int,
    now_iso: Callable[[], str],
) -> dict[str, int | str]:
    """Idempotently project one verified legacy analysis into canonical tables."""

    report = conn.execute(
        """
        SELECT ar.*, w.workbook_id AS current_workbook_id,
               w.file_name, w.size_bytes, w.mtime_ns, w.extractor
        FROM analysis_reports ar
        LEFT JOIN workbooks w ON w.workbook_id=ar.workbook_id
        WHERE ar.analysis_report_id=?
        """,
        (analysis_report_id,),
    ).fetchone()
    if not report:
        raise ValueError(f"analysis report not found: {analysis_report_id}")

    document_id, revision_id, revision_uid = _upsert_source_revision(conn, report, now_iso)
    analysis_uid = stable_uid("analysis", report["dataset"], report["source_path"], report["analysis_key"])
    analysis_verification = _verification_status(report["overall_status"])
    conn.execute(
        """
        INSERT INTO workbook_analyses(
            analysis_uid, public_analysis_id, document_id, revision_id,
            legacy_analysis_report_id, analysis_key, title, analysis_type,
            purpose, scope_text, analysis_status, verification_status,
            decision_text, consolidated_summary, limitations_json,
            analyzer_name, analyzer_version, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                  'legacy-analysis-v1', '1', ?, ?)
        ON CONFLICT(analysis_uid) DO UPDATE SET
            revision_id=excluded.revision_id,
            legacy_analysis_report_id=excluded.legacy_analysis_report_id,
            title=excluded.title,
            analysis_type=excluded.analysis_type,
            purpose=excluded.purpose,
            scope_text=excluded.scope_text,
            analysis_status=excluded.analysis_status,
            verification_status=excluded.verification_status,
            decision_text=excluded.decision_text,
            consolidated_summary=excluded.consolidated_summary,
            limitations_json=excluded.limitations_json,
            updated_at=excluded.updated_at
        """,
        (
            analysis_uid,
            public_id("ANALYSIS", analysis_uid),
            document_id,
            revision_id,
            analysis_report_id,
            str(report["analysis_key"]),
            str(report["title"] or ""),
            str(report["analysis_type"] or ""),
            str(report["purpose"] or ""),
            str(report["scope_text"] or ""),
            str(report["overall_status"] or ""),
            analysis_verification,
            str(report["overall_decision"] or ""),
            str(report["overall_summary"] or ""),
            str(report["limitations_json"] or "[]"),
            str(report["created_at"] or now_iso()),
            now_iso(),
        ),
    )
    workbook_analysis_id = int(
        conn.execute("SELECT workbook_analysis_id FROM workbook_analyses WHERE analysis_uid=?", (analysis_uid,)).fetchone()[0]
    )

    study_ids: dict[int, int] = {}
    study_uids: dict[int, str] = {}
    for review in conn.execute(
        "SELECT * FROM analysis_review_items WHERE analysis_report_id=? ORDER BY sort_order, review_item_id",
        (analysis_report_id,),
    ):
        study_uid = stable_uid("study", analysis_uid, review["review_key"])
        status = _verification_status(review["status"])
        conn.execute(
            """
            INSERT INTO knowledge_studies(
                study_uid, public_data_id, workbook_analysis_id, legacy_review_item_id,
                study_key, title, purpose, hypothesis, objective, design_type,
                comparison_basis, analysis_status, verification_status,
                comparability_status, confounding_status, decision_text,
                summary_text, limitations_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, '', '', ?, ?, ?, ?, ?,
                      'UNASSESSED', 'UNASSESSED', ?, ?, '[]', ?, ?)
            ON CONFLICT(study_uid) DO UPDATE SET
                legacy_review_item_id=excluded.legacy_review_item_id,
                title=excluded.title,
                objective=excluded.objective,
                design_type=excluded.design_type,
                comparison_basis=excluded.comparison_basis,
                analysis_status=excluded.analysis_status,
                verification_status=excluded.verification_status,
                decision_text=excluded.decision_text,
                summary_text=excluded.summary_text,
                updated_at=excluded.updated_at
            """,
            (
                study_uid,
                public_id("DATA", study_uid),
                workbook_analysis_id,
                int(review["review_item_id"]),
                str(review["review_key"]),
                str(review["title"] or ""),
                str(review["objective"] or ""),
                str(review["review_type"] or "UNSPECIFIED"),
                str(review["comparison_basis"] or ""),
                str(review["status"] or ""),
                status,
                str(review["decision_text"] or ""),
                str(review["summary_text"] or ""),
                now_iso(),
                now_iso(),
            ),
        )
        study_id = int(conn.execute("SELECT study_id FROM knowledge_studies WHERE study_uid=?", (study_uid,)).fetchone()[0])
        study_ids[int(review["review_item_id"])] = study_id
        study_uids[int(review["review_item_id"])] = study_uid

    arm_ids: dict[int, int] = {}
    arm_uids: dict[int, str] = {}
    for cohort in conn.execute(
        """
        SELECT c.*, r.analysis_report_id
        FROM analysis_cohorts c
        JOIN analysis_review_items r ON r.review_item_id=c.review_item_id
        WHERE r.analysis_report_id=?
        """,
        (analysis_report_id,),
    ):
        legacy_review_id = int(cohort["review_item_id"])
        study_id = study_ids[legacy_review_id]
        arm_uid = stable_uid("arm", study_uids[legacy_review_id], cohort["cohort_key"])
        conn.execute(
            """
            INSERT INTO knowledge_arms(
                arm_uid, study_id, legacy_cohort_id, arm_key, arm_role,
                label, condition_text, attributes_json, verification_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(arm_uid) DO UPDATE SET
                legacy_cohort_id=excluded.legacy_cohort_id,
                arm_role=excluded.arm_role,
                label=excluded.label,
                condition_text=excluded.condition_text,
                attributes_json=excluded.attributes_json,
                verification_status=excluded.verification_status
            """,
            (
                arm_uid,
                study_id,
                int(cohort["cohort_id"]),
                str(cohort["cohort_key"]),
                _arm_role(cohort["cohort_role"]),
                str(cohort["label"] or ""),
                str(cohort["condition_text"] or ""),
                str(cohort["attributes_json"] or "{}"),
                "NEEDS_REVIEW",
            ),
        )
        arm_id = int(conn.execute("SELECT arm_id FROM knowledge_arms WHERE arm_uid=?", (arm_uid,)).fetchone()[0])
        arm_ids[int(cohort["cohort_id"])] = arm_id
        arm_uids[int(cohort["cohort_id"])] = arm_uid

    outcome_ids: dict[int, int] = {}
    outcome_uids: dict[int, str] = {}
    for metric in conn.execute(
        """
        SELECT m.*, r.review_item_id
        FROM analysis_metrics m
        JOIN analysis_review_items r ON r.review_item_id=m.review_item_id
        WHERE r.analysis_report_id=?
        """,
        (analysis_report_id,),
    ):
        legacy_review_id = int(metric["review_item_id"])
        study_id = study_ids[legacy_review_id]
        outcome_key = str(metric["metric_key"])
        scope_key = str(metric["scope_key"] or "")
        canonical_key = outcome_key if not scope_key else f"{outcome_key}::{scope_key}"
        outcome_uid = stable_uid("outcome", study_uids[legacy_review_id], canonical_key)
        unit_id = _unit_id(conn, metric["unit"])
        concept_id = find_concept_id(conn, metric["label"], "OUTCOME")
        conn.execute(
            """
            INSERT INTO knowledge_outcomes(
                outcome_uid, study_id, legacy_metric_id, outcome_key, concept_id,
                original_label, outcome_domain, metric_type, unit_id, original_unit,
                definition_text, spec_text, verification_status, attributes_json
            ) VALUES (?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(outcome_uid) DO UPDATE SET
                legacy_metric_id=excluded.legacy_metric_id,
                concept_id=excluded.concept_id,
                original_label=excluded.original_label,
                metric_type=excluded.metric_type,
                unit_id=excluded.unit_id,
                original_unit=excluded.original_unit,
                definition_text=excluded.definition_text,
                spec_text=excluded.spec_text,
                verification_status=excluded.verification_status,
                attributes_json=excluded.attributes_json
            """,
            (
                outcome_uid,
                study_id,
                int(metric["metric_id"]),
                canonical_key,
                concept_id,
                str(metric["label"] or ""),
                str(metric["metric_type"] or ""),
                unit_id,
                str(metric["unit"] or ""),
                str(metric["definition_text"] or ""),
                str(metric["spec_text"] or ""),
                _verification_status(metric["status"]),
                _json({"legacyScope": scope_key, "notes": metric["notes_json"]}, {}),
            ),
        )
        outcome_id = int(conn.execute("SELECT outcome_id FROM knowledge_outcomes WHERE outcome_uid=?", (outcome_uid,)).fetchone()[0])
        outcome_ids[int(metric["metric_id"])] = outcome_id
        outcome_uids[int(metric["metric_id"])] = outcome_uid

    for value in conn.execute(
        """
        SELECT v.*, m.review_item_id
        FROM analysis_metric_values v
        JOIN analysis_metrics m ON m.metric_id=v.metric_id
        JOIN analysis_review_items r ON r.review_item_id=m.review_item_id
        WHERE r.analysis_report_id=?
        """,
        (analysis_report_id,),
    ):
        outcome_uid = outcome_uids[int(value["metric_id"])]
        arm_uid = arm_uids[int(value["cohort_id"])]
        observation_uid = stable_uid("observation", outcome_uid, arm_uid, "legacy-primary")
        conn.execute(
            """
            INSERT INTO knowledge_observations(
                observation_uid, outcome_id, arm_id, legacy_metric_value_id,
                observation_key, value_number, value_text, numerator, denominator,
                rate_ppm, min_value, max_value, average_value, sample_size,
                result_status, verification_status, details_json
            ) VALUES (?, ?, ?, ?, 'legacy-primary', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(observation_uid) DO UPDATE SET
                legacy_metric_value_id=excluded.legacy_metric_value_id,
                value_number=excluded.value_number,
                value_text=excluded.value_text,
                numerator=excluded.numerator,
                denominator=excluded.denominator,
                rate_ppm=excluded.rate_ppm,
                min_value=excluded.min_value,
                max_value=excluded.max_value,
                average_value=excluded.average_value,
                sample_size=excluded.sample_size,
                result_status=excluded.result_status,
                verification_status=excluded.verification_status,
                details_json=excluded.details_json
            """,
            (
                observation_uid,
                outcome_ids[int(value["metric_id"])],
                arm_ids[int(value["cohort_id"])],
                int(value["metric_value_id"]),
                value["value_number"],
                str(value["value_text"] or ""),
                value["numerator"],
                value["denominator"],
                value["rate_ppm"],
                value["min_value"],
                value["max_value"],
                value["average_value"],
                value["denominator"],
                str(value["result_status"] or ""),
                "NEEDS_REVIEW",
                str(value["details_json"] or "{}"),
            ),
        )

    comparison_uids: dict[int, str] = {}
    comparison_ids: dict[int, int] = {}
    for comparison in conn.execute(
        """
        SELECT c.*, m.review_item_id
        FROM analysis_comparisons c
        JOIN analysis_metrics m ON m.metric_id=c.metric_id
        JOIN analysis_review_items r ON r.review_item_id=m.review_item_id
        WHERE r.analysis_report_id=?
        """,
        (analysis_report_id,),
    ):
        legacy_review_id = int(comparison["review_item_id"])
        outcome_uid = outcome_uids[int(comparison["metric_id"])]
        canonical_comparison_key = f"{outcome_uid}::{comparison['comparison_key']}"
        comparison_uid = stable_uid(
            "comparison",
            study_uids[legacy_review_id],
            outcome_uid,
            comparison["comparison_key"],
        )
        conn.execute(
            """
            INSERT INTO knowledge_comparisons(
                comparison_uid, public_comparison_id, study_id, legacy_comparison_id,
                comparison_key, compared_arm_id, control_arm_id, design_type,
                validity_status, confounding_status, direction, summary_text,
                aggregation_eligible, verification_status, details_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'LEGACY_UNSPECIFIED',
                      'NEEDS_REVIEW', 'UNASSESSED', ?, ?, 0, 'NEEDS_REVIEW', ?)
            ON CONFLICT(comparison_uid) DO UPDATE SET
                legacy_comparison_id=excluded.legacy_comparison_id,
                compared_arm_id=excluded.compared_arm_id,
                control_arm_id=excluded.control_arm_id,
                direction=excluded.direction,
                summary_text=excluded.summary_text,
                aggregation_eligible=0,
                verification_status='NEEDS_REVIEW',
                details_json=excluded.details_json
            """,
            (
                comparison_uid,
                public_id("CMP", comparison_uid),
                study_ids[legacy_review_id],
                int(comparison["comparison_id"]),
                canonical_comparison_key,
                arm_ids[int(comparison["compared_cohort_id"])],
                arm_ids[int(comparison["control_cohort_id"])],
                str(comparison["direction"] or ""),
                str(comparison["summary_text"] or ""),
                str(comparison["details_json"] or "{}"),
            ),
        )
        canonical_comparison_id = int(
            conn.execute("SELECT comparison_id FROM knowledge_comparisons WHERE comparison_uid=?", (comparison_uid,)).fetchone()[0]
        )
        comparison_uids[int(comparison["comparison_id"])] = comparison_uid
        comparison_ids[int(comparison["comparison_id"])] = canonical_comparison_id
        effect_specs = [
            ("ABSOLUTE_DIFFERENCE", comparison["delta_value"], str(comparison["delta_unit"] or "")),
            ("RELATIVE_CHANGE_PERCENT", comparison["relative_delta_percent"], "%"),
        ]
        for effect_type, estimate, unit_symbol in effect_specs:
            if estimate is None:
                continue
            effect_uid = stable_uid("effect", comparison_uid, effect_type, "legacy-v1")
            conn.execute(
                """
                INSERT INTO knowledge_effects(
                    effect_uid, public_effect_id, comparison_id, outcome_id,
                    effect_type, estimate, unit_id, original_unit, formula_version,
                    calculation_text, direction, aggregation_eligible,
                    verification_status, details_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'legacy-v1', ?, ?, 0, 'NEEDS_REVIEW', '{}')
                ON CONFLICT(effect_uid) DO UPDATE SET
                    comparison_id=excluded.comparison_id,
                    outcome_id=excluded.outcome_id,
                    estimate=excluded.estimate,
                    unit_id=excluded.unit_id,
                    original_unit=excluded.original_unit,
                    calculation_text=excluded.calculation_text,
                    direction=excluded.direction,
                    aggregation_eligible=0,
                    verification_status='NEEDS_REVIEW'
                """,
                (
                    effect_uid,
                    public_id("EFF", effect_uid),
                    canonical_comparison_id,
                    outcome_ids[int(comparison["metric_id"])],
                    effect_type,
                    estimate,
                    _unit_id(conn, unit_symbol),
                    unit_symbol,
                    str(comparison["calculation_text"] or ""),
                    str(comparison["direction"] or ""),
                ),
            )

    claim_uids: dict[int, str] = {}
    for conclusion in conn.execute(
        "SELECT * FROM analysis_conclusions WHERE analysis_report_id=?",
        (analysis_report_id,),
    ):
        legacy_review_id = conclusion["review_item_id"]
        study_id = study_ids.get(int(legacy_review_id)) if legacy_review_id is not None else None
        owner_uid = study_uids.get(int(legacy_review_id), analysis_uid) if legacy_review_id is not None else analysis_uid
        claim_uid = stable_uid("claim", owner_uid, conclusion["conclusion_key"])
        conn.execute(
            """
            INSERT INTO knowledge_claims(
                claim_uid, public_claim_id, workbook_analysis_id, study_id,
                legacy_conclusion_id, claim_key, claim_type, claim_text,
                verdict, causal_strength, verification_status, limitations_json
            ) VALUES (?, ?, ?, ?, ?, ?, 'SOURCE_CONCLUSION', ?, ?,
                      'UNSPECIFIED', ?, ?)
            ON CONFLICT(claim_uid) DO UPDATE SET
                legacy_conclusion_id=excluded.legacy_conclusion_id,
                claim_text=excluded.claim_text,
                verdict=excluded.verdict,
                verification_status=excluded.verification_status,
                limitations_json=excluded.limitations_json
            """,
            (
                claim_uid,
                public_id("CLM", claim_uid),
                workbook_analysis_id,
                study_id,
                int(conclusion["conclusion_id"]),
                str(conclusion["conclusion_key"]),
                str(conclusion["conclusion_text"]),
                str(conclusion["verdict"] or ""),
                analysis_verification,
                str(conclusion["limitations_json"] or "[]"),
            ),
        )
        claim_uids[int(conclusion["conclusion_id"])] = claim_uid

    evidence_count = _sync_legacy_evidence(
        conn,
        analysis_report_id,
        revision_id,
        revision_uid,
        analysis_uid,
        study_uids,
        outcome_uids,
        comparison_uids,
        claim_uids,
        now_iso,
    )
    return {
        "analysisUid": analysis_uid,
        "workbookAnalysisId": workbook_analysis_id,
        "studies": len(study_ids),
        "arms": len(arm_ids),
        "outcomes": len(outcome_ids),
        "comparisons": len(comparison_ids),
        "evidence": evidence_count,
    }


def _sync_legacy_evidence(
    conn: sqlite3.Connection,
    analysis_report_id: int,
    revision_id: int,
    revision_uid: str,
    analysis_uid: str,
    study_uids: dict[int, str],
    outcome_uids: dict[int, str],
    comparison_uids: dict[int, str],
    claim_uids: dict[int, str],
    now_iso: Callable[[], str],
) -> int:
    count = 0
    for evidence in conn.execute(
        "SELECT * FROM analysis_evidence WHERE analysis_report_id=? ORDER BY evidence_id",
        (analysis_report_id,),
    ):
        links: list[tuple[str, str]] = []
        owner_parts: list[object] = [analysis_uid]
        if evidence["review_item_id"] is not None:
            uid = study_uids.get(int(evidence["review_item_id"]))
            if uid:
                links.append(("STUDY", uid))
                owner_parts.append(uid)
        else:
            links.append(("WORKBOOK_ANALYSIS", analysis_uid))
        if evidence["metric_id"] is not None:
            uid = outcome_uids.get(int(evidence["metric_id"]))
            if uid:
                links.append(("OUTCOME", uid))
                owner_parts.append(uid)
        if evidence["comparison_id"] is not None:
            uid = comparison_uids.get(int(evidence["comparison_id"]))
            if uid:
                links.append(("COMPARISON", uid))
                owner_parts.append(uid)
        if evidence["conclusion_id"] is not None:
            uid = claim_uids.get(int(evidence["conclusion_id"]))
            if uid:
                links.append(("CLAIM", uid))
                owner_parts.append(uid)
        evidence_uid = stable_uid(
            "evidence",
            revision_uid,
            *owner_parts,
            evidence["sheet_name"],
            evidence["range_address"],
            evidence["evidence_role"],
        )
        conn.execute(
            """
            INSERT INTO evidence_items(
                evidence_uid, public_evidence_id, revision_id, legacy_evidence_id,
                evidence_kind, sheet_name, start_row, start_col, end_row, end_col,
                range_address, evidence_role, source_text, note,
                verification_status, created_at
            ) VALUES (?, ?, ?, ?, 'CELL_RANGE', ?, ?, ?, ?, ?, ?, ?, ?, ?, 'VERIFIED', ?)
            ON CONFLICT(evidence_uid) DO UPDATE SET
                legacy_evidence_id=excluded.legacy_evidence_id,
                source_text=excluded.source_text,
                note=excluded.note,
                verification_status='VERIFIED'
            """,
            (
                evidence_uid,
                public_id("EVD", evidence_uid),
                revision_id,
                int(evidence["evidence_id"]),
                str(evidence["sheet_name"]),
                int(evidence["start_row"]),
                int(evidence["start_col"]),
                int(evidence["end_row"]),
                int(evidence["end_col"]),
                str(evidence["range_address"]),
                str(evidence["evidence_role"] or "SOURCE"),
                str(evidence["source_text"] or ""),
                str(evidence["note"] or ""),
                now_iso(),
            ),
        )
        canonical_evidence_id = int(conn.execute("SELECT evidence_id FROM evidence_items WHERE evidence_uid=?", (evidence_uid,)).fetchone()[0])
        for entity_type, entity_uid in links:
            conn.execute(
                """
                INSERT OR IGNORE INTO entity_evidence_links(
                    entity_type, entity_uid, evidence_id, evidence_role, claim_scope
                ) VALUES (?, ?, ?, ?, '')
                """,
                (entity_type, entity_uid, canonical_evidence_id, str(evidence["evidence_role"] or "SOURCE")),
            )
        count += 1
    return count


def migrate_all_legacy_analyses(conn: sqlite3.Connection, now_iso: Callable[[], str]) -> dict[str, Any]:
    ensure_knowledge_schema(conn, now_iso)
    if not _table_exists(conn, "analysis_reports"):
        return {"reports": 0, "studies": 0, "evidence": 0}
    report_ids = [int(row[0]) for row in conn.execute("SELECT analysis_report_id FROM analysis_reports ORDER BY analysis_report_id")]
    results = [sync_legacy_analysis_report(conn, report_id, now_iso) for report_id in report_ids]
    return {
        "reports": len(results),
        "studies": sum(int(result["studies"]) for result in results),
        "arms": sum(int(result["arms"]) for result in results),
        "outcomes": sum(int(result["outcomes"]) for result in results),
        "comparisons": sum(int(result["comparisons"]) for result in results),
        "evidence": sum(int(result["evidence"]) for result in results),
    }


def knowledge_counts(conn: sqlite3.Connection) -> dict[str, int]:
    tables = [
        "source_documents",
        "source_revisions",
        "workbook_analyses",
        "knowledge_studies",
        "knowledge_study_contexts",
        "knowledge_schema_candidates",
        "knowledge_concept_resolution_history",
        "knowledge_concept_alias_approval_history",
        "knowledge_factors",
        "knowledge_arms",
        "knowledge_outcomes",
        "knowledge_observations",
        "knowledge_measurement_series",
        "knowledge_measurement_points",
        "knowledge_comparisons",
        "knowledge_effects",
        "knowledge_claims",
        "evidence_items",
        "entity_evidence_links",
        "validation_issues",
        "review_decisions",
    ]
    return {
        table: int(conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
        for table in tables
        if _table_exists(conn, table)
    }


def validate_knowledge_integrity(conn: sqlite3.Connection) -> dict[str, Any]:
    all_foreign_key_errors = [tuple(row) for row in conn.execute("PRAGMA foreign_key_check")]
    canonical_tables = {
        "source_documents",
        "source_revisions",
        "knowledge_concepts",
        "knowledge_concept_aliases",
        "knowledge_units",
        "knowledge_schema_candidates",
        "knowledge_concept_resolution_history",
        "knowledge_concept_alias_approval_history",
        "workbook_analyses",
        "knowledge_studies",
        "knowledge_study_contexts",
        "knowledge_factors",
        "knowledge_arms",
        "knowledge_arm_factor_values",
        "knowledge_outcomes",
        "knowledge_observations",
        "knowledge_measurement_series",
        "knowledge_measurement_points",
        "knowledge_comparisons",
        "knowledge_effects",
        "evidence_items",
        "entity_evidence_links",
        "knowledge_claims",
        "validation_issues",
        "review_decisions",
    }
    foreign_key_errors = [row for row in all_foreign_key_errors if str(row[0]) in canonical_tables]
    legacy_foreign_key_warnings = [row for row in all_foreign_key_errors if str(row[0]) not in canonical_tables]
    invalid_effects = int(
        conn.execute(
            """
            SELECT COUNT(*)
            FROM knowledge_effects e
            JOIN knowledge_comparisons c ON c.comparison_id=e.comparison_id
            WHERE e.aggregation_eligible=1
              AND (
                  e.verification_status<>'VERIFIED'
                  OR c.validity_status<>'VALID'
                  OR c.confounding_status<>'NONE'
                  OR c.aggregation_eligible<>1
                  OR c.verification_status<>'VERIFIED'
              )
            """
        ).fetchone()[0]
    )
    orphan_links = int(
        conn.execute(
            """
            SELECT COUNT(*)
            FROM entity_evidence_links l
            LEFT JOIN evidence_items e ON e.evidence_id=l.evidence_id
            WHERE e.evidence_id IS NULL
            """
        ).fetchone()[0]
    )
    result = {
        "ok": not foreign_key_errors and invalid_effects == 0 and orphan_links == 0,
        "foreignKeyErrors": foreign_key_errors,
        "legacyForeignKeyWarnings": legacy_foreign_key_warnings,
        "invalidAggregationEffects": invalid_effects,
        "orphanEvidenceLinks": orphan_links,
        "counts": knowledge_counts(conn),
    }
    return result


def validate_analysis_integrity(
    conn: sqlite3.Connection,
    *,
    workbook_analysis_id: int,
) -> dict[str, Any]:
    """Validate one imported analysis without scanning the whole database.

    The importing connection must enforce SQLite foreign keys, so inserted and
    updated rows have already passed their declared constraints. A corpus run
    still performs ``validate_knowledge_integrity`` once after all workbook
    imports instead of repeating that multi-gigabyte scan per workbook.
    """

    foreign_keys_enabled = int(
        conn.execute("PRAGMA foreign_keys").fetchone()[0]
    )
    analysis = conn.execute(
        """
        SELECT workbook_analysis_id, revision_id
        FROM workbook_analyses
        WHERE workbook_analysis_id=?
        """,
        (workbook_analysis_id,),
    ).fetchone()
    missing_analysis = analysis is None
    invalid_effects = int(
        conn.execute(
            """
            SELECT COUNT(*)
            FROM knowledge_effects e
            JOIN knowledge_comparisons c
              ON c.comparison_id=e.comparison_id
            JOIN knowledge_studies s
              ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
              AND e.aggregation_eligible=1
              AND (
                  e.verification_status<>'VERIFIED'
                  OR c.validity_status<>'VALID'
                  OR c.confounding_status<>'NONE'
                  OR c.aggregation_eligible<>1
                  OR c.verification_status<>'VERIFIED'
              )
            """,
            (workbook_analysis_id,),
        ).fetchone()[0]
    )
    foreign_key_errors: list[tuple[Any, ...]] = []
    if not foreign_keys_enabled:
        foreign_key_errors.append(
            (
                "workbook_analyses",
                workbook_analysis_id,
                "foreign_keys_disabled",
            )
        )
    if missing_analysis:
        foreign_key_errors.append(
            (
                "workbook_analyses",
                workbook_analysis_id,
                "missing_analysis",
            )
        )

    counts: dict[str, int] = {
        "workbook_analyses": 0 if missing_analysis else 1,
    }
    if not missing_analysis:
        scoped_queries = {
            "knowledge_studies": """
                SELECT COUNT(*)
                FROM knowledge_studies
                WHERE workbook_analysis_id=?
            """,
            "knowledge_study_contexts": """
                SELECT COUNT(*)
                FROM knowledge_study_contexts c
                JOIN knowledge_studies s ON s.study_id=c.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_factors": """
                SELECT COUNT(*)
                FROM knowledge_factors f
                JOIN knowledge_studies s ON s.study_id=f.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_arms": """
                SELECT COUNT(*)
                FROM knowledge_arms a
                JOIN knowledge_studies s ON s.study_id=a.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_arm_factor_values": """
                SELECT COUNT(*)
                FROM knowledge_arm_factor_values af
                JOIN knowledge_arms a ON a.arm_id=af.arm_id
                JOIN knowledge_studies s ON s.study_id=a.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_outcomes": """
                SELECT COUNT(*)
                FROM knowledge_outcomes o
                JOIN knowledge_studies s ON s.study_id=o.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_observations": """
                SELECT COUNT(*)
                FROM knowledge_observations o
                JOIN knowledge_outcomes k ON k.outcome_id=o.outcome_id
                JOIN knowledge_studies s ON s.study_id=k.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_measurement_series": """
                SELECT COUNT(*)
                FROM knowledge_measurement_series ms
                JOIN knowledge_studies s ON s.study_id=ms.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_measurement_points": """
                SELECT COUNT(*)
                FROM knowledge_measurement_points mp
                JOIN knowledge_measurement_series ms
                  ON ms.series_id=mp.series_id
                JOIN knowledge_studies s ON s.study_id=ms.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_comparisons": """
                SELECT COUNT(*)
                FROM knowledge_comparisons c
                JOIN knowledge_studies s ON s.study_id=c.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_effects": """
                SELECT COUNT(*)
                FROM knowledge_effects e
                JOIN knowledge_comparisons c
                  ON c.comparison_id=e.comparison_id
                JOIN knowledge_studies s ON s.study_id=c.study_id
                WHERE s.workbook_analysis_id=?
            """,
            "knowledge_claims": """
                SELECT COUNT(*)
                FROM knowledge_claims
                WHERE workbook_analysis_id=?
            """,
            "evidence_items": """
                SELECT COUNT(*)
                FROM evidence_items
                WHERE revision_id=(
                    SELECT revision_id
                    FROM workbook_analyses
                    WHERE workbook_analysis_id=?
                )
            """,
        }
        counts.update(
            {
                table: int(
                    conn.execute(
                        sql,
                        (workbook_analysis_id,),
                    ).fetchone()[0]
                )
                for table, sql in scoped_queries.items()
            }
        )

    return {
        "ok": not foreign_key_errors and invalid_effects == 0,
        "scope": "WORKBOOK_ANALYSIS",
        "workbookAnalysisId": workbook_analysis_id,
        "foreignKeyErrors": foreign_key_errors,
        "legacyForeignKeyWarnings": [],
        "invalidAggregationEffects": invalid_effects,
        "orphanEvidenceLinks": 0,
        "counts": counts,
    }
