from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import json
import re
import sqlite3
import subprocess
import sys
import time
from collections import Counter, deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator

try:
    from inference_data_ai_schema import (
        ensure_knowledge_schema,
        knowledge_counts,
        migrate_all_legacy_analyses,
        sync_legacy_analysis_report,
        validate_knowledge_integrity,
    )
except ModuleNotFoundError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from inference_data_ai_schema import (
        ensure_knowledge_schema,
        knowledge_counts,
        migrate_all_legacy_analyses,
        sync_legacy_analysis_report,
        validate_knowledge_integrity,
    )

from inference_data_ai_source_ingest import (
    bridge_capture_to_canonical_source,
    capture_json_bytes,
    ensure_capture_v2_schema,
    extract_workbook as extract_openxml_workbook,
    import_capture as import_openxml_capture,
    verify_capture_database,
)
from inference_data_ai_study_import import import_study_manifest
from inference_data_ai_semantic_ai import (
    run_codex_locator,
    run_codex_locator_batch,
    run_codex_study_draft,
    validate_ai_study_draft,
    validate_locator_result,
)
from inference_data_ai_semantic_packets import (
    build_semantic_source_packets_from_db,
    packet_json_bytes as semantic_packet_json_bytes,
)
from inference_data_ai_table_first import (
    ANALYSIS_SCHEMA_VERSION as TABLE_FIRST_ANALYSIS_SCHEMA_VERSION,
    BUILDER_VERSION as TABLE_FIRST_BUILDER_VERSION,
    PROMPT_VERSION as TABLE_FIRST_PROMPT_VERSION,
    REQUEST_SCHEMA_VERSION as TABLE_FIRST_REQUEST_SCHEMA_VERSION,
    build_table_first_request,
    normalize_table_first_analysis,
    project_table_first_analysis,
    run_codex_table_first_analysis,
    table_first_prompt_stats,
    table_first_json_bytes,
    validate_table_first_analysis,
)
from inference_data_ai_table_first_html import build_table_first_html_report
from inference_data_ai_table_first_history import (
    build_history_answer,
    build_history_detail,
    build_history_index,
    build_history_pack,
    history_json_bytes,
    render_history_answer_markdown,
    run_history_acceptance,
    validate_history_answer,
)
from inference_data_ai_contextual_query import (
    build_contextual_query_request,
    contextual_json_bytes,
    render_contextual_answer_markdown,
    run_codex_contextual_query,
)
from inference_data_ai_relevance_query import (
    build_relevance_query_request,
    relevance_json_bytes,
    render_relevance_result_markdown,
    run_codex_relevance_query,
)
from inference_data_ai_study_contract import validate_study_manifest
from inference_data_ai_query import build_evidence_pack_from_db
from inference_data_ai_answer import (
    answer_json_bytes,
    build_evidence_answer,
    render_answer_markdown,
    validate_evidence_answer,
)
from inference_data_ai_evidence_detail import build_evidence_detail_from_db
from inference_data_ai_related import (
    build_related_studies_from_db,
    related_studies_json_bytes,
)
from inference_data_ai_concept_curation import (
    ConceptCurationError,
    list_canonical_concepts,
    list_schema_candidates,
    resolve_schema_candidate,
    upsert_human_concept_alias,
)
from inference_data_ai_review import (
    ReviewGateError,
    decide_comparison,
    get_review_detail,
    list_review_queue,
)
from inference_data_ai_acceptance import run_golden_question_acceptance
from inference_data_ai_workflow import ingest_workbook
from inference_data_ai_corpus_workflow import run_corpus_ingest
from inference_data_ai_form_preflight import run_form_preflight
from inference_data_ai_form_registry import (
    analyze_form_family,
    decide_form_family,
    reclassify_form_preflight_report,
    write_form_group_review,
)
from inference_data_ai_form_pipeline import run_form_pipeline_complete
from inference_data_ai_study_import import (
    AnalysisQuarantineError,
    make_database_evidence_checker,
    quarantine_canonical_analysis,
    resolve_manifest_revision,
    validate_comparison_representation_alignment,
    validate_conclusion_evidence,
    validate_factor_and_arm_evidence,
    validate_numeric_observation_evidence,
)


try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


SERVICE_DIR = Path(__file__).resolve().parent
REPO_ROOT = SERVICE_DIR.parents[1]
DEFAULT_INPUT_DIR = Path(r"D:\000. MyWorks\test\result\InputDataFinish")
OUTPUT_DIR = SERVICE_DIR / "outputs"
DEFAULT_DATASET = "InputDataFinish"
QUICK_INDEX_DIR = OUTPUT_DIR / "quick-index"
UNIVERSAL_GRID_DIR = OUTPUT_DIR / "universal-grid"
PACKET_DIR = OUTPUT_DIR / "packets"
LOG_DIR = OUTPUT_DIR / "logs"
CAPTURE_V2_DIR = OUTPUT_DIR / "capture-v2"
SEMANTIC_PACKET_DIR = OUTPUT_DIR / "semantic-source-packets"
SEMANTIC_LOCATOR_DIR = OUTPUT_DIR / "semantic-locators"
SEMANTIC_STUDY_DIR = OUTPUT_DIR / "semantic-study-drafts"
TABLE_FIRST_DIR = OUTPUT_DIR / "table-first"
MICROSPEAKER_INDEXER = REPO_ROOT / "MicroSpeaker_ProductTech_DB" / "tools" / "incremental_dataset_indexer.py"
COM_EXTRACTOR = REPO_ROOT / "JinoSupporter" / "JinoSupporter.Web" / "tools" / "input_data_excel_com_extract.py"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm"}

CONTEXT_TERMS = (
    "title",
    "purpose",
    "objective",
    "content",
    "condition",
    "standard",
    "spec",
    "result",
    "decision",
    "conclusion",
    "remark",
    "note",
    "problem",
    "reason",
    "before",
    "after",
    "normal",
    "test",
    "sample",
    "lot",
    "supplier",
    "material",
    "machine",
    "m/c",
    "jig",
    "base",
    "laser",
    "coating",
    "gauss",
    "tension",
    "function",
    "vision",
    "repair",
)

# These words only decide which source rows are most useful to include when a
# workbook is larger than the analysis-packet budget.  They never create a
# conclusion or replace the coordinate-based source evidence.
PACKET_PRIORITY_TERMS = CONTEXT_TERMS + (
    "결론",
    "판정",
    "결과",
    "목적",
    "조건",
    "기준",
    "사양",
    "비고",
    "정상",
    "불량",
    "개선",
    "시험",
    "검증",
)


def now_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "_", value.strip())
    return cleaned.strip("._") or "dataset"


def ensure_output_parent(path: Path) -> None:
    resolved = path.resolve()
    try:
        resolved.relative_to(SERVICE_DIR.resolve())
    except ValueError as exc:
        raise SystemExit(f"Output path must stay under {SERVICE_DIR}: {resolved}") from exc
    resolved.parent.mkdir(parents=True, exist_ok=True)


def service_output_path(value: str | None, default_path: Path) -> Path:
    path = Path(value) if value else default_path
    if not path.is_absolute():
        path = SERVICE_DIR / path
    path = path.resolve()
    ensure_output_parent(path)
    return path


def output_path_under_root(
    value: str | None,
    default_path: Path,
    allowed_root: str | Path,
) -> Path:
    root = Path(allowed_root).expanduser().resolve()
    path = Path(value) if value else default_path
    if not path.is_absolute():
        path = root / path
    resolved = path.expanduser().resolve()
    try:
        resolved.relative_to(root)
    except ValueError as exc:
        raise SystemExit(
            f"Output path must stay under {root}: {resolved}"
        ) from exc
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def database_scoped_output_path(
    value: str | None,
    default_path: Path,
    database_path: str | Path,
) -> Path:
    """Keep DB query artifacts beside the configured DB output tree."""

    if value is None or not Path(value).expanduser().is_absolute():
        return service_output_path(value, default_path)
    database = Path(database_path).expanduser().resolve()
    database_parent = database.parent
    if database_parent.name.casefold() in {
        "universal-grid",
        "table-first-history",
    }:
        allowed_root = database_parent.parent
    else:
        allowed_root = database_parent
    resolved = Path(value).expanduser().resolve()
    try:
        resolved.relative_to(SERVICE_DIR.resolve())
    except ValueError:
        return output_path_under_root(
            str(resolved),
            default_path,
            allowed_root,
        )
    return service_output_path(str(resolved), default_path)


def service_output_dir(value: str | None, default_dir: Path) -> Path:
    path = Path(value) if value else default_dir
    if not path.is_absolute():
        path = SERVICE_DIR / path
    path = path.resolve()
    try:
        path.relative_to(SERVICE_DIR.resolve())
    except ValueError as exc:
        raise SystemExit(f"Output directory must stay under {SERVICE_DIR}: {path}") from exc
    path.mkdir(parents=True, exist_ok=True)
    return path


def write_json(path: Path, data: Any) -> None:
    ensure_output_parent(path)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def print_json(data: Any) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def value_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


def compact(value: Any, max_len: int = 320) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text if len(text) <= max_len else text[: max_len - 3].rstrip() + "..."


def file_fingerprint(path: Path) -> str:
    try:
        stat = path.stat()
        raw = f"{path.resolve()}|{stat.st_size}|{stat.st_mtime_ns}"
    except OSError:
        raw = str(path.resolve())
    return hashlib.sha1(raw.encode("utf-8", errors="replace")).hexdigest()


def collect_excel_files(input_path: Path, limit: int = 0) -> list[Path]:
    if input_path.is_file():
        files = [input_path] if input_path.suffix.lower() in EXCEL_EXTENSIONS else []
    elif input_path.is_dir():
        files = sorted(
            path
            for path in input_path.rglob("*")
            if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS and not path.name.startswith("~$")
        )
    else:
        raise SystemExit(f"Input path does not exist: {input_path}")

    if limit > 0:
        files = files[:limit]
    return files


def run_command(cmd: list[str], log_path: Path, cwd: Path | None = None) -> int:
    ensure_output_parent(log_path)
    with log_path.open("a", encoding="utf-8") as log:
        log.write(f"\n[{now_iso()}] {' '.join(cmd)}\n")
        process = subprocess.Popen(
            cmd,
            cwd=str(cwd or REPO_ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        assert process.stdout is not None
        for line in process.stdout:
            print(line, end="")
            log.write(line)
        process.wait()
        log.write(f"[exit] {process.returncode}\n")
        return int(process.returncode)


@contextmanager
def connect_rw(db_path: Path) -> Iterator[sqlite3.Connection]:
    ensure_output_parent(db_path)
    conn = sqlite3.connect(str(db_path), timeout=60)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=60000")
    try:
        yield conn
    finally:
        conn.close()


@contextmanager
def connect_ro(db_path: Path) -> Iterator[sqlite3.Connection]:
    if not db_path.exists():
        raise SystemExit(f"DB not found: {db_path}")
    uri = "file:" + db_path.as_posix().replace("?", "%3f") + "?mode=ro"
    conn = sqlite3.connect(uri, uri=True, timeout=60)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=60000")
    try:
        yield conn
    finally:
        conn.close()


def table_exists(conn: sqlite3.Connection, table: str) -> bool:
    return (
        conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type IN ('table','view') AND name=? LIMIT 1",
            (table,),
        ).fetchone()
        is not None
    )


def dict_rows(conn: sqlite3.Connection, sql: str, params: Iterable[Any] = ()) -> list[dict[str, Any]]:
    return [dict(row) for row in conn.execute(sql, tuple(params))]


def first_dict(conn: sqlite3.Connection, sql: str, params: Iterable[Any] = ()) -> dict[str, Any] | None:
    row = conn.execute(sql, tuple(params)).fetchone()
    return dict(row) if row else None


def table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    return {str(row["name"]) for row in conn.execute(f"PRAGMA table_info({table})")}


def ensure_column(conn: sqlite3.Connection, table: str, column: str, definition: str) -> None:
    if column not in table_columns(conn, table):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def ensure_universal_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS runs (
            run_id INTEGER PRIMARY KEY AUTOINCREMENT,
            dataset TEXT NOT NULL,
            input_path TEXT NOT NULL,
            started_at TEXT NOT NULL,
            finished_at TEXT NOT NULL DEFAULT '',
            total_files INTEGER NOT NULL DEFAULT 0,
            succeeded INTEGER NOT NULL DEFAULT 0,
            failed INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS workbooks (
            workbook_id INTEGER PRIMARY KEY AUTOINCREMENT,
            dataset TEXT NOT NULL,
            source_path TEXT NOT NULL,
            file_name TEXT NOT NULL,
            extension TEXT NOT NULL,
            size_bytes INTEGER NOT NULL DEFAULT 0,
            mtime_ns INTEGER NOT NULL DEFAULT 0,
            fingerprint TEXT NOT NULL,
            extractor TEXT NOT NULL,
            status TEXT NOT NULL,
            error TEXT NOT NULL DEFAULT '',
            sheet_count INTEGER NOT NULL DEFAULT 0,
            total_rows INTEGER NOT NULL DEFAULT 0,
            total_cells INTEGER NOT NULL DEFAULT 0,
            non_empty_cells INTEGER NOT NULL DEFAULT 0,
            merge_count INTEGER NOT NULL DEFAULT 0,
            raw_json_path TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            UNIQUE(dataset, source_path)
        );

        CREATE TABLE IF NOT EXISTS worksheets (
            workbook_id INTEGER NOT NULL,
            sheet_index INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            visible INTEGER NOT NULL DEFAULT 1,
            used_top INTEGER NOT NULL DEFAULT 0,
            used_left INTEGER NOT NULL DEFAULT 0,
            used_bottom INTEGER NOT NULL DEFAULT 0,
            used_right INTEGER NOT NULL DEFAULT 0,
            row_count INTEGER NOT NULL DEFAULT 0,
            col_count INTEGER NOT NULL DEFAULT 0,
            non_empty_cells INTEGER NOT NULL DEFAULT 0,
            merge_count INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(workbook_id, sheet_index)
        );

        CREATE TABLE IF NOT EXISTS grid_sheet_rows (
            workbook_id INTEGER NOT NULL,
            sheet_index INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            row_number INTEGER NOT NULL,
            non_empty_count INTEGER NOT NULL DEFAULT 0,
            row_text TEXT NOT NULL DEFAULT '',
            cells_json TEXT NOT NULL DEFAULT '[]',
            PRIMARY KEY(workbook_id, sheet_index, row_number)
        );

        CREATE TABLE IF NOT EXISTS grid_sheet_cells (
            workbook_id INTEGER NOT NULL,
            sheet_index INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            row_number INTEGER NOT NULL,
            col_number INTEGER NOT NULL,
            col_label TEXT NOT NULL,
            address TEXT NOT NULL,
            value_text TEXT NOT NULL DEFAULT '',
            raw_value_text TEXT NOT NULL DEFAULT '',
            merge_role TEXT NOT NULL DEFAULT 'none',
            merge_address TEXT NOT NULL DEFAULT '',
            anchor_row INTEGER,
            anchor_col INTEGER,
            PRIMARY KEY(workbook_id, sheet_index, row_number, col_number)
        );

        CREATE TABLE IF NOT EXISTS merge_ranges (
            workbook_id INTEGER NOT NULL,
            sheet_index INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            address TEXT NOT NULL,
            top INTEGER NOT NULL,
            left_col INTEGER NOT NULL,
            bottom INTEGER NOT NULL,
            right_col INTEGER NOT NULL,
            row_span INTEGER NOT NULL,
            column_span INTEGER NOT NULL,
            anchor_value TEXT NOT NULL DEFAULT '',
            PRIMARY KEY(workbook_id, sheet_index, address)
        );

        CREATE TABLE IF NOT EXISTS ingest_items (
            ingest_item_id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL,
            dataset TEXT NOT NULL,
            source_path TEXT NOT NULL,
            source_fingerprint TEXT NOT NULL,
            raw_json_path TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL,
            workbook_id INTEGER,
            message TEXT NOT NULL DEFAULT '',
            started_at TEXT NOT NULL,
            finished_at TEXT NOT NULL DEFAULT '',
            UNIQUE(run_id, source_path)
        );

        CREATE TABLE IF NOT EXISTS schema_migrations (
            migration_name TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL
        );

        -- The universal-grid tables above preserve the source workbook.  The
        -- analysis tables below preserve a reusable, evidence-linked reading
        -- of that source without treating an HTML dashboard as the authority.
        CREATE TABLE IF NOT EXISTS analysis_reports (
            analysis_report_id INTEGER PRIMARY KEY AUTOINCREMENT,
            dataset TEXT NOT NULL,
            workbook_id INTEGER NOT NULL,
            source_path TEXT NOT NULL,
            workbook_fingerprint TEXT NOT NULL,
            analysis_key TEXT NOT NULL,
            title TEXT NOT NULL DEFAULT '',
            analysis_type TEXT NOT NULL DEFAULT '',
            purpose TEXT NOT NULL DEFAULT '',
            scope_text TEXT NOT NULL DEFAULT '',
            overall_status TEXT NOT NULL DEFAULT 'DRAFT',
            overall_decision TEXT NOT NULL DEFAULT '',
            overall_summary TEXT NOT NULL DEFAULT '',
            limitations_json TEXT NOT NULL DEFAULT '[]',
            dashboard_html_path TEXT NOT NULL DEFAULT '',
            dashboard_markdown_path TEXT NOT NULL DEFAULT '',
            manifest_path TEXT NOT NULL DEFAULT '',
            stale_reason TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(dataset, source_path, workbook_fingerprint, analysis_key)
        );

        CREATE TABLE IF NOT EXISTS analysis_review_items (
            review_item_id INTEGER PRIMARY KEY AUTOINCREMENT,
            analysis_report_id INTEGER NOT NULL,
            review_key TEXT NOT NULL,
            sort_order INTEGER NOT NULL DEFAULT 0,
            review_type TEXT NOT NULL DEFAULT '',
            title TEXT NOT NULL DEFAULT '',
            objective TEXT NOT NULL DEFAULT '',
            comparison_basis TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT '',
            decision_text TEXT NOT NULL DEFAULT '',
            summary_text TEXT NOT NULL DEFAULT '',
            notes_json TEXT NOT NULL DEFAULT '[]',
            UNIQUE(analysis_report_id, review_key),
            FOREIGN KEY(analysis_report_id) REFERENCES analysis_reports(analysis_report_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS analysis_cohorts (
            cohort_id INTEGER PRIMARY KEY AUTOINCREMENT,
            review_item_id INTEGER NOT NULL,
            cohort_key TEXT NOT NULL,
            cohort_role TEXT NOT NULL DEFAULT 'VARIANT',
            label TEXT NOT NULL DEFAULT '',
            condition_text TEXT NOT NULL DEFAULT '',
            sort_order INTEGER NOT NULL DEFAULT 0,
            attributes_json TEXT NOT NULL DEFAULT '{}',
            UNIQUE(review_item_id, cohort_key),
            FOREIGN KEY(review_item_id) REFERENCES analysis_review_items(review_item_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS analysis_metrics (
            metric_id INTEGER PRIMARY KEY AUTOINCREMENT,
            review_item_id INTEGER NOT NULL,
            metric_key TEXT NOT NULL,
            scope_key TEXT NOT NULL DEFAULT '',
            sort_order INTEGER NOT NULL DEFAULT 0,
            label TEXT NOT NULL DEFAULT '',
            metric_type TEXT NOT NULL DEFAULT '',
            unit TEXT NOT NULL DEFAULT '',
            spec_text TEXT NOT NULL DEFAULT '',
            definition_text TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT '',
            notes_json TEXT NOT NULL DEFAULT '[]',
            UNIQUE(review_item_id, metric_key, scope_key),
            FOREIGN KEY(review_item_id) REFERENCES analysis_review_items(review_item_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS analysis_metric_values (
            metric_value_id INTEGER PRIMARY KEY AUTOINCREMENT,
            metric_id INTEGER NOT NULL,
            cohort_id INTEGER NOT NULL,
            value_number REAL,
            value_text TEXT NOT NULL DEFAULT '',
            numerator REAL,
            denominator REAL,
            rate_ppm REAL,
            min_value REAL,
            max_value REAL,
            average_value REAL,
            result_status TEXT NOT NULL DEFAULT '',
            details_json TEXT NOT NULL DEFAULT '{}',
            UNIQUE(metric_id, cohort_id),
            FOREIGN KEY(metric_id) REFERENCES analysis_metrics(metric_id) ON DELETE CASCADE,
            FOREIGN KEY(cohort_id) REFERENCES analysis_cohorts(cohort_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS analysis_comparisons (
            comparison_id INTEGER PRIMARY KEY AUTOINCREMENT,
            metric_id INTEGER NOT NULL,
            comparison_key TEXT NOT NULL,
            compared_cohort_id INTEGER NOT NULL,
            control_cohort_id INTEGER NOT NULL,
            delta_value REAL,
            delta_unit TEXT NOT NULL DEFAULT '',
            relative_delta_percent REAL,
            direction TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT '',
            summary_text TEXT NOT NULL DEFAULT '',
            calculation_text TEXT NOT NULL DEFAULT '',
            details_json TEXT NOT NULL DEFAULT '{}',
            UNIQUE(metric_id, comparison_key),
            FOREIGN KEY(metric_id) REFERENCES analysis_metrics(metric_id) ON DELETE CASCADE,
            FOREIGN KEY(compared_cohort_id) REFERENCES analysis_cohorts(cohort_id),
            FOREIGN KEY(control_cohort_id) REFERENCES analysis_cohorts(cohort_id)
        );

        CREATE TABLE IF NOT EXISTS analysis_conclusions (
            conclusion_id INTEGER PRIMARY KEY AUTOINCREMENT,
            analysis_report_id INTEGER NOT NULL,
            review_item_id INTEGER,
            conclusion_key TEXT NOT NULL,
            sort_order INTEGER NOT NULL DEFAULT 0,
            verdict TEXT NOT NULL DEFAULT '',
            label TEXT NOT NULL DEFAULT '',
            conclusion_text TEXT NOT NULL DEFAULT '',
            scope_text TEXT NOT NULL DEFAULT '',
            limitations_json TEXT NOT NULL DEFAULT '[]',
            FOREIGN KEY(analysis_report_id) REFERENCES analysis_reports(analysis_report_id) ON DELETE CASCADE,
            FOREIGN KEY(review_item_id) REFERENCES analysis_review_items(review_item_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS analysis_evidence (
            evidence_id INTEGER PRIMARY KEY AUTOINCREMENT,
            analysis_report_id INTEGER NOT NULL,
            review_item_id INTEGER,
            metric_id INTEGER,
            comparison_id INTEGER,
            conclusion_id INTEGER,
            workbook_id INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            start_row INTEGER NOT NULL,
            start_col INTEGER NOT NULL,
            end_row INTEGER NOT NULL,
            end_col INTEGER NOT NULL,
            range_address TEXT NOT NULL,
            evidence_role TEXT NOT NULL DEFAULT 'SOURCE',
            note TEXT NOT NULL DEFAULT '',
            source_text TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(analysis_report_id) REFERENCES analysis_reports(analysis_report_id) ON DELETE CASCADE,
            FOREIGN KEY(review_item_id) REFERENCES analysis_review_items(review_item_id) ON DELETE CASCADE,
            FOREIGN KEY(metric_id) REFERENCES analysis_metrics(metric_id) ON DELETE CASCADE,
            FOREIGN KEY(comparison_id) REFERENCES analysis_comparisons(comparison_id) ON DELETE CASCADE,
            FOREIGN KEY(conclusion_id) REFERENCES analysis_conclusions(conclusion_id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_workbooks_dataset ON workbooks(dataset);
        CREATE INDEX IF NOT EXISTS idx_workbooks_file_name ON workbooks(file_name);
        CREATE INDEX IF NOT EXISTS idx_grid_rows_text ON grid_sheet_rows(row_text);
        CREATE INDEX IF NOT EXISTS idx_grid_cells_value ON grid_sheet_cells(value_text);
        CREATE INDEX IF NOT EXISTS idx_grid_cells_lookup ON grid_sheet_cells(workbook_id, sheet_name, row_number);
        CREATE INDEX IF NOT EXISTS idx_ingest_items_dataset_status ON ingest_items(dataset, status);
        CREATE INDEX IF NOT EXISTS idx_ingest_items_source ON ingest_items(source_path);
        CREATE INDEX IF NOT EXISTS idx_analysis_reports_workbook ON analysis_reports(workbook_id, overall_status);
        CREATE INDEX IF NOT EXISTS idx_analysis_reviews_report ON analysis_review_items(analysis_report_id, sort_order);
        CREATE INDEX IF NOT EXISTS idx_analysis_metrics_review ON analysis_metrics(review_item_id, sort_order);
        CREATE INDEX IF NOT EXISTS idx_analysis_values_metric ON analysis_metric_values(metric_id, cohort_id);
        CREATE INDEX IF NOT EXISTS idx_analysis_comparisons_metric ON analysis_comparisons(metric_id);
        CREATE INDEX IF NOT EXISTS idx_analysis_evidence_report ON analysis_evidence(analysis_report_id, review_item_id, metric_id);
        CREATE INDEX IF NOT EXISTS idx_analysis_evidence_source ON analysis_evidence(workbook_id, sheet_name, start_row, start_col);
        """
    )
    ensure_column(conn, "runs", "skipped", "INTEGER NOT NULL DEFAULT 0")
    ensure_column(conn, "runs", "options_json", "TEXT NOT NULL DEFAULT ''")
    conn.execute(
        "INSERT OR IGNORE INTO schema_migrations(migration_name, applied_at) VALUES (?, ?)",
        ("universal-grid-resumable-ingestion-v1", now_iso()),
    )
    conn.execute(
        "INSERT OR IGNORE INTO schema_migrations(migration_name, applied_at) VALUES (?, ?)",
        ("universal-analysis-layer-v1", now_iso()),
    )
    ensure_knowledge_schema(conn, now_iso)


def init_universal_db(db_path: Path, dataset: str) -> dict[str, Any]:
    with connect_rw(db_path) as conn:
        ensure_universal_schema(conn)
        conn.commit()
    return {"status": "ok", "dataset": dataset, "db": str(db_path)}


def delete_existing_workbook(conn: sqlite3.Connection, dataset: str, source_path: str) -> None:
    rows = conn.execute(
        "SELECT workbook_id FROM workbooks WHERE dataset=? AND source_path=?",
        (dataset, source_path),
    ).fetchall()
    for row in rows:
        workbook_id = int(row["workbook_id"])
        if table_exists(conn, "analysis_reports"):
            conn.execute(
                """
                UPDATE analysis_reports
                SET overall_status='STALE',
                    stale_reason='The source workbook was re-imported; re-validate or recreate this analysis.',
                    updated_at=?
                WHERE workbook_id=? AND overall_status NOT IN ('STALE', 'ARCHIVED')
                """,
                (now_iso(), workbook_id),
            )
        conn.execute("DELETE FROM merge_ranges WHERE workbook_id=?", (workbook_id,))
        conn.execute("DELETE FROM grid_sheet_cells WHERE workbook_id=?", (workbook_id,))
        conn.execute("DELETE FROM grid_sheet_rows WHERE workbook_id=?", (workbook_id,))
        conn.execute("DELETE FROM worksheets WHERE workbook_id=?", (workbook_id,))
        conn.execute("DELETE FROM workbooks WHERE workbook_id=?", (workbook_id,))


def excel_column_label(column: int) -> str:
    if column < 1:
        return ""
    label = ""
    current = column
    while current:
        current, remainder = divmod(current - 1, 26)
        label = chr(65 + remainder) + label
    return label


def grid_cell_address(row: int, column: int) -> str:
    return f"{excel_column_label(column)}{row}"


def read_com_grid_json(json_path: Path) -> dict[str, Any]:
    try:
        data = json.loads(json_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid COM grid JSON: {json_path}: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"COM grid JSON must be an object: {json_path}")
    return data


def integer_field(value: Any, label: str) -> int:
    if isinstance(value, bool):
        raise ValueError(f"{label} must be an integer, not a boolean")
    try:
        number = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be an integer: {value!r}") from exc
    return number


def normalized_source_path(value: Any) -> str:
    text = str(value or "").strip()
    return str(Path(text).resolve()) if text else ""


ANALYSIS_SCHEMA_VERSION = "universal-analysis-v1"
A1_CELL_PATTERN = re.compile(r"^(?P<column>[A-Za-z]+)(?P<row>[1-9]\d*)$")


def optional_number(value: Any, label: str) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        raise ValueError(f"{label} must be numeric, not a boolean")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be numeric: {value!r}") from exc


def required_text(value: Any, label: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"{label} is required")
    return text


def json_text(value: Any, default: Any, label: str) -> str:
    payload = default if value is None else value
    try:
        return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} is not JSON serializable") from exc


def expect_object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be an object")
    return value


def expect_list(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        raise ValueError(f"{label} must be a list")
    return value


def excel_column_number(label: str) -> int:
    column = 0
    for character in label.upper():
        if not ("A" <= character <= "Z"):
            raise ValueError(f"Invalid Excel column label: {label!r}")
        column = column * 26 + (ord(character) - ord("A") + 1)
    if column < 1:
        raise ValueError(f"Invalid Excel column label: {label!r}")
    return column


def parse_a1_cell(value: str, label: str) -> tuple[int, int]:
    match = A1_CELL_PATTERN.fullmatch(value.strip())
    if not match:
        raise ValueError(f"{label} must be an A1 cell reference: {value!r}")
    return int(match.group("row")), excel_column_number(match.group("column"))


def parse_analysis_range(sheet_name: str, value: Any, label: str) -> tuple[str, int, int, int, int, str]:
    """Parse one contiguous A1 range and return a canonical, sheet-qualified ref."""

    text = required_text(value, label)
    explicit_sheet = ""
    range_text = text
    if "!" in text:
        explicit_sheet, range_text = text.rsplit("!", 1)
        explicit_sheet = explicit_sheet.strip()
        if explicit_sheet.startswith("'") and explicit_sheet.endswith("'"):
            explicit_sheet = explicit_sheet[1:-1].replace("''", "'")
        if explicit_sheet and sheet_name and explicit_sheet != sheet_name:
            raise ValueError(f"{label} sheet mismatch: {explicit_sheet!r} != {sheet_name!r}")
    actual_sheet = explicit_sheet or required_text(sheet_name, f"{label}.sheet")
    parts = [part.strip() for part in range_text.split(":")]
    if len(parts) not in {1, 2} or not all(parts):
        raise ValueError(f"{label} must be one contiguous A1 range: {text!r}")
    start_row, start_col = parse_a1_cell(parts[0], label)
    end_row, end_col = parse_a1_cell(parts[-1], label)
    if start_row > end_row or start_col > end_col:
        raise ValueError(f"{label} has reversed bounds: {text!r}")
    start = grid_cell_address(start_row, start_col)
    end = grid_cell_address(end_row, end_col)
    canonical_range = start if start == end else f"{start}:{end}"
    return actual_sheet, start_row, start_col, end_row, end_col, canonical_range


def validate_analysis_evidence(
    conn: sqlite3.Connection,
    workbook_id: int,
    evidence: Any,
    label: str,
) -> dict[str, Any]:
    source = expect_object(evidence, label)
    sheet_name, start_row, start_col, end_row, end_col, range_address = parse_analysis_range(
        str(source.get("sheet") or ""),
        source.get("range"),
        label,
    )
    worksheet = first_dict(
        conn,
        """
        SELECT used_top, used_left, used_bottom, used_right
        FROM worksheets
        WHERE workbook_id=? AND sheet_name=?
        LIMIT 1
        """,
        (workbook_id, sheet_name),
    )
    if not worksheet:
        raise ValueError(f"{label} refers to a missing sheet: {sheet_name!r}")
    if not (
        int(worksheet["used_top"]) <= start_row <= end_row <= int(worksheet["used_bottom"])
        and int(worksheet["used_left"]) <= start_col <= end_col <= int(worksheet["used_right"])
    ):
        raise ValueError(f"{label} is outside the source UsedRange: {sheet_name}!{range_address}")
    expected_cells = (end_row - start_row + 1) * (end_col - start_col + 1)
    stored_cells = conn.execute(
        """
        SELECT COUNT(*)
        FROM grid_sheet_cells
        WHERE workbook_id=? AND sheet_name=?
          AND row_number BETWEEN ? AND ?
          AND col_number BETWEEN ? AND ?
        """,
        (workbook_id, sheet_name, start_row, end_row, start_col, end_col),
    ).fetchone()[0]
    if int(stored_cells) != expected_cells:
        raise ValueError(f"{label} is not fully represented by the stored source grid: {sheet_name}!{range_address}")
    return {
        "sheet": sheet_name,
        "startRow": start_row,
        "startCol": start_col,
        "endRow": end_row,
        "endCol": end_col,
        "range": range_address,
        "role": str(source.get("role") or "SOURCE").strip() or "SOURCE",
        "note": str(source.get("note") or "").strip(),
        "sourceText": str(source.get("sourceText") or "").strip(),
    }


def validate_com_grid_payload(
    data: dict[str, Any],
    expected_source: Path | None = None,
    expected_covered_cell_mode: str | None = None,
) -> dict[str, int]:
    """Validate a fixed-grid COM payload before it can replace source data."""

    if data.get("schemaVersion") != "input-data-com-grid-v1":
        raise ValueError(f"Unsupported COM grid schema: {data.get('schemaVersion')!r}")

    source_path = normalized_source_path(data.get("sourcePath"))
    if not source_path:
        raise ValueError("COM grid JSON is missing sourcePath")
    if expected_source is not None:
        expected_path = str(expected_source.resolve())
        if source_path != expected_path:
            raise ValueError(f"COM grid source mismatch: {source_path} != {expected_path}")
        stat = expected_source.stat()
        if integer_field(data.get("fileSize"), "fileSize") != stat.st_size:
            raise ValueError("COM grid fileSize does not match the current source file")
        if integer_field(data.get("mtimeNs"), "mtimeNs") != stat.st_mtime_ns:
            raise ValueError("COM grid mtimeNs does not match the current source file")

    if expected_covered_cell_mode is not None and data.get("coveredCellMode") != expected_covered_cell_mode:
        raise ValueError(
            "COM grid coveredCellMode does not match the requested mode: "
            f"{data.get('coveredCellMode')!r} != {expected_covered_cell_mode!r}"
        )

    sheets = data.get("sheets")
    if not isinstance(sheets, list):
        raise ValueError("COM grid JSON is missing a sheets list")

    total_rows = 0
    total_cells = 0
    total_non_empty = 0
    total_merges = 0
    seen_sheet_indexes: set[int] = set()
    for fallback_index, sheet in enumerate(sheets, start=1):
        if not isinstance(sheet, dict):
            raise ValueError(f"Sheet {fallback_index} is not an object")
        sheet_index = integer_field(sheet.get("sheetIndex") or fallback_index, f"sheet {fallback_index}.sheetIndex")
        if sheet_index in seen_sheet_indexes:
            raise ValueError(f"Duplicate sheetIndex in COM grid JSON: {sheet_index}")
        seen_sheet_indexes.add(sheet_index)

        used = sheet.get("usedRange")
        if not isinstance(used, dict):
            raise ValueError(f"Sheet {sheet_index} is missing usedRange")
        top = integer_field(used.get("top"), f"sheet {sheet_index}.usedRange.top")
        left = integer_field(used.get("left"), f"sheet {sheet_index}.usedRange.left")
        row_count = integer_field(used.get("rowCount"), f"sheet {sheet_index}.usedRange.rowCount")
        col_count = integer_field(used.get("columnCount"), f"sheet {sheet_index}.usedRange.columnCount")
        if top < 1 or left < 1 or row_count < 1 or col_count < 1:
            raise ValueError(f"Sheet {sheet_index} has invalid UsedRange dimensions")
        bottom = top + row_count - 1
        right = left + col_count - 1
        if integer_field(used.get("bottom"), f"sheet {sheet_index}.usedRange.bottom") != bottom:
            raise ValueError(f"Sheet {sheet_index} UsedRange.bottom does not match rowCount")
        if integer_field(used.get("right"), f"sheet {sheet_index}.usedRange.right") != right:
            raise ValueError(f"Sheet {sheet_index} UsedRange.right does not match columnCount")

        rows = sheet.get("rows")
        if not isinstance(rows, list) or len(rows) != row_count:
            raise ValueError(f"Sheet {sheet_index} does not contain one row record per UsedRange row")

        include_empty_cells = data.get("includeEmptyCells")
        coordinates: dict[tuple[int, int], dict[str, Any]] = {}
        non_empty = 0
        for offset, row in enumerate(rows):
            if not isinstance(row, dict):
                raise ValueError(f"Sheet {sheet_index} row {offset} is not an object")
            row_number = integer_field(row.get("rowNumber"), f"sheet {sheet_index}.rowNumber")
            expected_row = top + offset
            if row_number != expected_row:
                raise ValueError(f"Sheet {sheet_index} row sequence is not a fixed UsedRange grid")
            cells = row.get("cells")
            if not isinstance(cells, list):
                raise ValueError(f"Sheet {sheet_index} row {row_number} is missing cells")
            for cell in cells:
                if not isinstance(cell, dict):
                    raise ValueError(f"Sheet {sheet_index} cell at row {row_number} is not an object")
                cell_row = integer_field(cell.get("row"), f"sheet {sheet_index}.cell.row")
                cell_col = integer_field(cell.get("column"), f"sheet {sheet_index}.cell.column")
                if cell_row != row_number or not (left <= cell_col <= right):
                    raise ValueError(f"Sheet {sheet_index} contains a cell outside its fixed row/column grid")
                key = (cell_row, cell_col)
                if key in coordinates:
                    raise ValueError(f"Sheet {sheet_index} duplicates coordinate {grid_cell_address(cell_row, cell_col)}")
                address = str(cell.get("address") or "")
                expected_address = grid_cell_address(cell_row, cell_col)
                if address != expected_address:
                    raise ValueError(f"Sheet {sheet_index} cell address mismatch: {address!r} != {expected_address!r}")
                merge = cell.get("merge") or {"role": "none"}
                role = str(merge.get("role") or "none")
                if role not in {"none", "anchor", "covered"}:
                    raise ValueError(f"Sheet {sheet_index} has invalid merge role {role!r}")
                coordinates[key] = cell
                if value_text(cell.get("value")):
                    non_empty += 1

        if include_empty_cells is True and len(coordinates) != row_count * col_count:
            raise ValueError(f"Sheet {sheet_index} omitted coordinates despite includeEmptyCells=true")

        merges = sheet.get("merges") or []
        if not isinstance(merges, list):
            raise ValueError(f"Sheet {sheet_index} merges must be a list")
        expected_merge_cells: dict[tuple[int, int], tuple[str, int, int]] = {}
        seen_merges: set[tuple[int, int, int, int]] = set()
        for merge in merges:
            if not isinstance(merge, dict):
                raise ValueError(f"Sheet {sheet_index} merge is not an object")
            m_top = integer_field(merge.get("top"), f"sheet {sheet_index}.merge.top")
            m_left = integer_field(merge.get("left"), f"sheet {sheet_index}.merge.left")
            m_bottom = integer_field(merge.get("bottom"), f"sheet {sheet_index}.merge.bottom")
            m_right = integer_field(merge.get("right"), f"sheet {sheet_index}.merge.right")
            merge_key = (m_top, m_left, m_bottom, m_right)
            if merge_key in seen_merges:
                raise ValueError(f"Sheet {sheet_index} has duplicate merge range {merge_key}")
            seen_merges.add(merge_key)
            if not (top <= m_top <= m_bottom <= bottom and left <= m_left <= m_right <= right):
                raise ValueError(f"Sheet {sheet_index} merge range is outside UsedRange")
            if integer_field(merge.get("rowSpan"), f"sheet {sheet_index}.merge.rowSpan") != m_bottom - m_top + 1:
                raise ValueError(f"Sheet {sheet_index} merge rowSpan is inconsistent")
            if integer_field(merge.get("columnSpan"), f"sheet {sheet_index}.merge.columnSpan") != m_right - m_left + 1:
                raise ValueError(f"Sheet {sheet_index} merge columnSpan is inconsistent")
            address = str(merge.get("address") or "")
            expected_address = f"{grid_cell_address(m_top, m_left)}:{grid_cell_address(m_bottom, m_right)}"
            if address != expected_address:
                raise ValueError(f"Sheet {sheet_index} merge address mismatch: {address!r} != {expected_address!r}")
            for row_number in range(m_top, m_bottom + 1):
                for col_number in range(m_left, m_right + 1):
                    key = (row_number, col_number)
                    if key in expected_merge_cells:
                        raise ValueError(f"Sheet {sheet_index} has overlapping merge ranges")
                    role = "anchor" if key == (m_top, m_left) else "covered"
                    expected_merge_cells[key] = (address, m_top, m_left)

        for key, (merge_address, anchor_row, anchor_col) in expected_merge_cells.items():
            cell = coordinates.get(key)
            if cell is None:
                raise ValueError(f"Sheet {sheet_index} is missing merged coordinate {grid_cell_address(*key)}")
            merge = cell.get("merge") or {}
            expected_role = "anchor" if key == (anchor_row, anchor_col) else "covered"
            if merge.get("role") != expected_role or merge.get("address") != merge_address:
                raise ValueError(f"Sheet {sheet_index} merge metadata is inconsistent at {grid_cell_address(*key)}")
            anchor = merge.get("anchor") or {}
            if anchor.get("row") != anchor_row or anchor.get("column") != anchor_col:
                raise ValueError(f"Sheet {sheet_index} merge anchor is inconsistent at {grid_cell_address(*key)}")

        declared_merge_count = sheet.get("mergeCount")
        if declared_merge_count is not None and integer_field(declared_merge_count, f"sheet {sheet_index}.mergeCount") != len(merges):
            raise ValueError(f"Sheet {sheet_index} mergeCount does not match merges")

        total_rows += row_count
        total_cells += row_count * col_count
        total_non_empty += non_empty
        total_merges += len(merges)

    totals = data.get("totals") or {}
    expected_totals = {
        "sheetCount": len(sheets),
        "rowCount": total_rows,
        "cellCount": total_cells,
        "mergeCount": total_merges,
    }
    for name, expected in expected_totals.items():
        if name in totals and integer_field(totals.get(name), f"totals.{name}") != expected:
            raise ValueError(f"COM grid totals.{name} does not match the fixed-grid payload")
    if "nonEmptyCells" in totals and integer_field(totals.get("nonEmptyCells"), "totals.nonEmptyCells") != total_non_empty:
        raise ValueError("COM grid totals.nonEmptyCells does not match the payload")

    return {
        "sheets": len(sheets),
        "rows": total_rows,
        "cells": total_cells,
        "nonEmptyCells": total_non_empty,
        "merges": total_merges,
    }


def current_workbook_for_source(conn: sqlite3.Connection, dataset: str, source: Path) -> dict[str, Any] | None:
    return first_dict(
        conn,
        "SELECT * FROM workbooks WHERE dataset=? AND source_path=? LIMIT 1",
        (dataset, str(source.resolve())),
    )


def workbook_is_current(conn: sqlite3.Connection, dataset: str, source: Path) -> dict[str, Any] | None:
    workbook = current_workbook_for_source(conn, dataset, source)
    if not workbook or workbook.get("status") != "OK":
        return None
    raw_path = Path(str(workbook.get("raw_json_path") or ""))
    if not raw_path.is_file() or workbook.get("fingerprint") != file_fingerprint(source):
        return None
    return workbook


def raw_json_path_for_source(raw_dir: Path, source: Path) -> Path:
    return raw_dir / f"{safe_name(source.stem)}_{file_fingerprint(source)[:16]}.com-grid.json"


def raw_json_matches_source(
    json_path: Path,
    source: Path,
    covered_cell_mode: str,
    include_empty_cells: bool,
) -> bool:
    if not json_path.is_file():
        return False
    try:
        data = read_com_grid_json(json_path)
        if data.get("includeEmptyCells") is not include_empty_cells:
            return False
        validate_com_grid_payload(
            data,
            expected_source=source,
            expected_covered_cell_mode=covered_cell_mode,
        )
    except (OSError, ValueError):
        return False
    return True


def start_ingest_item(
    conn: sqlite3.Connection,
    run_id: int,
    dataset: str,
    source: Path,
    raw_json_path: Path,
    status: str,
    workbook_id: int | None = None,
    message: str = "",
) -> int:
    cur = conn.execute(
        """
        INSERT INTO ingest_items(
            run_id, dataset, source_path, source_fingerprint, raw_json_path, status,
            workbook_id, message, started_at, finished_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            run_id,
            dataset,
            str(source.resolve()),
            file_fingerprint(source),
            str(raw_json_path),
            status,
            workbook_id,
            message,
            now_iso(),
            now_iso() if status == "SKIPPED" else "",
        ),
    )
    return int(cur.lastrowid)


def finish_ingest_item(
    conn: sqlite3.Connection,
    ingest_item_id: int,
    status: str,
    workbook_id: int | None = None,
    message: str = "",
) -> None:
    conn.execute(
        """
        UPDATE ingest_items
        SET status=?, workbook_id=?, message=?, finished_at=?
        WHERE ingest_item_id=?
        """,
        (status, workbook_id, message, now_iso(), ingest_item_id),
    )


def import_com_json(
    conn: sqlite3.Connection,
    dataset: str,
    json_path: Path,
    expected_source: Path | None = None,
    expected_covered_cell_mode: str | None = None,
    verify_after_import: bool = False,
) -> dict[str, Any]:
    data = read_com_grid_json(json_path)
    validate_com_grid_payload(data, expected_source, expected_covered_cell_mode)
    conn.execute("SAVEPOINT universal_grid_import")
    try:
        result = import_com_grid_payload(conn, dataset, json_path, data)
        if verify_after_import:
            verification = verify_universal_workbook(conn, int(result["workbookId"]), data)
            if not verification["ok"]:
                raise ValueError("Post-import verification failed: " + "; ".join(verification["errors"]))
            result["verification"] = verification
    except Exception:
        conn.execute("ROLLBACK TO SAVEPOINT universal_grid_import")
        conn.execute("RELEASE SAVEPOINT universal_grid_import")
        raise
    conn.execute("RELEASE SAVEPOINT universal_grid_import")
    return result


def import_com_grid_payload(conn: sqlite3.Connection, dataset: str, json_path: Path, data: dict[str, Any]) -> dict[str, Any]:
    source_path = str(Path(data.get("sourcePath") or "").resolve()) if data.get("sourcePath") else ""
    source = Path(source_path) if source_path else Path(data.get("fileName") or json_path.stem)
    totals = data.get("totals") or {}
    sheets = data.get("sheets") or []

    size_bytes = int(data.get("fileSize") or (source.stat().st_size if source.exists() else 0))
    mtime_ns = int(data.get("mtimeNs") or (source.stat().st_mtime_ns if source.exists() else 0))
    fingerprint = file_fingerprint(source) if source.exists() else hashlib.sha1(str(source).encode("utf-8")).hexdigest()

    delete_existing_workbook(conn, dataset, source_path or str(source))
    cur = conn.execute(
        """
        INSERT INTO workbooks(
            dataset, source_path, file_name, extension, size_bytes, mtime_ns, fingerprint,
            extractor, status, error, sheet_count, total_rows, total_cells, non_empty_cells,
            merge_count, raw_json_path, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            dataset,
            source_path or str(source),
            data.get("fileName") or source.name,
            source.suffix.lower(),
            size_bytes,
            mtime_ns,
            fingerprint,
            "excel-com-grid-v1",
            "OK",
            "",
            int(totals.get("sheetCount") or len(sheets)),
            int(totals.get("rowCount") or 0),
            int(totals.get("cellCount") or 0),
            int(totals.get("nonEmptyCells") or 0),
            int(totals.get("mergeCount") or 0),
            str(json_path),
            now_iso(),
        ),
    )
    workbook_id = int(cur.lastrowid)

    worksheet_rows: list[tuple[Any, ...]] = []
    merge_rows: list[tuple[Any, ...]] = []
    row_rows: list[tuple[Any, ...]] = []
    cell_rows: list[tuple[Any, ...]] = []

    for fallback_index, sheet in enumerate(sheets, start=1):
        sheet_index = int(sheet.get("sheetIndex") or fallback_index)
        sheet_name = str(sheet.get("sheetName") or f"Sheet{sheet_index}")
        used = sheet.get("usedRange") or {}
        worksheet_rows.append(
            (
                workbook_id,
                sheet_index,
                sheet_name,
                1 if sheet.get("visible", True) else 0,
                int(used.get("top") or 0),
                int(used.get("left") or 0),
                int(used.get("bottom") or 0),
                int(used.get("right") or 0),
                int(used.get("rowCount") or 0),
                int(used.get("columnCount") or 0),
                int(sheet.get("nonEmptyCells") or 0),
                int(sheet.get("mergeCount") or 0),
            )
        )

        for merge in sheet.get("merges") or []:
            merge_rows.append(
                (
                    workbook_id,
                    sheet_index,
                    sheet_name,
                    str(merge.get("address") or ""),
                    int(merge.get("top") or 0),
                    int(merge.get("left") or 0),
                    int(merge.get("bottom") or 0),
                    int(merge.get("right") or 0),
                    int(merge.get("rowSpan") or 0),
                    int(merge.get("columnSpan") or 0),
                    value_text(merge.get("value")),
                )
            )

        for row in sheet.get("rows") or []:
            row_number = int(row.get("rowNumber") or 0)
            condensed_cells = []
            non_empty_values = []
            for cell in row.get("cells") or []:
                merge = cell.get("merge") or {}
                anchor = merge.get("anchor") or {}
                text = value_text(cell.get("value"))
                raw_text = value_text(cell.get("rawValue"))
                if text:
                    non_empty_values.append(text)
                col_number = int(cell.get("column") or 0)
                condensed_cells.append(
                    {
                        "address": str(cell.get("address") or ""),
                        "column": col_number,
                        "colLabel": str(cell.get("colLabel") or ""),
                        "value": text,
                        "mergeRole": str(merge.get("role") or "none"),
                        "mergeAddress": str(merge.get("address") or ""),
                    }
                )
                cell_rows.append(
                    (
                        workbook_id,
                        sheet_index,
                        sheet_name,
                        row_number,
                        col_number,
                        str(cell.get("colLabel") or ""),
                        str(cell.get("address") or ""),
                        text,
                        raw_text,
                        str(merge.get("role") or "none"),
                        str(merge.get("address") or ""),
                        anchor.get("row"),
                        anchor.get("column"),
                    )
                )

            row_text = " | ".join(non_empty_values)
            row_rows.append(
                (
                    workbook_id,
                    sheet_index,
                    sheet_name,
                    row_number,
                    int(row.get("nonEmptyCount") or len(non_empty_values)),
                    row_text,
                    json.dumps(condensed_cells, ensure_ascii=False, separators=(",", ":")),
                )
            )

    conn.executemany(
        """
        INSERT INTO worksheets(
            workbook_id, sheet_index, sheet_name, visible, used_top, used_left, used_bottom, used_right,
            row_count, col_count, non_empty_cells, merge_count
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        worksheet_rows,
    )
    conn.executemany(
        """
        INSERT INTO merge_ranges(
            workbook_id, sheet_index, sheet_name, address, top, left_col, bottom, right_col,
            row_span, column_span, anchor_value
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        merge_rows,
    )
    conn.executemany(
        """
        INSERT INTO grid_sheet_rows(
            workbook_id, sheet_index, sheet_name, row_number, non_empty_count, row_text, cells_json
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        row_rows,
    )
    conn.executemany(
        """
        INSERT INTO grid_sheet_cells(
            workbook_id, sheet_index, sheet_name, row_number, col_number, col_label, address,
            value_text, raw_value_text, merge_role, merge_address, anchor_row, anchor_col
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        cell_rows,
    )

    return {
        "workbookId": workbook_id,
        "fileName": data.get("fileName") or source.name,
        "sheets": len(worksheet_rows),
        "rows": len(row_rows),
        "cells": len(cell_rows),
        "merges": len(merge_rows),
    }


def verify_universal_workbook(conn: sqlite3.Connection, workbook_id: int, data: dict[str, Any]) -> dict[str, Any]:
    expected = validate_com_grid_payload(data)
    workbook = first_dict(conn, "SELECT * FROM workbooks WHERE workbook_id=? LIMIT 1", (workbook_id,))
    if not workbook:
        return {"workbookId": workbook_id, "ok": False, "errors": ["workbook is missing"]}

    actual = {
        "sheets": conn.execute("SELECT COUNT(*) FROM worksheets WHERE workbook_id=?", (workbook_id,)).fetchone()[0],
        "rows": conn.execute("SELECT COUNT(*) FROM grid_sheet_rows WHERE workbook_id=?", (workbook_id,)).fetchone()[0],
        "cells": conn.execute("SELECT COUNT(*) FROM grid_sheet_cells WHERE workbook_id=?", (workbook_id,)).fetchone()[0],
        "merges": conn.execute("SELECT COUNT(*) FROM merge_ranges WHERE workbook_id=?", (workbook_id,)).fetchone()[0],
    }
    errors: list[str] = []
    for key in ("sheets", "rows", "cells", "merges"):
        if actual[key] != expected[key]:
            errors.append(f"{key} mismatch: DB={actual[key]} JSON={expected[key]}")

    if int(workbook["sheet_count"]) != expected["sheets"]:
        errors.append("workbooks.sheet_count does not match JSON")
    if int(workbook["total_rows"]) != expected["rows"]:
        errors.append("workbooks.total_rows does not match JSON")
    if int(workbook["total_cells"]) != expected["cells"]:
        errors.append("workbooks.total_cells does not match JSON")
    if int(workbook["merge_count"]) != expected["merges"]:
        errors.append("workbooks.merge_count does not match JSON")

    for fallback_index, sheet in enumerate(data.get("sheets") or [], start=1):
        sheet_index = int(sheet.get("sheetIndex") or fallback_index)
        db_sheet = first_dict(
            conn,
            "SELECT * FROM worksheets WHERE workbook_id=? AND sheet_index=? LIMIT 1",
            (workbook_id, sheet_index),
        )
        if not db_sheet:
            errors.append(f"worksheet {sheet_index} is missing")
            continue
        used = sheet.get("usedRange") or {}
        expected_sheet = {
            "used_top": int(used.get("top") or 0),
            "used_left": int(used.get("left") or 0),
            "used_bottom": int(used.get("bottom") or 0),
            "used_right": int(used.get("right") or 0),
            "row_count": int(used.get("rowCount") or 0),
            "col_count": int(used.get("columnCount") or 0),
            "merge_count": len(sheet.get("merges") or []),
        }
        for key, expected_value in expected_sheet.items():
            if int(db_sheet[key]) != expected_value:
                errors.append(f"worksheet {sheet_index}.{key} mismatch")

    return {
        "workbookId": workbook_id,
        "fileName": workbook["file_name"],
        "ok": not errors,
        "expected": expected,
        "actual": actual,
        "errors": errors,
    }


def mark_failed_workbook(conn: sqlite3.Connection, dataset: str, source: Path, raw_json_path: Path, error: str) -> int:
    source_path = str(source.resolve())
    existing = current_workbook_for_source(conn, dataset, source)
    if existing:
        # Preserve the last successful grid. Per-run ingest_items records the
        # new failure without destroying source data that was already valid.
        return int(existing["workbook_id"])
    stat = source.stat()
    cur = conn.execute(
        """
        INSERT INTO workbooks(
            dataset, source_path, file_name, extension, size_bytes, mtime_ns, fingerprint,
            extractor, status, error, raw_json_path, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            dataset,
            source_path,
            source.name,
            source.suffix.lower(),
            stat.st_size,
            stat.st_mtime_ns,
            file_fingerprint(source),
            "excel-com-grid-v1",
            "FAILED",
            compact(error, 2000),
            str(raw_json_path),
            now_iso(),
        ),
    )
    return int(cur.lastrowid)


def cmd_init_db(args: argparse.Namespace) -> int:
    dataset = args.dataset or DEFAULT_DATASET
    db_path = service_output_path(
        args.db,
        UNIVERSAL_GRID_DIR / f"{safe_name(dataset)}.sqlite",
    )
    result = init_universal_db(db_path, dataset)
    print_json(result)
    return 0


def _capture_v2_files(args: argparse.Namespace) -> list[Path]:
    if args.pilot_manifest:
        manifest_path = Path(args.pilot_manifest).resolve()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        source_root = Path(args.input or manifest.get("sourceRoot") or DEFAULT_INPUT_DIR).resolve()
        files = [
            (source_root / str(item["relativePath"])).resolve()
            for item in manifest.get("workbooks", [])
        ]
        missing = [str(path) for path in files if not path.is_file()]
        if missing:
            raise SystemExit("Pilot manifest contains missing files:\n" + "\n".join(missing))
    else:
        files = [
            path
            for path in collect_excel_files(Path(args.input or DEFAULT_INPUT_DIR).resolve())
            if path.suffix.lower() == ".xlsx"
        ]
    offset = max(0, int(args.offset or 0))
    files = files[offset:]
    if args.limit > 0:
        files = files[: args.limit]
    return files


def _ordered_capture_v2_extractions(
    files: list[Path],
    *,
    workers: int,
) -> Iterator[tuple[Path, dict[str, Any] | None, Exception | None]]:
    """Extract a bounded number of workbooks concurrently in source order."""

    if workers < 1:
        raise ValueError("Capture workers must be positive.")

    def extract_one(
        source: Path,
    ) -> tuple[dict[str, Any] | None, Exception | None]:
        try:
            return extract_openxml_workbook(source), None
        except Exception as exc:
            return None, exc

    if workers == 1:
        for source in files:
            payload, error = extract_one(source)
            yield source, payload, error
        return

    with ThreadPoolExecutor(max_workers=workers) as executor:
        pending = deque(
            (
                source,
                executor.submit(extract_one, source),
            )
            for source in files[:workers]
        )
        next_index = len(pending)
        while pending:
            source, future = pending.popleft()
            payload, error = future.result()
            yield source, payload, error
            if next_index < len(files):
                next_source = files[next_index]
                pending.append(
                    (
                        next_source,
                        executor.submit(extract_one, next_source),
                    )
                )
                next_index += 1


def cmd_openxml_index(args: argparse.Namespace) -> int:
    files = _capture_v2_files(args)
    workers = int(args.workers or 1)
    if workers < 1:
        raise SystemExit("--workers must be positive.")
    db_path = service_output_path(
        args.db,
        UNIVERSAL_GRID_DIR / f"{safe_name(args.dataset)}.sqlite",
    )
    raw_dir = service_output_dir(args.raw_dir, CAPTURE_V2_DIR / "raw-json" / safe_name(args.dataset)) if args.raw_dir else None
    started_at = now_iso()
    with connect_rw(db_path) as conn:
        ensure_universal_schema(conn)
        ensure_capture_v2_schema(conn)
        conn.execute(
            """
            UPDATE capture_v2_runs
            SET finished_at=?, failed=CASE WHEN failed=0 THEN total_files ELSE failed END
            WHERE finished_at=''
            """,
            (started_at,),
        )
        cursor = conn.execute(
            """
            INSERT INTO capture_v2_runs(
                dataset, input_path, started_at, total_files, options_json
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (
                args.dataset,
                str(Path(args.input or DEFAULT_INPUT_DIR).resolve()),
                started_at,
                len(files),
                json.dumps(
                    {
                        "offset": args.offset,
                        "limit": args.limit,
                        "pilotManifest": str(Path(args.pilot_manifest).resolve()) if args.pilot_manifest else "",
                        "retainRawJson": raw_dir is not None,
                        "imageHandling": "IGNORED",
                        "workers": workers,
                    },
                    ensure_ascii=False,
                ),
            ),
        )
        run_id = int(cursor.lastrowid)
        conn.commit()
        succeeded = failed = skipped = reactivated = 0
        items: list[dict[str, Any]] = []
        for source, payload, extraction_error in _ordered_capture_v2_extractions(
            files,
            workers=workers,
        ):
            item_started = now_iso()
            action = "FAILED"
            message = ""
            content_sha256 = ""
            capture_revision_id: int | None = None
            canonical_revision_id: int | None = None
            conn.execute("SAVEPOINT capture_v2_cli_item")
            try:
                if extraction_error is not None:
                    raise extraction_error
                if payload is None:
                    raise RuntimeError(
                        "OpenXML extraction returned no workbook payload."
                    )
                content_sha256 = str(payload["source"]["contentSha256"])
                capture_result = import_openxml_capture(conn, payload, captured_at=item_started)
                capture_revision_id = int(capture_result["revisionId"])
                bridge = bridge_capture_to_canonical_source(
                    conn,
                    dataset=args.dataset,
                    payload=payload,
                    capture_result=capture_result,
                    captured_at=item_started,
                )
                canonical_revision_id = int(bridge["revisionId"])
                action = str(capture_result["action"])
                if raw_dir is not None:
                    path_key = hashlib.sha256(str(source).casefold().encode("utf-8")).hexdigest()[:12]
                    raw_path = raw_dir / f"{safe_name(source.stem)[:80]}_{path_key}.capture-v2.json"
                    raw_path.write_bytes(capture_json_bytes(payload))
                if action == "SKIPPED":
                    skipped += 1
                elif action == "REACTIVATED":
                    reactivated += 1
                    succeeded += 1
                else:
                    succeeded += 1
                conn.execute("RELEASE SAVEPOINT capture_v2_cli_item")
            except Exception as exc:
                conn.execute("ROLLBACK TO SAVEPOINT capture_v2_cli_item")
                conn.execute("RELEASE SAVEPOINT capture_v2_cli_item")
                failed += 1
                message = f"{type(exc).__name__}: {exc}"
            conn.execute(
                """
                INSERT INTO capture_v2_ingest_items(
                    run_id, source_path, content_sha256, action,
                    capture_revision_id, canonical_revision_id, message,
                    started_at, finished_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    str(source),
                    content_sha256,
                    action,
                    capture_revision_id,
                    canonical_revision_id,
                    message,
                    item_started,
                    now_iso(),
                ),
            )
            conn.commit()
            items.append(
                {
                    "sourcePath": str(source),
                    "action": action,
                    "message": message,
                    "captureRevisionId": capture_revision_id,
                    "canonicalRevisionId": canonical_revision_id,
                }
            )
        conn.execute(
            """
            UPDATE capture_v2_runs
            SET finished_at=?, succeeded=?, failed=?, skipped=?, reactivated=?
            WHERE run_id=?
            """,
            (now_iso(), succeeded, failed, skipped, reactivated, run_id),
        )
        conn.commit()
    result = {
        "status": "ok" if failed == 0 else "partial",
        "runId": run_id,
        "dataset": args.dataset,
        "db": str(db_path),
        "selected": len(files),
        "succeeded": succeeded,
        "failed": failed,
        "skipped": skipped,
        "reactivated": reactivated,
        "imageHandling": "IGNORED",
        "items": items,
    }
    print_json(result)
    return 0 if failed == 0 else 1


def cmd_verify_capture_v2(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if not table_exists(conn, "capture_v2_revisions"):
            raise SystemExit("capture-v2-verify requires a DB with Capture v2 tables.")
        result = verify_capture_database(
            conn,
            current_only=not args.all_revisions,
            verify_source_sha256=args.source_sha256,
        )
    print_json({"db": str(db_path), **result})
    return 0 if result["ok"] else 1


def _semantic_source(packet_set: dict[str, Any], dataset: str) -> dict[str, Any]:
    revision = packet_set["inventory"]["sourceRevision"]
    return {
        "dataset": dataset,
        "sourcePath": str(revision["sourcePath"]),
        "revisionUid": str(revision["revisionUid"]),
        "contentSha256": str(revision["contentSha256"]),
    }


def _semantic_workbook_summary(packet_set: dict[str, Any]) -> dict[str, Any]:
    inventory = packet_set["inventory"]
    return {
        **inventory["workbook"],
        "semanticCellCoverageComplete": bool(
            inventory.get("semanticCellCoverageComplete")
        ),
        "contentCompleteForManifest": bool(
            inventory.get("contentCompleteForManifest")
        ),
        "coverage": inventory["coverage"],
        "sheets": [
            {
                "sheetIndex": sheet["sheetIndex"],
                "title": sheet["title"],
                "sheetState": sheet["sheetState"],
                "status": sheet["status"],
                "hasTabularEvidence": sheet["hasTabularEvidence"],
                "nonEmptyCellCount": sheet["nonEmptyCellCount"],
                "formulaCellCount": sheet["formulaCellCount"],
                "mergeCount": sheet["mergeCount"],
                "contentBounds": sheet["contentBounds"],
                "sections": sheet["sections"],
            }
            for sheet in inventory["sheets"]
        ],
    }


def cmd_build_semantic_packets(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    packet_set = build_semantic_source_packets_from_db(
        db_path,
        revision_id=args.revision_id,
        source_path=args.source_path,
        max_cells=args.max_cells,
        max_rows=args.max_rows,
        empty_row_gap=args.empty_row_gap,
    )
    revision = packet_set["inventory"]["sourceRevision"]
    default_out = SEMANTIC_PACKET_DIR / f"{safe_name(revision['revisionUid'])}.json"
    output_path = service_output_path(args.out, default_out)
    output_path.write_bytes(semantic_packet_json_bytes(packet_set))
    print_json(
        {
            "status": "ok",
            "db": str(db_path),
            "packet": str(output_path),
            "revisionUid": revision["revisionUid"],
            "workbookStatus": packet_set["inventory"]["workbook"]["status"],
            "chunks": len(packet_set["chunks"]),
            "terminalPackets": len(packet_set["terminalPackets"]),
            "coverage": packet_set["inventory"]["coverage"],
        }
    )
    return 0


def cmd_build_semantic_packets_batch(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        revision_ids = [
            int(row[0])
            for row in conn.execute(
                """
                SELECT revision_id
                FROM capture_v2_revisions
                WHERE is_current=1
                ORDER BY revision_id
                """
            )
        ]
    if args.offset:
        revision_ids = revision_ids[args.offset :]
    if args.limit > 0:
        revision_ids = revision_ids[: args.limit]
    output_dir = service_output_dir(
        args.out_dir,
        SEMANTIC_PACKET_DIR,
    )

    def build_one(revision_id: int) -> dict[str, Any]:
        packet_set = build_semantic_source_packets_from_db(
            db_path,
            revision_id=revision_id,
            max_cells=args.max_cells,
            max_rows=args.max_rows,
            empty_row_gap=args.empty_row_gap,
        )
        revision = packet_set["inventory"]["sourceRevision"]
        output_path = output_dir / f"{safe_name(revision['revisionUid'])}.json"
        if output_path.is_file() and not args.force:
            existing = _load_semantic_packet(output_path)
            if (
                existing["inventory"]["sourceRevision"]["contentSha256"]
                == revision["contentSha256"]
                and existing["inventory"]["limits"]
                == packet_set["inventory"]["limits"]
            ):
                return {
                    "revisionId": revision_id,
                    "revisionUid": revision["revisionUid"],
                    "action": "SKIPPED",
                    "packet": str(output_path),
                    "chunks": len(existing["chunks"]),
                    "cells": existing["inventory"]["coverage"][
                        "packetCellCount"
                    ],
                }
        output_path.write_bytes(semantic_packet_json_bytes(packet_set))
        return {
            "revisionId": revision_id,
            "revisionUid": revision["revisionUid"],
            "action": "BUILT",
            "packet": str(output_path),
            "chunks": len(packet_set["chunks"]),
            "cells": packet_set["inventory"]["coverage"]["packetCellCount"],
        }

    items: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        futures = {
            executor.submit(build_one, revision_id): revision_id
            for revision_id in revision_ids
        }
        for future in as_completed(futures):
            revision_id = futures[future]
            try:
                items.append(future.result())
            except Exception as exc:
                failures.append(
                    {
                        "revisionId": revision_id,
                        "error": str(exc),
                    }
                )
    items.sort(key=lambda item: int(item["revisionId"]))
    print_json(
        {
            "status": "ok" if not failures else "partial",
            "db": str(db_path),
            "outputDir": str(output_dir),
            "selected": len(revision_ids),
            "built": sum(item["action"] == "BUILT" for item in items),
            "skipped": sum(item["action"] == "SKIPPED" for item in items),
            "failed": len(failures),
            "chunks": sum(int(item["chunks"]) for item in items),
            "cells": sum(int(item["cells"]) for item in items),
            "items": items,
            "failures": failures,
        }
    )
    return 0 if not failures else 1


def _load_semantic_packet(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise SystemExit(f"Semantic source packet not found: {path}")
    packet_set = json.loads(path.read_text(encoding="utf-8"))
    if packet_set.get("schemaVersion") != "semantic-source-packet-v1":
        raise SystemExit("Expected a semantic-source-packet-v1 packet set.")
    if packet_set.get("inventory", {}).get("coverage", {}).get("status") != "COMPLETE":
        raise SystemExit("Semantic packet coverage is not complete.")
    return packet_set


def _partition_semantic_locator_jobs(
    jobs: list[tuple[dict[str, Any], Path]],
    *,
    batch_size: int,
    batch_max_bytes: int,
) -> list[list[tuple[dict[str, Any], Path]]]:
    """Group source chunks without dropping a single oversized source chunk."""

    if batch_size < 1:
        raise ValueError("semantic locator batch size must be at least 1")
    if batch_max_bytes < 1:
        raise ValueError("semantic locator batch byte budget must be at least 1")
    batches: list[list[tuple[dict[str, Any], Path]]] = []
    current: list[tuple[dict[str, Any], Path]] = []
    current_bytes = 0
    for job in jobs:
        estimated_bytes = len(
            json.dumps(
                job[0],
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode("utf-8")
        )
        if current and (
            len(current) >= batch_size
            or current_bytes + estimated_bytes > batch_max_bytes
        ):
            batches.append(current)
            current = []
            current_bytes = 0
        current.append(job)
        current_bytes += estimated_bytes
    if current:
        batches.append(current)
    return batches


def cmd_semantic_locate(args: argparse.Namespace) -> int:
    packet_path = Path(args.packet).resolve()
    packet_set = _load_semantic_packet(packet_path)
    source = _semantic_source(packet_set, args.dataset)
    workbook = _semantic_workbook_summary(packet_set)
    chunks = list(packet_set["chunks"])
    requested_ids = set(args.chunk_id or [])
    if requested_ids:
        known_ids = {str(chunk["chunkId"]) for chunk in chunks}
        unknown = sorted(requested_ids - known_ids)
        if unknown:
            raise SystemExit("Unknown semantic chunk id(s): " + ", ".join(unknown))
        chunks = [chunk for chunk in chunks if str(chunk["chunkId"]) in requested_ids]
    if args.offset:
        chunks = chunks[args.offset :]
    if args.limit > 0:
        chunks = chunks[: args.limit]
    output_dir = service_output_dir(
        args.out_dir,
        SEMANTIC_LOCATOR_DIR / safe_name(source["revisionUid"]),
    )

    jobs: list[tuple[dict[str, Any], Path]] = []
    skipped = 0
    deterministic = 0

    def has_primary_text(chunk: dict[str, Any]) -> bool:
        return any(
            isinstance(
                cell.get("displayValue")
                if cell.get("displayValue") is not None
                else cell.get("rawValue"),
                str,
            )
            and str(
                cell.get("displayValue")
                if cell.get("displayValue") is not None
                else cell.get("rawValue")
            ).strip()
            for cell in chunk.get("cells", [])
        )

    for chunk in chunks:
        output_path = output_dir / f"{safe_name(chunk['chunkId'])}.locator.json"
        if output_path.is_file() and not args.force:
            try:
                existing = json.loads(output_path.read_text(encoding="utf-8"))
                validate_locator_result(
                    existing,
                    revision_uid=source["revisionUid"],
                    content_sha256=source["contentSha256"],
                    chunk=chunk,
                )
                skipped += 1
                continue
            except (OSError, json.JSONDecodeError, RuntimeError, ValueError):
                pass
        if not args.all_chunks_ai and not has_primary_text(chunk):
            result = {
                "schemaVersion": "semantic-locator-v1",
                "promptVersion": "semantic-locator-prompt-v1",
                "revisionUid": source["revisionUid"],
                "contentSha256": source["contentSha256"],
                "chunkId": str(chunk["chunkId"]),
                "status": "NO_CANDIDATE",
                "candidates": [],
                "notes": [
                    "Deterministic numeric/formula continuation: retained in Capture v2 "
                    "and the lossless packet for on-demand evidence retrieval."
                ],
            }
            validate_locator_result(
                result,
                revision_uid=source["revisionUid"],
                content_sha256=source["contentSha256"],
                chunk=chunk,
            )
            write_json(output_path, result)
            deterministic += 1
            continue
        jobs.append((chunk, output_path))

    batches = _partition_semantic_locator_jobs(
        jobs,
        batch_size=args.batch_size,
        batch_max_bytes=args.batch_max_bytes,
    )
    succeeded = 0
    failures: list[dict[str, str]] = []

    def run_batch(
        batch: list[tuple[dict[str, Any], Path]],
    ) -> list[str]:
        if len(batch) == 1:
            chunk, output_path = batch[0]
            run_codex_locator(
                source=source,
                workbook=workbook,
                chunk=chunk,
                output_path=output_path,
                model=args.model or None,
                reasoning_effort=args.reasoning_effort or None,
                timeout_seconds=args.timeout,
            )
            return [str(chunk["chunkId"])]
        chunks_in_batch = [chunk for chunk, _ in batch]
        output_paths = {
            str(chunk["chunkId"]): output_path
            for chunk, output_path in batch
        }
        run_codex_locator_batch(
            source=source,
            workbook=workbook,
            chunks=chunks_in_batch,
            output_paths=output_paths,
            model=args.model or None,
            reasoning_effort=args.reasoning_effort or None,
            timeout_seconds=args.timeout,
        )
        return [str(chunk["chunkId"]) for chunk in chunks_in_batch]

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        futures = {
            executor.submit(run_batch, batch): (batch_index, batch)
            for batch_index, batch in enumerate(batches, start=1)
        }
        for future in as_completed(futures):
            batch_index, batch = futures[future]
            try:
                succeeded += len(future.result())
            except Exception as exc:
                for chunk, output_path in batch:
                    failures.append(
                        {
                            "chunkId": str(chunk["chunkId"]),
                            "output": str(output_path),
                            "batch": str(batch_index),
                            "error": str(exc),
                        }
                    )
    result = {
        "status": "ok" if not failures else "partial",
        "packet": str(packet_path),
        "outputDir": str(output_dir),
        "selected": len(chunks),
        "succeeded": succeeded,
        "aiSubmitted": len(jobs),
        "aiCalls": len(batches),
        "batchSize": args.batch_size,
        "batchMaxBytes": args.batch_max_bytes,
        "deterministicNoCandidate": deterministic,
        "skipped": skipped,
        "failed": len(failures),
        "failures": failures,
    }
    print_json(result)
    return 0 if not failures else 1


def _terminal_study_manifest(
    packet_set: dict[str, Any],
    dataset: str,
) -> dict[str, Any]:
    source = _semantic_source(packet_set, dataset)
    revision = packet_set["inventory"]["sourceRevision"]
    status = str(packet_set["inventory"]["workbook"]["status"])
    return validate_study_manifest(
        {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                **source,
                "contentComplete": False,
            },
            "workbookAnalysis": {
                "key": f"tabular-evidence-{safe_name(revision['revisionUid']).lower()}",
                "title": str(revision["fileName"]),
                "summary": (
                    "No reviewable tabular evidence was captured. Embedded visual "
                    "content is outside the configured scope and was not assessed."
                ),
                "status": status,
                "verificationStatus": "EXCLUDED",
                "limitations": [
                    f"Capture v2 workbook status: {status}.",
                    "Only tabular cell evidence is in scope.",
                ],
                "evidence": [],
            },
            "studies": [],
        }
    )


def cmd_semantic_draft(args: argparse.Namespace) -> int:
    packet_path = Path(args.packet).resolve()
    packet_set = _load_semantic_packet(packet_path)
    source = _semantic_source(packet_set, args.dataset)
    revision = packet_set["inventory"]["sourceRevision"]
    default_out = (
        SEMANTIC_STUDY_DIR
        / f"{safe_name(revision['revisionUid'])}.study-draft.json"
    )
    output_path = service_output_path(args.out, default_out)
    workbook_status = str(packet_set["inventory"]["workbook"]["status"])
    if workbook_status in {"EMPTY_WORKBOOK", "NO_TABULAR_EVIDENCE"}:
        manifest = _terminal_study_manifest(packet_set, args.dataset)
        write_json(output_path, manifest)
        print_json(
            {
                "status": "excluded",
                "packet": str(packet_path),
                "manifest": str(output_path),
                "workbookStatus": workbook_status,
                "studies": 0,
                "aiExecuted": False,
            }
        )
        return 0

    locator_dir = Path(args.locator_dir).resolve()
    chunks_by_id = {
        str(chunk["chunkId"]): chunk for chunk in packet_set["chunks"]
    }
    locator_results: list[dict[str, Any]] = []
    missing: list[str] = []
    for chunk_id, chunk in chunks_by_id.items():
        locator_path = locator_dir / f"{safe_name(chunk_id)}.locator.json"
        if not locator_path.is_file():
            missing.append(chunk_id)
            continue
        result = json.loads(locator_path.read_text(encoding="utf-8"))
        locator_results.append(
            validate_locator_result(
                result,
                revision_uid=source["revisionUid"],
                content_sha256=source["contentSha256"],
                chunk=chunk,
            )
        )
    if missing:
        raise SystemExit(
            f"Semantic draft requires every locator result; missing {len(missing)} chunk(s)."
        )
    candidate_ids = {
        str(result["chunkId"])
        for result in locator_results
        if result["status"] in {"CANDIDATES", "NEEDS_REVIEW"}
        and result["candidates"]
    }
    if not candidate_ids:
        manifest = _terminal_study_manifest(packet_set, args.dataset)
        manifest["workbookAnalysis"]["status"] = "NO_SEMANTIC_CANDIDATE"
        manifest["workbookAnalysis"]["limitations"].append(
            "Every source chunk completed the locator pass without a study candidate."
        )
        write_json(output_path, manifest)
        print_json(
            {
                "status": "excluded",
                "packet": str(packet_path),
                "manifest": str(output_path),
                "workbookStatus": workbook_status,
                "studies": 0,
                "aiExecuted": False,
            }
        )
        return 0
    focused_chunks = [
        chunks_by_id[chunk_id] for chunk_id in sorted(candidate_ids)
    ]
    content_complete = bool(
        packet_set["inventory"].get("contentCompleteForManifest")
    )
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        canonical_revision = resolve_manifest_revision(conn, source)
        checker = make_database_evidence_checker(conn, canonical_revision)

        def source_claim_validator(draft: dict[str, Any]) -> None:
            validate_numeric_observation_evidence(
                conn,
                canonical_revision,
                draft,
            )
            validate_factor_and_arm_evidence(
                conn,
                canonical_revision,
                draft,
            )
            validate_comparison_representation_alignment(
                conn,
                canonical_revision,
                draft,
            )
            validate_conclusion_evidence(
                conn,
                canonical_revision,
                draft,
            )

        manifest = run_codex_study_draft(
            source=source,
            workbook=_semantic_workbook_summary(packet_set),
            locator_results=locator_results,
            focused_chunks=focused_chunks,
            content_complete=content_complete,
            output_path=output_path,
            evidence_checker=checker,
            additional_validator=source_claim_validator,
            model=args.model or None,
            reasoning_effort=args.reasoning_effort or None,
            timeout_seconds=args.timeout,
        )
    print_json(
        {
            "status": "ok",
            "packet": str(packet_path),
            "manifest": str(output_path),
            "locatorResults": len(locator_results),
            "focusedChunks": len(focused_chunks),
            "studies": len(manifest["studies"]),
            "verificationStatus": manifest["workbookAnalysis"]["verificationStatus"],
        }
    )
    return 0


def _load_table_first_request(path: Path) -> dict[str, Any]:
    try:
        request = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SystemExit(f"Invalid table-first request JSON: {path}") from exc
    if not isinstance(request, dict) or request.get(
        "schemaVersion"
    ) != TABLE_FIRST_REQUEST_SCHEMA_VERSION:
        raise SystemExit("Expected a table-first-request-v1 request.")
    if request.get("builderVersion") != TABLE_FIRST_BUILDER_VERSION:
        raise SystemExit(
            "Table-first request is stale; rebuild it with table-first-request."
        )
    return request


def cmd_table_first_request(args: argparse.Namespace) -> int:
    packet_path = Path(args.packet).expanduser().resolve()
    packet_set = _load_semantic_packet(packet_path)
    request = build_table_first_request(
        packet_set,
        max_preview_rows=args.max_preview_rows,
        max_preview_columns=args.max_preview_columns,
        max_value_samples=args.max_value_samples,
        term_dictionary_path=args.term_dictionary,
    )
    output_path = service_output_path(
        args.out,
        TABLE_FIRST_DIR / "requests" / f"{safe_name(request['requestId'])}.json",
    )
    output_path.write_bytes(table_first_json_bytes(request))
    packet_bytes = packet_path.stat().st_size
    request_bytes = output_path.stat().st_size
    prompt_stats = table_first_prompt_stats(request)
    print_json(
        {
            "status": "ok",
            "packet": str(packet_path),
            "request": str(output_path),
            "requestId": request["requestId"],
            "tables": len(request["tables"]),
            "textBlocks": len(request["textBlocks"]),
            "rawFrequencyResponseExclusions": len(
                (request.get("codeOwnedExclusions") or {}).get(
                    "rawFrequencyResponseTables"
                )
                or []
            ),
            "termDictionary": {
                "status": (request.get("codeOwnedTermDictionary") or {}).get(
                    "status"
                ),
                "definedTerms": (
                    request.get("codeOwnedTermDictionary") or {}
                ).get("definedTermCount"),
                "ignoreTerms": (
                    request.get("codeOwnedTermDictionary") or {}
                ).get("ignoreTermCount"),
                "aliasGroups": (
                    request.get("codeOwnedTermDictionary") or {}
                ).get("aliasGroupCount"),
            },
            "capturedPrimaryCells": request["workbook"][
                "capturedPrimaryCellCount"
            ],
            "packetBytes": packet_bytes,
            "requestBytes": request_bytes,
            **prompt_stats,
            "compressionRatio": (
                round(request_bytes / packet_bytes, 4) if packet_bytes else None
            ),
            "plannedAiCalls": 1 if request["tables"] else 0,
        }
    )
    return 0


def _nearest_rank_percentile(values: list[int], percentile: int) -> int:
    if not values:
        return 0
    ordered = sorted(values)
    rank = max(1, (len(ordered) * percentile + 99) // 100)
    return ordered[min(rank - 1, len(ordered) - 1)]


def _table_first_request_batch_report(
    *,
    packet_dir: Path,
    output_dir: Path,
    selected: int,
    workers: int,
    oversized_request_bytes: int,
    started_at: str,
    elapsed_seconds: float,
    items: list[dict[str, Any]],
    failures: list[dict[str, Any]],
    completed: bool,
) -> dict[str, Any]:
    ordered_items = sorted(items, key=lambda item: int(item["index"]))
    ordered_failures = sorted(failures, key=lambda item: int(item["index"]))
    request_sizes = [int(item["requestBytes"]) for item in ordered_items]
    prompt_sizes = [int(item["promptBytes"]) for item in ordered_items]
    formula_status_counts: Counter[str] = Counter(
        str(item.get("formulaStatus") or "UNKNOWN") for item in ordered_items
    )
    term_dictionary_status_counts: Counter[str] = Counter(
        str(item.get("termDictionaryStatus") or "UNKNOWN")
        for item in ordered_items
    )
    aggregate_status_counts: Counter[str] = Counter()
    for item in ordered_items:
        aggregate_status_counts.update(item.get("aggregateCheckCounts") or {})
    outliers = [
        {
            "index": item["index"],
            "fileName": item["fileName"],
            "request": item["request"],
            "requestBytes": item["requestBytes"],
            "promptBytes": item["promptBytes"],
            "promptTableCount": item["promptTableCount"],
            "tableCount": item["tableCount"],
            "textBlockCount": item["textBlockCount"],
            "reasons": item["outlierReasons"],
        }
        for item in ordered_items
        if item.get("outlierReasons")
    ]
    largest_requests = [
        {
            "index": item["index"],
            "fileName": item["fileName"],
            "request": item["request"],
            "requestBytes": item["requestBytes"],
            "tableCount": item["tableCount"],
            "capturedPrimaryCells": item["capturedPrimaryCells"],
        }
        for item in sorted(
            ordered_items,
            key=lambda item: int(item["requestBytes"]),
            reverse=True,
        )[:20]
    ]
    largest_prompts = [
        {
            "index": item["index"],
            "fileName": item["fileName"],
            "request": item["request"],
            "promptBytes": item["promptBytes"],
            "sourceTableCount": item["tableCount"],
            "promptTableCount": item["promptTableCount"],
            "repeatedOccurrenceCount": item["repeatedOccurrenceCount"],
        }
        for item in sorted(
            ordered_items,
            key=lambda item: int(item["promptBytes"]),
            reverse=True,
        )[:20]
    ]
    succeeded = len(ordered_items)
    return {
        "schemaVersion": "table-first-request-batch-report-v1",
        "status": (
            "running"
            if not completed
            else ("ok" if not ordered_failures else "partial")
        ),
        "builderVersion": TABLE_FIRST_BUILDER_VERSION,
        "promptVersion": TABLE_FIRST_PROMPT_VERSION,
        "packetDir": str(packet_dir),
        "outputDir": str(output_dir),
        "startedAt": started_at,
        "completedAt": now_iso() if completed else None,
        "elapsedSeconds": round(elapsed_seconds, 3),
        "workers": workers,
        "selected": selected,
        "completed": succeeded + len(ordered_failures),
        "succeeded": succeeded,
        "failed": len(ordered_failures),
        "prepared": sum(item["action"] == "PREPARED" for item in ordered_items),
        "reused": sum(item["action"] == "REUSED" for item in ordered_items),
        "aiCalls": 0,
        "plannedAiCalls": sum(bool(item["tableCount"]) for item in ordered_items),
        "noTables": sum(not item["tableCount"] for item in ordered_items),
        "tables": sum(int(item["tableCount"]) for item in ordered_items),
        "textBlocks": sum(int(item["textBlockCount"]) for item in ordered_items),
        "capturedPrimaryCells": sum(
            int(item["capturedPrimaryCells"]) for item in ordered_items
        ),
        "rawFrequencyResponseExclusions": sum(
            int(item["rawFrequencyResponseExclusionCount"])
            for item in ordered_items
        ),
        "rawFrequencyResponseExcludedCells": sum(
            int(item["rawFrequencyResponseExcludedCellCount"])
            for item in ordered_items
        ),
        "workbooksWithRawFrequencyResponseExclusions": sum(
            bool(item["rawFrequencyResponseExclusionCount"])
            for item in ordered_items
        ),
        "workbooksExcludedFromLearning": sum(
            bool(item.get("workbookLearningExcluded"))
            for item in ordered_items
        ),
        "workbookLearningExcludedCells": sum(
            int(item.get("workbookLearningExcludedCellCount") or 0)
            for item in ordered_items
        ),
        "requestBytes": {
            "total": sum(request_sizes),
            "average": (
                round(sum(request_sizes) / len(request_sizes), 1)
                if request_sizes
                else 0
            ),
            "median": _nearest_rank_percentile(request_sizes, 50),
            "p95": _nearest_rank_percentile(request_sizes, 95),
            "p99": _nearest_rank_percentile(request_sizes, 99),
            "maximum": max(request_sizes, default=0),
            "oversizedThreshold": oversized_request_bytes,
            "oversizedCount": sum(
                size > oversized_request_bytes for size in request_sizes
            ),
        },
        "promptBytes": {
            "total": sum(prompt_sizes),
            "average": (
                round(sum(prompt_sizes) / len(prompt_sizes), 1)
                if prompt_sizes
                else 0
            ),
            "median": _nearest_rank_percentile(prompt_sizes, 50),
            "p95": _nearest_rank_percentile(prompt_sizes, 95),
            "p99": _nearest_rank_percentile(prompt_sizes, 99),
            "maximum": max(prompt_sizes, default=0),
            "oversizedThreshold": oversized_request_bytes,
            "oversizedCount": sum(
                size > oversized_request_bytes for size in prompt_sizes
            ),
        },
        "aggregateCheckCounts": dict(sorted(aggregate_status_counts.items())),
        "formulaStatusCounts": dict(sorted(formula_status_counts.items())),
        "formulaCount": sum(int(item["formulaCount"]) for item in ordered_items),
        "formulaNonNumericCount": sum(
            int(item["formulaNonNumericCount"]) for item in ordered_items
        ),
        "formulaErrorCount": sum(
            int(item["formulaErrorCount"]) for item in ordered_items
        ),
        "termDictionaryStatusCounts": dict(
            sorted(term_dictionary_status_counts.items())
        ),
        "outlierCount": len(outliers),
        "outliers": outliers,
        "largestRequests": largest_requests,
        "largestPrompts": largest_prompts,
        "items": ordered_items,
        "failures": ordered_failures,
    }


def cmd_table_first_request_batch(args: argparse.Namespace) -> int:
    """Build and audit table-first requests without invoking any AI analysis."""

    packet_dir = Path(args.packet_dir).expanduser().resolve()
    if not packet_dir.is_dir():
        raise SystemExit(f"Semantic packet directory not found: {packet_dir}")
    packet_paths = sorted(
        path for path in packet_dir.glob("*.json") if path.is_file()
    )
    if args.offset:
        packet_paths = packet_paths[max(0, args.offset) :]
    if args.limit > 0:
        packet_paths = packet_paths[: args.limit]
    output_dir = service_output_dir(
        args.out_dir,
        TABLE_FIRST_DIR / "request-batch",
    )
    request_dir = service_output_dir(None, output_dir / "requests")
    report_path = output_dir / "request-batch-report.json"
    workers = max(1, args.workers)
    oversized_request_bytes = max(1, args.oversized_request_bytes)
    checkpoint_every = max(1, args.checkpoint_every)
    started_at = now_iso()
    started = time.perf_counter()

    def process_one(index: int, packet_path: Path) -> dict[str, Any]:
        item_started = time.perf_counter()
        packet_set = _load_semantic_packet(packet_path)
        request = build_table_first_request(
            packet_set,
            max_preview_rows=args.max_preview_rows,
            max_preview_columns=args.max_preview_columns,
            max_value_samples=args.max_value_samples,
            term_dictionary_path=args.term_dictionary,
        )
        request_bytes = table_first_json_bytes(request)
        prompt_stats = table_first_prompt_stats(request)
        request_path = request_dir / f"{safe_name(request['requestId'])}.json"
        action = "PREPARED"
        if request_path.is_file():
            try:
                if request_path.read_bytes() == request_bytes:
                    action = "REUSED"
            except OSError:
                pass
        if action == "PREPARED":
            request_path.write_bytes(request_bytes)

        raw_frequency_exclusions = (
            (request.get("codeOwnedExclusions") or {}).get(
                "rawFrequencyResponseTables"
            )
            or []
        )
        workbook_learning_exclusion = (
            (request.get("codeOwnedExclusions") or {}).get(
                "workbookLearningExclusion"
            )
            or {}
        )
        formula = request.get("formulaDerivation") or {}
        aggregate_status_counts: Counter[str] = Counter()
        for table in request.get("tables") or []:
            for check in table.get("aggregateChecks") or []:
                aggregate_status_counts[str(check.get("status") or "UNKNOWN")] += 1
        outlier_reasons: list[str] = []
        if not request.get("tables"):
            outlier_reasons.append(
                "WORKBOOK_EXCLUDED_RAW_FREQUENCY_RESPONSE"
                if workbook_learning_exclusion.get("excluded") is True
                else "NO_TABLES"
            )
        if prompt_stats["promptBytes"] > oversized_request_bytes:
            outlier_reasons.append("OVERSIZED_PROMPT")
        if int(formula.get("errorCount") or 0):
            outlier_reasons.append("FORMULA_ERRORS")
        if aggregate_status_counts["MISMATCH"]:
            outlier_reasons.append("AGGREGATE_MISMATCH")
        return {
            "index": index,
            "packet": str(packet_path),
            "fileName": request["source"]["fileName"],
            "requestId": request["requestId"],
            "request": str(request_path),
            "action": action,
            "elapsedSeconds": round(time.perf_counter() - item_started, 3),
            "packetBytes": packet_path.stat().st_size,
            "requestBytes": len(request_bytes),
            **prompt_stats,
            "capturedPrimaryCells": int(
                request["workbook"]["capturedPrimaryCellCount"]
            ),
            "tableCount": len(request["tables"]),
            "textBlockCount": len(request["textBlocks"]),
            "rawFrequencyResponseExclusionCount": len(
                raw_frequency_exclusions
            ),
            "rawFrequencyResponseExcludedCellCount": sum(
                int(exclusion.get("sourceCellCount") or 0)
                for exclusion in raw_frequency_exclusions
            ),
            "workbookLearningExcluded": (
                workbook_learning_exclusion.get("excluded") is True
            ),
            "workbookLearningExclusionReason": str(
                workbook_learning_exclusion.get("reason") or ""
            ),
            "workbookLearningExcludedCellCount": int(
                workbook_learning_exclusion.get(
                    "excludedCapturedPrimaryCellCount"
                )
                or 0
            ),
            "aggregateCheckCounts": dict(sorted(aggregate_status_counts.items())),
            "formulaStatus": str(formula.get("status") or "UNKNOWN"),
            "formulaCount": int(formula.get("formulaCount") or 0),
            "formulaNonNumericCount": int(
                formula.get("nonNumericCount") or 0
            ),
            "formulaErrorCount": int(formula.get("errorCount") or 0),
            "formulaErrorSamples": list(formula.get("errorSamples") or []),
            "termDictionaryStatus": str(
                (request.get("codeOwnedTermDictionary") or {}).get("status")
                or "UNKNOWN"
            ),
            "termDictionaryContentSha256": str(
                (request.get("codeOwnedTermDictionary") or {}).get(
                    "contentSha256"
                )
                or ""
            ),
            "outlierReasons": outlier_reasons,
        }

    items: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(process_one, index, packet_path): (index, packet_path)
            for index, packet_path in enumerate(packet_paths, start=1)
        }
        for future in as_completed(futures):
            index, packet_path = futures[future]
            try:
                item = future.result()
            except Exception as exc:
                failures.append(
                    {
                        "index": index,
                        "packet": str(packet_path),
                        "error": str(exc),
                    }
                )
                progress_action = "FAILED"
            else:
                items.append(item)
                progress_action = str(item["action"])
            completed_count = len(items) + len(failures)
            if completed_count % checkpoint_every == 0:
                report = _table_first_request_batch_report(
                    packet_dir=packet_dir,
                    output_dir=output_dir,
                    selected=len(packet_paths),
                    workers=workers,
                    oversized_request_bytes=oversized_request_bytes,
                    started_at=started_at,
                    elapsed_seconds=time.perf_counter() - started,
                    items=items,
                    failures=failures,
                    completed=False,
                )
                write_json(report_path, report)
                print(
                    f"[{completed_count}/{len(packet_paths)}] "
                    f"{packet_path.name}: {progress_action}",
                    file=sys.stderr,
                    flush=True,
                )

    report = _table_first_request_batch_report(
        packet_dir=packet_dir,
        output_dir=output_dir,
        selected=len(packet_paths),
        workers=workers,
        oversized_request_bytes=oversized_request_bytes,
        started_at=started_at,
        elapsed_seconds=time.perf_counter() - started,
        items=items,
        failures=failures,
        completed=True,
    )
    write_json(report_path, report)
    print_json(
        {
            key: value
            for key, value in report.items()
            if key
            not in {"items", "outliers", "largestRequests", "failures"}
        }
        | {
            "report": str(report_path),
            "failureCount": len(report["failures"]),
        }
    )
    return 0 if not failures else 1


def _no_table_first_analysis(request: dict[str, Any]) -> dict[str, Any]:
    exclusions = (
        (request.get("codeOwnedExclusions") or {}).get(
            "rawFrequencyResponseTables"
        )
        or []
    )
    return validate_table_first_analysis(
        {
            "schemaVersion": TABLE_FIRST_ANALYSIS_SCHEMA_VERSION,
            "promptVersion": TABLE_FIRST_PROMPT_VERSION,
            "requestId": request["requestId"],
            "revisionUid": request["source"]["revisionUid"],
            "status": "NO_TABLES",
            "workbookSummary": (
                "No AI-analyzable table remains after deterministic raw "
                "frequency-response exclusion."
                if exclusions
                else "No table candidate was found."
            ),
            "tables": [],
            "notes": [],
        },
        request=request,
    )


def _adapt_reusable_table_first_analysis(
    existing: dict[str, Any],
    *,
    request: dict[str, Any],
    previous_request: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Adapt a stored analysis when builder-v4 only removed raw tables."""

    exclusions = (
        (request.get("codeOwnedExclusions") or {}).get(
            "rawFrequencyResponseTables"
        )
        or []
    )
    if existing.get("requestId") == request.get("requestId"):
        return existing
    if (
        existing.get("schemaVersion") != TABLE_FIRST_ANALYSIS_SCHEMA_VERSION
        or existing.get("promptVersion") != TABLE_FIRST_PROMPT_VERSION
        or existing.get("revisionUid")
        != (request.get("source") or {}).get("revisionUid")
    ):
        return existing

    existing_tables = existing.get("tables")
    if not isinstance(existing_tables, list) or any(
        not isinstance(table, dict) for table in existing_tables
    ):
        return existing
    by_id = {
        str(table.get("tableId") or ""): table for table in existing_tables
    }
    if len(by_id) != len(existing_tables):
        return existing
    expected_ids = [
        str(table["tableId"]) for table in request.get("tables") or []
    ]
    if any(table_id not in by_id for table_id in expected_ids):
        return existing

    safe_builder_upgrade = False
    if isinstance(previous_request, dict):
        previous_tables = {
            str(table.get("tableId") or ""): table
            for table in previous_request.get("tables") or []
            if isinstance(table, dict)
        }
        safe_builder_upgrade = bool(
            previous_request.get("schemaVersion")
            == TABLE_FIRST_REQUEST_SCHEMA_VERSION
            and previous_request.get("builderVersion")
            in {"table-first-builder-v3", "table-first-builder-v4"}
            and previous_request.get("requestId") == existing.get("requestId")
            and previous_request.get("source") == request.get("source")
            and previous_request.get("textBlocks") == request.get("textBlocks")
            and all(
                previous_tables.get(table_id) == table
                for table_id, table in zip(
                    expected_ids,
                    request.get("tables") or [],
                    strict=True,
                )
            )
        )
    if not safe_builder_upgrade and not exclusions:
        return existing

    adapted = copy.deepcopy(existing)
    adapted["requestId"] = request["requestId"]
    adapted["tables"] = [by_id[table_id] for table_id in expected_ids]
    allowed_table_ids = set(expected_ids)
    allowed_text_ids = {
        str(block["textId"]) for block in request.get("textBlocks") or []
    }
    for table in adapted["tables"]:
        table["relatedTableIds"] = [
            table_id
            for table_id in table.get("relatedTableIds") or []
            if str(table_id) in allowed_table_ids
            and str(table_id) != str(table.get("tableId") or "")
        ]
        table["textLinks"] = [
            text_id
            for text_id in table.get("textLinks") or []
            if str(text_id) in allowed_text_ids
        ]
    if not expected_ids:
        adapted["status"] = "NO_TABLES"
        adapted["workbookSummary"] = (
            "No AI-analyzable table remains after deterministic raw "
            "frequency-response exclusion."
        )
    return adapted


def cmd_table_first_analyze(args: argparse.Namespace) -> int:
    request_path = Path(args.request).expanduser().resolve()
    request = _load_table_first_request(request_path)
    output_path = service_output_path(
        args.out,
        TABLE_FIRST_DIR
        / "analyses"
        / f"{safe_name(request['requestId'])}.json",
    )
    ai_calls = 0
    if request.get("tables"):
        analysis = run_codex_table_first_analysis(
            request=request,
            output_path=output_path,
            model=args.model,
            reasoning_effort=args.reasoning_effort,
            timeout_seconds=args.timeout,
        )
        ai_calls = 1
    else:
        analysis = _no_table_first_analysis(request)
        output_path.write_bytes(table_first_json_bytes(analysis))
    projection = project_table_first_analysis(request, analysis)
    projection_path = service_output_path(
        args.projection_out,
        TABLE_FIRST_DIR
        / "projections"
        / f"{safe_name(request['requestId'])}.json",
    )
    projection_path.write_bytes(table_first_json_bytes(projection))
    print_json(
        {
            "status": "ok",
            "request": str(request_path),
            "analysis": str(output_path),
            "projection": str(projection_path),
            "aiCalls": ai_calls,
            "tables": len(analysis["tables"]),
            "studies": len(projection["studies"]),
            "verificationStatus": projection["verificationStatus"],
            "queryEligibility": projection["queryEligibility"],
        }
    )
    return 0


def _table_first_item_audit(
    request: dict[str, Any],
    analysis: dict[str, Any],
) -> dict[str, Any]:
    confidence_counts: Counter[str] = Counter()
    table_type_counts: Counter[str] = Counter()
    aggregate_status_counts: Counter[str] = Counter()
    identity_group_count = 0
    duplicate_group_label_count = 0
    aggregate_metric_count = 0
    metric_count = 0
    relation_count = 0
    source_tables = {
        str(table["tableId"]): table for table in request.get("tables") or []
    }
    low_confidence_tables: list[dict[str, Any]] = []
    identity_pattern = re.compile(
        r"(?:#?\d+(?:\.\d+)?|"
        r"(?:position|posistion|pos|nozzle|sample|specimen|replicate)"
        r"\s*#?\d+)",
        flags=re.IGNORECASE,
    )
    aggregate_metric_pattern = re.compile(
        r"(?:min|minimum|max|maximum|avg|average|mean)",
        flags=re.IGNORECASE,
    )
    for table in analysis.get("tables") or []:
        confidence_counts[str(table.get("confidence") or "UNKNOWN")] += 1
        table_type_counts[str(table.get("type") or "UNKNOWN")] += 1
        labels = [
            str(group.get("label") or "").strip()
            for group in table.get("groups") or []
            if isinstance(group, dict)
        ]
        duplicate_group_label_count += len(labels) - len(set(labels))
        identity_group_count += sum(
            bool(identity_pattern.fullmatch(label)) for label in labels
        )
        metrics = [
            metric
            for metric in table.get("metrics") or []
            if isinstance(metric, dict)
        ]
        metric_count += len(metrics)
        aggregate_metric_count += sum(
            bool(
                aggregate_metric_pattern.fullmatch(
                    str(metric.get("name") or "").strip()
                )
            )
            for metric in metrics
        )
        relation_count += len(table.get("comparisonRelations") or [])
        if table.get("confidence") == "LOW":
            source_table = source_tables.get(str(table.get("tableId") or ""), {})
            low_confidence_tables.append(
                {
                    "tableId": table.get("tableId"),
                    "title": table.get("title"),
                    "type": table.get("type"),
                    "sheet": source_table.get("sheet"),
                    "range": source_table.get("range"),
                    "limitations": list(table.get("limitations") or []),
                }
            )
    for table in request.get("tables") or []:
        for check in table.get("aggregateChecks") or []:
            aggregate_status_counts[str(check.get("status") or "UNKNOWN")] += 1

    formula = request.get("formulaDerivation") or {}
    formula_status = str(formula.get("status") or "UNKNOWN")
    review_reasons: list[str] = []
    if confidence_counts["LOW"]:
        review_reasons.append("LOW_CONFIDENCE")
    if aggregate_status_counts["MISMATCH"]:
        review_reasons.append("AGGREGATE_MISMATCH")
    if identity_group_count:
        review_reasons.append("IDENTITY_AXIS_AS_GROUP")
    if duplicate_group_label_count:
        review_reasons.append("DUPLICATE_GROUP_LABEL")
    if aggregate_metric_count:
        review_reasons.append("AGGREGATE_LABEL_AS_METRIC")

    return {
        "analysisStatus": analysis["status"],
        "confidenceCounts": dict(sorted(confidence_counts.items())),
        "tableTypeCounts": dict(sorted(table_type_counts.items())),
        "metricCount": metric_count,
        "comparisonRelationCount": relation_count,
        "aggregateCheckCounts": dict(sorted(aggregate_status_counts.items())),
        "formulaStatus": formula_status,
        "formulaCount": int(formula.get("formulaCount") or 0),
        "derivedFormulaCount": int(formula.get("numericCount") or 0),
        "nonNumericFormulaCount": int(
            formula.get("nonNumericCount") or 0
        ),
        "appliedFormulaCount": int(
            formula.get(
                "appliedNumericCount",
                formula.get("numericCount") or 0,
            )
        ),
        "formulaErrorCount": int(formula.get("errorCount") or 0),
        "formulaErrorSamples": list(formula.get("errorSamples") or []),
        "identityGroupCount": identity_group_count,
        "duplicateGroupLabelCount": duplicate_group_label_count,
        "aggregateMetricCount": aggregate_metric_count,
        "reviewRecommended": bool(review_reasons),
        "reviewReasons": review_reasons,
        "lowConfidenceTables": low_confidence_tables,
    }


def _table_first_batch_report(
    *,
    packet_dir: Path,
    output_dir: Path,
    selected: int,
    workers: int,
    started_at: str,
    elapsed_seconds: float,
    items: list[dict[str, Any]],
    failures: list[dict[str, Any]],
    completed: bool,
) -> dict[str, Any]:
    confidence_counts: Counter[str] = Counter()
    table_type_counts: Counter[str] = Counter()
    aggregate_status_counts: Counter[str] = Counter()
    formula_status_counts: Counter[str] = Counter()
    term_dictionary_status_counts: Counter[str] = Counter()
    request_sizes: list[int] = []
    for item in items:
        confidence_counts.update(item.get("confidenceCounts") or {})
        table_type_counts.update(item.get("tableTypeCounts") or {})
        aggregate_status_counts.update(item.get("aggregateCheckCounts") or {})
        formula_status_counts[str(item.get("formulaStatus") or "UNKNOWN")] += 1
        term_dictionary_status_counts[
            str(item.get("termDictionaryStatus") or "UNKNOWN")
        ] += 1
        request_sizes.append(int(item.get("requestBytes") or 0))
    ordered_items = sorted(items, key=lambda item: int(item["index"]))
    ordered_failures = sorted(failures, key=lambda item: int(item["index"]))
    outliers = [
        {
            "index": item["index"],
            "fileName": item["fileName"],
            "analysis": item["analysis"],
            "reasons": item["reviewReasons"],
            "lowConfidenceTables": item["lowConfidenceTables"],
            "formula": {
                "status": item["formulaStatus"],
                "formulaCount": item["formulaCount"],
                "appliedFormulaCount": item["appliedFormulaCount"],
                "nonNumericFormulaCount": item[
                    "nonNumericFormulaCount"
                ],
                "errorCount": item["formulaErrorCount"],
                "errorSamples": item["formulaErrorSamples"],
            },
        }
        for item in ordered_items
        if item.get("reviewRecommended")
    ]
    succeeded = len(ordered_items)
    return {
        "schemaVersion": "table-first-batch-report-v1",
        "status": (
            "running"
            if not completed
            else ("ok" if not ordered_failures else "partial")
        ),
        "builderVersion": TABLE_FIRST_BUILDER_VERSION,
        "promptVersion": TABLE_FIRST_PROMPT_VERSION,
        "packetDir": str(packet_dir),
        "outputDir": str(output_dir),
        "startedAt": started_at,
        "completedAt": now_iso() if completed else None,
        "elapsedSeconds": round(elapsed_seconds, 3),
        "workers": workers,
        "selected": selected,
        "completed": succeeded + len(ordered_failures),
        "succeeded": succeeded,
        "failed": len(ordered_failures),
        "newAnalyses": sum(item["action"] == "ANALYZED" for item in ordered_items),
        "reused": sum(item["action"] == "REUSED" for item in ordered_items),
        "noTables": sum(
            item["analysisStatus"] == "NO_TABLES" for item in ordered_items
        ),
        "aiCalls": sum(int(item["aiCalls"]) for item in ordered_items),
        "tables": sum(int(item["tableCount"]) for item in ordered_items),
        "textBlocks": sum(int(item["textBlockCount"]) for item in ordered_items),
        "rawFrequencyResponseExclusions": sum(
            int(item.get("rawFrequencyResponseExclusionCount") or 0)
            for item in ordered_items
        ),
        "workbooksWithRawFrequencyResponseExclusions": sum(
            bool(item.get("rawFrequencyResponseExclusionCount"))
            for item in ordered_items
        ),
        "workbooksExcludedFromLearning": sum(
            bool(item.get("workbookLearningExcluded"))
            for item in ordered_items
        ),
        "workbookLearningExcludedCells": sum(
            int(item.get("workbookLearningExcludedCellCount") or 0)
            for item in ordered_items
        ),
        "requestBytes": {
            "total": sum(request_sizes),
            "average": (
                round(sum(request_sizes) / len(request_sizes), 1)
                if request_sizes
                else 0
            ),
            "maximum": max(request_sizes, default=0),
        },
        "confidenceCounts": dict(sorted(confidence_counts.items())),
        "tableTypeCounts": dict(sorted(table_type_counts.items())),
        "aggregateCheckCounts": dict(sorted(aggregate_status_counts.items())),
        "formulaStatusCounts": dict(sorted(formula_status_counts.items())),
        "termDictionaryStatusCounts": dict(
            sorted(term_dictionary_status_counts.items())
        ),
        "formulaCount": sum(
            int(item.get("formulaCount") or 0) for item in ordered_items
        ),
        "derivedFormulaCount": sum(
            int(item.get("derivedFormulaCount") or 0)
            for item in ordered_items
        ),
        "appliedFormulaCount": sum(
            int(item.get("appliedFormulaCount") or 0)
            for item in ordered_items
        ),
        "nonNumericFormulaCount": sum(
            int(item.get("nonNumericFormulaCount") or 0)
            for item in ordered_items
        ),
        "formulaErrorCount": sum(
            int(item.get("formulaErrorCount") or 0)
            for item in ordered_items
        ),
        "reviewRecommended": len(outliers),
        "outliers": outliers,
        "items": ordered_items,
        "failures": ordered_failures,
    }


def cmd_table_first_batch(args: argparse.Namespace) -> int:
    packet_dir = Path(args.packet_dir).expanduser().resolve()
    if not packet_dir.is_dir():
        raise SystemExit(f"Semantic packet directory not found: {packet_dir}")
    packet_paths = sorted(
        path for path in packet_dir.glob("*.json") if path.is_file()
    )
    if args.offset:
        packet_paths = packet_paths[max(0, args.offset) :]
    if args.limit > 0:
        packet_paths = packet_paths[: args.limit]
    output_dir = service_output_dir(
        args.out_dir,
        TABLE_FIRST_DIR / "batch",
    )
    request_dir = service_output_dir(None, output_dir / "requests")
    analysis_dir = service_output_dir(None, output_dir / "analyses")
    projection_dir = service_output_dir(None, output_dir / "projections")
    report_path = output_dir / "batch-report.json"
    workers = max(1, args.workers)
    started_at = now_iso()
    started = time.perf_counter()

    def process_one(index: int, packet_path: Path) -> dict[str, Any]:
        item_started = time.perf_counter()
        packet_set = _load_semantic_packet(packet_path)
        request = build_table_first_request(
            packet_set,
            max_preview_rows=args.max_preview_rows,
            max_preview_columns=args.max_preview_columns,
            max_value_samples=args.max_value_samples,
            term_dictionary_path=args.term_dictionary,
        )
        stem = safe_name(packet_path.stem)
        request_path = request_dir / f"{stem}.json"
        analysis_path = analysis_dir / f"{stem}.json"
        projection_path = projection_dir / f"{stem}.json"
        previous_request: dict[str, Any] | None = None
        if request_path.is_file():
            try:
                loaded_previous_request = json.loads(
                    request_path.read_text(encoding="utf-8")
                )
            except (OSError, json.JSONDecodeError):
                pass
            else:
                if isinstance(loaded_previous_request, dict):
                    previous_request = loaded_previous_request
        request_bytes = table_first_json_bytes(request)
        request_path.write_bytes(request_bytes)

        analysis: dict[str, Any] | None = None
        action = "ANALYZED"
        ai_calls = 0
        if analysis_path.is_file() and not args.force:
            try:
                existing = json.loads(analysis_path.read_text(encoding="utf-8"))
                reusable_existing = _adapt_reusable_table_first_analysis(
                    existing,
                    request=request,
                    previous_request=previous_request,
                )
                normalized_existing = normalize_table_first_analysis(
                    reusable_existing,
                    request=request,
                )
                analysis = validate_table_first_analysis(
                    normalized_existing,
                    request=request,
                )
            except (OSError, json.JSONDecodeError, RuntimeError, ValueError):
                analysis = None
            else:
                action = "REUSED"
                if normalized_existing != existing:
                    analysis_path.write_bytes(table_first_json_bytes(analysis))
        if analysis is None:
            if request.get("tables"):
                analysis = run_codex_table_first_analysis(
                    request=request,
                    output_path=analysis_path,
                    model=args.model or None,
                    reasoning_effort=args.reasoning_effort or None,
                    timeout_seconds=args.timeout,
                )
                ai_calls = 1
            else:
                analysis = _no_table_first_analysis(request)
                analysis_path.write_bytes(table_first_json_bytes(analysis))
                action = "NO_TABLES"
        projection = project_table_first_analysis(request, analysis)
        projection_path.write_bytes(table_first_json_bytes(projection))
        audit = _table_first_item_audit(request, analysis)
        raw_frequency_exclusions = (
            (request.get("codeOwnedExclusions") or {}).get(
                "rawFrequencyResponseTables"
            )
            or []
        )
        workbook_learning_exclusion = (
            (request.get("codeOwnedExclusions") or {}).get(
                "workbookLearningExclusion"
            )
            or {}
        )
        return {
            "index": index,
            "packet": str(packet_path),
            "fileName": request["source"]["fileName"],
            "requestId": request["requestId"],
            "request": str(request_path),
            "analysis": str(analysis_path),
            "projection": str(projection_path),
            "action": action,
            "aiCalls": ai_calls,
            "elapsedSeconds": round(time.perf_counter() - item_started, 3),
            "requestBytes": len(request_bytes),
            "capturedPrimaryCells": int(
                request["workbook"]["capturedPrimaryCellCount"]
            ),
            "tableCount": len(request["tables"]),
            "textBlockCount": len(request["textBlocks"]),
            "rawFrequencyResponseExclusionCount": len(
                raw_frequency_exclusions
            ),
            "rawFrequencyResponseExcludedCellCount": sum(
                int(exclusion.get("sourceCellCount") or 0)
                for exclusion in raw_frequency_exclusions
            ),
            "workbookLearningExcluded": (
                workbook_learning_exclusion.get("excluded") is True
            ),
            "workbookLearningExclusionReason": str(
                workbook_learning_exclusion.get("reason") or ""
            ),
            "workbookLearningExcludedCellCount": int(
                workbook_learning_exclusion.get(
                    "excludedCapturedPrimaryCellCount"
                )
                or 0
            ),
            "termDictionaryStatus": str(
                (request.get("codeOwnedTermDictionary") or {}).get("status")
                or "UNKNOWN"
            ),
            "termDictionaryContentSha256": str(
                (request.get("codeOwnedTermDictionary") or {}).get(
                    "contentSha256"
                )
                or ""
            ),
            "studyCount": len(projection["studies"]),
            **audit,
        }

    items: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(process_one, index, packet_path): (
                index,
                packet_path,
            )
            for index, packet_path in enumerate(packet_paths, start=1)
        }
        for future in as_completed(futures):
            index, packet_path = futures[future]
            try:
                item = future.result()
            except Exception as exc:
                failures.append(
                    {
                        "index": index,
                        "packet": str(packet_path),
                        "error": str(exc),
                    }
                )
                progress_action = "FAILED"
            else:
                items.append(item)
                progress_action = str(item["action"])
            report = _table_first_batch_report(
                packet_dir=packet_dir,
                output_dir=output_dir,
                selected=len(packet_paths),
                workers=workers,
                started_at=started_at,
                elapsed_seconds=time.perf_counter() - started,
                items=items,
                failures=failures,
                completed=False,
            )
            write_json(report_path, report)
            print(
                f"[{report['completed']}/{report['selected']}] "
                f"{packet_path.name}: {progress_action}",
                file=sys.stderr,
                flush=True,
            )

    report = _table_first_batch_report(
        packet_dir=packet_dir,
        output_dir=output_dir,
        selected=len(packet_paths),
        workers=workers,
        started_at=started_at,
        elapsed_seconds=time.perf_counter() - started,
        items=items,
        failures=failures,
        completed=True,
    )
    write_json(report_path, report)
    print_json(
        {
            key: value
            for key, value in report.items()
            if key not in {"items", "outliers", "failures"}
        }
        | {
            "report": str(report_path),
            "reviewRecommended": report["reviewRecommended"],
            "failureCount": len(report["failures"]),
        }
    )
    return 0 if not failures else 1


def cmd_table_first_html(args: argparse.Namespace) -> int:
    batch_dir = Path(args.batch_dir).expanduser().resolve()
    if not batch_dir.is_dir():
        raise SystemExit(f"Table-first batch directory not found: {batch_dir}")
    output_dir = (
        Path(args.out_dir).expanduser().resolve()
        if args.out_dir
        else batch_dir / "html-report"
    )
    result = build_table_first_html_report(
        batch_dir=batch_dir,
        output_dir=output_dir,
    )
    print_json(result)
    return 0


def cmd_semantic_validate_draft(args: argparse.Namespace) -> int:
    input_path = Path(args.input).resolve()
    if not input_path.is_file():
        raise SystemExit(f"Semantic Study draft not found: {input_path}")
    packet_set = _load_semantic_packet(Path(args.packet).resolve())
    source = _semantic_source(packet_set, args.dataset)
    content_complete = bool(
        packet_set["inventory"].get("contentCompleteForManifest")
    )
    draft = json.loads(input_path.read_text(encoding="utf-8"))
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        canonical_revision = resolve_manifest_revision(conn, source)
        normalized = validate_ai_study_draft(
            draft,
            source=source,
            content_complete=content_complete,
            evidence_checker=make_database_evidence_checker(
                conn,
                canonical_revision,
            ),
        )
        validate_numeric_observation_evidence(
            conn,
            canonical_revision,
            normalized,
        )
        validate_factor_and_arm_evidence(
            conn,
            canonical_revision,
            normalized,
        )
        validate_comparison_representation_alignment(
            conn,
            canonical_revision,
            normalized,
        )
        validate_conclusion_evidence(
            conn,
            canonical_revision,
            normalized,
        )
    output_path = service_output_path(
        args.out,
        SEMANTIC_STUDY_DIR
        / f"{safe_name(source['revisionUid'])}.study-draft.json",
    )
    write_json(output_path, normalized)
    print_json(
        {
            "status": "ok",
            "input": str(input_path),
            "manifest": str(output_path),
            "studies": len(normalized["studies"]),
            "verificationStatus": normalized["workbookAnalysis"][
                "verificationStatus"
            ],
        }
    )
    return 0


def cmd_evidence_query(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    evidence_pack = build_evidence_pack_from_db(db_path, args.question)
    if args.out:
        output_path = database_scoped_output_path(
            args.out,
            OUTPUT_DIR / "evidence-packs" / "query.json",
            db_path,
        )
        output_path.write_text(
            json.dumps(evidence_pack, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print_json(
            {
                "status": "ok",
                "db": str(db_path),
                "evidencePack": str(output_path),
                **evidence_pack["summary"],
            }
        )
    else:
        print_json(evidence_pack)
    return 0


def cmd_evidence_detail(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    detail = build_evidence_detail_from_db(
        db_path,
        args.evidence_id,
    )
    if args.out:
        output_path = database_scoped_output_path(
            args.out,
            OUTPUT_DIR / "evidence-details" / f"{safe_name(args.evidence_id)}.json",
            db_path,
        )
        write_json(output_path, detail)
        print_json(
            {
                "status": "ok",
                "db": str(db_path),
                "publicEvidenceId": detail["publicEvidenceId"],
                "evidenceDetail": str(output_path),
                "trustStatus": detail["trust"]["status"],
                "capturedCellCountInRange": detail["preview"][
                    "capturedCellCountInRange"
                ],
            }
        )
    else:
        print_json(detail)
    return 0


def cmd_related_studies(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    report = build_related_studies_from_db(
        db_path,
        args.target,
        limit=args.limit,
    )
    if args.out:
        output_path = database_scoped_output_path(
            args.out,
            OUTPUT_DIR
            / "related-studies"
            / f"{safe_name(report['targetIdentifier'])}.json",
            db_path,
        )
        output_path.write_bytes(related_studies_json_bytes(report))
        print_json(
            {
                "status": "ok",
                "db": str(db_path),
                "relatedStudies": str(output_path),
                **report["summary"],
                "imagesAnalyzed": False,
            }
        )
    else:
        print_json(report)
    return 0


def cmd_table_first_history_index(args: argparse.Namespace) -> int:
    report = build_history_index(
        Path(args.batch_dir).expanduser().resolve(),
        Path(args.db).expanduser().resolve(),
        require_complete=not args.allow_running,
    )
    print_json(report)
    return 0


def cmd_table_first_history_query(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    pack = build_history_pack(
        db_path,
        args.question,
        limit=args.limit,
    )
    answer = build_history_answer(pack)
    question_hash = hashlib.sha256(args.question.encode("utf-8")).hexdigest()[:16]
    pack_path = service_output_path(
        args.out_pack,
        OUTPUT_DIR / "table-first-history-answers" / f"{question_hash}.pack.json",
    )
    answer_path = service_output_path(
        args.out_json,
        OUTPUT_DIR / "table-first-history-answers" / f"{question_hash}.answer.json",
    )
    markdown_path = service_output_path(
        args.out_markdown,
        OUTPUT_DIR / "table-first-history-answers" / f"{question_hash}.answer.md",
    )
    pack_path.write_bytes(history_json_bytes(pack))
    answer_path.write_bytes(history_json_bytes(answer))
    markdown_path.write_text(
        render_history_answer_markdown(answer),
        encoding="utf-8",
    )
    print_json(
        {
            "status": "ok",
            "database": str(db_path),
            "answerStatus": answer["answerStatus"],
            "evidencePack": str(pack_path),
            "answerJson": str(answer_path),
            "answerMarkdown": str(markdown_path),
            **answer["coverage"],
        }
    )
    return 0


def cmd_table_first_contextual_query(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    request = build_contextual_query_request(
        db_path,
        args.question,
        candidate_limit=args.candidate_limit,
        detail_candidate_limit=args.detail_candidate_limit,
        max_fact_count=args.max_fact_count,
    )
    question_hash = hashlib.sha256(args.question.encode("utf-8")).hexdigest()[:16]
    request_path = service_output_path(
        args.out_request,
        OUTPUT_DIR
        / "table-first-contextual-answers"
        / f"{question_hash}.request.json",
    )
    answer_path = service_output_path(
        args.out_json,
        OUTPUT_DIR
        / "table-first-contextual-answers"
        / f"{question_hash}.answer.json",
    )
    markdown_path = service_output_path(
        args.out_markdown,
        OUTPUT_DIR
        / "table-first-contextual-answers"
        / f"{question_hash}.answer.md",
    )
    request_path.write_bytes(contextual_json_bytes(request))
    answer = run_codex_contextual_query(
        request=request,
        output_path=answer_path,
        model=args.model,
        reasoning_effort=args.reasoning_effort,
        timeout_seconds=args.timeout_seconds,
    )
    markdown_path.write_text(
        render_contextual_answer_markdown(answer),
        encoding="utf-8",
    )
    print_json(
        {
            "status": "ok",
            "database": str(db_path),
            "answerStatus": answer["answerStatus"],
            "contextRequest": str(request_path),
            "answerJson": str(answer_path),
            "answerMarkdown": str(markdown_path),
            **answer["coverage"],
        }
    )
    return 0


def cmd_table_first_relevance_query(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    request = build_relevance_query_request(
        db_path,
        args.question,
        candidate_limit=args.candidate_limit,
    )
    question_hash = hashlib.sha256(args.question.encode("utf-8")).hexdigest()[:16]
    request_path = service_output_path(
        args.out_request,
        OUTPUT_DIR
        / "table-first-relevance-answers"
        / f"{question_hash}.request.json",
    )
    answer_path = service_output_path(
        args.out_json,
        OUTPUT_DIR
        / "table-first-relevance-answers"
        / f"{question_hash}.answer.json",
    )
    markdown_path = service_output_path(
        args.out_markdown,
        OUTPUT_DIR
        / "table-first-relevance-answers"
        / f"{question_hash}.answer.md",
    )
    request_path.write_bytes(relevance_json_bytes(request))
    result = run_codex_relevance_query(
        request=request,
        output_path=answer_path,
        model=args.model,
        reasoning_effort=args.reasoning_effort,
        timeout_seconds=args.timeout_seconds,
    )
    markdown_path.write_text(
        render_relevance_result_markdown(result),
        encoding="utf-8",
    )
    print_json(
        {
            "status": "ok",
            "database": str(db_path),
            "answerStatus": result["answerStatus"],
            "relevanceRequest": str(request_path),
            "answerJson": str(answer_path),
            "answerMarkdown": str(markdown_path),
            **result["coverage"],
        }
    )
    return 0


def cmd_table_first_history_detail(args: argparse.Namespace) -> int:
    detail = build_history_detail(
        Path(args.db).expanduser().resolve(),
        args.evidence_id,
    )
    if args.out:
        output_path = service_output_path(
            args.out,
            OUTPUT_DIR
            / "table-first-history-details"
            / f"{safe_name(args.evidence_id)}.json",
        )
        output_path.write_bytes(history_json_bytes(detail))
        print_json(
            {
                "status": "ok",
                "database": str(Path(args.db).expanduser().resolve()),
                "publicEvidenceId": detail["publicEvidenceId"],
                "evidenceDetail": str(output_path),
                "trustStatus": detail["trust"]["status"],
            }
        )
    else:
        print_json(detail)
    return 0


def cmd_validate_table_first_history_answer(args: argparse.Namespace) -> int:
    pack = _load_json_object(args.pack, label="Table-first history pack")
    answer = _load_json_object(args.answer, label="Table-first history answer")
    validate_history_answer(answer, pack)
    print_json(
        {
            "status": "ok",
            "pack": str(Path(args.pack).expanduser().resolve()),
            "answer": str(Path(args.answer).expanduser().resolve()),
            "answerStatus": answer["answerStatus"],
            "evidencePackSha256": answer["evidencePackSha256"],
        }
    )
    return 0


def cmd_table_first_history_acceptance(args: argparse.Namespace) -> int:
    report = run_history_acceptance(
        Path(args.db).expanduser().resolve(),
        Path(args.manifest).expanduser().resolve(),
        query_limit=args.limit,
    )
    output_path = service_output_path(
        args.out,
        OUTPUT_DIR
        / "table-first-history"
        / "history-acceptance-report.json",
    )
    output_path.write_bytes(history_json_bytes(report))
    print_json(
        {
            "status": report["status"],
            "database": report["database"],
            "acceptanceReport": str(output_path),
            **report["summary"],
        }
    )
    return 0 if report["status"] == "PASS" else 1


def cmd_concept_candidates(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    try:
        with connect_ro(db_path) as connection:
            result = list_schema_candidates(
                connection,
                status=args.status,
                candidate_kind=args.kind,
                query=args.query,
                limit=args.limit,
            )
    except ConceptCurationError as exc:
        raise SystemExit(f"Concept curation unavailable: {exc}") from exc
    print_json(result)
    return 0


def cmd_concept_list(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    try:
        with connect_ro(db_path) as connection:
            result = list_canonical_concepts(
                connection,
                concept_kind=args.kind,
                lifecycle_status=args.status,
                query=args.query,
                limit=args.limit,
            )
    except ConceptCurationError as exc:
        raise SystemExit(f"Concept curation unavailable: {exc}") from exc
    print_json(result)
    return 0


def cmd_concept_resolve(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    if not db_path.is_file():
        raise SystemExit(f"Canonical DB not found: {db_path}")
    try:
        with connect_rw(db_path) as connection:
            result = resolve_schema_candidate(
                connection,
                candidate_uid=args.candidate_uid,
                action=args.action,
                reviewer=args.reviewer,
                note=args.note,
                now_iso=now_iso,
                canonical_name=args.canonical_name,
                concept_uid=args.concept_uid,
                alias=args.alias,
            )
            connection.commit()
    except ConceptCurationError as exc:
        raise SystemExit(f"Concept resolution rejected: {exc}") from exc
    print_json(result)
    return 0


def cmd_concept_alias_upsert(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    if not db_path.is_file():
        raise SystemExit(f"Canonical DB not found: {db_path}")
    try:
        with connect_rw(db_path) as connection:
            result = upsert_human_concept_alias(
                connection,
                concept_uid=args.concept_uid,
                alias=args.alias,
                reviewer=args.reviewer,
                note=args.note,
                now_iso=now_iso,
            )
            connection.commit()
    except ConceptCurationError as exc:
        raise SystemExit(f"Concept alias rejected: {exc}") from exc
    print_json(result)
    return 0


def cmd_review_queue(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    with connect_ro(db_path) as connection:
        queue = list_review_queue(connection, limit=args.limit)
    if args.out:
        output_path = service_output_path(
            args.out,
            OUTPUT_DIR / "human-review" / "queue.json",
        )
        write_json(output_path, queue)
        print_json(
            {
                "status": "ok",
                "db": str(db_path),
                "reviewQueue": str(output_path),
                "count": queue["count"],
                "imagesAnalyzed": False,
            }
        )
    else:
        print_json(queue)
    return 0


def cmd_review_detail(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    with connect_ro(db_path) as connection:
        detail = get_review_detail(connection, args.comparison_id)
    if args.out:
        output_path = service_output_path(
            args.out,
            OUTPUT_DIR
            / "human-review"
            / f"{safe_name(args.comparison_id)}.json",
        )
        write_json(output_path, detail)
        print_json(
            {
                "status": "ok",
                "db": str(db_path),
                "reviewDetail": str(output_path),
                "publicComparisonId": detail["publicComparisonId"],
                "approvalReady": detail["approvalReadiness"]["ready"],
                "imagesAnalyzed": False,
            }
        )
    else:
        print_json(detail)
    return 0


def cmd_review_decide(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    try:
        with connect_rw(db_path) as connection:
            connection.execute("PRAGMA foreign_keys=ON")
            result = decide_comparison(
                connection,
                args.comparison_id,
                decision=args.decision,
                reviewer=args.reviewer,
                reason=args.reason,
                study_comparability_status=args.study_comparability,
                study_confounding_status=args.study_confounding,
                comparison_validity_status=args.comparison_validity,
                comparison_confounding_status=args.comparison_confounding,
                matching_basis=args.matching_basis,
            )
            connection.commit()
    except ReviewGateError as exc:
        raise SystemExit(f"Review decision rejected: {exc}") from exc
    print_json(result)
    return 0


def cmd_golden_acceptance(args: argparse.Namespace) -> int:
    db_path = Path(args.db).expanduser().resolve()
    manifest_path = Path(args.manifest).expanduser().resolve()
    output_directory = service_output_dir(
        args.out_dir,
        OUTPUT_DIR / "golden-acceptance" / "current",
    )
    report = run_golden_question_acceptance(
        db_path,
        manifest_path,
        output_directory,
    )
    print_json(
        {
            "status": "ok",
            "db": str(db_path),
            "manifest": str(manifest_path),
            "acceptanceReport": str(
                output_directory / "acceptance-report.json"
            ),
            "overallStatus": report["overallStatus"],
            **report["summary"],
            "imagesAnalyzed": False,
        }
    )
    return 0


def _load_json_object(path: str | Path, *, label: str) -> dict[str, Any]:
    input_path = Path(path).expanduser().resolve()
    try:
        value = json.loads(input_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise SystemExit(f"{label} does not exist: {input_path}") from exc
    except json.JSONDecodeError as exc:
        raise SystemExit(f"{label} is not valid JSON: {input_path}") from exc
    if not isinstance(value, dict):
        raise SystemExit(f"{label} must contain one JSON object: {input_path}")
    return value


def cmd_evidence_answer(args: argparse.Namespace) -> int:
    database_path: Path | None = None
    if args.pack:
        if args.question:
            raise SystemExit("--question cannot be combined with --pack.")
        evidence_pack = _load_json_object(args.pack, label="Evidence pack")
        source_label = str(Path(args.pack).expanduser().resolve())
    else:
        if not args.db or not args.question:
            raise SystemExit(
                "Use either --pack, or both --db and --question."
            )
        database_path = Path(args.db).expanduser().resolve()
        evidence_pack = build_evidence_pack_from_db(
            database_path,
            args.question,
        )
        source_label = str(database_path)

    answer = build_evidence_answer(evidence_pack)
    validate_evidence_answer(answer, evidence_pack)
    question_hash = hashlib.sha256(
        str(evidence_pack["question"]).encode("utf-8")
    ).hexdigest()[:16]
    default_json_path = (
        OUTPUT_DIR / "evidence-answers" / f"{question_hash}.answer.json"
    )
    default_markdown_path = (
        OUTPUT_DIR / "evidence-answers" / f"{question_hash}.answer.md"
    )
    if database_path is None:
        json_path = service_output_path(args.out_json, default_json_path)
        markdown_path = service_output_path(
            args.out_markdown,
            default_markdown_path,
        )
    else:
        json_path = database_scoped_output_path(
            args.out_json,
            default_json_path,
            database_path,
        )
        markdown_path = database_scoped_output_path(
            args.out_markdown,
            default_markdown_path,
            database_path,
        )
    json_path.write_bytes(answer_json_bytes(answer))
    markdown_path.write_text(
        render_answer_markdown(answer),
        encoding="utf-8",
    )
    print_json(
        {
            "status": "ok",
            "source": source_label,
            "answerStatus": answer["answerStatus"],
            "answerJson": str(json_path),
            "answerMarkdown": str(markdown_path),
            **answer["coverage"],
        }
    )
    return 0


def cmd_validate_evidence_answer(args: argparse.Namespace) -> int:
    evidence_pack = _load_json_object(args.pack, label="Evidence pack")
    answer_path = Path(args.answer).expanduser().resolve()
    answer = _load_json_object(answer_path, label="Evidence answer")
    validate_evidence_answer(answer, evidence_pack)
    print_json(
        {
            "status": "ok",
            "answer": str(answer_path),
            "answerStatus": answer["answerStatus"],
            "evidencePackSha256": answer["evidencePackSha256"],
        }
    )
    return 0


def _print_ingest_progress(event: dict[str, Any]) -> None:
    print(
        "PROGRESS_JSON "
        + json.dumps(
            event,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ),
        file=sys.stderr,
        flush=True,
    )


def cmd_form_preflight(args: argparse.Namespace) -> int:
    output_root = (
        Path(args.output_root).expanduser().resolve()
        if args.output_root
        else SERVICE_DIR.resolve()
    )
    output_path = output_path_under_root(
        args.out,
        output_root / "form-preflight" / "latest.json",
        output_root,
    )
    cancel_path = (
        output_path_under_root(
            args.cancel_file,
            output_root / "form-preflight" / "cancel.request",
            output_root,
        )
        if args.cancel_file
        else None
    )
    result = run_form_preflight(
        database_path=Path(args.db).expanduser().resolve(),
        source_root=Path(args.input).expanduser().resolve(),
        output_path=output_path,
        dataset=args.dataset,
        com_timeout_seconds=args.com_timeout_seconds,
        progress_callback=_print_ingest_progress,
        inspect_auth_dialog=args.inspect_auth_dialog,
        dismiss_auth_dialog=args.dismiss_auth_dialog,
        auth_dialog_title=args.auth_dialog_title,
        auth_dialog_class=args.auth_dialog_class,
        auth_dialog_button=args.auth_dialog_button,
        cancel_file=cancel_path,
        retry_failed_captures=args.retry_failed_captures,
    )
    print_json(
        {
            "status": result["status"],
            "report": str(output_path),
            "manifest": result["knownFormManifestPath"],
            **result["summary"],
        }
    )
    return 0


def _form_registry_paths(
    args: argparse.Namespace,
) -> tuple[Path, Path, Path]:
    output_root = (
        Path(args.output_root).expanduser().resolve()
        if args.output_root
        else SERVICE_DIR.resolve()
    )
    report_path = output_path_under_root(
        args.report,
        output_root / "form-preflight" / "latest.json",
        output_root,
    )
    review_path = output_path_under_root(
        getattr(args, "review_out", None),
        output_root / "form-preflight" / "group-review.latest.json",
        output_root,
    )
    return output_root, report_path, review_path


def cmd_form_group_review(args: argparse.Namespace) -> int:
    _, report_path, review_path = _form_registry_paths(args)
    if not report_path.is_file():
        raise SystemExit(
            "Form preflight report does not exist: "
            + str(report_path)
        )
    review = write_form_group_review(
        database_path=Path(args.db).expanduser().resolve(),
        report_path=report_path,
        output_path=review_path,
    )
    print_json(
        {
            "status": "COMPLETED",
            "review": str(review_path),
            **review["summary"],
        }
    )
    return 0


def cmd_form_family_analyze(args: argparse.Namespace) -> int:
    output_root, report_path, review_path = _form_registry_paths(args)
    if not report_path.is_file():
        raise SystemExit(
            "Form preflight report does not exist: "
            + str(report_path)
        )
    safe_family = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        args.family_id,
    ).strip("._")
    if not safe_family:
        raise SystemExit("family-id must contain a safe identifier.")
    contract_path = output_path_under_root(
        args.out,
        (
            output_root
            / "form-preflight"
            / "contracts"
            / f"{safe_family}.json"
        ),
        output_root,
    )
    result = analyze_form_family(
        database_path=Path(args.db).expanduser().resolve(),
        report_path=report_path,
        family_id=args.family_id,
        output_path=contract_path,
        codex_executable=args.codex,
        reasoning_effort=args.reasoning_effort,
        timeout_seconds=args.timeout,
    )
    write_form_group_review(
        database_path=Path(args.db).expanduser().resolve(),
        report_path=report_path,
        output_path=review_path,
    )
    print_json(
        {
            **result,
            "review": str(review_path),
        }
    )
    return 0


def cmd_form_family_decide(args: argparse.Namespace) -> int:
    _, report_path, review_path = _form_registry_paths(args)
    if not report_path.is_file():
        raise SystemExit(
            "Form preflight report does not exist: "
            + str(report_path)
        )
    result = decide_form_family(
        database_path=Path(args.db).expanduser().resolve(),
        report_path=report_path,
        family_id=args.family_id,
        decision=args.decision,
        reviewer=args.reviewer,
        display_name=args.display_name,
        linked_form_signature_id=args.linked_form_signature_id,
        notes=args.notes,
    )
    report = reclassify_form_preflight_report(
        database_path=Path(args.db).expanduser().resolve(),
        report_path=report_path,
    )
    review = write_form_group_review(
        database_path=Path(args.db).expanduser().resolve(),
        report_path=report_path,
        output_path=review_path,
    )
    print_json(
        {
            **result,
            "report": str(report_path),
            "manifest": report["knownFormManifestPath"],
            "review": str(review_path),
            "knownForms": report["summary"]["knownForms"],
            "pendingGroups": review["summary"]["pendingCount"],
        }
    )
    return 0


def cmd_form_pipeline_complete(args: argparse.Namespace) -> int:
    output_root = (
        Path(args.output_root).expanduser().resolve()
        if args.output_root
        else SERVICE_DIR.resolve()
    )
    result = run_form_pipeline_complete(
        database_path=Path(args.db).expanduser().resolve(),
        source_root=Path(args.input).expanduser().resolve(),
        output_root=output_root,
        reviewer=args.reviewer,
        dataset=args.dataset,
        analysis_workers=args.analysis_workers,
        reasoning_effort=args.reasoning_effort,
        analysis_timeout_seconds=args.analysis_timeout,
        com_timeout_seconds=args.com_timeout_seconds,
        codex_executable=args.codex,
        exclude_on_analysis_error=args.exclude_on_analysis_error,
        run_corpus=not args.skip_corpus,
        max_families=args.max_families,
        draft_monolithic_max_bytes=(
            args.draft_monolithic_max_bytes
        ),
        progress_callback=_print_ingest_progress,
    )
    print_json(
        {
            "status": result["status"],
            "result": result["resultPath"],
            "report": result["reportPath"],
            "review": result["reviewPath"],
            "manifest": result["manifestPath"],
            "preflight": result["preflightSummary"],
            "formGroups": result["reviewSummary"],
            "analysisErrors": len(result["errors"]),
            "corpus": (
                result["corpus"]["summary"]
                if isinstance(result.get("corpus"), dict)
                else None
            ),
        }
    )
    return 0


def cmd_ingest_workbook(args: argparse.Namespace) -> int:
    result = ingest_workbook(
        database_path=Path(args.db).expanduser().resolve(),
        source_path=Path(args.input).expanduser().resolve(),
        artifact_root=service_output_dir(
            args.artifact_root,
            OUTPUT_DIR / "incremental-ingest",
        ),
        dataset=args.dataset,
        resume=not args.no_resume,
        max_cells=args.max_cells,
        max_rows=args.max_rows,
        empty_row_gap=args.empty_row_gap,
        locator_workers=args.workers,
        locator_batch_size=args.batch_size,
        locator_batch_max_bytes=args.batch_max_bytes,
        draft_monolithic_max_bytes=args.draft_monolithic_max_bytes,
        draft_fragment_max_chunks=args.draft_fragment_max_chunks,
        draft_fragment_max_cells=args.draft_fragment_max_cells,
        draft_fragment_max_bytes=args.draft_fragment_max_bytes,
        draft_fragment_workers=args.draft_fragment_workers,
        derive_formula_values=args.derive_formula_values,
        repair_rejected_draft=args.repair_rejected_draft,
        repair_unselected_source=args.repair_unselected_source,
        model=args.model,
        reasoning_effort=args.reasoning_effort,
        locator_timeout_seconds=args.locator_timeout,
        draft_timeout_seconds=args.draft_timeout,
        capture_backend=args.capture_backend,
        covered_cell_mode=args.covered_cell_mode,
        include_hidden_sheets=args.include_hidden_sheets,
        inspect_auth_dialog=args.inspect_auth_dialog,
        dismiss_auth_dialog=args.dismiss_auth_dialog,
        auth_dialog_title=args.auth_dialog_title,
        auth_dialog_class=args.auth_dialog_class,
        auth_dialog_button=args.auth_dialog_button,
        auth_dialog_timeout_seconds=args.auth_dialog_timeout,
        progress_callback=_print_ingest_progress,
    )
    print_json(result)
    return 0


def cmd_ingest_corpus(args: argparse.Namespace) -> int:
    artifact_root = service_output_dir(
        args.artifact_root,
        OUTPUT_DIR / "corpus-ingest",
    )
    journal_path = (
        service_output_path(
            args.journal,
            artifact_root / "corpus-journal.json",
        )
        if args.journal
        else artifact_root / "corpus-journal.json"
    )
    include_relative_paths = None
    if args.source_manifest:
        manifest = _load_json_object(
            args.source_manifest,
            label="Corpus source manifest",
        )
        workbooks = manifest.get("workbooks")
        if not isinstance(workbooks, list):
            raise SystemExit(
                "Corpus source manifest must contain a workbooks array."
            )
        include_relative_paths = []
        source_root = Path(args.input).expanduser().resolve()
        for index, workbook in enumerate(workbooks):
            if not isinstance(workbook, dict) or not str(
                workbook.get("relativePath") or ""
            ).strip():
                raise SystemExit(
                    "Corpus source manifest workbooks"
                    f"[{index}].relativePath is required."
                )
            relative_path = str(workbook["relativePath"]).strip()
            include_relative_paths.append(relative_path)
            expected_sha256 = str(
                workbook.get("contentSha256") or ""
            ).strip().lower()
            if expected_sha256:
                source_path = (
                    source_root / Path(relative_path)
                ).resolve()
                try:
                    source_path.relative_to(source_root)
                except ValueError as exc:
                    raise SystemExit(
                        "Corpus source manifest path escapes the input root: "
                        + relative_path
                    ) from exc
                if not source_path.is_file():
                    raise SystemExit(
                        "Corpus source manifest file is missing: "
                        + str(source_path)
                    )
                digest = hashlib.sha256()
                with source_path.open("rb") as stream:
                    for chunk in iter(
                        lambda: stream.read(1024 * 1024),
                        b"",
                    ):
                        digest.update(chunk)
                if digest.hexdigest() != expected_sha256:
                    raise SystemExit(
                        "Corpus source changed after form preflight: "
                        + relative_path
                    )
    result = run_corpus_ingest(
        database_path=Path(args.db).expanduser().resolve(),
        source_root=Path(args.input).expanduser().resolve(),
        artifact_root=artifact_root,
        journal_path=journal_path,
        dataset=args.dataset,
        resume=not args.no_resume,
        retry_failed=args.retry_failed,
        inventory_only=args.inventory_only,
        include_relative_paths=include_relative_paths,
        offset=args.offset,
        limit=args.limit,
        workbook_workers=args.workbook_workers,
        com_workers=args.com_workers,
        packet_workers=args.packet_workers,
        ai_workers=args.ai_workers,
        db_workers=args.db_workers,
        ingest_options={
            "max_cells": args.max_cells,
            "max_rows": args.max_rows,
            "empty_row_gap": args.empty_row_gap,
            "locator_workers": args.locator_workers,
            "locator_batch_size": args.batch_size,
            "locator_batch_max_bytes": args.batch_max_bytes,
            "draft_monolithic_max_bytes": (
                args.draft_monolithic_max_bytes
            ),
            "draft_fragment_max_chunks": (
                args.draft_fragment_max_chunks
            ),
            "draft_fragment_max_cells": args.draft_fragment_max_cells,
            "draft_fragment_max_bytes": args.draft_fragment_max_bytes,
            "draft_fragment_workers": args.draft_fragment_workers,
            "derive_formula_values": args.derive_formula_values,
            "repair_rejected_draft": args.repair_rejected_draft,
            "repair_unselected_source": args.repair_unselected_source,
            "model": args.model,
            "reasoning_effort": args.reasoning_effort,
            "locator_timeout_seconds": args.locator_timeout,
            "draft_timeout_seconds": args.draft_timeout,
            "capture_backend": args.capture_backend,
            "covered_cell_mode": args.covered_cell_mode,
            "include_hidden_sheets": args.include_hidden_sheets,
            "inspect_auth_dialog": args.inspect_auth_dialog,
            "dismiss_auth_dialog": args.dismiss_auth_dialog,
            "auth_dialog_title": args.auth_dialog_title,
            "auth_dialog_class": args.auth_dialog_class,
            "auth_dialog_button": args.auth_dialog_button,
            "auth_dialog_timeout_seconds": args.auth_dialog_timeout,
            "progress_callback": _print_ingest_progress,
        },
    )
    result_path = service_output_path(
        args.out,
        artifact_root / f"{safe_name(result['runId'])}.result.json",
    )
    write_json(result_path, result)
    print_json(
        {
            "status": result["status"],
            "sourceRoot": result["sourceRoot"],
            "journal": result["journalPath"],
            "result": str(result_path),
            "imagesAnalyzed": False,
            **result["summary"],
        }
    )
    return 0


def cmd_quick_index(args: argparse.Namespace) -> int:
    dataset = args.dataset or DEFAULT_DATASET
    input_path = Path(args.input).resolve()
    db_path = service_output_path(args.db, QUICK_INDEX_DIR / f"{safe_name(dataset)}.sqlite")
    html_path = service_output_path(args.html, QUICK_INDEX_DIR / f"{safe_name(dataset)}_dashboard.html")
    log_path = service_output_path(args.log, LOG_DIR / f"{safe_name(dataset)}_quick_index.log")

    if not MICROSPEAKER_INDEXER.exists():
        raise SystemExit(f"Indexer not found: {MICROSPEAKER_INDEXER}")

    cmd = [
        sys.executable,
        str(MICROSPEAKER_INDEXER),
        "--input-file" if input_path.is_file() else "--input-dir",
        str(input_path),
        "--dataset",
        dataset,
        "--db",
        str(db_path),
        "--html",
        str(html_path),
    ]
    if args.limit > 0:
        cmd += ["--limit", str(args.limit)]
    if args.force:
        cmd.append("--force")
    if args.no_html:
        cmd.append("--no-html")

    code = run_command(cmd, log_path, cwd=MICROSPEAKER_INDEXER.parent)
    if code != 0:
        return code
    print_json({"status": "ok", "db": str(db_path), "html": "" if args.no_html else str(html_path), "log": str(log_path)})
    return 0


def cmd_com_index(args: argparse.Namespace) -> int:
    dataset = args.dataset or DEFAULT_DATASET
    input_path = Path(args.input).resolve()
    files = collect_excel_files(input_path)
    if not files:
        raise SystemExit(f"No Excel files found: {input_path}")
    if not COM_EXTRACTOR.exists():
        raise SystemExit(f"COM extractor not found: {COM_EXTRACTOR}")

    db_path = service_output_path(args.db, UNIVERSAL_GRID_DIR / f"{safe_name(dataset)}.sqlite")
    raw_dir = service_output_dir(args.raw_dir, UNIVERSAL_GRID_DIR / "raw-json" / safe_name(dataset))
    log_path = service_output_path(args.log, LOG_DIR / f"{safe_name(dataset)}_com_index.log")

    options = {
        "coveredCellMode": args.covered_cell_mode,
        "includeHidden": bool(args.include_hidden),
        "sparse": bool(args.sparse),
        "force": bool(args.force),
        "reuseJson": bool(args.reuse_json),
        "verifyAfterImport": bool(args.verify_after_import),
    }
    succeeded = 0
    failed = 0
    skipped = 0
    with connect_rw(db_path) as conn:
        ensure_universal_schema(conn)
        current: list[tuple[Path, dict[str, Any]]] = []
        pending: list[tuple[Path, dict[str, Any] | None]] = []
        for source in files:
            existing = workbook_is_current(conn, dataset, source) if not args.force else None
            if existing is not None:
                current.append((source, existing))
            else:
                pending.append((source, current_workbook_for_source(conn, dataset, source)))

        targets = pending[: args.limit] if args.limit > 0 else pending
        deferred = len(pending) - len(targets)
        run_id = conn.execute(
            """
            INSERT INTO runs(dataset, input_path, started_at, total_files, options_json)
            VALUES (?, ?, ?, ?, ?)
            """,
            (dataset, str(input_path), now_iso(), len(current) + len(targets), json.dumps(options, ensure_ascii=False, sort_keys=True)),
        ).lastrowid

        for source, workbook in current:
            start_ingest_item(
                conn,
                int(run_id),
                dataset,
                source,
                Path(str(workbook["raw_json_path"])),
                "SKIPPED",
                workbook_id=int(workbook["workbook_id"]),
                message="Unchanged source fingerprint; retained the current verified raw grid.",
            )
            skipped += 1
        conn.commit()

        for index, (source, previous_workbook) in enumerate(targets, start=1):
            raw_json = raw_json_path_for_source(raw_dir, source)
            previous_workbook_id = int(previous_workbook["workbook_id"]) if previous_workbook else None
            ingest_item_id = start_ingest_item(
                conn,
                int(run_id),
                dataset,
                source,
                raw_json,
                "EXTRACTING",
                workbook_id=previous_workbook_id,
            )
            conn.commit()
            print(f"[{index}/{len(targets)}] {source.name}")
            try:
                if not args.reuse_json or not raw_json_matches_source(
                    raw_json,
                    source,
                    args.covered_cell_mode,
                    include_empty_cells=not args.sparse,
                ):
                    cmd = [
                        sys.executable,
                        str(COM_EXTRACTOR),
                        "--input",
                        str(source),
                        "--output",
                        str(raw_json),
                        "--covered-cell-mode",
                        args.covered_cell_mode,
                    ]
                    if args.include_hidden:
                        cmd.append("--include-hidden")
                    if args.sparse:
                        cmd.append("--sparse")
                    code = run_command(cmd, log_path, cwd=COM_EXTRACTOR.parent)
                    if code != 0:
                        raise RuntimeError(f"COM extractor exited with code {code}")

                result = import_com_json(
                    conn,
                    dataset,
                    raw_json,
                    expected_source=source,
                    expected_covered_cell_mode=args.covered_cell_mode,
                    verify_after_import=args.verify_after_import,
                )
                finish_ingest_item(
                    conn,
                    ingest_item_id,
                    "IMPORTED",
                    workbook_id=int(result["workbookId"]),
                    message="Imported fixed grid, rows, cells, and merge metadata.",
                )
                conn.commit()
                succeeded += 1
                print(f"stored workbookId={result['workbookId']} rows={result['rows']} cells={result['cells']} merges={result['merges']}")
            except Exception as exc:
                failed += 1
                finish_ingest_item(
                    conn,
                    ingest_item_id,
                    "FAILED",
                    workbook_id=previous_workbook_id,
                    message=compact(str(exc), 2000),
                )
                conn.commit()
                retained = f"; retained workbookId={previous_workbook_id}" if previous_workbook_id is not None else ""
                print(f"FAILED: {source.name}: {exc}{retained}")

        conn.execute(
            "UPDATE runs SET finished_at=?, succeeded=?, failed=?, skipped=? WHERE run_id=?",
            (now_iso(), succeeded, failed, skipped, run_id),
        )
        conn.commit()

    print_json(
        {
            "status": "ok" if failed == 0 else "partial",
            "dataset": dataset,
            "db": str(db_path),
            "rawJsonDir": str(raw_dir),
            "log": str(log_path),
            "discovered": len(files),
            "processed": len(targets),
            "succeeded": succeeded,
            "failed": failed,
            "skipped": skipped,
            "remaining": deferred,
            "resume": not args.force,
        }
    )
    return 0 if failed == 0 else 1


def inspect_quick_db(conn: sqlite3.Connection) -> dict[str, Any]:
    tables = [
        "files",
        "sheets",
        "sheet_rows",
        "sheet_cells",
        "metric_candidates",
        "measurement_stats",
        "comparison_pairs",
        "term_hits",
    ]
    counts = {table: conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] for table in tables if table_exists(conn, table)}
    status = dict_rows(conn, "SELECT status, COUNT(*) AS count FROM files GROUP BY status ORDER BY status") if table_exists(conn, "files") else []
    structures = (
        dict_rows(
            conn,
            "SELECT structure_family, structure_confidence, COUNT(*) AS count FROM files GROUP BY structure_family, structure_confidence ORDER BY count DESC LIMIT 30",
        )
        if table_exists(conn, "files")
        else []
    )
    return {"dbType": "quick-index", "counts": counts, "status": status, "structures": structures}


def inspect_universal_db(conn: sqlite3.Connection) -> dict[str, Any]:
    tables = [
        "runs",
        "ingest_items",
        "workbooks",
        "worksheets",
        "grid_sheet_rows",
        "grid_sheet_cells",
        "merge_ranges",
        "analysis_reports",
        "analysis_review_items",
        "analysis_cohorts",
        "analysis_metrics",
        "analysis_metric_values",
        "analysis_comparisons",
        "analysis_conclusions",
        "analysis_evidence",
    ]
    counts = {table: conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] for table in tables if table_exists(conn, table)}
    status = dict_rows(conn, "SELECT status, COUNT(*) AS count FROM workbooks GROUP BY status ORDER BY status")
    ingest_status = (
        dict_rows(conn, "SELECT status, COUNT(*) AS count FROM ingest_items GROUP BY status ORDER BY status")
        if table_exists(conn, "ingest_items")
        else []
    )
    recent = dict_rows(
        conn,
        """
        SELECT workbook_id, dataset, file_name, status, sheet_count, total_rows, total_cells, non_empty_cells, merge_count
        FROM workbooks
        ORDER BY workbook_id DESC
        LIMIT 10
        """,
    )
    analysis_status = (
        dict_rows(
            conn,
            "SELECT overall_status, COUNT(*) AS count FROM analysis_reports GROUP BY overall_status ORDER BY overall_status",
        )
        if table_exists(conn, "analysis_reports")
        else []
    )
    return {
        "dbType": "universal-grid",
        "counts": counts,
        "status": status,
        "ingestStatus": ingest_status,
        "analysisStatus": analysis_status,
        "recentWorkbooks": recent,
    }


def verify_universal_db(conn: sqlite3.Connection, dataset: str | None = None, limit: int = 0) -> dict[str, Any]:
    sql = "SELECT * FROM workbooks WHERE status='OK'"
    params: list[Any] = []
    if dataset:
        sql += " AND dataset=?"
        params.append(dataset)
    sql += " ORDER BY workbook_id"
    if limit > 0:
        sql += " LIMIT ?"
        params.append(limit)

    results = []
    for workbook in dict_rows(conn, sql, params):
        raw_json_path = Path(str(workbook.get("raw_json_path") or ""))
        if not raw_json_path.is_file():
            results.append(
                {
                    "workbookId": workbook["workbook_id"],
                    "fileName": workbook["file_name"],
                    "ok": False,
                    "errors": ["raw JSON artifact is missing"],
                }
            )
            continue
        try:
            payload = read_com_grid_json(raw_json_path)
            verification = verify_universal_workbook(conn, int(workbook["workbook_id"]), payload)
            source = Path(str(workbook["source_path"]))
            verification["sourceChanged"] = source.exists() and workbook["fingerprint"] != file_fingerprint(source)
            results.append(verification)
        except (OSError, ValueError) as exc:
            results.append(
                {
                    "workbookId": workbook["workbook_id"],
                    "fileName": workbook["file_name"],
                    "ok": False,
                    "errors": [str(exc)],
                }
            )

    return {
        "dataset": dataset or "",
        "checked": len(results),
        "valid": sum(1 for item in results if item.get("ok")),
        "invalid": sum(1 for item in results if not item.get("ok")),
        "sourceChanged": sum(1 for item in results if item.get("sourceChanged")),
        "workbooks": results,
    }


def read_analysis_manifest(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid analysis manifest JSON: {path}: {exc}") from exc
    return expect_object(data, f"analysis manifest {path}")


def resolve_analysis_workbook(
    conn: sqlite3.Connection,
    source: dict[str, Any],
    dataset_hint: str | None = None,
) -> dict[str, Any]:
    workbook_id_value = source.get("workbookId")
    source_path = normalized_source_path(source.get("sourcePath"))
    requested_dataset = str(source.get("dataset") or dataset_hint or "").strip()
    if workbook_id_value is None and not source_path:
        raise ValueError("analysis source requires workbookId or sourcePath")

    workbook: dict[str, Any] | None = None
    if workbook_id_value is not None:
        workbook_id = integer_field(workbook_id_value, "analysis source.workbookId")
        workbook = first_dict(conn, "SELECT * FROM workbooks WHERE workbook_id=? LIMIT 1", (workbook_id,))
        if not workbook:
            raise ValueError(f"analysis source workbookId does not exist: {workbook_id}")
    if source_path:
        params: list[Any] = [source_path]
        sql = "SELECT * FROM workbooks WHERE source_path=?"
        if requested_dataset:
            sql += " AND dataset=?"
            params.append(requested_dataset)
        sql += " LIMIT 1"
        by_path = first_dict(conn, sql, params)
        if not by_path:
            raise ValueError(f"analysis sourcePath is not imported into the Universal DB: {source_path}")
        if workbook and int(workbook["workbook_id"]) != int(by_path["workbook_id"]):
            raise ValueError("analysis source workbookId and sourcePath refer to different workbooks")
        workbook = by_path

    assert workbook is not None
    if requested_dataset and workbook["dataset"] != requested_dataset:
        raise ValueError(f"analysis source dataset mismatch: {workbook['dataset']!r} != {requested_dataset!r}")
    if workbook["status"] != "OK":
        raise ValueError(f"analysis source workbook is not usable: status={workbook['status']!r}")
    expected_fingerprint = str(source.get("fingerprint") or "").strip()
    if expected_fingerprint and expected_fingerprint != workbook["fingerprint"]:
        raise ValueError("analysis source fingerprint does not match the current Universal DB workbook")
    return workbook


def metric_value_base_number(value: dict[str, Any], label: str) -> float | None:
    for field in ("ratePpm", "average", "valueNumber"):
        number = optional_number(value.get(field), f"{label}.{field}")
        if number is not None:
            return number
    return None


def validate_metric_value(value: Any, label: str) -> dict[str, Any]:
    item = expect_object(value, label)
    numerator = optional_number(item.get("numerator"), f"{label}.numerator")
    denominator = optional_number(item.get("denominator"), f"{label}.denominator")
    rate_ppm = optional_number(item.get("ratePpm"), f"{label}.ratePpm")
    if (numerator is None) != (denominator is None):
        raise ValueError(f"{label} must provide numerator and denominator together")
    if denominator is not None:
        if denominator <= 0:
            raise ValueError(f"{label}.denominator must be positive")
        if numerator is not None and numerator < 0:
            raise ValueError(f"{label}.numerator must not be negative")
        if numerator is not None and rate_ppm is not None:
            calculated = numerator * 1_000_000 / denominator
            if abs(calculated - rate_ppm) > 0.51:
                raise ValueError(
                    f"{label}.ratePpm does not match numerator/denominator: {rate_ppm} != {calculated}"
                )
    for field in ("valueNumber", "min", "max", "average"):
        optional_number(item.get(field), f"{label}.{field}")
    return item


def manifest_evidence_list(value: Any, label: str) -> list[dict[str, Any]]:
    items = expect_list(value, label)
    return [expect_object(item, f"{label}[{index}]") for index, item in enumerate(items)]


def validate_analysis_conclusions(
    conn: sqlite3.Connection,
    workbook_id: int,
    conclusions: Any,
    label: str,
) -> None:
    keys: set[str] = set()
    for index, item_value in enumerate(expect_list(conclusions, label)):
        item = expect_object(item_value, f"{label}[{index}]")
        key = required_text(item.get("key"), f"{label}[{index}].key")
        if key in keys:
            raise ValueError(f"{label} duplicates conclusion key: {key}")
        keys.add(key)
        required_text(item.get("verdict"), f"{label}[{index}].verdict")
        required_text(item.get("text"), f"{label}[{index}].text")
        evidence = manifest_evidence_list(item.get("evidence"), f"{label}[{index}].evidence")
        if not evidence:
            raise ValueError(f"{label}[{index}] must have at least one evidence reference")
        for evidence_index, evidence_item in enumerate(evidence):
            validate_analysis_evidence(conn, workbook_id, evidence_item, f"{label}[{index}].evidence[{evidence_index}]")


def validate_analysis_manifest(
    conn: sqlite3.Connection,
    data: dict[str, Any],
    dataset_hint: str | None = None,
) -> dict[str, Any]:
    if data.get("schemaVersion") != ANALYSIS_SCHEMA_VERSION:
        raise ValueError(f"Unsupported analysis manifest schema: {data.get('schemaVersion')!r}")
    source = expect_object(data.get("source"), "analysis source")
    workbook = resolve_analysis_workbook(conn, source, dataset_hint)
    report = expect_object(data.get("report"), "analysis report")
    required_text(report.get("key"), "analysis report.key")
    required_text(report.get("title"), "analysis report.title")
    required_text(report.get("type"), "analysis report.type")
    report_evidence = manifest_evidence_list(report.get("evidence"), "analysis report.evidence")
    if not report_evidence:
        raise ValueError("analysis report must have at least one evidence reference")
    for index, evidence in enumerate(report_evidence):
        validate_analysis_evidence(conn, int(workbook["workbook_id"]), evidence, f"analysis report.evidence[{index}]")
    validate_analysis_conclusions(conn, int(workbook["workbook_id"]), report.get("conclusions"), "analysis report.conclusions")

    review_keys: set[str] = set()
    reviews = expect_list(data.get("reviews"), "analysis reviews")
    if not reviews:
        raise ValueError("analysis manifest must contain at least one review")
    for review_index, review_value in enumerate(reviews):
        review = expect_object(review_value, f"analysis reviews[{review_index}]")
        review_key = required_text(review.get("key"), f"analysis reviews[{review_index}].key")
        if review_key in review_keys:
            raise ValueError(f"analysis manifest duplicates review key: {review_key}")
        review_keys.add(review_key)
        required_text(review.get("title"), f"analysis reviews[{review_index}].title")
        required_text(review.get("type"), f"analysis reviews[{review_index}].type")

        review_evidence = manifest_evidence_list(review.get("evidence", []), f"analysis reviews[{review_index}].evidence")
        for evidence_index, evidence in enumerate(review_evidence):
            validate_analysis_evidence(
                conn,
                int(workbook["workbook_id"]),
                evidence,
                f"analysis reviews[{review_index}].evidence[{evidence_index}]",
            )

        cohorts = expect_list(review.get("cohorts"), f"analysis reviews[{review_index}].cohorts")
        if not cohorts:
            raise ValueError(f"analysis reviews[{review_index}] must have comparison cohorts")
        cohort_keys: set[str] = set()
        for cohort_index, cohort_value in enumerate(cohorts):
            cohort = expect_object(cohort_value, f"analysis reviews[{review_index}].cohorts[{cohort_index}]")
            cohort_key = required_text(cohort.get("key"), f"analysis reviews[{review_index}].cohorts[{cohort_index}].key")
            if cohort_key in cohort_keys:
                raise ValueError(f"analysis reviews[{review_index}] duplicates cohort key: {cohort_key}")
            cohort_keys.add(cohort_key)
            required_text(cohort.get("role"), f"analysis reviews[{review_index}].cohorts[{cohort_index}].role")
            required_text(cohort.get("label"), f"analysis reviews[{review_index}].cohorts[{cohort_index}].label")

        metric_keys: set[tuple[str, str]] = set()
        metrics = expect_list(review.get("metrics"), f"analysis reviews[{review_index}].metrics")
        if not metrics:
            raise ValueError(f"analysis reviews[{review_index}] must have at least one metric")
        for metric_index, metric_value in enumerate(metrics):
            metric = expect_object(metric_value, f"analysis reviews[{review_index}].metrics[{metric_index}]")
            metric_key = required_text(metric.get("key"), f"analysis reviews[{review_index}].metrics[{metric_index}].key")
            scope_key = str(metric.get("scope") or "").strip()
            identity = (metric_key, scope_key)
            if identity in metric_keys:
                raise ValueError(f"analysis reviews[{review_index}] duplicates metric key/scope: {identity}")
            metric_keys.add(identity)
            required_text(metric.get("label"), f"analysis reviews[{review_index}].metrics[{metric_index}].label")
            required_text(metric.get("type"), f"analysis reviews[{review_index}].metrics[{metric_index}].type")
            metric_evidence = manifest_evidence_list(metric.get("evidence"), f"analysis reviews[{review_index}].metrics[{metric_index}].evidence")
            if not metric_evidence:
                raise ValueError(f"analysis reviews[{review_index}].metrics[{metric_index}] must have evidence")
            for evidence_index, evidence in enumerate(metric_evidence):
                validate_analysis_evidence(
                    conn,
                    int(workbook["workbook_id"]),
                    evidence,
                    f"analysis reviews[{review_index}].metrics[{metric_index}].evidence[{evidence_index}]",
                )

            values_by_cohort: dict[str, dict[str, Any]] = {}
            for value_index, value_value in enumerate(expect_list(metric.get("values"), f"analysis reviews[{review_index}].metrics[{metric_index}].values")):
                value = validate_metric_value(
                    value_value,
                    f"analysis reviews[{review_index}].metrics[{metric_index}].values[{value_index}]",
                )
                cohort_key = required_text(
                    value.get("cohort"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].values[{value_index}].cohort",
                )
                if cohort_key not in cohort_keys:
                    raise ValueError(f"metric value refers to unknown cohort: {cohort_key}")
                if cohort_key in values_by_cohort:
                    raise ValueError(f"metric has duplicate value for cohort: {cohort_key}")
                values_by_cohort[cohort_key] = value

            comparison_keys: set[str] = set()
            for comparison_index, comparison_value in enumerate(expect_list(metric.get("comparisons", []), f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons")):
                comparison = expect_object(
                    comparison_value,
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}]",
                )
                comparison_key = required_text(
                    comparison.get("key"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].key",
                )
                if comparison_key in comparison_keys:
                    raise ValueError(f"metric has duplicate comparison key: {comparison_key}")
                comparison_keys.add(comparison_key)
                compared = required_text(
                    comparison.get("comparedCohort"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].comparedCohort",
                )
                control = required_text(
                    comparison.get("controlCohort"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].controlCohort",
                )
                if compared not in values_by_cohort or control not in values_by_cohort:
                    raise ValueError(f"comparison {comparison_key} must reference metric values for both cohorts")
                if compared == control:
                    raise ValueError(f"comparison {comparison_key} cannot compare a cohort with itself")
                delta = optional_number(
                    comparison.get("deltaValue"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].deltaValue",
                )
                relative = optional_number(
                    comparison.get("relativeDeltaPercent"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].relativeDeltaPercent",
                )
                compared_base = metric_value_base_number(values_by_cohort[compared], "comparison compared value")
                control_base = metric_value_base_number(values_by_cohort[control], "comparison control value")
                if delta is not None and compared_base is not None and control_base is not None:
                    calculated_delta = compared_base - control_base
                    if abs(calculated_delta - delta) > 1.1:
                        raise ValueError(
                            f"comparison {comparison_key} delta does not match its cohort values: {delta} != {calculated_delta}"
                        )
                    if relative is not None and control_base != 0:
                        calculated_relative = calculated_delta * 100 / control_base
                        if abs(calculated_relative - relative) > 0.2:
                            raise ValueError(
                                f"comparison {comparison_key} relative delta does not match its cohort values: "
                                f"{relative} != {calculated_relative}"
                            )
                comparison_evidence = manifest_evidence_list(
                    comparison.get("evidence"),
                    f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].evidence",
                )
                if not comparison_evidence:
                    raise ValueError(f"comparison {comparison_key} must have evidence")
                for evidence_index, evidence in enumerate(comparison_evidence):
                    validate_analysis_evidence(
                        conn,
                        int(workbook["workbook_id"]),
                        evidence,
                        f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons[{comparison_index}].evidence[{evidence_index}]",
                    )

        validate_analysis_conclusions(
            conn,
            int(workbook["workbook_id"]),
            review.get("conclusions"),
            f"analysis reviews[{review_index}].conclusions",
        )

    return {"dataset": workbook["dataset"], "workbook": workbook, "report": report}


def insert_analysis_evidence(
    conn: sqlite3.Connection,
    analysis_report_id: int,
    workbook_id: int,
    evidence_items: list[dict[str, Any]],
    *,
    review_item_id: int | None = None,
    metric_id: int | None = None,
    comparison_id: int | None = None,
    conclusion_id: int | None = None,
) -> int:
    inserted = 0
    for index, item in enumerate(evidence_items):
        evidence = validate_analysis_evidence(conn, workbook_id, item, f"analysis evidence[{index}]")
        conn.execute(
            """
            INSERT INTO analysis_evidence(
                analysis_report_id, review_item_id, metric_id, comparison_id, conclusion_id,
                workbook_id, sheet_name, start_row, start_col, end_row, end_col, range_address,
                evidence_role, note, source_text
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                analysis_report_id,
                review_item_id,
                metric_id,
                comparison_id,
                conclusion_id,
                workbook_id,
                evidence["sheet"],
                evidence["startRow"],
                evidence["startCol"],
                evidence["endRow"],
                evidence["endCol"],
                evidence["range"],
                evidence["role"],
                evidence["note"],
                evidence["sourceText"],
            ),
        )
        inserted += 1
    return inserted


def insert_analysis_conclusions(
    conn: sqlite3.Connection,
    analysis_report_id: int,
    workbook_id: int,
    conclusions: list[Any],
    *,
    review_item_id: int | None = None,
) -> int:
    inserted = 0
    for index, item_value in enumerate(conclusions):
        item = expect_object(item_value, f"analysis conclusion[{index}]")
        cur = conn.execute(
            """
            INSERT INTO analysis_conclusions(
                analysis_report_id, review_item_id, conclusion_key, sort_order,
                verdict, label, conclusion_text, scope_text, limitations_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                analysis_report_id,
                review_item_id,
                required_text(item.get("key"), f"analysis conclusion[{index}].key"),
                integer_field(item.get("sortOrder") or index + 1, f"analysis conclusion[{index}].sortOrder"),
                required_text(item.get("verdict"), f"analysis conclusion[{index}].verdict"),
                str(item.get("label") or "").strip(),
                required_text(item.get("text"), f"analysis conclusion[{index}].text"),
                str(item.get("scope") or "").strip(),
                json_text(item.get("limitations"), [], f"analysis conclusion[{index}].limitations"),
            ),
        )
        conclusion_id = int(cur.lastrowid)
        insert_analysis_evidence(
            conn,
            analysis_report_id,
            workbook_id,
            manifest_evidence_list(item.get("evidence"), f"analysis conclusion[{index}].evidence"),
            review_item_id=review_item_id,
            conclusion_id=conclusion_id,
        )
        inserted += 1
    return inserted


def import_analysis_manifest(
    conn: sqlite3.Connection,
    manifest_path: Path,
    data: dict[str, Any],
    dataset_hint: str | None = None,
) -> dict[str, Any]:
    context = validate_analysis_manifest(conn, data, dataset_hint)
    workbook = context["workbook"]
    report = context["report"]
    source_path = str(workbook["source_path"])
    dataset = str(workbook["dataset"])
    analysis_key = required_text(report.get("key"), "analysis report.key")
    artifacts = expect_object(report.get("artifacts", {}), "analysis report.artifacts")

    conn.execute("SAVEPOINT universal_analysis_import")
    try:
        conn.execute(
            """
            DELETE FROM analysis_reports
            WHERE dataset=? AND source_path=? AND workbook_fingerprint=? AND analysis_key=?
            """,
            (dataset, source_path, workbook["fingerprint"], analysis_key),
        )
        cur = conn.execute(
            """
            INSERT INTO analysis_reports(
                dataset, workbook_id, source_path, workbook_fingerprint, analysis_key,
                title, analysis_type, purpose, scope_text, overall_status, overall_decision,
                overall_summary, limitations_json, dashboard_html_path, dashboard_markdown_path,
                manifest_path, stale_reason, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?)
            """,
            (
                dataset,
                int(workbook["workbook_id"]),
                source_path,
                workbook["fingerprint"],
                analysis_key,
                required_text(report.get("title"), "analysis report.title"),
                required_text(report.get("type"), "analysis report.type"),
                str(report.get("purpose") or "").strip(),
                str(report.get("scope") or "").strip(),
                str(report.get("status") or "VERIFIED").strip().upper(),
                str(report.get("decision") or "").strip(),
                str(report.get("summary") or "").strip(),
                json_text(report.get("limitations"), [], "analysis report.limitations"),
                str(artifacts.get("html") or "").strip(),
                str(artifacts.get("markdown") or "").strip(),
                str(manifest_path.resolve()),
                now_iso(),
                now_iso(),
            ),
        )
        analysis_report_id = int(cur.lastrowid)
        evidence_count = insert_analysis_evidence(
            conn,
            analysis_report_id,
            int(workbook["workbook_id"]),
            manifest_evidence_list(report.get("evidence"), "analysis report.evidence"),
        )
        conclusion_count = insert_analysis_conclusions(
            conn,
            analysis_report_id,
            int(workbook["workbook_id"]),
            expect_list(report.get("conclusions"), "analysis report.conclusions"),
        )

        review_count = 0
        cohort_count = 0
        metric_count = 0
        metric_value_count = 0
        comparison_count = 0
        for review_index, review_value in enumerate(expect_list(data.get("reviews"), "analysis reviews")):
            review = expect_object(review_value, f"analysis reviews[{review_index}]")
            cur = conn.execute(
                """
                INSERT INTO analysis_review_items(
                    analysis_report_id, review_key, sort_order, review_type, title,
                    objective, comparison_basis, status, decision_text, summary_text, notes_json
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    analysis_report_id,
                    required_text(review.get("key"), f"analysis reviews[{review_index}].key"),
                    integer_field(review.get("sortOrder") or review_index + 1, f"analysis reviews[{review_index}].sortOrder"),
                    required_text(review.get("type"), f"analysis reviews[{review_index}].type"),
                    required_text(review.get("title"), f"analysis reviews[{review_index}].title"),
                    str(review.get("objective") or "").strip(),
                    str(review.get("comparisonBasis") or "").strip(),
                    str(review.get("status") or "").strip().upper(),
                    str(review.get("decision") or "").strip(),
                    str(review.get("summary") or "").strip(),
                    json_text(review.get("notes"), [], f"analysis reviews[{review_index}].notes"),
                ),
            )
            review_item_id = int(cur.lastrowid)
            review_count += 1
            evidence_count += insert_analysis_evidence(
                conn,
                analysis_report_id,
                int(workbook["workbook_id"]),
                manifest_evidence_list(review.get("evidence", []), f"analysis reviews[{review_index}].evidence"),
                review_item_id=review_item_id,
            )

            cohort_ids: dict[str, int] = {}
            for cohort_index, cohort_value in enumerate(expect_list(review.get("cohorts"), f"analysis reviews[{review_index}].cohorts")):
                cohort = expect_object(cohort_value, f"analysis reviews[{review_index}].cohorts[{cohort_index}]")
                cohort_key = required_text(cohort.get("key"), f"analysis reviews[{review_index}].cohorts[{cohort_index}].key")
                cur = conn.execute(
                    """
                    INSERT INTO analysis_cohorts(
                        review_item_id, cohort_key, cohort_role, label, condition_text, sort_order, attributes_json
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        review_item_id,
                        cohort_key,
                        required_text(cohort.get("role"), f"analysis reviews[{review_index}].cohorts[{cohort_index}].role").upper(),
                        required_text(cohort.get("label"), f"analysis reviews[{review_index}].cohorts[{cohort_index}].label"),
                        str(cohort.get("condition") or "").strip(),
                        integer_field(cohort.get("sortOrder") or cohort_index + 1, f"analysis reviews[{review_index}].cohorts[{cohort_index}].sortOrder"),
                        json_text(cohort.get("attributes"), {}, f"analysis reviews[{review_index}].cohorts[{cohort_index}].attributes"),
                    ),
                )
                cohort_ids[cohort_key] = int(cur.lastrowid)
                cohort_count += 1

            for metric_index, metric_value in enumerate(expect_list(review.get("metrics"), f"analysis reviews[{review_index}].metrics")):
                metric = expect_object(metric_value, f"analysis reviews[{review_index}].metrics[{metric_index}]")
                cur = conn.execute(
                    """
                    INSERT INTO analysis_metrics(
                        review_item_id, metric_key, scope_key, sort_order, label, metric_type,
                        unit, spec_text, definition_text, status, notes_json
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        review_item_id,
                        required_text(metric.get("key"), f"analysis reviews[{review_index}].metrics[{metric_index}].key"),
                        str(metric.get("scope") or "").strip(),
                        integer_field(metric.get("sortOrder") or metric_index + 1, f"analysis reviews[{review_index}].metrics[{metric_index}].sortOrder"),
                        required_text(metric.get("label"), f"analysis reviews[{review_index}].metrics[{metric_index}].label"),
                        required_text(metric.get("type"), f"analysis reviews[{review_index}].metrics[{metric_index}].type"),
                        str(metric.get("unit") or "").strip(),
                        str(metric.get("spec") or "").strip(),
                        str(metric.get("definition") or "").strip(),
                        str(metric.get("status") or "").strip().upper(),
                        json_text(metric.get("notes"), [], f"analysis reviews[{review_index}].metrics[{metric_index}].notes"),
                    ),
                )
                metric_id = int(cur.lastrowid)
                metric_count += 1
                evidence_count += insert_analysis_evidence(
                    conn,
                    analysis_report_id,
                    int(workbook["workbook_id"]),
                    manifest_evidence_list(metric.get("evidence"), f"analysis reviews[{review_index}].metrics[{metric_index}].evidence"),
                    review_item_id=review_item_id,
                    metric_id=metric_id,
                )

                for value_index, value_value in enumerate(expect_list(metric.get("values"), f"analysis reviews[{review_index}].metrics[{metric_index}].values")):
                    value = validate_metric_value(value_value, f"analysis reviews[{review_index}].metrics[{metric_index}].values[{value_index}]")
                    cohort_key = required_text(value.get("cohort"), f"analysis metric value[{value_index}].cohort")
                    conn.execute(
                        """
                        INSERT INTO analysis_metric_values(
                            metric_id, cohort_id, value_number, value_text, numerator, denominator,
                            rate_ppm, min_value, max_value, average_value, result_status, details_json
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            metric_id,
                            cohort_ids[cohort_key],
                            optional_number(value.get("valueNumber"), "analysis metric value.valueNumber"),
                            str(value.get("valueText") or "").strip(),
                            optional_number(value.get("numerator"), "analysis metric value.numerator"),
                            optional_number(value.get("denominator"), "analysis metric value.denominator"),
                            optional_number(value.get("ratePpm"), "analysis metric value.ratePpm"),
                            optional_number(value.get("min"), "analysis metric value.min"),
                            optional_number(value.get("max"), "analysis metric value.max"),
                            optional_number(value.get("average"), "analysis metric value.average"),
                            str(value.get("status") or "").strip().upper(),
                            json_text(value.get("details"), {}, "analysis metric value.details"),
                        ),
                    )
                    metric_value_count += 1

                for comparison_index, comparison_value in enumerate(expect_list(metric.get("comparisons", []), f"analysis reviews[{review_index}].metrics[{metric_index}].comparisons")):
                    comparison = expect_object(comparison_value, f"analysis comparison[{comparison_index}]")
                    cur = conn.execute(
                        """
                        INSERT INTO analysis_comparisons(
                            metric_id, comparison_key, compared_cohort_id, control_cohort_id,
                            delta_value, delta_unit, relative_delta_percent, direction, status,
                            summary_text, calculation_text, details_json
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            metric_id,
                            required_text(comparison.get("key"), f"analysis comparison[{comparison_index}].key"),
                            cohort_ids[required_text(comparison.get("comparedCohort"), f"analysis comparison[{comparison_index}].comparedCohort")],
                            cohort_ids[required_text(comparison.get("controlCohort"), f"analysis comparison[{comparison_index}].controlCohort")],
                            optional_number(comparison.get("deltaValue"), "analysis comparison.deltaValue"),
                            str(comparison.get("deltaUnit") or metric.get("unit") or "").strip(),
                            optional_number(comparison.get("relativeDeltaPercent"), "analysis comparison.relativeDeltaPercent"),
                            str(comparison.get("direction") or "").strip().upper(),
                            str(comparison.get("status") or "").strip().upper(),
                            str(comparison.get("summary") or "").strip(),
                            str(comparison.get("calculation") or "").strip(),
                            json_text(comparison.get("details"), {}, "analysis comparison.details"),
                        ),
                    )
                    comparison_id = int(cur.lastrowid)
                    comparison_count += 1
                    evidence_count += insert_analysis_evidence(
                        conn,
                        analysis_report_id,
                        int(workbook["workbook_id"]),
                        manifest_evidence_list(comparison.get("evidence"), f"analysis comparison[{comparison_index}].evidence"),
                        review_item_id=review_item_id,
                        metric_id=metric_id,
                        comparison_id=comparison_id,
                    )

            conclusion_count += insert_analysis_conclusions(
                conn,
                analysis_report_id,
                int(workbook["workbook_id"]),
                expect_list(review.get("conclusions"), f"analysis reviews[{review_index}].conclusions"),
                review_item_id=review_item_id,
            )

        verification = verify_analysis_report(conn, analysis_report_id)
        if not verification["ok"]:
            raise ValueError("Post-import analysis verification failed: " + "; ".join(verification["errors"]))
        canonical = sync_legacy_analysis_report(conn, analysis_report_id, now_iso)
    except Exception:
        conn.execute("ROLLBACK TO SAVEPOINT universal_analysis_import")
        conn.execute("RELEASE SAVEPOINT universal_analysis_import")
        raise
    conn.execute("RELEASE SAVEPOINT universal_analysis_import")
    evidence_count = conn.execute(
        "SELECT COUNT(*) FROM analysis_evidence WHERE analysis_report_id=?",
        (analysis_report_id,),
    ).fetchone()[0]
    return {
        "analysisReportId": analysis_report_id,
        "analysisKey": analysis_key,
        "workbookId": int(workbook["workbook_id"]),
        "reviews": review_count,
        "cohorts": cohort_count,
        "metrics": metric_count,
        "metricValues": metric_value_count,
        "comparisons": comparison_count,
        "conclusions": conclusion_count,
        "evidence": evidence_count,
        "verification": verification,
        "canonical": canonical,
    }


def verify_analysis_report(conn: sqlite3.Connection, analysis_report_id: int) -> dict[str, Any]:
    report = first_dict(conn, "SELECT * FROM analysis_reports WHERE analysis_report_id=? LIMIT 1", (analysis_report_id,))
    if not report:
        return {"analysisReportId": analysis_report_id, "ok": False, "errors": ["analysis report is missing"]}
    errors: list[str] = []
    workbook = first_dict(conn, "SELECT * FROM workbooks WHERE workbook_id=? LIMIT 1", (report["workbook_id"],))
    if not workbook:
        errors.append("source workbook record is missing; analysis is stale")
    else:
        if workbook["fingerprint"] != report["workbook_fingerprint"]:
            errors.append("source workbook fingerprint differs; analysis is stale")
        if workbook["status"] != "OK":
            errors.append(f"source workbook status is {workbook['status']!r}")

    evidence_rows = dict_rows(
        conn,
        "SELECT * FROM analysis_evidence WHERE analysis_report_id=? ORDER BY evidence_id",
        (analysis_report_id,),
    )
    if not evidence_rows:
        errors.append("analysis report has no evidence")
    if workbook:
        for evidence in evidence_rows:
            try:
                validate_analysis_evidence(
                    conn,
                    int(workbook["workbook_id"]),
                    {"sheet": evidence["sheet_name"], "range": evidence["range_address"]},
                    f"stored evidence {evidence['evidence_id']}",
                )
            except ValueError as exc:
                errors.append(str(exc))

    missing_metrics = dict_rows(
        conn,
        """
        SELECT m.metric_id, m.label
        FROM analysis_metrics m
        LEFT JOIN analysis_evidence e ON e.metric_id=m.metric_id
        WHERE m.review_item_id IN (
            SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?
        )
        GROUP BY m.metric_id
        HAVING COUNT(e.evidence_id)=0
        """,
        (analysis_report_id,),
    )
    errors.extend(f"metric has no evidence: {row['label'] or row['metric_id']}" for row in missing_metrics)
    missing_comparisons = dict_rows(
        conn,
        """
        SELECT c.comparison_id, c.comparison_key
        FROM analysis_comparisons c
        JOIN analysis_metrics m ON m.metric_id=c.metric_id
        LEFT JOIN analysis_evidence e ON e.comparison_id=c.comparison_id
        WHERE m.review_item_id IN (
            SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?
        )
        GROUP BY c.comparison_id
        HAVING COUNT(e.evidence_id)=0
        """,
        (analysis_report_id,),
    )
    errors.extend(f"comparison has no evidence: {row['comparison_key']}" for row in missing_comparisons)
    missing_conclusions = dict_rows(
        conn,
        """
        SELECT c.conclusion_id, c.conclusion_key
        FROM analysis_conclusions c
        LEFT JOIN analysis_evidence e ON e.conclusion_id=c.conclusion_id
        WHERE c.analysis_report_id=?
        GROUP BY c.conclusion_id
        HAVING COUNT(e.evidence_id)=0
        """,
        (analysis_report_id,),
    )
    errors.extend(f"conclusion has no evidence: {row['conclusion_key']}" for row in missing_conclusions)

    values = dict_rows(
        conn,
        """
        SELECT v.metric_value_id, v.numerator, v.denominator, v.rate_ppm
        FROM analysis_metric_values v
        JOIN analysis_metrics m ON m.metric_id=v.metric_id
        JOIN analysis_review_items r ON r.review_item_id=m.review_item_id
        WHERE r.analysis_report_id=?
        """,
        (analysis_report_id,),
    )
    for value in values:
        numerator = value["numerator"]
        denominator = value["denominator"]
        rate_ppm = value["rate_ppm"]
        if numerator is not None and denominator is not None and rate_ppm is not None:
            if float(denominator) <= 0 or abs(float(numerator) * 1_000_000 / float(denominator) - float(rate_ppm)) > 0.51:
                errors.append(f"metric value {value['metric_value_id']} has inconsistent ppm rate")

    comparisons = dict_rows(
        conn,
        """
        SELECT c.comparison_id, c.comparison_key, c.delta_value, c.relative_delta_percent,
               compared.rate_ppm AS compared_rate, compared.average_value AS compared_average,
               compared.value_number AS compared_value,
               control.rate_ppm AS control_rate, control.average_value AS control_average,
               control.value_number AS control_value
        FROM analysis_comparisons c
        JOIN analysis_metric_values compared ON compared.metric_id=c.metric_id AND compared.cohort_id=c.compared_cohort_id
        JOIN analysis_metric_values control ON control.metric_id=c.metric_id AND control.cohort_id=c.control_cohort_id
        JOIN analysis_metrics m ON m.metric_id=c.metric_id
        JOIN analysis_review_items r ON r.review_item_id=m.review_item_id
        WHERE r.analysis_report_id=?
        """,
        (analysis_report_id,),
    )
    for comparison in comparisons:
        compared_value = next(
            (comparison[key] for key in ("compared_rate", "compared_average", "compared_value") if comparison[key] is not None),
            None,
        )
        control_value = next(
            (comparison[key] for key in ("control_rate", "control_average", "control_value") if comparison[key] is not None),
            None,
        )
        if comparison["delta_value"] is not None and compared_value is not None and control_value is not None:
            calculated_delta = float(compared_value) - float(control_value)
            if abs(calculated_delta - float(comparison["delta_value"])) > 1.1:
                errors.append(f"comparison {comparison['comparison_key']} has inconsistent delta")
            if comparison["relative_delta_percent"] is not None and float(control_value) != 0:
                calculated_relative = calculated_delta * 100 / float(control_value)
                if abs(calculated_relative - float(comparison["relative_delta_percent"])) > 0.2:
                    errors.append(f"comparison {comparison['comparison_key']} has inconsistent relative delta")

    return {
        "analysisReportId": analysis_report_id,
        "analysisKey": report["analysis_key"],
        "workbookId": report["workbook_id"],
        "ok": not errors,
        "errors": errors,
        "counts": {
            "reviews": conn.execute("SELECT COUNT(*) FROM analysis_review_items WHERE analysis_report_id=?", (analysis_report_id,)).fetchone()[0],
            "cohorts": conn.execute(
                """
                SELECT COUNT(*) FROM analysis_cohorts
                WHERE review_item_id IN (SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?)
                """,
                (analysis_report_id,),
            ).fetchone()[0],
            "metrics": conn.execute(
                """
                SELECT COUNT(*) FROM analysis_metrics
                WHERE review_item_id IN (SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?)
                """,
                (analysis_report_id,),
            ).fetchone()[0],
            "comparisons": len(comparisons),
            "conclusions": conn.execute("SELECT COUNT(*) FROM analysis_conclusions WHERE analysis_report_id=?", (analysis_report_id,)).fetchone()[0],
            "evidence": len(evidence_rows),
        },
    }


def inspect_analysis_reports(
    conn: sqlite3.Connection,
    analysis_report_id: int | None = None,
    workbook_id: int | None = None,
    dataset: str | None = None,
) -> dict[str, Any]:
    sql = """
        SELECT a.*, w.file_name, w.status AS workbook_status, w.fingerprint AS current_workbook_fingerprint
        FROM analysis_reports a
        LEFT JOIN workbooks w ON w.workbook_id=a.workbook_id
        WHERE 1=1
    """
    params: list[Any] = []
    if analysis_report_id is not None:
        sql += " AND a.analysis_report_id=?"
        params.append(analysis_report_id)
    if workbook_id is not None:
        sql += " AND a.workbook_id=?"
        params.append(workbook_id)
    if dataset:
        sql += " AND a.dataset=?"
        params.append(dataset)
    sql += " ORDER BY a.analysis_report_id"
    reports = dict_rows(conn, sql, params)
    for report in reports:
        report_id = int(report["analysis_report_id"])
        report["isCurrent"] = (
            report.get("workbook_status") == "OK"
            and report.get("current_workbook_fingerprint") == report.get("workbook_fingerprint")
        )
        report["reviewItems"] = dict_rows(
            conn,
            """
            SELECT review_item_id, review_key, sort_order, review_type, title, comparison_basis,
                   status, decision_text, summary_text
            FROM analysis_review_items
            WHERE analysis_report_id=?
            ORDER BY sort_order, review_item_id
            """,
            (report_id,),
        )
        report["conclusions"] = dict_rows(
            conn,
            """
            SELECT conclusion_id, review_item_id, conclusion_key, verdict, label, conclusion_text, scope_text
            FROM analysis_conclusions
            WHERE analysis_report_id=?
            ORDER BY review_item_id, sort_order, conclusion_id
            """,
            (report_id,),
        )
        report["counts"] = {
            "cohorts": conn.execute(
                """
                SELECT COUNT(*) FROM analysis_cohorts
                WHERE review_item_id IN (SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?)
                """,
                (report_id,),
            ).fetchone()[0],
            "metrics": conn.execute(
                """
                SELECT COUNT(*) FROM analysis_metrics
                WHERE review_item_id IN (SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?)
                """,
                (report_id,),
            ).fetchone()[0],
            "metricValues": conn.execute(
                """
                SELECT COUNT(*) FROM analysis_metric_values
                WHERE metric_id IN (
                    SELECT metric_id FROM analysis_metrics WHERE review_item_id IN (
                        SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?
                    )
                )
                """,
                (report_id,),
            ).fetchone()[0],
            "comparisons": conn.execute(
                """
                SELECT COUNT(*) FROM analysis_comparisons
                WHERE metric_id IN (
                    SELECT metric_id FROM analysis_metrics WHERE review_item_id IN (
                        SELECT review_item_id FROM analysis_review_items WHERE analysis_report_id=?
                    )
                )
                """,
                (report_id,),
            ).fetchone()[0],
            "evidence": conn.execute("SELECT COUNT(*) FROM analysis_evidence WHERE analysis_report_id=?", (report_id,)).fetchone()[0],
        }
    return {"analysisReports": reports}


def decode_json_text(value: Any, default: Any) -> Any:
    try:
        return json.loads(str(value or ""))
    except json.JSONDecodeError:
        return default


def analysis_evidence_rows(
    conn: sqlite3.Connection,
    analysis_report_id: int,
    where_sql: str = "",
    params: Iterable[Any] = (),
) -> list[dict[str, Any]]:
    rows = dict_rows(
        conn,
        """
        SELECT evidence_id, review_item_id, metric_id, comparison_id, conclusion_id,
               sheet_name, range_address, evidence_role, note, source_text
        FROM analysis_evidence
        WHERE analysis_report_id=?
        """
        + where_sql
        + " ORDER BY evidence_id",
        (analysis_report_id, *tuple(params)),
    )
    return rows


def build_analysis_export(conn: sqlite3.Connection, analysis_report_id: int) -> dict[str, Any]:
    report = first_dict(
        conn,
        """
        SELECT a.*, w.file_name, w.status AS workbook_status, w.fingerprint AS current_workbook_fingerprint
        FROM analysis_reports a
        LEFT JOIN workbooks w ON w.workbook_id=a.workbook_id
        WHERE a.analysis_report_id=?
        LIMIT 1
        """,
        (analysis_report_id,),
    )
    if not report:
        raise ValueError(f"Analysis report not found: {analysis_report_id}")
    export_report = {
        "analysisReportId": int(report["analysis_report_id"]),
        "key": report["analysis_key"],
        "dataset": report["dataset"],
        "workbookId": int(report["workbook_id"]),
        "fileName": report.get("file_name") or "",
        "sourcePath": report["source_path"],
        "workbookFingerprint": report["workbook_fingerprint"],
        "isCurrent": report.get("workbook_status") == "OK" and report.get("current_workbook_fingerprint") == report["workbook_fingerprint"],
        "title": report["title"],
        "type": report["analysis_type"],
        "purpose": report["purpose"],
        "scope": report["scope_text"],
        "status": report["overall_status"],
        "decision": report["overall_decision"],
        "summary": report["overall_summary"],
        "limitations": decode_json_text(report["limitations_json"], []),
        "artifacts": {
            "html": report["dashboard_html_path"],
            "markdown": report["dashboard_markdown_path"],
            "manifest": report["manifest_path"],
        },
        "evidence": analysis_evidence_rows(
            conn,
            analysis_report_id,
            " AND review_item_id IS NULL AND metric_id IS NULL AND comparison_id IS NULL AND conclusion_id IS NULL",
        ),
        "conclusions": [],
    }
    report_conclusions = dict_rows(
        conn,
        """
        SELECT conclusion_id, conclusion_key, sort_order, verdict, label, conclusion_text, scope_text, limitations_json
        FROM analysis_conclusions
        WHERE analysis_report_id=? AND review_item_id IS NULL
        ORDER BY sort_order, conclusion_id
        """,
        (analysis_report_id,),
    )
    for conclusion in report_conclusions:
        conclusion["limitations"] = decode_json_text(conclusion.pop("limitations_json"), [])
        conclusion["evidence"] = analysis_evidence_rows(
            conn,
            analysis_report_id,
            " AND conclusion_id=?",
            (conclusion["conclusion_id"],),
        )
        export_report["conclusions"].append(conclusion)

    reviews: list[dict[str, Any]] = []
    review_rows = dict_rows(
        conn,
        """
        SELECT * FROM analysis_review_items
        WHERE analysis_report_id=?
        ORDER BY sort_order, review_item_id
        """,
        (analysis_report_id,),
    )
    for review in review_rows:
        review_id = int(review["review_item_id"])
        export_review = {
            "reviewItemId": review_id,
            "key": review["review_key"],
            "sortOrder": review["sort_order"],
            "title": review["title"],
            "type": review["review_type"],
            "objective": review["objective"],
            "comparisonBasis": review["comparison_basis"],
            "status": review["status"],
            "decision": review["decision_text"],
            "summary": review["summary_text"],
            "notes": decode_json_text(review["notes_json"], []),
            "evidence": analysis_evidence_rows(
                conn,
                analysis_report_id,
                " AND review_item_id=? AND metric_id IS NULL AND comparison_id IS NULL AND conclusion_id IS NULL",
                (review_id,),
            ),
            "cohorts": [],
            "metrics": [],
            "conclusions": [],
        }
        cohorts = dict_rows(
            conn,
            """
            SELECT cohort_id, cohort_key, cohort_role, label, condition_text, sort_order, attributes_json
            FROM analysis_cohorts
            WHERE review_item_id=?
            ORDER BY sort_order, cohort_id
            """,
            (review_id,),
        )
        for cohort in cohorts:
            cohort["attributes"] = decode_json_text(cohort.pop("attributes_json"), {})
        export_review["cohorts"] = cohorts

        metrics = dict_rows(
            conn,
            """
            SELECT metric_id, metric_key, scope_key, sort_order, label, metric_type, unit,
                   spec_text, definition_text, status, notes_json
            FROM analysis_metrics
            WHERE review_item_id=?
            ORDER BY sort_order, metric_id
            """,
            (review_id,),
        )
        for metric in metrics:
            metric_id = int(metric["metric_id"])
            metric["notes"] = decode_json_text(metric.pop("notes_json"), [])
            metric["evidence"] = analysis_evidence_rows(
                conn,
                analysis_report_id,
                " AND metric_id=? AND comparison_id IS NULL",
                (metric_id,),
            )
            values = dict_rows(
                conn,
                """
                SELECT v.metric_value_id, c.cohort_key, c.cohort_role, c.label AS cohort_label,
                       v.value_number, v.value_text, v.numerator, v.denominator, v.rate_ppm,
                       v.min_value, v.max_value, v.average_value, v.result_status, v.details_json
                FROM analysis_metric_values v
                JOIN analysis_cohorts c ON c.cohort_id=v.cohort_id
                WHERE v.metric_id=?
                ORDER BY c.sort_order, v.metric_value_id
                """,
                (metric_id,),
            )
            for value in values:
                value["details"] = decode_json_text(value.pop("details_json"), {})
            metric["values"] = values
            comparisons = dict_rows(
                conn,
                """
                SELECT c.comparison_id, c.comparison_key,
                       compared.cohort_key AS compared_cohort_key, compared.label AS compared_cohort_label,
                       control.cohort_key AS control_cohort_key, control.label AS control_cohort_label,
                       c.delta_value, c.delta_unit, c.relative_delta_percent, c.direction, c.status,
                       c.summary_text, c.calculation_text, c.details_json
                FROM analysis_comparisons c
                JOIN analysis_cohorts compared ON compared.cohort_id=c.compared_cohort_id
                JOIN analysis_cohorts control ON control.cohort_id=c.control_cohort_id
                WHERE c.metric_id=?
                ORDER BY c.comparison_id
                """,
                (metric_id,),
            )
            for comparison in comparisons:
                comparison["details"] = decode_json_text(comparison.pop("details_json"), {})
                comparison["evidence"] = analysis_evidence_rows(
                    conn,
                    analysis_report_id,
                    " AND comparison_id=?",
                    (comparison["comparison_id"],),
                )
            metric["comparisons"] = comparisons
        export_review["metrics"] = metrics

        conclusions = dict_rows(
            conn,
            """
            SELECT conclusion_id, conclusion_key, sort_order, verdict, label, conclusion_text, scope_text, limitations_json
            FROM analysis_conclusions
            WHERE review_item_id=?
            ORDER BY sort_order, conclusion_id
            """,
            (review_id,),
        )
        for conclusion in conclusions:
            conclusion["limitations"] = decode_json_text(conclusion.pop("limitations_json"), [])
            conclusion["evidence"] = analysis_evidence_rows(
                conn,
                analysis_report_id,
                " AND conclusion_id=?",
                (conclusion["conclusion_id"],),
            )
        export_review["conclusions"] = conclusions
        reviews.append(export_review)
    return {
        "schemaVersion": "universal-analysis-export-v1",
        "exportedAt": now_iso(),
        "report": export_report,
        "reviews": reviews,
    }


def cmd_import_analysis(args: argparse.Namespace) -> int:
    manifest_path = Path(args.input).resolve()
    if not manifest_path.is_file():
        raise SystemExit(f"Analysis manifest not found: {manifest_path}")
    db_path = service_output_path(args.db, UNIVERSAL_GRID_DIR / f"{safe_name(args.dataset or DEFAULT_DATASET)}.sqlite")
    data = read_analysis_manifest(manifest_path)
    with connect_rw(db_path) as conn:
        ensure_universal_schema(conn)
        result = import_analysis_manifest(conn, manifest_path, data, args.dataset or None)
        conn.commit()
    print_json({"status": "ok", "db": str(db_path), "manifest": str(manifest_path), **result})
    return 0


def cmd_verify_analysis(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if not table_exists(conn, "analysis_reports"):
            raise SystemExit("analysis-verify requires a Universal DB with the analysis schema.")
        reports = dict_rows(
            conn,
            "SELECT analysis_report_id FROM analysis_reports WHERE (? IS NULL OR analysis_report_id=?) ORDER BY analysis_report_id",
            (args.report_id, args.report_id),
        )
        results = [verify_analysis_report(conn, int(report["analysis_report_id"])) for report in reports]
    result = {
        "checked": len(results),
        "valid": sum(1 for item in results if item["ok"]),
        "invalid": sum(1 for item in results if not item["ok"]),
        "reports": results,
    }
    print_json(result)
    return 0 if result["invalid"] == 0 else 1


def cmd_inspect_analysis(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if not table_exists(conn, "analysis_reports"):
            raise SystemExit("analysis-inspect requires a Universal DB with the analysis schema.")
        result = inspect_analysis_reports(conn, args.report_id, args.workbook_id, args.dataset or None)
    print_json(result)
    return 0


def cmd_export_analysis(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if not table_exists(conn, "analysis_reports"):
            raise SystemExit("analysis-export requires a Universal DB with the analysis schema.")
        data = build_analysis_export(conn, args.report_id)
    out_path = service_output_path(args.out, OUTPUT_DIR / "analysis-exports" / f"analysis_report_{args.report_id}.json")
    write_json(out_path, data)
    print_json({"status": "ok", "analysisReportId": args.report_id, "export": str(out_path)})
    return 0


def cmd_migrate_knowledge(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_rw(db_path) as conn:
        ensure_universal_schema(conn)
        migrated = migrate_all_legacy_analyses(conn, now_iso)
        integrity = validate_knowledge_integrity(conn)
        if integrity["ok"]:
            conn.commit()
        else:
            conn.rollback()
    print_json({"status": "ok" if integrity["ok"] else "invalid", "db": str(db_path), "migrated": migrated, "integrity": integrity})
    return 0 if integrity["ok"] else 1


def cmd_inspect_knowledge(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if not table_exists(conn, "knowledge_studies"):
            raise SystemExit("knowledge-inspect requires a DB initialized with the canonical knowledge schema.")
        integrity = validate_knowledge_integrity(conn)
        result = {"db": str(db_path), "counts": knowledge_counts(conn), "integrity": integrity}
    print_json(result)
    return 0 if integrity["ok"] else 1


def cmd_import_study(args: argparse.Namespace) -> int:
    manifest_path = Path(args.input).resolve()
    if not manifest_path.is_file():
        raise SystemExit(f"Canonical study manifest not found: {manifest_path}")
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    db_path = Path(args.db).resolve()
    with connect_rw(db_path) as conn:
        ensure_universal_schema(conn)
        ensure_capture_v2_schema(conn)
        imported = import_study_manifest(conn, data, now_iso=now_iso)
        integrity = validate_knowledge_integrity(conn)
        if not integrity["ok"]:
            conn.rollback()
            print_json({"status": "invalid", "db": str(db_path), "manifest": str(manifest_path), "integrity": integrity})
            return 1
        conn.commit()
    print_json(
        {
            "status": "ok",
            "db": str(db_path),
            "manifest": str(manifest_path),
            "imported": imported,
            "integrity": integrity,
        }
    )
    return 0


def cmd_quarantine_analysis(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_rw(db_path) as conn:
        if not table_exists(conn, "workbook_analyses"):
            raise SystemExit(
                "analysis-quarantine requires a DB initialized with the "
                "canonical knowledge schema."
            )
        conn.execute("BEGIN IMMEDIATE")
        try:
            quarantined = quarantine_canonical_analysis(
                conn,
                public_analysis_id=args.public_analysis_id,
                reason=args.reason,
                now_iso=now_iso,
            )
            integrity = validate_knowledge_integrity(conn)
            if not integrity["ok"]:
                conn.rollback()
                print_json(
                    {
                        "status": "invalid",
                        "db": str(db_path),
                        "quarantine": quarantined,
                        "integrity": integrity,
                    }
                )
                return 1
            conn.commit()
        except AnalysisQuarantineError as exc:
            conn.rollback()
            raise SystemExit(str(exc)) from exc
        except Exception:
            conn.rollback()
            raise
    print_json(
        {
            "status": "ok",
            "db": str(db_path),
            "quarantine": quarantined,
            "integrity": integrity,
        }
    )
    return 0


def cmd_inspect_db(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if table_exists(conn, "files"):
            result = inspect_quick_db(conn)
        elif table_exists(conn, "workbooks"):
            result = inspect_universal_db(conn)
        else:
            raise SystemExit("Unknown DB shape. Expected quick-index or universal-grid tables.")
    print_json(result)
    return 0


def cmd_verify_universal_db(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if not table_exists(conn, "workbooks"):
            raise SystemExit("verify-universal-db requires a universal-grid SQLite database.")
        result = verify_universal_db(conn, args.dataset or None, args.limit)
    print_json(result)
    return 0 if result["invalid"] == 0 else 1


def cmd_search(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    query = f"%{args.q}%"
    with connect_ro(db_path) as conn:
        if table_exists(conn, "files"):
            rows = dict_rows(
                conn,
                """
                SELECT file_id, file_name, status, models, categories, structure_family, structure_confidence,
                       metric_candidate_count, measurement_stat_count, comparison_pair_count, term_summary
                FROM files
                WHERE file_name LIKE ?
                   OR models LIKE ?
                   OR categories LIKE ?
                   OR term_summary LIKE ?
                   OR path LIKE ?
                ORDER BY file_id
                LIMIT ?
                """,
                (query, query, query, query, query, args.limit),
            )
            print_json({"dbType": "quick-index", "matches": rows})
        elif table_exists(conn, "workbooks"):
            rows = dict_rows(
                conn,
                """
                SELECT DISTINCT w.workbook_id, w.file_name, w.status, w.sheet_count, w.total_rows,
                                w.total_cells, w.non_empty_cells, w.merge_count
                FROM workbooks w
                LEFT JOIN grid_sheet_rows r ON r.workbook_id=w.workbook_id
                WHERE w.file_name LIKE ?
                   OR w.source_path LIKE ?
                   OR r.row_text LIKE ?
                ORDER BY w.workbook_id
                LIMIT ?
                """,
                (query, query, query, args.limit),
            )
            print_json({"dbType": "universal-grid", "matches": rows})
        else:
            raise SystemExit("Unknown DB shape. Expected quick-index or universal-grid tables.")
    return 0


def row_ref(sheet_name: Any, row_number: Any) -> str:
    sheet = str(sheet_name or "").strip()
    try:
        row = int(row_number)
    except (TypeError, ValueError):
        row = 0
    return f"{sheet}!{row}" if sheet else str(row)


def parse_evidence_refs(evidence: str) -> list[dict[str, Any]]:
    refs = []
    for match in re.finditer(r"(?P<sheet>[^!;\r\n]+)!(?P<row>\d+)", evidence or ""):
        sheet = re.sub(r"^(?:vs|and|or|,|\(|\)|/|-)+\s*", "", match.group("sheet").strip(), flags=re.IGNORECASE)
        row_number = int(match.group("row"))
        refs.append({"rowId": row_ref(sheet, row_number), "sheetName": sheet, "rowNumber": row_number})
    seen = set()
    unique = []
    for ref in refs:
        key = ref["rowId"]
        if key not in seen:
            unique.append(ref)
            seen.add(key)
    return unique


def context_rows_from_rows(rows: list[dict[str, Any]], max_rows: int = 120) -> list[dict[str, Any]]:
    out = []
    for index, row in enumerate(rows):
        text = str(row.get("row_text") or "")
        lower = text.lower()
        if not text.strip():
            continue
        if index < 12 or any(term in lower for term in CONTEXT_TERMS):
            out.append(
                {
                    "rowId": row_ref(row.get("sheet_name"), row.get("row_number")),
                    "sheetName": row.get("sheet_name"),
                    "rowNumber": row.get("row_number"),
                    "rowText": compact(text, 420),
                }
            )
        if len(out) >= max_rows:
            break
    return out


def rate_percent(value: Any) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number * 100.0 if abs(number) <= 1.5 else number


def reviewcase_contract() -> dict[str, Any]:
    return {
        "instruction": "Generate zero, one, or more source-backed ReviewCases. Candidate rows are hints only; sheet rows/cells are source authority.",
        "requiredEvidenceRule": "Every changedFactor, outcome, condition, and numeric claim must cite row/cell evidence from this packet.",
        "packetCompletenessRule": "When packetSelection.dataTruncated is true, do not issue VERIFIED, CAN_USE, or a causal conclusion from omitted rows/cells. Use NEEDS_REVIEW and name the packet limit.",
        "reviewCaseFields": [
            "reviewCaseId",
            "sourceWorkbook",
            "modelReview",
            "changedFactors",
            "outcomes",
            "verification",
        ],
        "verificationStatuses": ["verified", "needs_review", "excluded"],
    }


def packet_cell_has_source_value(cell: dict[str, Any]) -> bool:
    """Keep meaningful cells (including numeric zero) in compact AI packets."""
    value = cell.get("value")
    if value is not None and str(value).strip() != "":
        return True
    return str(cell.get("mergeRole") or cell.get("merge_role") or "").lower() == "anchor"


def compact_packet_cells(cells: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Remove empty grid padding without changing values or source coordinates."""
    return [
        {
            "address": cell.get("address") or "",
            "column": cell.get("column", cell.get("col_number")),
            "colLabel": cell.get("colLabel", cell.get("col_label", "")),
            "value": cell.get("value", cell.get("value_text", "")),
            "rawValue": cell.get("rawValue", cell.get("raw_value_text", "")),
            "mergeRole": cell.get("mergeRole", cell.get("merge_role", "none")),
            "mergeAddress": cell.get("mergeAddress", cell.get("merge_address", "")),
        }
        for cell in cells
        if packet_cell_has_source_value(cell)
    ]


def packet_priority_where_sql() -> tuple[str, list[str]]:
    clauses = ["LOWER(row_text) LIKE ?" for _ in PACKET_PRIORITY_TERMS]
    return "(" + " OR ".join(clauses) + ")", [f"%{term.lower()}%" for term in PACKET_PRIORITY_TERMS]


def compact_universal_packet_rows(
    conn: sqlite3.Connection,
    workbook_id: int,
    sheets: list[dict[str, Any]],
    row_limit: int,
    cell_limit: int,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Return a bounded, stratified, non-empty representation of the source grid.

    The old packet placed a full row grid *and* a flat-cell copy in the JSON.
    Large sheets could therefore exceed the intended cell budget before Codex
    started reading.  This selection keeps header, context and final rows from
    each sheet and emits each retained value once with its original address.
    """
    source_row_count = int(
        conn.execute(
            "SELECT COUNT(*) FROM grid_sheet_rows WHERE workbook_id=? AND non_empty_count > 0",
            (workbook_id,),
        ).fetchone()[0]
    )
    source_cell_count = int(
        conn.execute(
            "SELECT COUNT(*) FROM grid_sheet_cells WHERE workbook_id=? AND (value_text <> '' OR merge_role='anchor')",
            (workbook_id,),
        ).fetchone()[0]
    )
    source_merge_count = int(
        conn.execute("SELECT COUNT(*) FROM merge_ranges WHERE workbook_id=?", (workbook_id,)).fetchone()[0]
    )
    row_limit = max(1, row_limit)
    cell_limit = max(1, cell_limit)
    active_sheets = [sheet for sheet in sheets if int(sheet.get("non_empty_cells") or 0) > 0]
    selected: dict[tuple[int, int], dict[str, Any]] = {}
    priority_where, priority_params = packet_priority_where_sql()

    def add_rows(rows: list[dict[str, Any]], budget: int) -> None:
        for row in rows:
            if len(selected) >= row_limit or budget <= 0:
                return
            key = (int(row["sheet_index"]), int(row["row_number"]))
            if key not in selected:
                selected[key] = row
                budget -= 1

    remaining_global = row_limit
    for sheet_position, sheet in enumerate(active_sheets):
        sheets_left = len(active_sheets) - sheet_position
        sheet_budget = max(1, remaining_global // max(1, sheets_left))
        sheet_index = int(sheet["sheet_index"])
        base_sql = """
            SELECT sheet_index, sheet_name, row_number, non_empty_count, row_text, cells_json
            FROM grid_sheet_rows
            WHERE workbook_id=? AND sheet_index=? AND non_empty_count > 0
        """
        head_limit = min(8, max(1, sheet_budget // 4))
        head_rows = dict_rows(conn, base_sql + " ORDER BY row_number LIMIT ?", (workbook_id, sheet_index, head_limit))
        before = len(selected); add_rows(head_rows, sheet_budget); sheet_budget -= len(selected) - before

        tail_limit = min(8, max(0, sheet_budget // 3))
        if tail_limit:
            tail_rows = dict_rows(conn, base_sql + " ORDER BY row_number DESC LIMIT ?", (workbook_id, sheet_index, tail_limit))
            before = len(selected); add_rows(list(reversed(tail_rows)), sheet_budget); sheet_budget -= len(selected) - before

        if sheet_budget:
            priority_rows = dict_rows(
                conn,
                base_sql + f" AND {priority_where} ORDER BY row_number LIMIT ?",
                (workbook_id, sheet_index, *priority_params, sheet_budget),
            )
            before = len(selected); add_rows(priority_rows, sheet_budget); sheet_budget -= len(selected) - before

        # Use evenly spaced non-empty rows to keep tables represented even
        # where they contain neither an English nor a Korean context keyword.
        if sheet_budget:
            row_count = max(1, int(sheet.get("row_count") or 1))
            step = max(1, row_count // max(1, sheet_budget))
            sampled_rows = dict_rows(
                conn,
                base_sql + " AND ((row_number - ?) % ?) = 0 ORDER BY row_number LIMIT ?",
                (workbook_id, sheet_index, int(sheet.get("used_top") or 1), step, sheet_budget * 2),
            )
            before = len(selected); add_rows(sampled_rows, sheet_budget); sheet_budget -= len(selected) - before

        # A short sheet can still have fewer rows than the sampling interval.
        if sheet_budget:
            fallback_rows = dict_rows(conn, base_sql + " ORDER BY row_number LIMIT ?", (workbook_id, sheet_index, sheet_budget * 2))
            before = len(selected); add_rows(fallback_rows, sheet_budget); sheet_budget -= len(selected) - before
        remaining_global = row_limit - len(selected)
        if remaining_global <= 0:
            break

    rows = [selected[key] for key in sorted(selected)]
    decoded_rows: list[dict[str, Any]] = []
    for row in rows:
        try:
            original_cells = json.loads(row.pop("cells_json") or "[]")
        except json.JSONDecodeError:
            original_cells = []
        compact_cells = compact_packet_cells(original_cells)
        row["rowId"] = row_ref(row["sheet_name"], row["row_number"])
        row["sourceNonEmptyCount"] = int(row.get("non_empty_count") or 0)
        row["cells"] = compact_cells
        decoded_rows.append(row)

    full_cells_by_row = {str(row["rowId"]): list(row["cells"]) for row in decoded_rows}
    available_selected_cells = sum(len(cells) for cells in full_cells_by_row.values())
    # Reserve a fair first slice for every selected row before consuming the
    # remaining budget. A final-result row therefore is not erased by a wide
    # header table earlier in the workbook.
    per_row_limit = max(1, cell_limit // max(1, len(decoded_rows)))
    remaining_cells = cell_limit
    for row in decoded_rows:
        full_cells = full_cells_by_row[str(row["rowId"])]
        keep = min(len(full_cells), per_row_limit, remaining_cells)
        row["cells"] = full_cells[:keep]
        row["includedCellCount"] = keep
        remaining_cells -= keep
    for row in decoded_rows:
        if remaining_cells <= 0:
            break
        full_cells = full_cells_by_row[str(row["rowId"])]
        current_count = len(row["cells"])
        extra = min(max(0, len(full_cells) - current_count), remaining_cells)
        if extra:
            row["cells"].extend(full_cells[current_count : current_count + extra])
            row["includedCellCount"] = len(row["cells"])
            remaining_cells -= extra

    included_cells = sum(len(row["cells"]) for row in decoded_rows)
    selection = {
        "mode": "compact-stratified-v1",
        "rowLimit": row_limit,
        "cellLimit": cell_limit,
        "availableNonEmptyRows": source_row_count,
        "availableNonEmptyCells": source_cell_count,
        "availableMergeRanges": source_merge_count,
        "selectedRows": len(decoded_rows),
        "includedCells": included_cells,
        "rowTruncated": len(decoded_rows) < source_row_count,
        "cellTruncated": included_cells < available_selected_cells,
    }
    selection["dataTruncated"] = bool(selection["rowTruncated"] or selection["cellTruncated"])
    return decoded_rows, selection


def build_quick_packet(conn: sqlite3.Connection, file_id: int, row_limit: int, cell_limit: int, candidate_limit: int) -> dict[str, Any]:
    file_row = first_dict(
        conn,
        """
        SELECT f.*,
               (SELECT COUNT(*) FROM sheet_rows sr WHERE sr.file_id=f.file_id) AS sheet_row_count,
               (SELECT COUNT(*) FROM sheet_cells sc WHERE sc.file_id=f.file_id) AS sheet_cell_count
        FROM files f
        WHERE f.file_id=?
        LIMIT 1
        """,
        (file_id,),
    )
    if not file_row:
        raise SystemExit(f"file_id not found: {file_id}")

    sheets = dict_rows(
        conn,
        "SELECT sheet_name, row_count, col_count, non_empty_count, sample_text FROM sheets WHERE file_id=? ORDER BY sheet_name",
        (file_id,),
    )
    rows = dict_rows(
        conn,
        """
        SELECT sheet_name, row_number, non_empty_count, row_text, cells_json
        FROM sheet_rows
        WHERE file_id=?
        ORDER BY sheet_name, row_number
        LIMIT ?
        """,
        (file_id, row_limit),
    )

    row_map = {row_ref(row["sheet_name"], row["row_number"]): row for row in rows}
    for row in rows:
        row["rowId"] = row_ref(row["sheet_name"], row["row_number"])
        row["cells"] = []

    cells = dict_rows(
        conn,
        """
        SELECT sheet_name, row_number, col_number, col_label, cell_value
        FROM sheet_cells
        WHERE file_id=?
        ORDER BY sheet_name, row_number, col_number
        LIMIT ?
        """,
        (file_id, cell_limit),
    )
    for cell in cells:
        rid = row_ref(cell["sheet_name"], cell["row_number"])
        if rid not in row_map:
            continue
        row_map[rid]["cells"].append(
            {
                "cellId": f"{cell['sheet_name']}!{cell['col_label']}{cell['row_number']}",
                "colNumber": cell["col_number"],
                "colLabel": cell["col_label"],
                "value": cell["cell_value"],
            }
        )

    pair_candidates = dict_rows(
        conn,
        """
        SELECT pair_id, table_title, compare_item, control_condition, test_condition,
               control_input, control_ng, control_rate, test_input, test_ng, test_rate,
               delta_rate, improvement_rate, effect_direction, evidence, pair_confidence
        FROM comparison_pairs
        WHERE file_id=?
        ORDER BY pair_id
        LIMIT ?
        """,
        (file_id, candidate_limit),
    )
    for row in pair_candidates:
        row["controlRatePercent"] = rate_percent(row.get("control_rate"))
        row["testRatePercent"] = rate_percent(row.get("test_rate"))
        row["deltaRatePercentPoint"] = rate_percent(row.get("delta_rate"))
        row["evidenceRows"] = parse_evidence_refs(str(row.get("evidence") or ""))

    metric_candidates = dict_rows(
        conn,
        """
        SELECT metric_id, sheet_name, row_number, table_title, condition_label,
               input_qty, ok_qty, ng_qty, ng_rate, detail, raw_row, parse_confidence
        FROM metric_candidates
        WHERE file_id=?
        ORDER BY metric_id
        LIMIT ?
        """,
        (file_id, candidate_limit),
    )
    for row in metric_candidates:
        row["rowId"] = row_ref(row["sheet_name"], row["row_number"])
        row["ngRatePercent"] = rate_percent(row.get("ng_rate"))

    measurement_candidates = dict_rows(
        conn,
        """
        SELECT stat_id, sheet_name, row_number, item_label, condition_label, spec,
               min_value, max_value, avg_value, sample_count, violation_count,
               raw_row, parse_confidence
        FROM measurement_stats
        WHERE file_id=?
        ORDER BY stat_id
        LIMIT ?
        """,
        (file_id, candidate_limit),
    )
    for row in measurement_candidates:
        row["rowId"] = row_ref(row["sheet_name"], row["row_number"])

    term_hints = dict_rows(
        conn,
        """
        SELECT term_raw, term_type, normalized_name, korean_desc, hit_count, example_context
        FROM term_hits
        WHERE file_id=?
        ORDER BY hit_count DESC, term_raw
        LIMIT ?
        """,
        (file_id, candidate_limit),
    )

    notes = [
        "This packet came from the quick-index DB.",
        "pairCandidates, metricCandidates, measurementCandidates, and termHints are hints only.",
        "Use sheetRows and cells as the source authority.",
    ]
    if len(rows) >= row_limit:
        notes.append("sheetRows were truncated by rowLimit. Increase rowLimit before final generation.")
    if len(cells) >= cell_limit:
        notes.append("attached cells were truncated by cellLimit.")

    return {
        "schemaVersion": "inference-data-ai-reviewcase-packet-v1",
        "createdAt": now_iso(),
        "sourceDbType": "quick-index",
        "notes": notes,
        "reviewCaseContract": reviewcase_contract(),
        "file": file_row,
        "sheets": sheets,
        "contextRows": context_rows_from_rows(rows),
        "sheetRows": rows,
        "pairCandidates": pair_candidates,
        "metricCandidates": metric_candidates,
        "measurementCandidates": measurement_candidates,
        "termHints": term_hints,
    }


def build_universal_packet(conn: sqlite3.Connection, workbook_id: int, row_limit: int, cell_limit: int) -> dict[str, Any]:
    workbook = first_dict(conn, "SELECT * FROM workbooks WHERE workbook_id=? LIMIT 1", (workbook_id,))
    if not workbook:
        raise SystemExit(f"workbook_id not found: {workbook_id}")
    sheets = dict_rows(conn, "SELECT * FROM worksheets WHERE workbook_id=? ORDER BY sheet_index", (workbook_id,))
    rows, selection = compact_universal_packet_rows(conn, workbook_id, sheets, row_limit, cell_limit)
    merges = dict_rows(
        conn,
        """
        SELECT sheet_index, sheet_name, address, top, left_col, bottom, right_col, row_span, column_span, anchor_value
        FROM merge_ranges
        WHERE workbook_id=?
        ORDER BY sheet_index, top, left_col
        LIMIT 500
        """,
        (workbook_id,),
    )

    notes = [
        "This packet came from the universal-grid DB.",
        "No business meaning is assumed at this layer.",
        "Use row/cell coordinates and merge ranges as source authority.",
        "sheetRows retain non-empty source values once; empty grid padding and the duplicate flat cell array are intentionally omitted.",
    ]
    if selection["dataTruncated"]:
        notes.append(
            "Packet is incomplete because row/cell limits were reached. Do not approve or infer a causal conclusion from omitted data; use NEEDS_REVIEW and state the limit."
        )
    if len(merges) < selection["availableMergeRanges"]:
        notes.append("mergeRanges were truncated at 500 entries; omitted merge metadata must not be assumed absent.")
        selection["mergeTruncated"] = True
        selection["dataTruncated"] = True
    else:
        selection["mergeTruncated"] = False

    return {
        "schemaVersion": "inference-data-ai-reviewcase-packet-v2",
        "createdAt": now_iso(),
        "sourceDbType": "universal-grid",
        "notes": notes,
        "reviewCaseContract": reviewcase_contract(),
        "workbook": workbook,
        "sheets": sheets,
        "packetSelection": selection,
        "contextRows": context_rows_from_rows(rows),
        "sheetRows": rows,
        "mergeRanges": merges,
        "candidateHints": [],
    }


def cmd_build_packet(args: argparse.Namespace) -> int:
    db_path = Path(args.db).resolve()
    with connect_ro(db_path) as conn:
        if table_exists(conn, "files"):
            if args.file_id is None:
                raise SystemExit("--file-id is required for quick-index DB.")
            packet = build_quick_packet(conn, args.file_id, args.row_limit, args.cell_limit, args.candidate_limit)
            default_out = PACKET_DIR / f"file_{args.file_id}_reviewcase_packet.json"
        elif table_exists(conn, "workbooks"):
            workbook_id = args.workbook_id if args.workbook_id is not None else args.file_id
            if workbook_id is None:
                raise SystemExit("--workbook-id is required for universal-grid DB.")
            packet = build_universal_packet(conn, workbook_id, args.row_limit, args.cell_limit)
            default_out = PACKET_DIR / f"workbook_{workbook_id}_reviewcase_packet.json"
        else:
            raise SystemExit("Unknown DB shape. Expected quick-index or universal-grid tables.")

    out_path = service_output_path(args.out, default_out)
    write_json(out_path, packet)
    print_json({"status": "ok", "packet": str(out_path)})
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="CLI-only Inference Data AI service for mixed Excel report DB and ReviewCase packet generation."
    )
    sub = parser.add_subparsers(dest="command", required=True)

    init_db = sub.add_parser("init-db", help="Create the universal grid SQLite schema under this service folder.")
    init_db.add_argument("--dataset", default=DEFAULT_DATASET)
    init_db.add_argument("--db", help="Output DB path. Must stay under this service folder.")
    init_db.set_defaults(func=cmd_init_db)

    quick = sub.add_parser("quick-index", help="Run the existing MicroSpeaker quick indexer into this service folder.")
    quick.add_argument("--input", default=str(DEFAULT_INPUT_DIR), help="Excel folder path.")
    quick.add_argument("--dataset", default=DEFAULT_DATASET)
    quick.add_argument("--db", help="Output SQLite path. Must stay under this service folder.")
    quick.add_argument("--html", help="Output dashboard path. Must stay under this service folder.")
    quick.add_argument("--log", help="Output log path. Must stay under this service folder.")
    quick.add_argument("--limit", type=int, default=0)
    quick.add_argument("--force", action="store_true")
    quick.add_argument("--no-html", action="store_true")
    quick.set_defaults(func=cmd_quick_index)

    com = sub.add_parser("com-index", help="Extract Excel through COM and import into the universal grid DB.")
    com.add_argument("--input", default=str(DEFAULT_INPUT_DIR), help="Excel file or folder path.")
    com.add_argument("--dataset", default=DEFAULT_DATASET)
    com.add_argument("--db", help="Output SQLite path. Must stay under this service folder.")
    com.add_argument("--raw-dir", help="Raw COM JSON output directory. Must stay under this service folder.")
    com.add_argument("--log", help="Output log path. Must stay under this service folder.")
    com.add_argument("--limit", type=int, default=0, help="Maximum new or changed files to process after resume filtering.")
    com.add_argument("--include-hidden", action="store_true")
    com.add_argument("--sparse", action="store_true")
    com.add_argument("--force", action="store_true", help="Re-extract and re-import files even when the source fingerprint is unchanged.")
    com.add_argument("--reuse-json", action="store_true", help="Reuse only a validated raw JSON artifact matching the current source and options.")
    com.add_argument("--verify-after-import", action="store_true", help="Verify DB row/cell/merge counts before replacing the prior workbook.")
    com.add_argument("--covered-cell-mode", choices=["blank", "anchor", "raw"], default="blank")
    com.set_defaults(func=cmd_com_index)

    openxml = sub.add_parser(
        "openxml-index",
        help="Capture DRM-free XLSX sources with SHA-256, formulas, styles, dimensions, and merges; images are ignored.",
    )
    openxml.add_argument("--input", help="XLSX file or folder. A pilot manifest sourceRoot is used when omitted.")
    openxml.add_argument("--dataset", default=DEFAULT_DATASET)
    openxml.add_argument("--db", help="Output SQLite path. Must stay under this service folder.")
    openxml.add_argument("--pilot-manifest", help="Optional representative-pilot JSON containing relativePath entries.")
    openxml.add_argument("--offset", type=int, default=0)
    openxml.add_argument("--limit", type=int, default=0)
    openxml.add_argument(
        "--workers",
        type=int,
        default=1,
        help=(
            "Parallel OpenXML readers; DB imports remain serialized in "
            "deterministic source order."
        ),
    )
    openxml.add_argument("--raw-dir", help="Optional retained Capture v2 JSON directory under the service folder.")
    openxml.set_defaults(func=cmd_openxml_index)

    capture_verify = sub.add_parser(
        "capture-v2-verify",
        help="Verify Capture v2 stored counts, canonical bridges, and optionally current source SHA-256.",
    )
    capture_verify.add_argument("--db", required=True)
    capture_verify.add_argument("--all-revisions", action="store_true")
    capture_verify.add_argument("--source-sha256", action="store_true")
    capture_verify.set_defaults(func=cmd_verify_capture_v2)

    semantic_packets = sub.add_parser(
        "semantic-packets",
        help="Build lossless, domain-neutral AI source chunks from one Capture v2 revision.",
    )
    semantic_packets.add_argument("--db", required=True)
    semantic_source = semantic_packets.add_mutually_exclusive_group(required=True)
    semantic_source.add_argument("--revision-id", type=int)
    semantic_source.add_argument("--source-path")
    semantic_packets.add_argument("--max-cells", type=int, default=400)
    semantic_packets.add_argument("--max-rows", type=int, default=50)
    semantic_packets.add_argument("--empty-row-gap", type=int, default=3)
    semantic_packets.add_argument("--out")
    semantic_packets.set_defaults(func=cmd_build_semantic_packets)

    semantic_packet_batch = sub.add_parser(
        "semantic-packets-batch",
        help="Build resumable semantic source packets for current Capture v2 revisions in parallel.",
    )
    semantic_packet_batch.add_argument("--db", required=True)
    semantic_packet_batch.add_argument("--offset", type=int, default=0)
    semantic_packet_batch.add_argument("--limit", type=int, default=0)
    semantic_packet_batch.add_argument("--workers", type=int, default=3)
    semantic_packet_batch.add_argument("--max-cells", type=int, default=400)
    semantic_packet_batch.add_argument("--max-rows", type=int, default=50)
    semantic_packet_batch.add_argument("--empty-row-gap", type=int, default=3)
    semantic_packet_batch.add_argument("--out-dir")
    semantic_packet_batch.add_argument("--force", action="store_true")
    semantic_packet_batch.set_defaults(func=cmd_build_semantic_packets_batch)

    table_first_request = sub.add_parser(
        "table-first-request",
        help=(
            "Build one compact table/text inventory from a lossless semantic "
            "source packet without calling AI."
        ),
    )
    table_first_request.add_argument("--packet", required=True)
    table_first_request.add_argument("--out")
    table_first_request.add_argument("--max-preview-rows", type=int, default=12)
    table_first_request.add_argument(
        "--max-preview-columns",
        type=int,
        default=16,
    )
    table_first_request.add_argument("--max-value-samples", type=int, default=3)
    table_first_request.add_argument("--term-dictionary")
    table_first_request.set_defaults(func=cmd_table_first_request)

    table_first_request_batch = sub.add_parser(
        "table-first-request-batch",
        help=(
            "Build and audit table-first requests for a packet directory "
            "without calling AI."
        ),
    )
    table_first_request_batch.add_argument("--packet-dir", required=True)
    table_first_request_batch.add_argument("--out-dir")
    table_first_request_batch.add_argument("--offset", type=int, default=0)
    table_first_request_batch.add_argument("--limit", type=int, default=0)
    table_first_request_batch.add_argument("--workers", type=int, default=3)
    table_first_request_batch.add_argument(
        "--max-preview-rows", type=int, default=12
    )
    table_first_request_batch.add_argument(
        "--max-preview-columns",
        type=int,
        default=16,
    )
    table_first_request_batch.add_argument(
        "--max-value-samples", type=int, default=3
    )
    table_first_request_batch.add_argument("--term-dictionary")
    table_first_request_batch.add_argument(
        "--oversized-request-bytes",
        type=int,
        default=240000,
        help="Flag requests larger than this many UTF-8 JSON bytes.",
    )
    table_first_request_batch.add_argument(
        "--checkpoint-every",
        type=int,
        default=25,
        help="Rewrite the resumable audit report after this many completions.",
    )
    table_first_request_batch.set_defaults(func=cmd_table_first_request_batch)

    table_first_analyze = sub.add_parser(
        "table-first-analyze",
        help=(
            "Classify one compact workbook request in one AI call and write a "
            "non-approved deterministic study/evidence projection."
        ),
    )
    table_first_analyze.add_argument("--request", required=True)
    table_first_analyze.add_argument("--out")
    table_first_analyze.add_argument("--projection-out")
    table_first_analyze.add_argument("--model")
    table_first_analyze.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="low",
    )
    table_first_analyze.add_argument("--timeout", type=int, default=600)
    table_first_analyze.set_defaults(func=cmd_table_first_analyze)

    table_first_batch = sub.add_parser(
        "table-first-batch",
        help=(
            "Build, analyze, project, and audit a resumable directory of "
            "semantic workbook packets with one AI call per workbook."
        ),
    )
    table_first_batch.add_argument("--packet-dir", required=True)
    table_first_batch.add_argument("--out-dir")
    table_first_batch.add_argument("--offset", type=int, default=0)
    table_first_batch.add_argument("--limit", type=int, default=0)
    table_first_batch.add_argument("--workers", type=int, default=3)
    table_first_batch.add_argument("--max-preview-rows", type=int, default=12)
    table_first_batch.add_argument(
        "--max-preview-columns",
        type=int,
        default=16,
    )
    table_first_batch.add_argument("--max-value-samples", type=int, default=3)
    table_first_batch.add_argument("--term-dictionary")
    table_first_batch.add_argument("--model")
    table_first_batch.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="low",
    )
    table_first_batch.add_argument("--timeout", type=int, default=600)
    table_first_batch.add_argument("--force", action="store_true")
    table_first_batch.set_defaults(func=cmd_table_first_batch)

    table_first_html = sub.add_parser(
        "table-first-html",
        help=(
            "Render a static index and per-workbook HTML report from a "
            "completed table-first analysis batch without calling AI."
        ),
    )
    table_first_html.add_argument("--batch-dir", required=True)
    table_first_html.add_argument("--out-dir")
    table_first_html.set_defaults(func=cmd_table_first_html)

    table_first_history_index = sub.add_parser(
        "table-first-history-index",
        help=(
            "Build one searchable SQLite history index from completed "
            "table-first projections without calling AI or rendering HTML."
        ),
    )
    table_first_history_index.add_argument("--batch-dir", required=True)
    table_first_history_index.add_argument("--db", required=True)
    table_first_history_index.add_argument(
        "--allow-running",
        action="store_true",
        help="Index only artifacts already present in a running batch.",
    )
    table_first_history_index.set_defaults(func=cmd_table_first_history_index)

    table_first_history_query = sub.add_parser(
        "table-first-history-query",
        help=(
            "Search table-first Study histories and create a deterministic "
            "Korean evidence answer without another AI call."
        ),
    )
    table_first_history_query.add_argument("--db", required=True)
    table_first_history_query.add_argument("--question", required=True)
    table_first_history_query.add_argument("--limit", type=int, default=30)
    table_first_history_query.add_argument("--out-pack")
    table_first_history_query.add_argument("--out-json")
    table_first_history_query.add_argument("--out-markdown")
    table_first_history_query.set_defaults(func=cmd_table_first_history_query)

    table_first_contextual_query = sub.add_parser(
        "table-first-contextual-query",
        help=(
            "Retrieve broad table-first candidates, ask AI to understand the "
            "question context and direct evidence relation, and create a "
            "concise citation-bound Korean answer."
        ),
    )
    table_first_contextual_query.add_argument("--db", required=True)
    table_first_contextual_query.add_argument("--question", required=True)
    table_first_contextual_query.add_argument(
        "--candidate-limit", type=int, default=40
    )
    table_first_contextual_query.add_argument(
        "--detail-candidate-limit", type=int, default=18
    )
    table_first_contextual_query.add_argument(
        "--max-fact-count", type=int, default=240
    )
    table_first_contextual_query.add_argument("--model")
    table_first_contextual_query.add_argument(
        "--reasoning-effort",
        choices=("minimal", "low", "medium", "high", "xhigh"),
        default="medium",
    )
    table_first_contextual_query.add_argument(
        "--timeout-seconds", type=int, default=600
    )
    table_first_contextual_query.add_argument("--out-request")
    table_first_contextual_query.add_argument("--out-json")
    table_first_contextual_query.add_argument("--out-markdown")
    table_first_contextual_query.set_defaults(
        func=cmd_table_first_contextual_query
    )

    table_first_relevance_query = sub.add_parser(
        "table-first-relevance-query",
        help=(
            "Retrieve broad table-first candidates and ask AI only whether "
            "each Study is needed for the question, without interpreting results."
        ),
    )
    table_first_relevance_query.add_argument("--db", required=True)
    table_first_relevance_query.add_argument("--question", required=True)
    table_first_relevance_query.add_argument(
        "--candidate-limit", type=int, default=200
    )
    table_first_relevance_query.add_argument("--model")
    table_first_relevance_query.add_argument(
        "--reasoning-effort",
        choices=("minimal", "low", "medium", "high", "xhigh"),
        default="medium",
    )
    table_first_relevance_query.add_argument(
        "--timeout-seconds", type=int, default=600
    )
    table_first_relevance_query.add_argument("--out-request")
    table_first_relevance_query.add_argument("--out-json")
    table_first_relevance_query.add_argument("--out-markdown")
    table_first_relevance_query.set_defaults(
        func=cmd_table_first_relevance_query
    )

    table_first_history_detail = sub.add_parser(
        "table-first-history-detail",
        help="Resolve one TF-EVD citation to its source and saved request preview.",
    )
    table_first_history_detail.add_argument("--db", required=True)
    table_first_history_detail.add_argument("--evidence-id", required=True)
    table_first_history_detail.add_argument("--out")
    table_first_history_detail.set_defaults(func=cmd_table_first_history_detail)

    table_first_history_validate = sub.add_parser(
        "table-first-history-answer-validate",
        help="Validate deterministic wording and the exact TF-EVD citation set.",
    )
    table_first_history_validate.add_argument("--pack", required=True)
    table_first_history_validate.add_argument("--answer", required=True)
    table_first_history_validate.set_defaults(
        func=cmd_validate_table_first_history_answer
    )

    table_first_history_acceptance = sub.add_parser(
        "table-first-history-acceptance",
        help=(
            "Run the representative golden-question acceptance suite against "
            "a completed table-first history index without calling AI."
        ),
    )
    table_first_history_acceptance.add_argument("--db", required=True)
    table_first_history_acceptance.add_argument(
        "--manifest",
        default=str(SERVICE_DIR / "pilot" / "representative-pilot-v1.json"),
    )
    table_first_history_acceptance.add_argument("--limit", type=int, default=30)
    table_first_history_acceptance.add_argument("--out")
    table_first_history_acceptance.set_defaults(
        func=cmd_table_first_history_acceptance
    )

    semantic_locate = sub.add_parser(
        "semantic-locate",
        help="Run resumable, parallel, read-only AI locator passes for semantic source chunks.",
    )
    semantic_locate.add_argument("--packet", required=True)
    semantic_locate.add_argument("--dataset", default=DEFAULT_DATASET)
    semantic_locate.add_argument("--chunk-id", action="append")
    semantic_locate.add_argument("--offset", type=int, default=0)
    semantic_locate.add_argument("--limit", type=int, default=0)
    semantic_locate.add_argument("--workers", type=int, default=3)
    semantic_locate.add_argument(
        "--batch-size",
        type=int,
        default=6,
        help="Maximum independent source chunks per AI call.",
    )
    semantic_locate.add_argument(
        "--batch-max-bytes",
        type=int,
        default=240000,
        help="Maximum estimated serialized source bytes per AI call; a single oversized chunk is retained.",
    )
    semantic_locate.add_argument("--model")
    semantic_locate.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="medium",
    )
    semantic_locate.add_argument("--timeout", type=int, default=900)
    semantic_locate.add_argument("--out-dir")
    semantic_locate.add_argument("--force", action="store_true")
    semantic_locate.add_argument(
        "--all-chunks-ai",
        action="store_true",
        help="Send numeric-only continuation chunks to AI too; normally they remain available for on-demand evidence retrieval.",
    )
    semantic_locate.set_defaults(func=cmd_semantic_locate)

    semantic_draft = sub.add_parser(
        "semantic-draft",
        help="Consolidate complete locator results into a source-validated, non-self-approved Study draft.",
    )
    semantic_draft.add_argument("--packet", required=True)
    semantic_draft.add_argument("--locator-dir", required=True)
    semantic_draft.add_argument("--db", required=True)
    semantic_draft.add_argument("--dataset", default=DEFAULT_DATASET)
    semantic_draft.add_argument("--model")
    semantic_draft.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="medium",
    )
    semantic_draft.add_argument("--timeout", type=int, default=1800)
    semantic_draft.add_argument("--out")
    semantic_draft.set_defaults(func=cmd_semantic_draft)

    semantic_validate = sub.add_parser(
        "semantic-validate-draft",
        help="Normalize a saved AI draft and verify its source identity, ranges, and numeric evidence.",
    )
    semantic_validate.add_argument("--input", required=True)
    semantic_validate.add_argument("--packet", required=True)
    semantic_validate.add_argument("--db", required=True)
    semantic_validate.add_argument("--dataset", default=DEFAULT_DATASET)
    semantic_validate.add_argument("--out")
    semantic_validate.set_defaults(func=cmd_semantic_validate_draft)

    evidence_query = sub.add_parser(
        "evidence-query",
        help="Build a domain-neutral evidence pack and separate answer-eligible effects from excluded candidates.",
    )
    evidence_query.add_argument("--db", required=True)
    evidence_query.add_argument("--question", required=True)
    evidence_query.add_argument("--out")
    evidence_query.set_defaults(func=cmd_evidence_query)

    evidence_detail = sub.add_parser(
        "evidence-detail",
        help="Read one stable EVD ID from its exact current Capture v2 revision for source-table preview.",
    )
    evidence_detail.add_argument("--db", required=True)
    evidence_detail.add_argument("--evidence-id", required=True)
    evidence_detail.add_argument("--out")
    evidence_detail.set_defaults(func=cmd_evidence_detail)

    related_studies = sub.add_parser(
        "related-studies",
        help="Find exact-content duplicates and lexically related current Studies for one DATA ID or revision UID.",
    )
    related_studies.add_argument("--db", required=True)
    related_studies.add_argument("--target", required=True)
    related_studies.add_argument("--limit", type=int, default=25)
    related_studies.add_argument("--out")
    related_studies.set_defaults(func=cmd_related_studies)

    concept_candidates = sub.add_parser(
        "concept-candidates",
        help=(
            "List and filter immutable schema-candidate identities for "
            "explicit human concept curation."
        ),
    )
    concept_candidates.add_argument("--db", required=True)
    concept_candidates.add_argument(
        "--status",
        choices=["OPEN", "APPROVED", "MERGED", "REJECTED", "ALL"],
        default="OPEN",
    )
    concept_candidates.add_argument("--kind")
    concept_candidates.add_argument("--query")
    concept_candidates.add_argument("--limit", type=int, default=500)
    concept_candidates.set_defaults(func=cmd_concept_candidates)

    concept_list = sub.add_parser(
        "concept-list",
        help=(
            "List canonical concepts and the human/seed aliases consumed "
            "by evidence retrieval."
        ),
    )
    concept_list.add_argument("--db", required=True)
    concept_list.add_argument(
        "--status",
        choices=["ACTIVE", "DEPRECATED", "ALL"],
        default="ACTIVE",
    )
    concept_list.add_argument("--kind")
    concept_list.add_argument("--query")
    concept_list.add_argument("--limit", type=int, default=500)
    concept_list.set_defaults(func=cmd_concept_list)

    concept_resolve = sub.add_parser(
        "concept-resolve",
        help=(
            "Atomically CREATE, MERGE, or REJECT one OPEN CONCEPT:* "
            "candidate with immutable human provenance."
        ),
    )
    concept_resolve.add_argument("--db", required=True)
    concept_resolve.add_argument("--candidate-uid", required=True)
    concept_resolve.add_argument(
        "--action",
        required=True,
        choices=["CREATE", "MERGE", "REJECT"],
    )
    concept_resolve.add_argument("--canonical-name", default="")
    concept_resolve.add_argument("--concept-uid", default="")
    concept_resolve.add_argument("--alias", default="")
    concept_resolve.add_argument("--reviewer", required=True)
    concept_resolve.add_argument("--note", required=True)
    concept_resolve.set_defaults(func=cmd_concept_resolve)

    concept_alias = sub.add_parser(
        "concept-alias-upsert",
        help=(
            "Upsert one stable HUMAN_APPROVED alias on an ACTIVE concept "
            "with immutable approval history."
        ),
    )
    concept_alias.add_argument("--db", required=True)
    concept_alias.add_argument("--concept-uid", required=True)
    concept_alias.add_argument("--alias", required=True)
    concept_alias.add_argument("--reviewer", required=True)
    concept_alias.add_argument("--note", required=True)
    concept_alias.set_defaults(func=cmd_concept_alias_upsert)

    review_queue = sub.add_parser(
        "review-queue",
        help="List current fail-closed comparison drafts that require an explicit human decision.",
    )
    review_queue.add_argument("--db", required=True)
    review_queue.add_argument("--limit", type=int, default=500)
    review_queue.add_argument("--out")
    review_queue.set_defaults(func=cmd_review_queue)

    review_detail = sub.add_parser(
        "review-detail",
        help="Show one CMP draft, paired values, direct EVD ranges, and approval blockers without changing the DB.",
    )
    review_detail.add_argument("--db", required=True)
    review_detail.add_argument("--comparison-id", required=True)
    review_detail.add_argument("--out")
    review_detail.set_defaults(func=cmd_review_detail)

    review_decide = sub.add_parser(
        "review-decide",
        help="Record an explicit human CMP decision; approval recalculates effects only from verified current-revision evidence.",
    )
    review_decide.add_argument("--db", required=True)
    review_decide.add_argument("--comparison-id", required=True)
    review_decide.add_argument(
        "--decision",
        required=True,
        choices=["APPROVE", "REJECT", "EXCLUDE", "RETURN_TO_REVIEW"],
    )
    review_decide.add_argument("--reviewer", required=True)
    review_decide.add_argument("--reason", required=True)
    review_decide.add_argument(
        "--study-comparability",
        choices=["VALID", "PARTIAL", "INVALID", "UNASSESSED"],
    )
    review_decide.add_argument(
        "--study-confounding",
        choices=["NONE", "POSSIBLE", "CONFOUNDED", "UNASSESSED"],
    )
    review_decide.add_argument(
        "--comparison-validity",
        choices=["VALID", "NEEDS_REVIEW", "INVALID", "EXCLUDED"],
    )
    review_decide.add_argument(
        "--comparison-confounding",
        choices=["NONE", "POSSIBLE", "CONFOUNDED", "UNASSESSED"],
    )
    review_decide.add_argument("--matching-basis")
    review_decide.set_defaults(func=cmd_review_decide)

    golden_acceptance = sub.add_parser(
        "golden-acceptance",
        help="Run every representative pilot question and distinguish pending ingestion from retrieval failures.",
    )
    golden_acceptance.add_argument("--db", required=True)
    golden_acceptance.add_argument(
        "--manifest",
        default=str(SERVICE_DIR / "pilot" / "representative-pilot-v1.json"),
    )
    golden_acceptance.add_argument("--out-dir")
    golden_acceptance.set_defaults(func=cmd_golden_acceptance)

    evidence_answer = sub.add_parser(
        "evidence-answer",
        help="Create a deterministic Korean answer from a canonical evidence pack; only answer-eligible effects can produce quantitative claims.",
    )
    evidence_answer_source = evidence_answer.add_mutually_exclusive_group(
        required=True
    )
    evidence_answer_source.add_argument(
        "--pack",
        help="Existing canonical evidence-pack JSON.",
    )
    evidence_answer_source.add_argument(
        "--db",
        help="Canonical SQLite DB; requires --question.",
    )
    evidence_answer.add_argument(
        "--question",
        help="Domain-neutral natural-language question used with --db.",
    )
    evidence_answer.add_argument("--out-json")
    evidence_answer.add_argument("--out-markdown")
    evidence_answer.set_defaults(func=cmd_evidence_answer)

    evidence_answer_validate = sub.add_parser(
        "evidence-answer-validate",
        help="Reject any answer whose values, IDs, citations, or deterministic wording differ from its evidence pack.",
    )
    evidence_answer_validate.add_argument("--pack", required=True)
    evidence_answer_validate.add_argument("--answer", required=True)
    evidence_answer_validate.set_defaults(func=cmd_validate_evidence_answer)

    form_group_review = sub.add_parser(
        "form-group-review",
        help=(
            "Group captured new/similar Excel layouts into stable review "
            "families and write the current human-decision state."
        ),
    )
    form_group_review.add_argument("--db", required=True)
    form_group_review.add_argument("--report", required=True)
    form_group_review.add_argument(
        "--out",
        dest="review_out",
    )
    form_group_review.add_argument("--output-root")
    form_group_review.set_defaults(func=cmd_form_group_review)

    form_family_analyze = sub.add_parser(
        "form-family-analyze",
        help=(
            "Analyze one form-family representative and validation samples "
            "with AI, then hold the extraction contract for human approval."
        ),
    )
    form_family_analyze.add_argument("--db", required=True)
    form_family_analyze.add_argument("--report", required=True)
    form_family_analyze.add_argument("--family-id", required=True)
    form_family_analyze.add_argument("--out")
    form_family_analyze.add_argument("--review-out")
    form_family_analyze.add_argument("--output-root")
    form_family_analyze.add_argument("--codex")
    form_family_analyze.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="medium",
    )
    form_family_analyze.add_argument(
        "--timeout",
        type=int,
        default=900,
    )
    form_family_analyze.set_defaults(func=cmd_form_family_analyze)

    form_family_decide = sub.add_parser(
        "form-family-decide",
        help=(
            "Record a human form-family decision and immediately rebuild "
            "the preflight report and full-processing manifest."
        ),
    )
    form_family_decide.add_argument("--db", required=True)
    form_family_decide.add_argument("--report", required=True)
    form_family_decide.add_argument("--family-id", required=True)
    form_family_decide.add_argument(
        "--decision",
        choices=["REGISTER_NEW", "LINK_EXISTING", "EXCLUDE"],
        required=True,
    )
    form_family_decide.add_argument("--reviewer", required=True)
    form_family_decide.add_argument("--display-name", default="")
    form_family_decide.add_argument(
        "--linked-form-signature-id",
        default="",
    )
    form_family_decide.add_argument("--notes", default="")
    form_family_decide.add_argument("--review-out")
    form_family_decide.add_argument("--output-root")
    form_family_decide.set_defaults(func=cmd_form_family_decide)

    form_pipeline_complete = sub.add_parser(
        "form-pipeline-complete",
        help=(
            "Resume COM preflight, analyze and auto-decide every pending "
            "form family with checkpoints, rebuild the manifest, and run "
            "the selected COM corpus through completion."
        ),
    )
    form_pipeline_complete.add_argument("--db", required=True)
    form_pipeline_complete.add_argument("--input", required=True)
    form_pipeline_complete.add_argument("--output-root", required=True)
    form_pipeline_complete.add_argument("--reviewer", required=True)
    form_pipeline_complete.add_argument(
        "--dataset",
        default=DEFAULT_DATASET,
    )
    form_pipeline_complete.add_argument(
        "--analysis-workers",
        type=int,
        default=2,
    )
    form_pipeline_complete.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="low",
    )
    form_pipeline_complete.add_argument(
        "--analysis-timeout",
        type=int,
        default=900,
    )
    form_pipeline_complete.add_argument(
        "--com-timeout-seconds",
        type=float,
        default=300.0,
    )
    form_pipeline_complete.add_argument("--codex")
    form_pipeline_complete.add_argument(
        "--exclude-on-analysis-error",
        action="store_true",
        help=(
            "Fail closed by marking a family excluded after its AI call "
            "fails; source files and captures are preserved."
        ),
    )
    form_pipeline_complete.add_argument(
        "--skip-corpus",
        action="store_true",
        help="Stop after preflight, family decisions, and manifest rebuild.",
    )
    form_pipeline_complete.add_argument(
        "--max-families",
        type=int,
        default=0,
        help=(
            "Process at most this many pending families for a bounded "
            "verification run; zero means all and is required for corpus."
        ),
    )
    form_pipeline_complete.add_argument(
        "--draft-monolithic-max-bytes",
        type=int,
        default=400_000,
        help=(
            "Use staged source-complete drafting when the exact one-call "
            "prompt exceeds this byte limit."
        ),
    )
    form_pipeline_complete.set_defaults(
        func=cmd_form_pipeline_complete
    )

    form_preflight = sub.add_parser(
        "form-preflight",
        help=(
            "Capture a local Excel archive through read-only COM, compare "
            "form structure with analyzed workbooks, and stop before AI."
        ),
    )
    form_preflight.add_argument("--db", required=True)
    form_preflight.add_argument("--input", required=True)
    form_preflight.add_argument("--out")
    form_preflight.add_argument(
        "--output-root",
        help=(
            "Explicit allowed root for --out. WPF supplies its configured "
            "output root; defaults to the service directory."
        ),
    )
    form_preflight.add_argument("--dataset", default=DEFAULT_DATASET)
    form_preflight.add_argument(
        "--cancel-file",
        help=(
            "Stop cooperatively when this marker file exists; the current "
            "dedicated COM worker is terminated and partial results remain."
        ),
    )
    form_preflight.add_argument(
        "--com-timeout-seconds",
        type=float,
        default=300.0,
        help=(
            "Maximum seconds allowed for one isolated Excel COM workbook; "
            "a timeout is recorded as one failed item and the run continues."
        ),
    )
    form_preflight.add_argument(
        "--retry-failed-captures",
        action="store_true",
        help=(
            "Retry unchanged files whose previous preflight result was "
            "CAPTURE_FAILED; default reuses the failure without reopening "
            "Excel."
        ),
    )
    form_preflight.add_argument(
        "--inspect-auth-dialog",
        action="store_true",
    )
    form_preflight.add_argument(
        "--dismiss-auth-dialog",
        action="store_true",
    )
    form_preflight.add_argument("--auth-dialog-title", default="")
    form_preflight.add_argument("--auth-dialog-class", default="")
    form_preflight.add_argument("--auth-dialog-button", default="")
    form_preflight.set_defaults(func=cmd_form_preflight)

    ingest_xlsx = sub.add_parser(
        "ingest-workbook",
        help="Durably capture and analyze one Excel workbook into the review-gated canonical DB; COM is available for DRM sources and images are ignored.",
    )
    ingest_xlsx.add_argument("--db", required=True)
    ingest_xlsx.add_argument("--input", required=True)
    ingest_xlsx.add_argument("--artifact-root")
    ingest_xlsx.add_argument("--dataset", default=DEFAULT_DATASET)
    ingest_xlsx.add_argument(
        "--capture-backend",
        choices=["openxml", "com"],
        default="openxml",
        help=(
            "Use Excel COM for DRM/policy-protected files or OpenXML for "
            "existing decrypted XLSX files."
        ),
    )
    ingest_xlsx.add_argument(
        "--covered-cell-mode",
        choices=["blank", "anchor", "raw"],
        default="blank",
    )
    ingest_xlsx.add_argument(
        "--exclude-hidden-sheets",
        dest="include_hidden_sheets",
        action="store_false",
        default=True,
    )
    ingest_xlsx.add_argument(
        "--inspect-auth-dialog",
        action="store_true",
        help="Report Excel-owned dialog metadata without closing it.",
    )
    ingest_xlsx.add_argument(
        "--dismiss-auth-dialog",
        action="store_true",
        help=(
            "Click only the exact Excel-owned dialog button identified by "
            "the three auth-dialog match arguments."
        ),
    )
    ingest_xlsx.add_argument("--auth-dialog-title", default="")
    ingest_xlsx.add_argument("--auth-dialog-class", default="")
    ingest_xlsx.add_argument("--auth-dialog-button", default="")
    ingest_xlsx.add_argument(
        "--auth-dialog-timeout",
        type=float,
        default=30.0,
    )
    ingest_xlsx.add_argument("--no-resume", action="store_true")
    ingest_xlsx.add_argument("--max-cells", type=int, default=400)
    ingest_xlsx.add_argument("--max-rows", type=int, default=50)
    ingest_xlsx.add_argument("--empty-row-gap", type=int, default=3)
    ingest_xlsx.add_argument("--workers", type=int, default=3)
    ingest_xlsx.add_argument("--batch-size", type=int, default=6)
    ingest_xlsx.add_argument("--batch-max-bytes", type=int, default=240000)
    ingest_xlsx.add_argument(
        "--draft-monolithic-max-bytes",
        type=int,
        default=400000,
    )
    ingest_xlsx.add_argument(
        "--draft-fragment-max-chunks",
        type=int,
        default=8,
    )
    ingest_xlsx.add_argument(
        "--draft-fragment-max-cells",
        type=int,
        default=2000,
    )
    ingest_xlsx.add_argument(
        "--draft-fragment-max-bytes",
        type=int,
        default=400000,
    )
    ingest_xlsx.add_argument(
        "--draft-fragment-workers",
        type=int,
        default=3,
    )
    ingest_xlsx.add_argument(
        "--derive-formula-values",
        action="store_true",
        help=(
            "Evaluate only the restricted same-sheet A1 formula grammar "
            "into a checksum-validated derived overlay; Capture v2 remains "
            "unchanged."
        ),
    )
    ingest_xlsx.add_argument(
        "--repair-rejected-draft",
        action="store_true",
        help=(
            "On resume only, run one source-backed AI repair from the "
            "current rejected manifest instead of repeating the original "
            "exact draft request."
        ),
    )
    ingest_xlsx.add_argument(
        "--repair-unselected-source",
        action="store_true",
        help=(
            "Promote exact required source cells missed by the locator to "
            "NEEDS_REVIEW candidate sections instead of failing selection."
        ),
    )
    ingest_xlsx.add_argument("--model")
    ingest_xlsx.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="medium",
    )
    ingest_xlsx.add_argument("--locator-timeout", type=int, default=900)
    ingest_xlsx.add_argument("--draft-timeout", type=int, default=1800)
    ingest_xlsx.set_defaults(func=cmd_ingest_workbook)

    ingest_corpus = sub.add_parser(
        "ingest-corpus",
        help="Durably account for and ingest a deterministic Excel corpus slice through the same review-gated workflow.",
    )
    ingest_corpus.add_argument("--db", required=True)
    ingest_corpus.add_argument("--input", required=True)
    ingest_corpus.add_argument("--artifact-root")
    ingest_corpus.add_argument(
        "--source-manifest",
        help="Optional JSON workbooks[].relativePath selection, such as the representative pilot manifest.",
    )
    ingest_corpus.add_argument("--journal")
    ingest_corpus.add_argument("--out")
    ingest_corpus.add_argument("--dataset", default=DEFAULT_DATASET)
    ingest_corpus.add_argument(
        "--capture-backend",
        choices=["openxml", "com"],
        default="openxml",
    )
    ingest_corpus.add_argument(
        "--covered-cell-mode",
        choices=["blank", "anchor", "raw"],
        default="blank",
    )
    ingest_corpus.add_argument(
        "--exclude-hidden-sheets",
        dest="include_hidden_sheets",
        action="store_false",
        default=True,
    )
    ingest_corpus.add_argument(
        "--inspect-auth-dialog",
        action="store_true",
    )
    ingest_corpus.add_argument(
        "--dismiss-auth-dialog",
        action="store_true",
    )
    ingest_corpus.add_argument("--auth-dialog-title", default="")
    ingest_corpus.add_argument("--auth-dialog-class", default="")
    ingest_corpus.add_argument("--auth-dialog-button", default="")
    ingest_corpus.add_argument(
        "--auth-dialog-timeout",
        type=float,
        default=30.0,
    )
    ingest_corpus.add_argument("--offset", type=int, default=0)
    ingest_corpus.add_argument("--limit", type=int, default=0)
    ingest_corpus.add_argument("--workbook-workers", type=int, default=1)
    ingest_corpus.add_argument("--com-workers", type=int, default=1)
    ingest_corpus.add_argument("--packet-workers", type=int, default=3)
    ingest_corpus.add_argument("--ai-workers", type=int, default=3)
    ingest_corpus.add_argument("--db-workers", type=int, default=1)
    ingest_corpus.add_argument("--locator-workers", type=int, default=3)
    ingest_corpus.add_argument("--retry-failed", action="store_true")
    ingest_corpus.add_argument(
        "--inventory-only",
        action="store_true",
        help="Hash and journal the complete deterministic inventory without ingesting any workbook.",
    )
    ingest_corpus.add_argument("--no-resume", action="store_true")
    ingest_corpus.add_argument("--max-cells", type=int, default=400)
    ingest_corpus.add_argument("--max-rows", type=int, default=50)
    ingest_corpus.add_argument("--empty-row-gap", type=int, default=3)
    ingest_corpus.add_argument("--batch-size", type=int, default=6)
    ingest_corpus.add_argument(
        "--batch-max-bytes",
        type=int,
        default=240000,
    )
    ingest_corpus.add_argument(
        "--draft-monolithic-max-bytes",
        type=int,
        default=400000,
    )
    ingest_corpus.add_argument(
        "--draft-fragment-max-chunks",
        type=int,
        default=8,
    )
    ingest_corpus.add_argument(
        "--draft-fragment-max-cells",
        type=int,
        default=2000,
    )
    ingest_corpus.add_argument(
        "--draft-fragment-max-bytes",
        type=int,
        default=400000,
    )
    ingest_corpus.add_argument(
        "--draft-fragment-workers",
        type=int,
        default=3,
    )
    ingest_corpus.add_argument(
        "--derive-formula-values",
        action="store_true",
        help=(
            "Opt in every selected workbook to restricted deterministic "
            "formula overlays; any unsupported formula fails closed."
        ),
    )
    ingest_corpus.add_argument(
        "--repair-rejected-draft",
        action="store_true",
        help=(
            "On retry only, repair each current rejected manifest from its "
            "validator error; initial PENDING workbooks remain exact."
        ),
    )
    ingest_corpus.add_argument(
        "--repair-unselected-source",
        action="store_true",
        help=(
            "On retry, promote exact required source cells missed by the "
            "locator to deterministic NEEDS_REVIEW candidate sections."
        ),
    )
    ingest_corpus.add_argument("--model")
    ingest_corpus.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        default="medium",
    )
    ingest_corpus.add_argument("--locator-timeout", type=int, default=900)
    ingest_corpus.add_argument("--draft-timeout", type=int, default=1800)
    ingest_corpus.set_defaults(func=cmd_ingest_corpus)

    inspect = sub.add_parser("inspect-db", help="Inspect quick-index or universal-grid SQLite counts.")
    inspect.add_argument("--db", required=True)
    inspect.set_defaults(func=cmd_inspect_db)

    verify_universal = sub.add_parser("verify-universal-db", help="Validate universal-grid records against their raw COM JSON artifacts.")
    verify_universal.add_argument("--db", required=True)
    verify_universal.add_argument("--dataset", help="Optional dataset filter.")
    verify_universal.add_argument("--limit", type=int, default=0)
    verify_universal.set_defaults(func=cmd_verify_universal_db)

    analysis_import = sub.add_parser(
        "analysis-import",
        help="Import a reusable, evidence-linked analysis summary into a Universal DB.",
    )
    analysis_import.add_argument("--input", required=True, help="Analysis manifest JSON path.")
    analysis_import.add_argument("--db", help="Universal SQLite path. Must stay under this service folder.")
    analysis_import.add_argument("--dataset", help="Optional expected dataset; the manifest source remains authoritative.")
    analysis_import.set_defaults(func=cmd_import_analysis)

    analysis_verify = sub.add_parser(
        "analysis-verify",
        help="Verify analysis evidence ranges, rate arithmetic, comparison deltas, and source freshness.",
    )
    analysis_verify.add_argument("--db", required=True)
    analysis_verify.add_argument("--report-id", type=int, help="Optional analysis report ID; verifies all when omitted.")
    analysis_verify.set_defaults(func=cmd_verify_analysis)

    analysis_inspect = sub.add_parser(
        "analysis-inspect",
        help="List reusable analysis summaries and their review/cohort/conclusion metadata.",
    )
    analysis_inspect.add_argument("--db", required=True)
    analysis_inspect.add_argument("--report-id", type=int)
    analysis_inspect.add_argument("--workbook-id", type=int)
    analysis_inspect.add_argument("--dataset")
    analysis_inspect.set_defaults(func=cmd_inspect_analysis)

    analysis_export = sub.add_parser(
        "analysis-export",
        help="Export one reusable analysis report with cohorts, metrics, comparisons, conclusions, and evidence.",
    )
    analysis_export.add_argument("--db", required=True)
    analysis_export.add_argument("--report-id", required=True, type=int)
    analysis_export.add_argument("--out", help="Output JSON path. Must stay under this service folder.")
    analysis_export.set_defaults(func=cmd_export_analysis)

    knowledge_migrate = sub.add_parser(
        "knowledge-migrate",
        help="Project legacy analysis reports into the canonical Study/Comparison/Evidence schema.",
    )
    knowledge_migrate.add_argument("--db", required=True)
    knowledge_migrate.set_defaults(func=cmd_migrate_knowledge)

    knowledge_inspect = sub.add_parser(
        "knowledge-inspect",
        help="Inspect canonical knowledge counts and integrity constraints.",
    )
    knowledge_inspect.add_argument("--db", required=True)
    knowledge_inspect.set_defaults(func=cmd_inspect_knowledge)

    study_import = sub.add_parser(
        "study-import",
        help="Validate and import a canonical, source-backed, domain-neutral Study manifest.",
    )
    study_import.add_argument("--input", required=True)
    study_import.add_argument("--db", required=True)
    study_import.set_defaults(func=cmd_import_study)

    analysis_quarantine = sub.add_parser(
        "analysis-quarantine",
        help=(
            "Atomically hide one current unverified canonical analysis from "
            "queries while preserving its data and evidence for audit."
        ),
    )
    analysis_quarantine.add_argument("--db", required=True)
    analysis_quarantine.add_argument(
        "--public-analysis-id",
        required=True,
    )
    analysis_quarantine.add_argument("--reason", required=True)
    analysis_quarantine.set_defaults(func=cmd_quarantine_analysis)

    search = sub.add_parser("search", help="Search quick-index or universal-grid DB.")
    search.add_argument("--db", required=True)
    search.add_argument("--q", required=True)
    search.add_argument("--limit", type=int, default=30)
    search.set_defaults(func=cmd_search)

    packet = sub.add_parser("build-packet", help="Build one source-backed AI ReviewCase packet from a DB record.")
    packet.add_argument("--db", required=True)
    packet.add_argument("--file-id", type=int, help="quick-index file_id. Also accepted as universal workbook_id if --workbook-id is omitted.")
    packet.add_argument("--workbook-id", type=int, help="universal-grid workbook_id.")
    packet.add_argument("--out", help="Output packet path. Must stay under this service folder.")
    packet.add_argument("--row-limit", type=int, default=1200)
    packet.add_argument("--cell-limit", type=int, default=20000)
    packet.add_argument("--candidate-limit", type=int, default=300)
    packet.set_defaults(func=cmd_build_packet)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
