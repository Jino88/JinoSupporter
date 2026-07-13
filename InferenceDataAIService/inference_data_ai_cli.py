from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import sqlite3
import subprocess
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator


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
    rows = dict_rows(
        conn,
        """
        SELECT sheet_index, sheet_name, row_number, non_empty_count, row_text, cells_json
        FROM grid_sheet_rows
        WHERE workbook_id=?
        ORDER BY sheet_index, row_number
        LIMIT ?
        """,
        (workbook_id, row_limit),
    )
    for row in rows:
        row["rowId"] = row_ref(row["sheet_name"], row["row_number"])
        try:
            row["cells"] = json.loads(row.pop("cells_json") or "[]")
        except json.JSONDecodeError:
            row["cells"] = []

    cells = dict_rows(
        conn,
        """
        SELECT sheet_index, sheet_name, row_number, col_number, col_label, address,
               value_text, raw_value_text, merge_role, merge_address, anchor_row, anchor_col
        FROM grid_sheet_cells
        WHERE workbook_id=?
        ORDER BY sheet_index, row_number, col_number
        LIMIT ?
        """,
        (workbook_id, cell_limit),
    )
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
    ]
    if len(rows) >= row_limit:
        notes.append("sheetRows were truncated by rowLimit. Increase rowLimit before final generation.")
    if len(cells) >= cell_limit:
        notes.append("sheetCells were truncated by cellLimit.")

    return {
        "schemaVersion": "inference-data-ai-reviewcase-packet-v1",
        "createdAt": now_iso(),
        "sourceDbType": "universal-grid",
        "notes": notes,
        "reviewCaseContract": reviewcase_contract(),
        "workbook": workbook,
        "sheets": sheets,
        "contextRows": context_rows_from_rows(rows),
        "sheetRows": rows,
        "sheetCells": cells,
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
