#!/usr/bin/env python3
"""Derive strictly numeric, source-backed review facts from numeric-capture.sqlite.

This stage deliberately has no prose generator and no quality/release decision.
Its only automatic comparison is a defect-rate observation between an explicit
Test row and an explicit Normal row from the *same numeric table and calendar
date*.  Everything else remains an observed fact or NEEDS_REVIEW.
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import sqlite3
import sys
from collections import Counter, defaultdict
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl.utils.cell import range_boundaries

import inference_data_ai_structure_scan as structure


REVIEW_VERSION = "numeric-review-v1"


def utc_now() -> str:
    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


def atomic_write_text(path: Path, value: str, *, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(value, encoding=encoding)
    os.replace(temporary, path)


def atomic_write_json(path: Path, value: object) -> None:
    atomic_write_text(path, json.dumps(value, ensure_ascii=False, indent=2) + "\n")


def resolve_batch_directory(service_dir: Path, batch_id: str) -> Path:
    return structure.resolve_batch_directory(service_dir, structure.safe_batch_id(batch_id))


def open_review_db(path: Path) -> sqlite3.Connection:
    if not path.is_file():
        raise ValueError(f"Numeric capture database is missing: {path}")
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys=ON")
    connection.execute("PRAGMA journal_mode=WAL")
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS numeric_review_metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS numeric_table_reviews (
            table_id INTEGER PRIMARY KEY REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
            profile_name TEXT NOT NULL,
            extraction_status TEXT NOT NULL,
            reason_code TEXT NOT NULL DEFAULT '',
            extracted_fact_count INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_numeric_table_reviews_status
            ON numeric_table_reviews(profile_name, extraction_status);

        CREATE TABLE IF NOT EXISTS numeric_review_facts (
            fact_id INTEGER PRIMARY KEY,
            workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
            sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
            table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
            row_index INTEGER NOT NULL,
            measurement_date TEXT NOT NULL DEFAULT '',
            metric_type TEXT NOT NULL,
            condition_role TEXT NOT NULL,
            condition_label TEXT NOT NULL,
            input_value_text TEXT NOT NULL,
            input_value REAL,
            total_ng_value_text TEXT NOT NULL,
            total_ng_value REAL,
            reported_ng_rate_text TEXT NOT NULL,
            reported_ng_rate REAL,
            computed_ng_rate REAL,
            fact_status TEXT NOT NULL,
            validation_code TEXT NOT NULL DEFAULT '',
            date_row_index INTEGER,
            date_column_index INTEGER,
            condition_column_index INTEGER,
            input_column_index INTEGER,
            total_ng_column_index INTEGER,
            rate_column_index INTEGER,
            UNIQUE(table_id, row_index, metric_type)
        );
        CREATE INDEX IF NOT EXISTS idx_numeric_review_facts_pair
            ON numeric_review_facts(table_id, measurement_date, condition_role, metric_type);

        CREATE TABLE IF NOT EXISTS measurement_summary_facts (
            fact_id INTEGER PRIMARY KEY,
            workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
            sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
            table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
            row_index INTEGER NOT NULL,
            measurement_date TEXT NOT NULL DEFAULT '',
            condition_label TEXT NOT NULL,
            sample_value_text TEXT NOT NULL,
            sample_value REAL,
            average_value_text TEXT NOT NULL,
            average_value REAL,
            minimum_value_text TEXT NOT NULL,
            minimum_value REAL,
            maximum_value_text TEXT NOT NULL,
            maximum_value REAL,
            fact_status TEXT NOT NULL,
            validation_code TEXT NOT NULL DEFAULT '',
            date_row_index INTEGER,
            date_column_index INTEGER,
            sample_column_index INTEGER,
            average_column_index INTEGER,
            minimum_column_index INTEGER,
            maximum_column_index INTEGER,
            UNIQUE(table_id, row_index)
        );
        CREATE INDEX IF NOT EXISTS idx_measurement_summary_facts_workbook
            ON measurement_summary_facts(workbook_id, measurement_date);

        CREATE TABLE IF NOT EXISTS test_normal_comparisons (
            comparison_id INTEGER PRIMARY KEY,
            workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
            sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
            table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
            measurement_date TEXT NOT NULL,
            metric_type TEXT NOT NULL,
            test_fact_id INTEGER REFERENCES numeric_review_facts(fact_id) ON DELETE SET NULL,
            normal_fact_id INTEGER REFERENCES numeric_review_facts(fact_id) ON DELETE SET NULL,
            test_ng_rate REAL,
            normal_ng_rate REAL,
            absolute_delta REAL,
            relative_ratio REAL,
            comparison_status TEXT NOT NULL,
            validation_code TEXT NOT NULL DEFAULT '',
            UNIQUE(table_id, measurement_date, metric_type, test_fact_id)
        );
        CREATE INDEX IF NOT EXISTS idx_test_normal_comparisons_status
            ON test_normal_comparisons(comparison_status, measurement_date);
        """
    )
    connection.execute(
        "INSERT OR REPLACE INTO numeric_review_metadata(key, value) VALUES ('schemaVersion', ?)",
        ("numeric-review-db-v1",),
    )
    connection.execute(
        "INSERT OR REPLACE INTO numeric_review_metadata(key, value) VALUES ('reviewVersion', ?)", (REVIEW_VERSION,))
    connection.commit()
    return connection


def explicit_role(label: str) -> str | None:
    """Identify only labels that explicitly say Test or Normal.

    No synonym such as Control is accepted.  This prevents a title, filename, or
    vague text from silently becoming the Normal comparison row.
    """
    value = label.casefold().strip()
    if re.search(r"(?<![a-z])normal(?![a-z])", value):
        return "NORMAL"
    if re.search(r"(?<![a-z])test(?![a-z])", value):
        return "TEST"
    return None


def single_column(header_rows: Iterable[sqlite3.Row], token: str) -> int | None:
    columns = {int(row["column_index"]) for row in header_rows if structure.header_token(row["label_text"]) == token}
    return next(iter(columns)) if len(columns) == 1 else None


def defect_columns(header_rows: Iterable[sqlite3.Row]) -> tuple[int, int, int | None, bool] | None:
    """Find a conservative Input/Total NG schema.

    Some source tables label the total only as ``NG``.  That fallback is safe
    only when the same header band has exactly one NG column *and* an explicit
    ``NG rate`` column; otherwise an NG breakdown column could be mistaken for
    the total and the table remains unclassified.
    """
    headers = list(header_rows)
    input_columns = sorted({int(row["column_index"]) for row in headers if structure.header_token(row["label_text"]) == "INPUT"})
    total_ng_columns = sorted({int(row["column_index"]) for row in headers if structure.header_token(row["label_text"]) == "TOTAL_NG"})
    rate_columns = sorted({int(row["column_index"]) for row in headers if structure.header_token(row["label_text"]) == "NG_RATE"})
    if len(input_columns) == 1 and total_ng_columns:
        input_column = input_columns[0]
        after_input = [column for column in total_ng_columns if column > input_column]
        total_ng_column = after_input[0] if after_input else total_ng_columns[0]
        following_rates = [column for column in rate_columns if column > total_ng_column]
        rate_column = following_rates[0] if following_rates else (rate_columns[0] if len(rate_columns) == 1 else None)
        repeated_section = len(total_ng_columns) > 1 or len(rate_columns) > 1
        return input_column, total_ng_column, rate_column, repeated_section
    input_column = input_columns[0] if len(input_columns) == 1 else None
    rate_column = rate_columns[0] if len(rate_columns) == 1 else None
    if input_column is None or rate_column is None:
        return None
    ng_columns = {int(row["column_index"]) for row in headers if structure.header_token(row["label_text"]) == "NG_CUE"}
    if len(ng_columns) != 1:
        return None
    return input_column, next(iter(ng_columns)), rate_column, False


def measurement_columns(header_rows: Iterable[sqlite3.Row]) -> tuple[int | None, int, int, int] | None:
    """Find a numeric Sample/Average/Min/Max observation matrix.

    ``Sample`` is intentionally optional because several reports call it
    ``No sample`` or omit it entirely.  Average/Min/Max must be uniquely
    mapped; otherwise the table is not summarized automatically.
    """
    headers = list(header_rows)
    sample_column = single_column(headers, "SAMPLE")
    average_column = single_column(headers, "AVERAGE")
    minimum_column = single_column(headers, "MIN")
    maximum_column = single_column(headers, "MAX")
    if average_column is None or minimum_column is None or maximum_column is None:
        return None
    return sample_column, average_column, minimum_column, maximum_column


def row_dates(connection: sqlite3.Connection, sheet_id: int, table: sqlite3.Row) -> dict[int, tuple[str, int, int]]:
    """Resolve direct date cells and dates carried down through merged data cells."""
    values = {
        int(row["row_index"]): (str(row["date_value"]), int(row["row_index"]), int(row["column_index"]))
        for row in connection.execute(
            """
            SELECT row_index, column_index, date_value
            FROM date_cells
            WHERE sheet_id=? AND row_index BETWEEN ? AND ?
            """,
            (sheet_id, table["start_row"], table["end_row"]),
        )
    }
    all_dates = {
        (int(row["row_index"]), int(row["column_index"])): str(row["date_value"])
        for row in connection.execute("SELECT row_index, column_index, date_value FROM date_cells WHERE sheet_id=?", (sheet_id,))
    }
    for merge in connection.execute("SELECT range_ref FROM captured_merge_ranges WHERE sheet_id=?", (sheet_id,)):
        try:
            minimum_column, minimum_row, maximum_column, maximum_row = range_boundaries(str(merge["range_ref"]))
        except ValueError:
            continue
        source_date = all_dates.get((minimum_row, minimum_column))
        if not source_date:
            continue
        for row in range(max(minimum_row, int(table["start_row"])), min(maximum_row, int(table["end_row"])) + 1):
            values.setdefault(row, (source_date, minimum_row, minimum_column))
    return values


def value_for_column(values: dict[tuple[int, int], sqlite3.Row], row: int, column: int | None) -> tuple[str, float | None]:
    if column is None:
        return "", None
    item = values.get((row, column))
    if item is None:
        return "", None
    return str(item["value_text"]), float(item["numeric_value"])


def extract_defect_rate_table(connection: sqlite3.Connection, table: sqlite3.Row) -> list[int]:
    sheet_id = int(table["sheet_id"])
    workbook_id = int(table["workbook_id"])
    headers = list(
        connection.execute(
            "SELECT * FROM numeric_table_labels WHERE table_id=? AND label_role='HEADER' ORDER BY row_index, column_index",
            (table["table_id"],),
        )
    )
    columns = defect_columns(headers)
    if columns is None:
        connection.execute(
            """
            INSERT INTO numeric_table_reviews(table_id, profile_name, extraction_status, reason_code, extracted_fact_count, updated_at)
            VALUES (?, 'DEFECT_RATE', 'NEEDS_REVIEW', 'REQUIRED_HEADER_AMBIGUOUS_OR_MISSING', 0, ?)
            """,
            (table["table_id"], utc_now()),
        )
        return []
    input_column, total_ng_column, rate_column, repeated_section = columns
    values = {
        (int(row["row_index"]), int(row["column_index"])): row
        for row in connection.execute(
            """
            SELECT row_index, column_index, value_text, numeric_value
            FROM numeric_cells
            WHERE sheet_id=? AND row_index BETWEEN ? AND ?
            """,
            (sheet_id, table["start_row"], table["end_row"]),
        )
    }
    labels_by_row: dict[int, list[sqlite3.Row]] = defaultdict(list)
    for label in connection.execute(
        "SELECT * FROM numeric_table_labels WHERE table_id=? AND label_role='ROW_LABEL' ORDER BY row_index, column_index",
        (table["table_id"],),
    ):
        labels_by_row[int(label["row_index"])].append(label)
    date_by_row = row_dates(connection, sheet_id, table)
    facts: list[int] = []
    for row in range(int(table["start_row"]), int(table["end_row"]) + 1):
        role_label = next(
            ((explicit_role(str(label["label_text"])), label) for label in labels_by_row.get(row, []) if explicit_role(str(label["label_text"]))),
            None,
        )
        if role_label is None:
            continue
        role, label = role_label
        assert role is not None
        input_text, input_value = value_for_column(values, row, input_column)
        total_text, total_value = value_for_column(values, row, total_ng_column)
        rate_text, reported_rate = value_for_column(values, row, rate_column)
        date_info = date_by_row.get(row)
        validation: list[str] = []
        if repeated_section:
            validation.append("REPEATED_DEFECT_HEADER_SECTION")
        if date_info is None:
            validation.append("DATE_MISSING")
        rate_usable = False
        if input_value is None or total_value is None:
            validation.append("INPUT_OR_TOTAL_NG_MISSING")
        elif input_value <= 0:
            validation.append("INPUT_NOT_POSITIVE")
        elif total_value < 0:
            validation.append("TOTAL_NG_NEGATIVE")
        else:
            rate_usable = True
        computed_rate = total_value / input_value if rate_usable else None
        status = "OBSERVED" if not validation else "NEEDS_REVIEW"
        cursor = connection.execute(
            """
            INSERT INTO numeric_review_facts(
                workbook_id, sheet_id, table_id, row_index, measurement_date, metric_type,
                condition_role, condition_label, input_value_text, input_value,
                total_ng_value_text, total_ng_value, reported_ng_rate_text, reported_ng_rate,
                computed_ng_rate, fact_status, validation_code, date_row_index, date_column_index,
                condition_column_index, input_column_index, total_ng_column_index, rate_column_index
            ) VALUES (?, ?, ?, ?, ?, 'DEFECT_RATE', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                workbook_id,
                sheet_id,
                table["table_id"],
                row,
                "" if date_info is None else date_info[0],
                role,
                str(label["label_text"]),
                input_text,
                input_value,
                total_text,
                total_value,
                rate_text,
                reported_rate,
                computed_rate,
                status,
                ";".join(validation),
                None if date_info is None else date_info[1],
                None if date_info is None else date_info[2],
                int(label["column_index"]),
                input_column,
                total_ng_column,
                rate_column,
            ),
        )
        facts.append(int(cursor.lastrowid))
    status = "EXTRACTED" if facts else "NEEDS_REVIEW"
    reason = "" if facts else "EXPLICIT_TEST_OR_NORMAL_ROW_NOT_FOUND"
    connection.execute(
        """
        INSERT INTO numeric_table_reviews(table_id, profile_name, extraction_status, reason_code, extracted_fact_count, updated_at)
        VALUES (?, 'DEFECT_RATE', ?, ?, ?, ?)
        """,
        (table["table_id"], status, reason, len(facts), utc_now()),
    )
    return facts


def extract_measurement_summary_table(connection: sqlite3.Connection, table: sqlite3.Row) -> list[int]:
    """Capture Sample/Average/Max/Min rows as numeric observations only."""
    sheet_id = int(table["sheet_id"])
    workbook_id = int(table["workbook_id"])
    headers = list(
        connection.execute(
            "SELECT * FROM numeric_table_labels WHERE table_id=? AND label_role='HEADER' ORDER BY row_index, column_index",
            (table["table_id"],),
        )
    )
    columns = measurement_columns(headers)
    if columns is None:
        connection.execute(
            """
            INSERT INTO numeric_table_reviews(table_id, profile_name, extraction_status, reason_code, extracted_fact_count, updated_at)
            VALUES (?, 'MEASUREMENT_SUMMARY', 'NEEDS_REVIEW', 'REQUIRED_HEADER_AMBIGUOUS_OR_MISSING', 0, ?)
            """,
            (table["table_id"], utc_now()),
        )
        return []
    sample_column, average_column, minimum_column, maximum_column = columns
    values = {
        (int(row["row_index"]), int(row["column_index"])): row
        for row in connection.execute(
            """
            SELECT row_index, column_index, value_text, numeric_value
            FROM numeric_cells
            WHERE sheet_id=? AND row_index BETWEEN ? AND ?
            """,
            (sheet_id, table["start_row"], table["end_row"]),
        )
    }
    labels_by_row: dict[int, list[str]] = defaultdict(list)
    for label in connection.execute(
        "SELECT row_index, label_text FROM numeric_table_labels WHERE table_id=? AND label_role='ROW_LABEL' ORDER BY row_index, column_index",
        (table["table_id"],),
    ):
        labels_by_row[int(label["row_index"])].append(str(label["label_text"]))
    date_by_row = row_dates(connection, sheet_id, table)
    facts: list[int] = []
    for row in range(int(table["start_row"]), int(table["end_row"]) + 1):
        average_text, average_value = value_for_column(values, row, average_column)
        minimum_text, minimum_value = value_for_column(values, row, minimum_column)
        maximum_text, maximum_value = value_for_column(values, row, maximum_column)
        if average_value is None and minimum_value is None and maximum_value is None:
            continue
        sample_text, sample_value = value_for_column(values, row, sample_column)
        date_info = date_by_row.get(row)
        validation: list[str] = []
        if average_value is None or minimum_value is None or maximum_value is None:
            validation.append("AVERAGE_MIN_OR_MAX_MISSING")
        elif minimum_value > maximum_value:
            validation.append("MINIMUM_GREATER_THAN_MAXIMUM")
        elif average_value < minimum_value or average_value > maximum_value:
            validation.append("AVERAGE_OUTSIDE_MIN_MAX")
        if sample_column is not None and sample_value is None:
            validation.append("SAMPLE_MISSING")
        status = "OBSERVED" if not validation else "NEEDS_REVIEW"
        cursor = connection.execute(
            """
            INSERT INTO measurement_summary_facts(
                workbook_id, sheet_id, table_id, row_index, measurement_date, condition_label,
                sample_value_text, sample_value, average_value_text, average_value,
                minimum_value_text, minimum_value, maximum_value_text, maximum_value,
                fact_status, validation_code, date_row_index, date_column_index,
                sample_column_index, average_column_index, minimum_column_index, maximum_column_index
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                workbook_id,
                sheet_id,
                table["table_id"],
                row,
                "" if date_info is None else date_info[0],
                " | ".join(labels_by_row.get(row, [])),
                sample_text,
                sample_value,
                average_text,
                average_value,
                minimum_text,
                minimum_value,
                maximum_text,
                maximum_value,
                status,
                ";".join(validation),
                None if date_info is None else date_info[1],
                None if date_info is None else date_info[2],
                sample_column,
                average_column,
                minimum_column,
                maximum_column,
            ),
        )
        facts.append(int(cursor.lastrowid))
    status = "EXTRACTED" if facts else "NEEDS_REVIEW"
    reason = "" if facts else "NUMERIC_SUMMARY_ROW_NOT_FOUND"
    connection.execute(
        """
        INSERT INTO numeric_table_reviews(table_id, profile_name, extraction_status, reason_code, extracted_fact_count, updated_at)
        VALUES (?, 'MEASUREMENT_SUMMARY', ?, ?, ?, ?)
        """,
        (table["table_id"], status, reason, len(facts), utc_now()),
    )
    return facts


def create_comparisons(connection: sqlite3.Connection) -> None:
    facts = list(
        connection.execute(
            """
            SELECT * FROM numeric_review_facts
            WHERE metric_type='DEFECT_RATE' AND measurement_date <> ''
            ORDER BY table_id, measurement_date, condition_role, fact_id
            """
        )
    )
    groups: dict[tuple[int, str], list[sqlite3.Row]] = defaultdict(list)
    for fact in facts:
        groups[(int(fact["table_id"]), str(fact["measurement_date"]))].append(fact)
    for (_, measurement_date), rows in groups.items():
        tests = [item for item in rows if item["condition_role"] == "TEST"]
        normals = [item for item in rows if item["condition_role"] == "NORMAL"]
        for test in tests:
            status = "VALID"
            validation = "SAME_TABLE_SAME_DATE_EXPLICIT_TEST_NORMAL"
            normal: sqlite3.Row | None = None
            if test["fact_status"] != "OBSERVED":
                status, validation = "NO_COMPARISON_NEEDS_REVIEW", "TEST_FACT_INVALID"
            elif not normals:
                status, validation = "NO_SAME_DAY_NORMAL", "NORMAL_NOT_FOUND_FOR_SAME_DATE"
            elif len(normals) > 1:
                status, validation = "NORMAL_AMBIGUOUS", "MULTIPLE_NORMAL_ROWS_FOR_SAME_DATE"
            elif len(tests) > 1:
                status, validation = "TEST_AMBIGUOUS", "MULTIPLE_TEST_ROWS_FOR_SAME_DATE"
            elif normals[0]["fact_status"] != "OBSERVED":
                status, validation = "NO_COMPARISON_NEEDS_REVIEW", "NORMAL_FACT_INVALID"
            else:
                normal = normals[0]
            test_rate = test["computed_ng_rate"]
            normal_rate = None if normal is None else normal["computed_ng_rate"]
            delta = None if normal_rate is None else float(test_rate) - float(normal_rate)
            ratio = None if normal_rate in (None, 0) else float(test_rate) / float(normal_rate)
            connection.execute(
                """
                INSERT INTO test_normal_comparisons(
                    workbook_id, sheet_id, table_id, measurement_date, metric_type,
                    test_fact_id, normal_fact_id, test_ng_rate, normal_ng_rate,
                    absolute_delta, relative_ratio, comparison_status, validation_code
                ) VALUES (?, ?, ?, ?, 'DEFECT_RATE', ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    test["workbook_id"],
                    test["sheet_id"],
                    test["table_id"],
                    measurement_date,
                    test["fact_id"],
                    None if normal is None else normal["fact_id"],
                    test_rate,
                    normal_rate,
                    delta,
                    ratio,
                    status,
                    validation,
                ),
            )


def rebuild_reviews(connection: sqlite3.Connection) -> None:
    connection.execute("DELETE FROM test_normal_comparisons")
    connection.execute("DELETE FROM numeric_review_facts")
    connection.execute("DELETE FROM measurement_summary_facts")
    connection.execute("DELETE FROM numeric_table_reviews")
    tables = list(
        connection.execute(
            """
            SELECT table_candidate.*, sheet.workbook_id
            FROM numeric_table_candidates AS table_candidate
            JOIN captured_sheets AS sheet ON sheet.sheet_id=table_candidate.sheet_id
            JOIN capture_workbooks AS workbook ON workbook.workbook_id=sheet.workbook_id
            WHERE workbook.capture_status='CAPTURED'
            ORDER BY table_candidate.table_id
            """
        )
    )
    for table in tables:
        headers = list(
            connection.execute(
                "SELECT * FROM numeric_table_labels WHERE table_id=? AND label_role='HEADER' ORDER BY row_index, column_index",
                (table["table_id"],),
            )
        )
        if defect_columns(headers) is not None:
            extract_defect_rate_table(connection, table)
        elif measurement_columns(headers) is not None:
            extract_measurement_summary_table(connection, table)
        else:
            connection.execute(
                """
                INSERT INTO numeric_table_reviews(table_id, profile_name, extraction_status, reason_code, extracted_fact_count, updated_at)
                VALUES (?, ?, 'NOT_IMPLEMENTED', 'PROFILE_PENDING', 0, ?)
                """,
                (table["table_id"], table["candidate_type"], utc_now()),
            )
    create_comparisons(connection)
    connection.commit()


def write_outputs(batch_dir: Path, connection: sqlite3.Connection) -> dict[str, object]:
    table_statuses = Counter(
        f"{row['profile_name']}:{row['extraction_status']}"
        for row in connection.execute("SELECT profile_name, extraction_status FROM numeric_table_reviews")
    )
    comparison_statuses = Counter(
        str(row[0]) for row in connection.execute("SELECT comparison_status FROM test_normal_comparisons")
    )
    defect_fact_count = int(connection.execute("SELECT COUNT(*) FROM numeric_review_facts").fetchone()[0])
    measurement_fact_count = int(connection.execute("SELECT COUNT(*) FROM measurement_summary_facts").fetchone()[0])
    comparison_count = int(connection.execute("SELECT COUNT(*) FROM test_normal_comparisons").fetchone()[0])
    summary = {
        "schemaVersion": "numeric-review-summary-v1",
        "reviewVersion": REVIEW_VERSION,
        "generatedAt": utc_now(),
        "usesCom": False,
        "numericFactCount": defect_fact_count + measurement_fact_count,
        "defectRateFactCount": defect_fact_count,
        "measurementSummaryFactCount": measurement_fact_count,
        "testNormalComparisonCount": comparison_count,
        "tableProfileStatusCounts": dict(sorted(table_statuses.items())),
        "comparisonStatusCounts": dict(sorted(comparison_statuses.items())),
        "database": "numeric-capture.sqlite",
        "limitations": [
            "Only explicit Test and Normal labels inside a defect-rate numeric table are compared.",
            "A comparison requires the same workbook table and normalized calendar date.",
            "No release, quality, improvement, causality, or other narrative decision is generated.",
        ],
    }
    atomic_write_json(batch_dir / "numeric-review-summary.json", summary)
    rows = [
        {
            "relativePath": str(row["relative_path"]),
            "measurementDate": str(row["measurement_date"]),
            "testRate": "" if row["test_ng_rate"] is None else f"{float(row['test_ng_rate']):.10g}",
            "normalRate": "" if row["normal_ng_rate"] is None else f"{float(row['normal_ng_rate']):.10g}",
            "difference": "" if row["absolute_delta"] is None else f"{float(row['absolute_delta']):.10g}",
            "status": str(row["comparison_status"]),
        }
        for row in connection.execute(
            """
            SELECT workbook.relative_path, comparison.measurement_date, comparison.test_ng_rate,
                   comparison.normal_ng_rate, comparison.absolute_delta, comparison.comparison_status
            FROM test_normal_comparisons AS comparison
            JOIN capture_workbooks AS workbook ON workbook.workbook_id=comparison.workbook_id
            ORDER BY workbook.relative_path, comparison.measurement_date
            """
        )
    ]
    columns = list(rows[0]) if rows else ["relativePath", "measurementDate", "status"]
    temporary = batch_dir / "numeric-review.csv.tmp"
    with temporary.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)
    os.replace(temporary, batch_dir / "numeric-review.csv")
    count_rows = "".join(f"<li>{html.escape(name)}: {count}</li>" for name, count in sorted(comparison_statuses.items()))
    report_rows = "".join(
        "<tr>" + "".join(f"<td>{html.escape(row[column])}</td>" for column in columns) + "</tr>" for row in rows
    )
    document = f"""<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>숫자 검토 추출</title>
<style>body{{font-family:Segoe UI,sans-serif;margin:24px;color:#172b4d}}table{{border-collapse:collapse;width:100%;font-size:12px}}th,td{{border:1px solid #d7dde7;padding:6px;text-align:left;vertical-align:top}}th{{background:#edf2f7}}</style></head>
<body><h1>숫자 검토 추출</h1><p>같은 날짜의 명시적 Test–Normal 불량률만 비교합니다. 이 단계는 품질 판정이나 서술형 결론을 만들지 않습니다.</p>
<h2>비교 상태</h2><ul>{count_rows}</ul><p><a href='numeric-review.csv'>CSV</a> · <a href='numeric-review-summary.json'>JSON</a></p>
<table><thead><tr>{''.join(f'<th>{html.escape(column)}</th>' for column in columns)}</tr></thead><tbody>{report_rows}</tbody></table></body></html>"""
    atomic_write_text(batch_dir / "numeric-review.html", document)
    return summary


def run(args: argparse.Namespace) -> int:
    service_dir = Path(args.service_dir).resolve()
    if not service_dir.is_dir():
        raise ValueError(f"Service directory does not exist: {service_dir}")
    batch_dir = resolve_batch_directory(service_dir, args.structure_batch)
    connection = open_review_db(batch_dir / "numeric-capture.sqlite")
    try:
        rebuild_reviews(connection)
        summary = write_outputs(batch_dir, connection)
    finally:
        connection.close()
    print(json.dumps({"status": "ok", "batchDirectory": str(batch_dir), "summary": summary}, ensure_ascii=False))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build source-backed numeric facts and same-date Test–Normal defect-rate comparisons.")
    parser.add_argument("--service-dir", required=True)
    parser.add_argument("--structure-batch", required=True)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run(args)
    except (ValueError, RuntimeError) as exc:
        print(f"numeric-review error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
