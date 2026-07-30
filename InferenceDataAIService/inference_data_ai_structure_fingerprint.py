"""Deterministic structural fingerprints for Capture v2 workbooks.

The fingerprint intentionally excludes numeric cell values and source identity
from its digest.  It is suitable for retrieving extraction recipes, but it is
not by itself permission to persist extracted data; recipe validation remains
the final gate.
"""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import unicodedata
from collections import Counter
from typing import Any, Iterable


FINGERPRINT_SCHEMA_VERSION = "excel-structure-fingerprint-v2"
FINGERPRINT_ENGINE_VERSION = "deterministic-structure-fingerprint-v2.0"
HEADER_SCAN_ROW_LIMIT = 12
MAX_ANCHORS_PER_SHEET = 160

_TOKEN_PATTERN = re.compile(r"[0-9A-Z가-힣]+(?:[._/-][0-9A-Z가-힣]+)*")
_CELL_REFERENCE_PATTERN = re.compile(
    r"(?<![A-Z0-9_])(?:'[^']+'!)?\$?[A-Z]{1,3}\$?[0-9]+",
    re.IGNORECASE,
)
_RANGE_REFERENCE_PATTERN = re.compile(
    r"(?<![A-Z0-9_])(?:'[^']+'!)?\$?[A-Z]{1,3}\$?[0-9]+:"
    r"\$?[A-Z]{1,3}\$?[0-9]+",
    re.IGNORECASE,
)
_NUMBER_PATTERN = re.compile(r"(?<![A-Z_])[-+]?\d+(?:\.\d+)?(?:E[-+]?\d+)?", re.IGNORECASE)
_VARIABLE_TEXT_PATTERNS = (
    re.compile(r"^\d+(?:[.,]\d+)?$"),
    re.compile(r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}$"),
    re.compile(r"^\d{1,2}:\d{2}(?::\d{2})?$"),
)


class StructureFingerprintError(ValueError):
    """Raised when a Capture v2 payload cannot produce a fingerprint."""


def _canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _sha256(value: Any) -> str:
    return hashlib.sha256(_canonical_bytes(value)).hexdigest()


def normalize_text(value: Any) -> str:
    """Return a locale-neutral label form used by recipes and fingerprints."""

    if value is None:
        return ""
    if isinstance(value, dict) and "value" in value:
        value = value["value"]
    text = unicodedata.normalize("NFKC", str(value))
    return " ".join(text.strip().upper().split())


def text_tokens(values: Iterable[Any]) -> list[str]:
    result: set[str] = set()
    for value in values:
        normalized = normalize_text(value)
        for match in _TOKEN_PATTERN.finditer(normalized):
            token = match.group(0).strip("._/-")
            if len(token) < 2 or token.isdigit():
                continue
            result.add(token)
    return sorted(result)


def _is_variable_text(text: str) -> bool:
    return not text or any(pattern.fullmatch(text) for pattern in _VARIABLE_TEXT_PATTERNS)


def _bucket(value: int) -> str:
    if value <= 0:
        return "0"
    bounds = (1, 2, 4, 8, 16, 32, 64, 96, 128, 192, 256, 384, 512, 768, 1024)
    lower = 1
    for upper in bounds:
        if value <= upper:
            return str(upper) if lower == upper else f"{lower}-{upper}"
        lower = upper + 1
    return "1025+"


def _bounds(sheet: dict[str, Any]) -> dict[str, int]:
    raw = sheet.get("contentBounds") or sheet.get("usedBounds") or {}
    min_row = int(raw.get("minRow") or 1)
    min_column = int(raw.get("minColumn") or 1)
    row_count = int(raw.get("rowCount") or 0)
    column_count = int(raw.get("columnCount") or 0)
    max_row = int(raw.get("maxRow") or (min_row + max(row_count - 1, 0)))
    max_column = int(
        raw.get("maxColumn") or (min_column + max(column_count - 1, 0))
    )
    return {
        "minRow": min_row,
        "minColumn": min_column,
        "maxRow": max_row,
        "maxColumn": max_column,
        "rowCount": row_count,
        "columnCount": column_count,
    }


def _relative_band(index: int, minimum: int, maximum: int) -> str:
    span = max(maximum - minimum, 1)
    ratio = (index - minimum) / span
    if ratio <= 0.20:
        return "START"
    if ratio >= 0.80:
        return "END"
    return "MIDDLE"


def _cell_value(cell: dict[str, Any]) -> Any:
    value = cell.get("displayValue")
    if value is None:
        value = cell.get("rawValue")
    return value


def _anchor_sketches(
    sheet: dict[str, Any],
    bounds: dict[str, int],
) -> list[dict[str, Any]]:
    first_rows_max = bounds["minRow"] + HEADER_SCAN_ROW_LIMIT - 1
    by_key: dict[tuple[str, str, str, bool], int] = {}
    for cell in sheet.get("cells") or []:
        value = _cell_value(cell)
        if not isinstance(value, str):
            continue
        text = normalize_text(value)
        if _is_variable_text(text):
            continue
        row = int(cell.get("row") or 0)
        column = int(cell.get("column") or 0)
        style = cell.get("style") or {}
        font = style.get("font") or {}
        structural = bool(
            row <= first_rows_max
            or cell.get("mergeRole") == "anchor"
            or font.get("bold")
        )
        if not structural:
            continue
        key = (
            text,
            _relative_band(row, bounds["minRow"], bounds["maxRow"]),
            _relative_band(column, bounds["minColumn"], bounds["maxColumn"]),
            cell.get("mergeRole") == "anchor",
        )
        by_key[key] = by_key.get(key, 0) + 1
    sketches = [
        {
            "text": key[0],
            "rowBand": key[1],
            "columnBand": key[2],
            "merged": key[3],
            "occurrences": count,
        }
        for key, count in sorted(by_key.items())
    ]
    return sketches[:MAX_ANCHORS_PER_SHEET]


def _merge_geometry(
    sheet: dict[str, Any],
    bounds: dict[str, int],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for merged in sheet.get("mergedRanges") or []:
        min_row = int(merged.get("minRow") or 0)
        min_column = int(merged.get("minColumn") or 0)
        max_row = int(merged.get("maxRow") or min_row)
        max_column = int(merged.get("maxColumn") or min_column)
        result.append(
            {
                "rowBand": _relative_band(
                    min_row,
                    bounds["minRow"],
                    bounds["maxRow"],
                ),
                "columnBand": _relative_band(
                    min_column,
                    bounds["minColumn"],
                    bounds["maxColumn"],
                ),
                "height": max_row - min_row + 1,
                "width": max_column - min_column + 1,
            }
        )
    return sorted(
        result,
        key=lambda item: (
            item["rowBand"],
            item["columnBand"],
            item["height"],
            item["width"],
        ),
    )


def _formula_pattern_hashes(sheet: dict[str, Any]) -> list[str]:
    patterns: set[str] = set()
    for cell in sheet.get("cells") or []:
        formula = cell.get("formula")
        if not formula:
            continue
        normalized = normalize_text(formula)
        normalized = _RANGE_REFERENCE_PATTERN.sub("RANGE", normalized)
        normalized = _CELL_REFERENCE_PATTERN.sub("CELL", normalized)
        normalized = _NUMBER_PATTERN.sub("N", normalized)
        patterns.add(hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:16])
    return sorted(patterns)


def _format_role(number_format: Any) -> str:
    text = normalize_text(number_format)
    if not text or text == "GENERAL":
        return "GENERAL"
    if "%" in text:
        return "PERCENT"
    if any(token in text for token in ("YY", "MM", "DD")):
        return "DATE"
    if any(token in text for token in ("0", "#", "?")):
        return "NUMBER"
    return "OTHER"


def _column_role_sketches(
    sheet: dict[str, Any],
    bounds: dict[str, int],
) -> list[str]:
    counters: dict[int, Counter[str]] = {}
    formats: dict[int, Counter[str]] = {}
    numeric_or_formula_rows = [
        int(cell.get("row") or 0)
        for cell in sheet.get("cells") or []
        if cell.get("formula")
        or (
            isinstance(_cell_value(cell), (int, float))
            and not isinstance(_cell_value(cell), bool)
        )
    ]
    data_start_row = (
        min(numeric_or_formula_rows)
        if numeric_or_formula_rows
        else bounds["minRow"] + HEADER_SCAN_ROW_LIMIT
    )
    for cell in sheet.get("cells") or []:
        row = int(cell.get("row") or 0)
        column = int(cell.get("column") or 0)
        if not column:
            continue
        if row < data_start_row:
            continue
        value = _cell_value(cell)
        if cell.get("formula"):
            kind = "FORMULA"
        elif isinstance(value, bool):
            kind = "BOOLEAN"
        elif isinstance(value, (int, float)):
            kind = "NUMBER"
        elif value is None or value == "":
            continue
        else:
            kind = "TEXT"
        counters.setdefault(column, Counter())[kind] += 1
        formats.setdefault(column, Counter())[_format_role(cell.get("numberFormat"))] += 1

    result: list[str] = []
    for column in range(bounds["minColumn"], bounds["maxColumn"] + 1):
        relative = column - bounds["minColumn"] + 1
        kind = counters.get(column, Counter()).most_common(1)
        format_role = formats.get(column, Counter()).most_common(1)
        result.append(
            f"C{relative}:"
            f"{kind[0][0] if kind else 'EMPTY'}:"
            f"{format_role[0][0] if format_role else 'GENERAL'}"
        )
    return result


def _estimate_header_depth(
    sheet: dict[str, Any],
    bounds: dict[str, int],
) -> int:
    numeric_rows: list[int] = []
    for cell in sheet.get("cells") or []:
        value = _cell_value(cell)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        numeric_rows.append(int(cell.get("row") or bounds["minRow"]))
    if not numeric_rows:
        return min(bounds["rowCount"], HEADER_SCAN_ROW_LIMIT)
    return max(0, min(min(numeric_rows) - bounds["minRow"], HEADER_SCAN_ROW_LIMIT))


def _sheet_fingerprint(sheet: dict[str, Any], fallback_index: int) -> dict[str, Any]:
    bounds = _bounds(sheet)
    anchors = _anchor_sketches(sheet, bounds)
    formula_hashes = list(
        sheet.get("_precomputedFormulaPatternHashes")
        or _formula_pattern_hashes(sheet)
    )
    column_roles = list(
        sheet.get("_precomputedColumnRoleSketch")
        or _column_role_sketches(sheet, bounds)
    )
    number_format_roles = list(
        sheet.get("_precomputedNumberFormatRoles")
        or sorted(
            {
                _format_role(cell.get("numberFormat"))
                for cell in sheet.get("cells") or []
                if _cell_value(cell) is not None or cell.get("formula")
            }
        )
    )
    header_depth = int(
        sheet.get("_precomputedHeaderDepth")
        if sheet.get("_precomputedHeaderDepth") is not None
        else _estimate_header_depth(sheet, bounds)
    )
    return {
        "sheetKey": f"S{int(sheet.get('sheetIndex') or fallback_index)}",
        "titleTokens": text_tokens([sheet.get("title") or ""]),
        "sheetState": str(sheet.get("sheetState") or "visible").lower(),
        "tabular": bool(sheet.get("hasTabularEvidence")),
        "usedRangeBucket": {
            "rows": _bucket(bounds["rowCount"]),
            "columns": _bucket(bounds["columnCount"]),
        },
        "rowCount": bounds["rowCount"],
        "columnCount": bounds["columnCount"],
        "mergedGeometry": _merge_geometry(sheet, bounds),
        "anchorSketches": anchors,
        "headerRoleSketch": column_roles,
        "tableRegionSketches": [
            {
                "headerDepth": header_depth,
                "columnCount": bounds["columnCount"],
                "rowCountBucket": _bucket(bounds["rowCount"]),
                "repeatMode": "ROWS" if bool(sheet.get("hasTabularEvidence")) else "NONE",
            }
        ],
        "formulaPatternHashes": formula_hashes,
        "numberFormatRoles": number_format_roles,
    }


def build_structure_fingerprint(payload: dict[str, Any]) -> dict[str, Any]:
    """Build a deterministic structure-only fingerprint from Capture v2 JSON."""

    workbook = payload.get("workbook")
    if not isinstance(workbook, dict):
        raise StructureFingerprintError("Capture payload requires workbook.")
    sheets = workbook.get("sheets")
    if not isinstance(sheets, list):
        raise StructureFingerprintError("Capture workbook requires sheets.")
    source = payload.get("source") or {}
    source_sha = str(source.get("contentSha256") or "")
    sheet_fingerprints = [
        _sheet_fingerprint(sheet, index)
        for index, sheet in enumerate(sheets, start=1)
    ]
    structural_core = {
        "engineVersion": FINGERPRINT_ENGINE_VERSION,
        "workbook": {
            "status": str(workbook.get("status") or ""),
            "sheetCount": int(workbook.get("sheetCount") or len(sheets)),
            "visibleSheetCount": sum(
                sheet["sheetState"] == "visible" for sheet in sheet_fingerprints
            ),
            "tabularSheetCount": int(
                workbook.get("tabularSheetCount")
                or sum(sheet["tabular"] for sheet in sheet_fingerprints)
            ),
        },
        "sheets": sheet_fingerprints,
    }
    return {
        "schemaVersion": FINGERPRINT_SCHEMA_VERSION,
        "sourceSha256": source_sha,
        **structural_core,
        "fingerprintSha256": _sha256(structural_core),
    }


def _json_value(value: Any) -> Any:
    if value is None:
        return None
    try:
        return json.loads(str(value))
    except (TypeError, json.JSONDecodeError):
        return None


def _database_anchor_cells(
    connection: sqlite3.Connection,
    sheet_id: int,
    header_max_row: int,
) -> list[dict[str, Any]]:
    base_query = """
        SELECT row_index, column_index, coordinate, raw_value_json,
               display_value_json, number_format, merge_range, merge_role
        FROM capture_v2_cells
        WHERE sheet_id=?
          AND (
            row_index<=?
            OR merge_role='anchor'
            OR json_extract(style_json, '$.font.bold')=1
          )
        ORDER BY row_index, column_index
    """
    try:
        rows = connection.execute(
            base_query,
            (sheet_id, header_max_row),
        ).fetchall()
    except sqlite3.OperationalError:
        rows = connection.execute(
            """
            SELECT row_index, column_index, coordinate, raw_value_json,
                   display_value_json, number_format, merge_range, merge_role
            FROM capture_v2_cells
            WHERE sheet_id=? AND (row_index<=? OR merge_role='anchor')
            ORDER BY row_index, column_index
            """,
            (sheet_id, header_max_row),
        ).fetchall()
    return [
        {
            "row": int(row[0]),
            "column": int(row[1]),
            "coordinate": str(row[2]),
            "rawValue": _json_value(row[3]),
            "displayValue": _json_value(row[4]),
            "numberFormat": str(row[5] or "General"),
            "mergeRange": row[6],
            "mergeRole": str(row[7] or "none"),
            "style": {
                "font": {
                    "bold": (
                        int(row[0]) > header_max_row
                        and str(row[7] or "none") != "anchor"
                    )
                }
            },
        }
        for row in rows
    ]


def _database_formula_patterns(
    connection: sqlite3.Connection,
    sheet_id: int,
) -> list[str]:
    cells = [
        {"formula": str(row[0])}
        for row in connection.execute(
            """
            SELECT DISTINCT formula_text
            FROM capture_v2_cells
            WHERE sheet_id=? AND formula_text IS NOT NULL
            ORDER BY formula_text
            """,
            (sheet_id,),
        )
    ]
    return _formula_pattern_hashes({"cells": cells})


def _database_column_roles(
    connection: sqlite3.Connection,
    sheet_id: int,
    bounds: dict[str, int],
) -> tuple[list[str], int, list[str]]:
    first_data_row = connection.execute(
        """
        SELECT MIN(row_index)
        FROM capture_v2_cells
        WHERE sheet_id=?
          AND (
            formula_text IS NOT NULL
            OR json_type(
                 CASE
                   WHEN json_type(display_value_json)<>'null'
                     THEN display_value_json
                   ELSE raw_value_json
                 END
               )
               IN ('integer','real')
          )
        """,
        (sheet_id,),
    ).fetchone()[0]
    data_start_row = (
        int(first_data_row)
        if first_data_row is not None
        else bounds["minRow"] + HEADER_SCAN_ROW_LIMIT
    )
    grouped = connection.execute(
        """
        SELECT
            column_index,
            CASE
              WHEN formula_text IS NOT NULL THEN 'FORMULA'
              WHEN json_type(
                     CASE
                       WHEN json_type(display_value_json)<>'null'
                         THEN display_value_json
                       ELSE raw_value_json
                     END
                   )
                   IN ('integer','real') THEN 'NUMBER'
              WHEN json_type(
                     CASE
                       WHEN json_type(display_value_json)<>'null'
                         THEN display_value_json
                       ELSE raw_value_json
                     END
                   )
                   IN ('true','false') THEN 'BOOLEAN'
              WHEN json_type(
                     CASE
                       WHEN json_type(display_value_json)<>'null'
                         THEN display_value_json
                       ELSE raw_value_json
                     END
                   )
                   IS NULL THEN 'EMPTY'
              ELSE 'TEXT'
            END AS value_kind,
            number_format,
            COUNT(*)
        FROM capture_v2_cells
        WHERE sheet_id=? AND row_index>=?
        GROUP BY column_index, value_kind, number_format
        ORDER BY column_index, value_kind, number_format
        """,
        (sheet_id, data_start_row),
    ).fetchall()
    kind_counts: dict[int, Counter[str]] = {}
    format_counts: dict[int, Counter[str]] = {}
    for column, kind, number_format, count in grouped:
        if str(kind) == "EMPTY":
            continue
        column_number = int(column)
        amount = int(count)
        kind_counts.setdefault(column_number, Counter())[str(kind)] += amount
        role = _format_role(number_format)
        format_counts.setdefault(column_number, Counter())[role] += amount
    roles: list[str] = []
    for column in range(bounds["minColumn"], bounds["maxColumn"] + 1):
        relative = column - bounds["minColumn"] + 1
        kind = kind_counts.get(column, Counter()).most_common(1)
        format_role = format_counts.get(column, Counter()).most_common(1)
        roles.append(
            f"C{relative}:"
            f"{kind[0][0] if kind else 'EMPTY'}:"
            f"{format_role[0][0] if format_role else 'GENERAL'}"
        )
    header_depth = max(
        0,
        min(data_start_row - bounds["minRow"], HEADER_SCAN_ROW_LIMIT),
    )
    format_roles = sorted(
        {
            _format_role(row[0])
            for row in connection.execute(
                """
                SELECT DISTINCT number_format
                FROM capture_v2_cells
                WHERE sheet_id=?
                  AND (
                    formula_text IS NOT NULL
                    OR json_type(raw_value_json)<>'null'
                    OR json_type(display_value_json)<>'null'
                  )
                """,
                (sheet_id,),
            )
        }
    )
    return roles, header_depth, format_roles


def build_structure_fingerprint_from_database(
    connection: sqlite3.Connection,
    capture_revision_id: int,
) -> dict[str, Any]:
    """Build the same v2 contract from normalized Capture v2 SQLite rows.

    Queries aggregate column roles in SQLite and only materialize structural
    header/anchor cells, so a catalog can be built without loading millions of
    source values into Python.
    """

    workbook_row = connection.execute(
        """
        SELECT r.content_sha256, w.workbook_status, w.sheet_count,
               w.tabular_sheet_count
        FROM capture_v2_revisions r
        JOIN capture_v2_workbooks w ON w.revision_id=r.revision_id
        WHERE r.revision_id=?
        """,
        (int(capture_revision_id),),
    ).fetchone()
    if workbook_row is None:
        raise StructureFingerprintError(
            f"Capture revision is missing: {capture_revision_id}"
        )
    sheet_rows = connection.execute(
        """
        SELECT sheet_id, sheet_index, title, sheet_state,
               has_tabular_evidence, used_bounds_json, content_bounds_json
        FROM capture_v2_sheets
        WHERE revision_id=?
        ORDER BY sheet_index
        """,
        (int(capture_revision_id),),
    ).fetchall()
    sheets: list[dict[str, Any]] = []
    for row in sheet_rows:
        sheet_id = int(row[0])
        used_bounds = _json_value(row[5])
        content_bounds = _json_value(row[6])
        bounds = _bounds(
            {
                "usedBounds": used_bounds,
                "contentBounds": content_bounds,
            }
        )
        column_roles, header_depth, number_format_roles = _database_column_roles(
            connection,
            sheet_id,
            bounds,
        )
        merged_ranges = [
            {
                "address": str(merged[0]),
                "minRow": int(merged[1]),
                "minColumn": int(merged[2]),
                "maxRow": int(merged[3]),
                "maxColumn": int(merged[4]),
                "anchor": str(merged[5]),
            }
            for merged in connection.execute(
                """
                SELECT address, min_row, min_column, max_row, max_column,
                       anchor_coordinate
                FROM capture_v2_merged_ranges
                WHERE sheet_id=?
                ORDER BY min_row, min_column, max_row, max_column
                """,
                (sheet_id,),
            )
        ]
        sheets.append(
            {
                "sheetIndex": int(row[1]),
                "title": str(row[2]),
                "sheetState": str(row[3]),
                "hasTabularEvidence": bool(row[4]),
                "usedBounds": used_bounds,
                "contentBounds": content_bounds,
                "mergedRanges": merged_ranges,
                "cells": _database_anchor_cells(
                    connection,
                    sheet_id,
                    bounds["minRow"] + HEADER_SCAN_ROW_LIMIT - 1,
                ),
                "_precomputedFormulaPatternHashes": _database_formula_patterns(
                    connection,
                    sheet_id,
                ),
                "_precomputedColumnRoleSketch": column_roles,
                "_precomputedNumberFormatRoles": number_format_roles,
                "_precomputedHeaderDepth": header_depth,
            }
        )
    return build_structure_fingerprint(
        {
            "source": {"contentSha256": str(workbook_row[0])},
            "workbook": {
                "status": str(workbook_row[1]),
                "sheetCount": int(workbook_row[2]),
                "tabularSheetCount": int(workbook_row[3]),
                "sheets": sheets,
            },
        }
    )


def validate_structure_fingerprint(fingerprint: dict[str, Any]) -> None:
    if fingerprint.get("schemaVersion") != FINGERPRINT_SCHEMA_VERSION:
        raise StructureFingerprintError("Unsupported fingerprint schemaVersion.")
    digest = fingerprint.get("fingerprintSha256")
    if not isinstance(digest, str) or len(digest) != 64:
        raise StructureFingerprintError("Fingerprint requires a SHA-256 digest.")
    structural_core = {
        "engineVersion": fingerprint.get("engineVersion"),
        "workbook": fingerprint.get("workbook"),
        "sheets": fingerprint.get("sheets"),
    }
    if _sha256(structural_core) != digest:
        raise StructureFingerprintError("Fingerprint digest does not match payload.")


__all__ = [
    "FINGERPRINT_ENGINE_VERSION",
    "FINGERPRINT_SCHEMA_VERSION",
    "StructureFingerprintError",
    "build_structure_fingerprint",
    "build_structure_fingerprint_from_database",
    "normalize_text",
    "text_tokens",
    "validate_structure_fingerprint",
]
