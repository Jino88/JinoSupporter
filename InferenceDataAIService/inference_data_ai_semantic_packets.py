"""Build lossless, domain-neutral source packets from Capture v2 SQLite.

The packetizer performs no semantic classification.  It copies every
non-empty captured cell exactly once into deterministic, bounded chunks while
retaining source provenance and structural metadata.
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
from collections import Counter
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator


PACKET_SCHEMA_VERSION = "semantic-source-packet-v1"
REQUIRED_TABLES = {
    "capture_v2_documents",
    "capture_v2_revisions",
    "capture_v2_workbooks",
    "capture_v2_sheets",
    "capture_v2_cells",
    "capture_v2_merged_ranges",
    "capture_v2_row_dimensions",
    "capture_v2_column_dimensions",
}


class SemanticPacketError(RuntimeError):
    """Raised when a Capture v2 database cannot produce trustworthy packets."""


def _stable_id(prefix: str, *parts: object) -> str:
    text = "\x1f".join(str(part) for part in parts)
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()[:24]
    return f"{prefix}_{digest}"


def _json_value(value: str | None) -> Any:
    if value is None:
        return None
    try:
        return json.loads(value)
    except json.JSONDecodeError as error:
        raise SemanticPacketError("Capture v2 contains invalid JSON metadata.") from error


def _rows(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> list[dict[str, Any]]:
    cursor = connection.execute(sql, tuple(parameters))
    columns = [description[0] for description in cursor.description or ()]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def _first(
    connection: sqlite3.Connection,
    sql: str,
    parameters: Iterable[Any] = (),
) -> dict[str, Any] | None:
    rows = _rows(connection, sql, parameters)
    return rows[0] if rows else None


def _validate_capture_schema(connection: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )
    }
    missing = sorted(REQUIRED_TABLES - tables)
    if missing:
        raise SemanticPacketError(
            "Capture v2 database is missing required tables: " + ", ".join(missing)
        )


@contextmanager
def connect_capture_v2_readonly(database_path: str | Path) -> Iterator[sqlite3.Connection]:
    """Open a Capture v2 database in SQLite's enforced read-only mode."""

    path = Path(database_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    connection = sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True)
    try:
        connection.execute("PRAGMA query_only=ON")
        yield connection
    finally:
        connection.close()


def _resolve_revision(
    connection: sqlite3.Connection,
    revision_id: int | None,
    source_path: str | Path | None,
) -> dict[str, Any]:
    conditions: list[str] = []
    parameters: list[Any] = []
    if revision_id is not None:
        conditions.append("r.revision_id=?")
        parameters.append(int(revision_id))
    if source_path is not None:
        conditions.append("d.source_path=?")
        parameters.append(str(Path(source_path).expanduser().resolve()))
    if revision_id is None:
        conditions.append("r.is_current=1")

    candidates = _rows(
        connection,
        f"""
        SELECT
            r.revision_id, r.revision_uid, r.document_id,
            r.content_sha256, r.capture_contract, r.extractor_name,
            r.extractor_version, r.size_bytes, r.mtime_ns,
            r.capture_status, r.is_current, r.capture_json_sha256,
            r.captured_at, r.stale_at,
            d.source_path, d.file_name, d.source_kind,
            w.workbook_status, w.is_truly_empty, w.sheet_count,
            w.nonempty_sheet_count, w.tabular_sheet_count, w.metadata_json
        FROM capture_v2_revisions r
        JOIN capture_v2_documents d ON d.document_id=r.document_id
        JOIN capture_v2_workbooks w ON w.revision_id=r.revision_id
        WHERE {" AND ".join(conditions)}
        ORDER BY r.revision_id
        """,
        parameters,
    )
    if not candidates:
        raise SemanticPacketError("No matching Capture v2 source revision was found.")
    if len(candidates) != 1:
        raise SemanticPacketError(
            "Revision selection is ambiguous; provide revision_id or source_path."
        )
    return candidates[0]


def _source_revision_payload(revision: dict[str, Any]) -> dict[str, Any]:
    return {
        "revisionId": int(revision["revision_id"]),
        "revisionUid": str(revision["revision_uid"]),
        "documentId": int(revision["document_id"]),
        "sourcePath": str(revision["source_path"]),
        "fileName": str(revision["file_name"]),
        "sourceKind": str(revision["source_kind"]),
        "contentSha256": str(revision["content_sha256"]),
        "captureContract": str(revision["capture_contract"]),
        "extractor": {
            "name": str(revision["extractor_name"]),
            "version": str(revision["extractor_version"]),
        },
        "sizeBytes": int(revision["size_bytes"]),
        "mtimeNs": int(revision["mtime_ns"]),
        "captureStatus": str(revision["capture_status"]),
        "isCurrent": bool(revision["is_current"]),
        "captureJsonSha256": str(revision["capture_json_sha256"]),
        "capturedAt": str(revision["captured_at"]),
        "staleAt": revision["stale_at"],
        "captureLimitations": {
            "formulaEvaluation": False,
            "embeddedVisualContentPolicy": "OUT_OF_SCOPE_NOT_CAPTURED",
            "displayValueIsStoredValueWithNumberFormat": True,
        },
    }


def _sheet_records(
    connection: sqlite3.Connection,
    revision_id: int,
) -> list[dict[str, Any]]:
    return _rows(
        connection,
        """
        SELECT
            sheet_id, sheet_index, title, sheet_state, capture_status,
            is_truly_empty, has_tabular_evidence, nonempty_cell_count,
            structural_cell_count, captured_cell_count, formula_cell_count,
            merge_count, used_bounds_json, content_bounds_json,
            freeze_panes, auto_filter, metadata_json
        FROM capture_v2_sheets
        WHERE revision_id=?
        ORDER BY sheet_index
        """,
        (revision_id,),
    )


def _row_dimensions(
    connection: sqlite3.Connection,
    sheet_id: int,
) -> list[dict[str, Any]]:
    return [
        {
            "row": int(row["row_index"]),
            "height": row["height"],
            "hidden": bool(row["hidden"]),
        }
        for row in _rows(
            connection,
            """
            SELECT row_index, height, hidden
            FROM capture_v2_row_dimensions
            WHERE sheet_id=?
            ORDER BY row_index
            """,
            (sheet_id,),
        )
    ]


def _column_dimensions(
    connection: sqlite3.Connection,
    sheet_id: int,
) -> list[dict[str, Any]]:
    return [
        {
            "key": str(row["dimension_key"]),
            "minColumn": int(row["min_column"]),
            "maxColumn": int(row["max_column"]),
            "width": row["width"],
            "hidden": bool(row["hidden"]),
        }
        for row in _rows(
            connection,
            """
            SELECT dimension_key, min_column, max_column, width, hidden
            FROM capture_v2_column_dimensions
            WHERE sheet_id=?
            ORDER BY min_column, max_column, dimension_key
            """,
            (sheet_id,),
        )
    ]


def _merged_ranges(
    connection: sqlite3.Connection,
    sheet_id: int,
) -> list[dict[str, Any]]:
    return [
        {
            "address": str(row["address"]),
            "minRow": int(row["min_row"]),
            "minColumn": int(row["min_column"]),
            "maxRow": int(row["max_row"]),
            "maxColumn": int(row["max_column"]),
            "anchor": str(row["anchor_coordinate"]),
        }
        for row in _rows(
            connection,
            """
            SELECT
                address, min_row, min_column, max_row, max_column,
                anchor_coordinate
            FROM capture_v2_merged_ranges
            WHERE sheet_id=?
            ORDER BY min_row, min_column, max_row, max_column
            """,
            (sheet_id,),
        )
    ]


def _nonempty_cells(
    connection: sqlite3.Connection,
    sheet: dict[str, Any],
    revision_uid: str,
    row_dimensions: list[dict[str, Any]],
    column_dimensions: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    rows_by_index = {dimension["row"]: dimension for dimension in row_dimensions}

    cells: list[dict[str, Any]] = []
    for row in _rows(
        connection,
        """
        SELECT
            row_index, column_index, coordinate, raw_value_json, formula_text,
            cached_value_json, display_value_json, data_type,
            cached_data_type, number_format, style_id, style_json,
            merge_range, merge_role
        FROM capture_v2_cells
        WHERE sheet_id=?
          AND (
              raw_value_json IS NOT NULL
              OR formula_text IS NOT NULL
              OR cached_value_json IS NOT NULL
              OR display_value_json IS NOT NULL
          )
        ORDER BY row_index, column_index
        """,
        (sheet["sheet_id"],),
    ):
        row_index = int(row["row_index"])
        column_index = int(row["column_index"])
        row_dimension = rows_by_index.get(row_index)
        matching_columns = [
            dimension
            for dimension in column_dimensions
            if dimension["minColumn"] <= column_index <= dimension["maxColumn"]
        ]
        column_dimension = matching_columns[0] if matching_columns else None
        coordinate = str(row["coordinate"])
        cells.append(
            {
                "sourceCellKey": (
                    f"{revision_uid}:{int(sheet['sheet_index'])}:{coordinate}"
                ),
                "row": row_index,
                "column": column_index,
                "coordinate": coordinate,
                "rawValue": _json_value(row["raw_value_json"]),
                "formula": row["formula_text"],
                "cachedValue": _json_value(row["cached_value_json"]),
                "displayValue": _json_value(row["display_value_json"]),
                "dataType": str(row["data_type"]),
                "cachedDataType": row["cached_data_type"],
                "numberFormat": str(row["number_format"]),
                "styleId": int(row["style_id"]),
                "style": _json_value(row["style_json"]) or {},
                "mergeRange": row["merge_range"],
                "mergeRole": str(row["merge_role"]),
                "hidden": {
                    "sheet": str(sheet["sheet_state"]) != "visible",
                    "row": bool(row_dimension and row_dimension["hidden"]),
                    "column": bool(column_dimension and column_dimension["hidden"]),
                },
                "rowDimension": row_dimension,
                "columnDimension": column_dimension,
            }
        )
    return cells


def _sheet_payload(
    sheet: dict[str, Any],
    row_dimensions: list[dict[str, Any]],
    column_dimensions: list[dict[str, Any]],
    merged_ranges: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "sheetId": int(sheet["sheet_id"]),
        "sheetIndex": int(sheet["sheet_index"]),
        "title": str(sheet["title"]),
        "sheetState": str(sheet["sheet_state"]),
        "status": str(sheet["capture_status"]),
        "isTrulyEmpty": bool(sheet["is_truly_empty"]),
        "hasTabularEvidence": bool(sheet["has_tabular_evidence"]),
        "nonEmptyCellCount": int(sheet["nonempty_cell_count"]),
        "structuralCellCount": int(sheet["structural_cell_count"]),
        "capturedCellCount": int(sheet["captured_cell_count"]),
        "formulaCellCount": int(sheet["formula_cell_count"]),
        "mergeCount": int(sheet["merge_count"]),
        "usedBounds": _json_value(sheet["used_bounds_json"]),
        "contentBounds": _json_value(sheet["content_bounds_json"]),
        "freezePanes": sheet["freeze_panes"],
        "autoFilter": sheet["auto_filter"],
        "metadata": _json_value(sheet["metadata_json"]) or {},
        "rowDimensions": row_dimensions,
        "columnDimensions": column_dimensions,
        "mergedRanges": merged_ranges,
    }


def _sections(
    cells: list[dict[str, Any]],
    empty_row_gap: int,
) -> list[list[tuple[int, list[dict[str, Any]]]]]:
    grouped_rows: list[tuple[int, list[dict[str, Any]]]] = []
    for cell in cells:
        if not grouped_rows or grouped_rows[-1][0] != cell["row"]:
            grouped_rows.append((int(cell["row"]), [cell]))
        else:
            grouped_rows[-1][1].append(cell)

    sections: list[list[tuple[int, list[dict[str, Any]]]]] = []
    for row_index, row_cells in grouped_rows:
        if (
            not sections
            or row_index - sections[-1][-1][0] - 1 >= empty_row_gap
        ):
            sections.append([])
        sections[-1].append((row_index, row_cells))
    return sections


def _bounds(cells: list[dict[str, Any]]) -> dict[str, int]:
    rows = [int(cell["row"]) for cell in cells]
    columns = [int(cell["column"]) for cell in cells]
    return {
        "minRow": min(rows),
        "minColumn": min(columns),
        "maxRow": max(rows),
        "maxColumn": max(columns),
        "rowCount": len(set(rows)),
        "cellCount": len(cells),
    }


def _column_label(column: int) -> str:
    label = ""
    value = column
    while value:
        value, remainder = divmod(value - 1, 26)
        label = chr(ord("A") + remainder) + label
    return label


def _range_from_bounds(bounds: dict[str, int]) -> str:
    start = f"{_column_label(bounds['minColumn'])}{bounds['minRow']}"
    end = f"{_column_label(bounds['maxColumn'])}{bounds['maxRow']}"
    return start if start == end else f"{start}:{end}"


def _intersects(bounds: dict[str, int], merged_range: dict[str, Any]) -> bool:
    return not (
        int(merged_range["maxRow"]) < bounds["minRow"]
        or int(merged_range["minRow"]) > bounds["maxRow"]
        or int(merged_range["maxColumn"]) < bounds["minColumn"]
        or int(merged_range["minColumn"]) > bounds["maxColumn"]
    )


def _chunk_packet(
    source_revision: dict[str, Any],
    sheet: dict[str, Any],
    section_index: int,
    chunk_index_in_section: int,
    cells: list[dict[str, Any]],
    split_reason: str,
    row_segment: dict[str, int] | None,
) -> dict[str, Any]:
    bounds = _bounds(cells)
    packet_id = _stable_id(
        "source_chunk",
        source_revision["revisionUid"],
        sheet["sheetIndex"],
        section_index,
        chunk_index_in_section,
        bounds["minRow"],
        bounds["minColumn"],
        bounds["maxRow"],
        bounds["maxColumn"],
    )
    row_dimensions = [
        dimension
        for dimension in sheet["rowDimensions"]
        if bounds["minRow"] <= int(dimension["row"]) <= bounds["maxRow"]
    ]
    column_dimensions = [
        dimension
        for dimension in sheet["columnDimensions"]
        if not (
            int(dimension["maxColumn"]) < bounds["minColumn"]
            or int(dimension["minColumn"]) > bounds["maxColumn"]
        )
    ]
    merged_ranges = [
        merged_range
        for merged_range in sheet["mergedRanges"]
        if _intersects(bounds, merged_range)
    ]
    style_dictionary: dict[str, dict[str, Any]] = {}
    packet_cells: list[dict[str, Any]] = []
    for source_cell in cells:
        cell = dict(source_cell)
        style_key = str(cell["styleId"])
        style = cell.pop("style", {})
        previous_style = style_dictionary.setdefault(style_key, style)
        if previous_style != style:
            raise SemanticPacketError(
                f"Capture v2 styleId {style_key} maps to conflicting style payloads."
            )
        cell.pop("rowDimension", None)
        cell.pop("columnDimension", None)
        if cell.get("formula"):
            cell["valueSource"] = (
                "FORMULA_CACHED"
                if cell.get("cachedValue") is not None
                else "FORMULA_NO_CACHE"
            )
        else:
            cell["valueSource"] = "RAW"
        cell["primary"] = True
        cell["contextOnly"] = False
        packet_cells.append(cell)
    return {
        "schemaVersion": PACKET_SCHEMA_VERSION,
        "packetType": "SOURCE_CHUNK",
        "packetId": packet_id,
        "chunkId": packet_id,
        "sourceRevision": source_revision,
        "sheet": {
            "sheetId": sheet["sheetId"],
            "sheetIndex": sheet["sheetIndex"],
            "title": sheet["title"],
            "sheetState": sheet["sheetState"],
            "status": sheet["status"],
            "isTrulyEmpty": sheet["isTrulyEmpty"],
            "hasTabularEvidence": sheet["hasTabularEvidence"],
            "usedBounds": sheet["usedBounds"],
            "contentBounds": sheet["contentBounds"],
            "freezePanes": sheet["freezePanes"],
            "autoFilter": sheet["autoFilter"],
            "metadata": sheet["metadata"],
        },
        "sectionIndex": section_index,
        "chunkIndexInSection": chunk_index_in_section,
        "splitReason": split_reason,
        "rowSegment": row_segment,
        "bounds": bounds,
        "primaryRange": _range_from_bounds(bounds),
        "truncated": False,
        "coverageOwnerRule": "Each non-empty Capture v2 source cell is primary in exactly one chunk.",
        "rowDimensions": row_dimensions,
        "columnDimensions": column_dimensions,
        "mergedRanges": merged_ranges,
        "styleDictionary": style_dictionary,
        "cells": packet_cells,
    }


def _chunk_section(
    source_revision: dict[str, Any],
    sheet: dict[str, Any],
    section_index: int,
    rows: list[tuple[int, list[dict[str, Any]]]],
    max_cells: int,
    max_rows: int,
) -> list[dict[str, Any]]:
    chunks: list[dict[str, Any]] = []
    pending_rows: list[tuple[int, list[dict[str, Any]]]] = []
    pending_cell_count = 0

    def flush_pending() -> None:
        nonlocal pending_rows, pending_cell_count
        if not pending_rows:
            return
        chunk_cells = [
            cell
            for _, row_cells in pending_rows
            for cell in row_cells
        ]
        chunks.append(
            _chunk_packet(
                source_revision,
                sheet,
                section_index,
                len(chunks) + 1,
                chunk_cells,
                "ROW_LIMIT" if len(pending_rows) >= max_rows else "CELL_LIMIT",
                None,
            )
        )
        pending_rows = []
        pending_cell_count = 0

    for row_index, row_cells in rows:
        if len(row_cells) > max_cells:
            flush_pending()
            segment_count = (len(row_cells) + max_cells - 1) // max_cells
            for segment_offset in range(0, len(row_cells), max_cells):
                segment = row_cells[segment_offset : segment_offset + max_cells]
                segment_index = segment_offset // max_cells + 1
                chunks.append(
                    _chunk_packet(
                        source_revision,
                        sheet,
                        section_index,
                        len(chunks) + 1,
                        segment,
                        "WIDE_ROW",
                        {
                            "row": row_index,
                            "segmentIndex": segment_index,
                            "segmentCount": segment_count,
                            "minColumn": int(segment[0]["column"]),
                            "maxColumn": int(segment[-1]["column"]),
                        },
                    )
                )
            continue

        exceeds_rows = len(pending_rows) >= max_rows
        exceeds_cells = pending_cell_count + len(row_cells) > max_cells
        if pending_rows and (exceeds_rows or exceeds_cells):
            flush_pending()
        pending_rows.append((row_index, row_cells))
        pending_cell_count += len(row_cells)

    flush_pending()
    _add_section_context(chunks, rows)
    return chunks


def _context_cell(source_cell: dict[str, Any]) -> dict[str, Any]:
    cell = dict(source_cell)
    cell.pop("style", None)
    cell.pop("rowDimension", None)
    cell.pop("columnDimension", None)
    if cell.get("formula"):
        cell["valueSource"] = (
            "FORMULA_CACHED"
            if cell.get("cachedValue") is not None
            else "FORMULA_NO_CACHE"
        )
    else:
        cell["valueSource"] = "RAW"
    cell["primary"] = False
    cell["contextOnly"] = True
    return cell


def _add_section_context(
    chunks: list[dict[str, Any]],
    rows: list[tuple[int, list[dict[str, Any]]]],
    *,
    header_row_count: int = 3,
    label_cell_count: int = 2,
) -> None:
    """Repeat bounded header/label cells without changing primary ownership."""

    header_cells = [
        cell
        for _, row_cells in rows[:header_row_count]
        for cell in row_cells
    ]
    row_lookup = {row_index: row_cells for row_index, row_cells in rows}
    for chunk in chunks:
        owned = {
            str(cell["sourceCellKey"])
            for cell in chunk["cells"]
        }
        candidates = [
            cell
            for cell in header_cells
            if int(cell["row"]) < int(chunk["bounds"]["minRow"])
            and int(chunk["bounds"]["minColumn"])
            <= int(cell["column"])
            <= int(chunk["bounds"]["maxColumn"])
        ]
        for row_index in range(
            int(chunk["bounds"]["minRow"]),
            int(chunk["bounds"]["maxRow"]) + 1,
        ):
            candidates.extend(row_lookup.get(row_index, [])[:label_cell_count])
        seen: set[str] = set()
        context_cells: list[dict[str, Any]] = []
        for source_cell in candidates:
            key = str(source_cell["sourceCellKey"])
            if key in owned or key in seen:
                continue
            seen.add(key)
            context_cells.append(_context_cell(source_cell))
        chunk["contextCells"] = context_cells
        chunk["contextPolicy"] = {
            "headerRowCount": header_row_count,
            "labelCellCount": label_cell_count,
            "primaryOwnershipUnchanged": True,
        }


def _terminal_packet(
    source_revision: dict[str, Any],
    terminal_status: str,
    scope: dict[str, Any],
) -> dict[str, Any]:
    return {
        "schemaVersion": PACKET_SCHEMA_VERSION,
        "packetType": "TERMINAL",
        "packetId": _stable_id(
            "terminal",
            source_revision["revisionUid"],
            terminal_status,
            scope.get("sheetIndex", "workbook"),
        ),
        "sourceRevision": source_revision,
        "terminalStatus": terminal_status,
        "scope": scope,
        "cells": [],
    }


def validate_packet_coverage(
    expected_source_cell_keys: Iterable[str],
    chunks: Iterable[dict[str, Any]],
) -> dict[str, Any]:
    expected = list(expected_source_cell_keys)
    actual = [
        str(cell["sourceCellKey"])
        for chunk in chunks
        for cell in chunk.get("cells") or []
    ]
    counts = Counter(actual)
    duplicates = sorted(key for key, count in counts.items() if count > 1)
    expected_set = set(expected)
    actual_set = set(actual)
    missing = sorted(expected_set - actual_set)
    unexpected = sorted(actual_set - expected_set)
    complete = (
        len(expected) == len(expected_set)
        and not duplicates
        and not missing
        and not unexpected
        and len(actual) == len(expected)
    )
    return {
        "status": "COMPLETE" if complete else "INVALID",
        "expectedCellCount": len(expected),
        "packetCellCount": len(actual),
        "uniquePacketCellCount": len(actual_set),
        "duplicateCellKeys": duplicates,
        "missingCellKeys": missing,
        "unexpectedCellKeys": unexpected,
    }


def build_semantic_source_packets(
    connection: sqlite3.Connection,
    *,
    revision_id: int | None = None,
    source_path: str | Path | None = None,
    max_cells: int = 400,
    max_rows: int = 50,
    empty_row_gap: int = 3,
) -> dict[str, Any]:
    """Build an inventory, terminal packets, and bounded universal chunks."""

    if max_cells < 1:
        raise ValueError("max_cells must be at least 1.")
    if max_rows < 1:
        raise ValueError("max_rows must be at least 1.")
    if empty_row_gap < 1:
        raise ValueError("empty_row_gap must be at least 1.")

    _validate_capture_schema(connection)
    revision = _resolve_revision(connection, revision_id, source_path)
    source_revision = _source_revision_payload(revision)
    sheets: list[dict[str, Any]] = []
    chunks: list[dict[str, Any]] = []
    terminal_packets: list[dict[str, Any]] = []
    expected_keys: list[str] = []

    for sheet_record in _sheet_records(connection, int(revision["revision_id"])):
        row_dimensions = _row_dimensions(connection, int(sheet_record["sheet_id"]))
        column_dimensions = _column_dimensions(connection, int(sheet_record["sheet_id"]))
        merged_ranges = _merged_ranges(connection, int(sheet_record["sheet_id"]))
        sheet = _sheet_payload(
            sheet_record,
            row_dimensions,
            column_dimensions,
            merged_ranges,
        )
        cells = _nonempty_cells(
            connection,
            sheet_record,
            str(revision["revision_uid"]),
            row_dimensions,
            column_dimensions,
        )
        declared_nonempty = int(sheet_record["nonempty_cell_count"])
        if len(cells) != declared_nonempty:
            raise SemanticPacketError(
                f"Capture v2 sheet {sheet['sheetIndex']} declares "
                f"{declared_nonempty} non-empty cells but exposes {len(cells)}."
            )
        expected_keys.extend(str(cell["sourceCellKey"]) for cell in cells)

        section_summaries: list[dict[str, Any]] = []
        for section_index, section_rows in enumerate(
            _sections(cells, empty_row_gap),
            start=1,
        ):
            section_chunks = _chunk_section(
                source_revision,
                sheet,
                section_index,
                section_rows,
                max_cells,
                max_rows,
            )
            chunks.extend(section_chunks)
            section_cells = [
                cell
                for _, row_cells in section_rows
                for cell in row_cells
            ]
            section_summaries.append(
                {
                    "sectionIndex": section_index,
                    "bounds": _bounds(section_cells),
                    "chunkIds": [chunk["packetId"] for chunk in section_chunks],
                }
            )

        if (
            str(revision["workbook_status"]) == "CAPTURED"
            and sheet["status"] == "NO_TABULAR_EVIDENCE"
        ):
            terminal_packets.append(
                _terminal_packet(
                    source_revision,
                    "NO_TABULAR_EVIDENCE",
                    {
                        "type": "SHEET",
                        "sheetId": sheet["sheetId"],
                        "sheetIndex": sheet["sheetIndex"],
                        "title": sheet["title"],
                    },
                )
            )

        sheet["packetNonEmptyCellCount"] = len(cells)
        sheet["sections"] = section_summaries
        sheets.append(sheet)

    workbook_status = str(revision["workbook_status"])
    if workbook_status in {"EMPTY_WORKBOOK", "NO_TABULAR_EVIDENCE"}:
        terminal_packets.append(
            _terminal_packet(
                source_revision,
                workbook_status,
                {"type": "WORKBOOK"},
            )
        )

    for chunk_index, chunk in enumerate(chunks, start=1):
        chunk["chunkIndex"] = chunk_index

    coverage = validate_packet_coverage(expected_keys, chunks)
    if coverage["status"] != "COMPLETE":
        raise SemanticPacketError(
            "Packet coverage invariant failed; no partial inventory was returned."
        )

    inventory = {
        "schemaVersion": PACKET_SCHEMA_VERSION,
        "packetType": "INVENTORY",
        "packetId": _stable_id("inventory", source_revision["revisionUid"]),
        "sourceRevision": source_revision,
        "workbook": {
            "status": workbook_status,
            "isTrulyEmpty": bool(revision["is_truly_empty"]),
            "sheetCount": int(revision["sheet_count"]),
            "nonEmptySheetCount": int(revision["nonempty_sheet_count"]),
            "tabularSheetCount": int(revision["tabular_sheet_count"]),
            "metadata": _json_value(revision["metadata_json"]) or {},
        },
        "limits": {
            "maxCells": max_cells,
            "maxRows": max_rows,
            "emptyRowGap": empty_row_gap,
        },
        "sheets": sheets,
        "chunkIds": [chunk["packetId"] for chunk in chunks],
        "terminalPacketIds": [
            packet["packetId"] for packet in terminal_packets
        ],
        "coverage": coverage,
        "semanticCellCoverageComplete": coverage["status"] == "COMPLETE",
        "contentCompleteForManifest": (
            coverage["status"] == "COMPLETE" and workbook_status == "CAPTURED"
        ),
    }
    return {
        "schemaVersion": PACKET_SCHEMA_VERSION,
        "inventory": inventory,
        "terminalPackets": terminal_packets,
        "chunks": chunks,
    }


def build_semantic_source_packets_from_db(
    database_path: str | Path,
    **options: Any,
) -> dict[str, Any]:
    with connect_capture_v2_readonly(database_path) as connection:
        return build_semantic_source_packets(connection, **options)


def packet_json_bytes(packet_set: dict[str, Any]) -> bytes:
    return (
        json.dumps(
            packet_set,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n"
    ).encode("utf-8")


__all__ = [
    "PACKET_SCHEMA_VERSION",
    "SemanticPacketError",
    "build_semantic_source_packets",
    "build_semantic_source_packets_from_db",
    "connect_capture_v2_readonly",
    "packet_json_bytes",
    "validate_packet_coverage",
]
