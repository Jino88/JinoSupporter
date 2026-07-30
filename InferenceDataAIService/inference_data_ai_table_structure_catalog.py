"""Build reusable table/block structures from table-first request artifacts."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from collections import defaultdict
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_structure_fingerprint import normalize_text, text_tokens


TABLE_STRUCTURE_FINGERPRINT_SCHEMA_VERSION = "excel-table-structure-fingerprint-v1"
TABLE_STRUCTURE_CATALOG_SCHEMA_VERSION = "excel-table-structure-catalog-v1"
TABLE_STRUCTURE_ENGINE_VERSION = "table-first-block-structure-v1.0"

_COORDINATE_PATTERN = re.compile(r"^([A-Z]{1,3})([0-9]+)$", re.IGNORECASE)
_RANGE_PATTERN = re.compile(
    r"^([A-Z]{1,3})([0-9]+):([A-Z]{1,3})([0-9]+)$",
    re.IGNORECASE,
)


class TableStructureCatalogError(RuntimeError):
    """Raised when a table-first request has no safe structural identity."""


def _canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _digest(value: Any) -> str:
    return hashlib.sha256(_canonical_bytes(value)).hexdigest()


def _bucket(value: int) -> str:
    if value <= 0:
        return "0"
    upper_bounds = (1, 2, 3, 4, 6, 8, 12, 16, 24, 32, 48, 64, 96, 128)
    for upper in upper_bounds:
        if value <= upper:
            return f"LE-{upper}"
    return "GT-128"


def _column_number(label: str) -> int:
    result = 0
    for character in label.upper():
        if not "A" <= character <= "Z":
            raise TableStructureCatalogError(f"Invalid column label: {label}")
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _coordinate(value: Any) -> tuple[int, int]:
    match = _COORDINATE_PATTERN.fullmatch(str(value or ""))
    if match is None:
        raise TableStructureCatalogError(f"Invalid cell coordinate: {value}")
    return int(match.group(2)), _column_number(match.group(1))


def _merge_shape(
    value: Any,
    *,
    min_row: int,
    min_column: int,
) -> dict[str, int] | None:
    if not value:
        return None
    match = _RANGE_PATTERN.fullmatch(str(value))
    if match is None:
        return None
    first_column = _column_number(match.group(1))
    first_row = int(match.group(2))
    last_column = _column_number(match.group(3))
    last_row = int(match.group(4))
    return {
        "relativeRow": first_row - min_row,
        "relativeColumn": first_column - min_column,
        "height": last_row - first_row + 1,
        "width": last_column - first_column + 1,
    }


def _format_role(value: Any) -> str:
    text = normalize_text(value)
    if not text or text == "GENERAL":
        return "GENERAL"
    if "%" in text:
        return "PERCENT"
    if any(token in text for token in ("YY", "MM", "DD")):
        return "DATE"
    if any(token in text for token in ("0", "#", "?")):
        return "NUMBER"
    return "OTHER"


def table_structure_fingerprint(table: dict[str, Any]) -> dict[str, Any]:
    """Fingerprint a deterministic table-first block without source values."""

    bounds = table.get("bounds")
    if not isinstance(bounds, dict):
        raise TableStructureCatalogError("Table bounds are missing.")
    min_row = int(bounds.get("minRow") or 0)
    min_column = int(bounds.get("minColumn") or 0)
    max_row = int(bounds.get("maxRow") or 0)
    max_column = int(bounds.get("maxColumn") or 0)
    if min_row < 1 or min_column < 1 or max_row < min_row or max_column < min_column:
        raise TableStructureCatalogError("Table bounds are invalid.")

    merge_shapes: dict[str, dict[str, int]] = {}
    preview_patterns: list[dict[str, Any]] = []
    header_values: list[str] = []
    for row in table.get("previewRows") or []:
        row_number = int(row.get("row") or 0)
        cells: list[dict[str, Any]] = []
        for cell in row.get("cells") or []:
            cell_row, cell_column = _coordinate(cell.get("coordinate"))
            kind = str(cell.get("kind") or "TEXT").upper()
            cells.append(
                {
                    "relativeRow": cell_row - min_row,
                    "relativeColumn": cell_column - min_column,
                    "kind": kind,
                    "merged": bool(cell.get("mergeRange")),
                }
            )
            shape = _merge_shape(
                cell.get("mergeRange"),
                min_row=min_row,
                min_column=min_column,
            )
            if shape is not None:
                merge_shapes[
                    json.dumps(shape, sort_keys=True, separators=(",", ":"))
                ] = shape
            if kind == "TEXT" and cell.get("value") is not None:
                header_values.append(str(cell["value"]))
        preview_patterns.append(
            {
                "relativeRow": row_number - min_row,
                "cells": sorted(
                    cells,
                    key=lambda item: (
                        item["relativeColumn"],
                        item["kind"],
                    ),
                ),
                "omittedCellBucket": _bucket(
                    int(row.get("omittedCellCount") or 0)
                ),
            }
        )

    numeric_columns: list[dict[str, Any]] = []
    for column in table.get("numericColumns") or []:
        column_number = _column_number(str(column.get("column") or ""))
        numeric_columns.append(
            {
                "relativeColumn": column_number - min_column,
                "columnRole": str(column.get("columnRole") or "UNSPECIFIED"),
                "numberFormatRoles": sorted(
                    {
                        _format_role(value)
                        for value in column.get("numberFormats") or []
                    }
                ),
                "numericCountBucket": _bucket(
                    int(column.get("numericCount") or 0)
                ),
            }
        )
        header_values.extend(
            str(value) for value in column.get("headerTexts") or []
        )

    numeric_series: list[dict[str, Any]] = []
    for series in table.get("numericSeries") or []:
        member_columns = list(
            series.get("columnIds")
            or series.get("memberColumnIds")
            or series.get("columns")
            or []
        )
        numeric_series.append(
            {
                "columnRole": str(series.get("columnRole") or "UNSPECIFIED"),
                "memberCountBucket": _bucket(len(member_columns)),
                "numberFormatRoles": sorted(
                    {
                        _format_role(value)
                        for value in series.get("numberFormats") or []
                    }
                ),
            }
        )

    row_label_shapes = [
        {
            "relativeRow": int(row.get("row") or 0) - min_row,
            "labelCountBucket": _bucket(len(row.get("labels") or [])),
        }
        for row in table.get("rowLabels") or []
    ]
    structural_core = {
        "engineVersion": TABLE_STRUCTURE_ENGINE_VERSION,
        "rowCountBucket": _bucket(max_row - min_row + 1),
        "columnCountBucket": _bucket(max_column - min_column + 1),
        "sourceCellCountBucket": _bucket(int(table.get("sourceCellCount") or 0)),
        "numericCellCountBucket": _bucket(int(table.get("numericCellCount") or 0)),
        "numericColumnCount": int(table.get("numericColumnCount") or 0),
        "previewPatterns": preview_patterns,
        "mergeShapes": sorted(
            merge_shapes.values(),
            key=lambda item: (
                item["relativeRow"],
                item["relativeColumn"],
                item["height"],
                item["width"],
            ),
        ),
        "rowLabelShapes": row_label_shapes,
        "numericColumns": sorted(
            numeric_columns,
            key=lambda item: (
                item["relativeColumn"],
                item["columnRole"],
            ),
        ),
        "numericSeries": sorted(
            numeric_series,
            key=lambda item: (
                item["columnRole"],
                item["memberCountBucket"],
            ),
        ),
        "aggregateCheckCountBucket": _bucket(
            len(table.get("aggregateChecks") or [])
        ),
    }
    return {
        "schemaVersion": TABLE_STRUCTURE_FINGERPRINT_SCHEMA_VERSION,
        **structural_core,
        "headerTokens": text_tokens(header_values),
        "fingerprintSha256": _digest(structural_core),
    }


def build_table_structure_catalog(
    *,
    table_first_batch_root: str | Path,
    limit: int | None = None,
    generated_at: str | None = None,
) -> dict[str, Any]:
    batch_root = Path(table_first_batch_root).expanduser().resolve()
    request_root = batch_root / "requests"
    if not request_root.is_dir():
        raise FileNotFoundError(request_root)
    paths = sorted(request_root.glob("*.json"))
    if limit is not None:
        paths = paths[: max(int(limit), 0)]
    groups: dict[str, dict[str, Any]] = {}
    workbook_structure_ids: dict[str, set[str]] = defaultdict(set)
    table_count = 0
    for request_path in paths:
        with request_path.open("r", encoding="utf-8") as stream:
            request = json.load(stream)
        analysis_path = batch_root / "analyses" / request_path.name
        analysis_tables: dict[str, dict[str, Any]] = {}
        if analysis_path.is_file():
            with analysis_path.open("r", encoding="utf-8") as stream:
                analysis = json.load(stream)
            analysis_tables = {
                str(item.get("tableId") or ""): item
                for item in analysis.get("tables") or []
                if item.get("tableId")
            }
        source = request.get("source") or {}
        workbook_key = str(
            source.get("contentSha256")
            or source.get("revisionUid")
            or request_path.name
        )
        for table in request.get("tables") or []:
            table_count += 1
            semantic = analysis_tables.get(str(table.get("tableId") or ""), {})
            fingerprint = table_structure_fingerprint(table)
            digest = fingerprint["fingerprintSha256"]
            group = groups.setdefault(
                digest,
                {
                    "tableStructureId": f"table-structure-{digest[:20]}",
                    "fingerprintSha256": digest,
                    "fingerprint": fingerprint,
                    "members": [],
                },
            )
            group["members"].append(
                {
                    "fileName": str(source.get("fileName") or ""),
                    "sourcePath": str(source.get("sourcePath") or ""),
                    "contentSha256": str(source.get("contentSha256") or ""),
                    "revisionUid": str(source.get("revisionUid") or ""),
                    "requestFile": request_path.name,
                    "tableId": str(table.get("tableId") or ""),
                    "sheetIndex": int(table.get("sheetIndex") or 0),
                    "sheet": str(table.get("sheet") or ""),
                    "range": str(table.get("range") or ""),
                    "semanticType": str(semantic.get("type") or "UNASSESSED"),
                    "semanticConfidence": str(
                        semantic.get("confidence") or "UNASSESSED"
                    ),
                }
            )
            workbook_structure_ids[workbook_key].add(
                group["tableStructureId"]
            )

    structures: list[dict[str, Any]] = []
    for group in groups.values():
        members = sorted(
            group["members"],
            key=lambda item: (
                item["fileName"],
                item["sheetIndex"],
                item["range"],
                item["tableId"],
            ),
        )
        workbook_count = len(
            {
                member["contentSha256"] or member["revisionUid"]
                for member in members
            }
        )
        semantic_type_counts = Counter(
            str(member["semanticType"]) for member in members
        )
        dominant_semantic_type = (
            semantic_type_counts.most_common(1)[0][0]
            if semantic_type_counts
            else "UNASSESSED"
        )
        quantitative = (
            group["fingerprint"]["numericCellCountBucket"] != "0"
            or int(group["fingerprint"]["numericColumnCount"]) > 0
        )
        structures.append(
            {
                **group,
                "tableCount": len(members),
                "workbookCount": workbook_count,
                "quantitative": quantitative,
                "semanticTypeCounts": dict(
                    sorted(semantic_type_counts.items())
                ),
                "dominantSemanticType": dominant_semantic_type,
                "semanticConsistency": (
                    round(
                        semantic_type_counts[dominant_semantic_type]
                        / len(members),
                        6,
                    )
                    if members
                    else 0.0
                ),
                "members": members,
            }
        )
    structures.sort(
        key=lambda item: (
            -int(item["tableCount"]),
            -int(item["workbookCount"]),
            str(item["fingerprintSha256"]),
        )
    )
    reusable = [
        structure
        for structure in structures
        if int(structure["workbookCount"]) > 1
    ]
    reusable_quantitative = [
        structure
        for structure in reusable
        if bool(structure["quantitative"])
        and str(structure["dominantSemanticType"]) != "TEXT"
    ]
    workbooks_with_reuse = {
        member["contentSha256"] or member["revisionUid"]
        for structure in reusable
        for member in structure["members"]
    }
    workbooks_with_quantitative_reuse = {
        member["contentSha256"] or member["revisionUid"]
        for structure in reusable_quantitative
        for member in structure["members"]
    }
    return {
        "schemaVersion": TABLE_STRUCTURE_CATALOG_SCHEMA_VERSION,
        "engineVersion": TABLE_STRUCTURE_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            "tableFirstBatchRoot": str(batch_root),
            "requestCount": len(paths),
            "limit": limit,
            "aiCalls": 0,
        },
        "summary": {
            "requestCount": len(paths),
            "tableCount": table_count,
            "tableStructureCount": len(structures),
            "reusableTableStructureCount": len(reusable),
            "tablesInReusableStructures": sum(
                int(item["tableCount"]) for item in reusable
            ),
            "workbooksWithReusableTableStructure": len(workbooks_with_reuse),
            "reusableQuantitativeStructureCount": len(
                reusable_quantitative
            ),
            "quantitativeTablesInReusableStructures": sum(
                int(item["tableCount"])
                for item in reusable_quantitative
            ),
            "workbooksWithReusableQuantitativeStructure": len(
                workbooks_with_quantitative_reuse
            ),
            "largestTableStructureTableCount": max(
                (int(item["tableCount"]) for item in structures),
                default=0,
            ),
            "largestTableStructureWorkbookCount": max(
                (int(item["workbookCount"]) for item in structures),
                default=0,
            ),
        },
        "structures": structures,
    }


def write_table_structure_catalog(
    catalog: dict[str, Any],
    output_path: str | Path,
) -> Path:
    target = Path(output_path).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.{os.getpid()}.tmp")
    temporary.write_bytes(_canonical_bytes(catalog) + b"\n")
    os.replace(temporary, target)
    return target


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build an AI-free table/block structure catalog."
    )
    parser.add_argument("--table-first-batch-root", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--limit", type=int)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    catalog = build_table_structure_catalog(
        table_first_batch_root=arguments.table_first_batch_root,
        limit=arguments.limit,
    )
    output = write_table_structure_catalog(catalog, arguments.out)
    print(
        json.dumps(
            {
                "output": str(output),
                "summary": catalog["summary"],
                "aiCalls": 0,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = [
    "TABLE_STRUCTURE_CATALOG_SCHEMA_VERSION",
    "TABLE_STRUCTURE_ENGINE_VERSION",
    "TABLE_STRUCTURE_FINGERPRINT_SCHEMA_VERSION",
    "TableStructureCatalogError",
    "build_table_structure_catalog",
    "table_structure_fingerprint",
    "write_table_structure_catalog",
]
