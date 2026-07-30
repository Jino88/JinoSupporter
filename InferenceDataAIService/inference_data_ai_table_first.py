"""Compact, table-first semantic analysis for captured Excel workbooks.

The module intentionally separates responsibilities:

* captured cells remain the authoritative value/evidence store;
* deterministic code builds bounded table/text inventories and numeric facts;
* AI labels table purpose, study grouping, groups, and metrics once per workbook.

It does not create per-cell semantic dispositions and it never self-approves a
comparison for aggregation.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import copy
from collections import defaultdict
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

from inference_data_ai_formula_derivation import (
    FORMULA_EVALUATOR_VERSION,
    FormulaDerivationError,
    apply_formula_overlay_to_chunks,
    derive_formula_overlay,
)
from inference_data_ai_term_dictionary import (
    TermDictionaryAdapter,
    load_term_dictionary_adapter,
)


REQUEST_SCHEMA_VERSION = "table-first-request-v1"
ANALYSIS_SCHEMA_VERSION = "table-first-analysis-v1"
PROJECTION_SCHEMA_VERSION = "table-first-projection-v1"
BUILDER_VERSION = "table-first-builder-v8"
PROMPT_VERSION = "table-first-analysis-prompt-v4"

TABLE_TYPES = ("COMPARISON", "DESCRIPTIVE", "SUPPORTING", "TEXT")
CONFIDENCE_LEVELS = ("HIGH", "MEDIUM", "LOW")
GROUP_ROLES = (
    "CONTROL",
    "REFERENCE",
    "COMPARATOR",
    "TREATMENT",
    "TEST",
    "BEFORE",
    "AFTER",
    "OTHER",
    "UNASSESSED",
)


class TableFirstError(RuntimeError):
    """Raised when the compact request or AI response breaks its contract."""


def _stable_id(prefix: str, *parts: object) -> str:
    material = "\x1f".join(str(part) for part in parts)
    digest = hashlib.sha256(material.encode("utf-8")).hexdigest()[:20]
    return f"{prefix}_{digest}"


def _column_label(column: int) -> str:
    result = ""
    value = int(column)
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _range(
    min_row: int,
    min_column: int,
    max_row: int,
    max_column: int,
) -> str:
    start = f"{_column_label(min_column)}{min_row}"
    end = f"{_column_label(max_column)}{max_row}"
    return start if start == end else f"{start}:{end}"


def _text(value: object, *, limit: int = 240) -> str:
    if value is None:
        return ""
    if isinstance(value, dict) and value.get("type") in {
        "date",
        "datetime",
        "time",
        "timedelta",
    }:
        value = value.get("value")
    result = re.sub(r"\s+", " ", str(value)).strip()
    if len(result) <= limit:
        return result
    return result[: max(0, limit - 1)] + "…"


def _cell_number(cell: dict[str, Any]) -> float | None:
    value = cell.get("cachedValue") if cell.get("formula") else cell.get("rawValue")
    if value is None and not cell.get("formula"):
        value = cell.get("displayValue")
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    result = float(value)
    return result if math.isfinite(result) else None


def _cell_display(cell: dict[str, Any]) -> str:
    for key in ("displayValue", "cachedValue", "rawValue", "formula"):
        value = cell.get(key)
        if value is not None:
            return _text(value, limit=160)
    return ""


def _cell_kind(cell: dict[str, Any]) -> str:
    if cell.get("formula"):
        return "FORMULA"
    if _cell_number(cell) is not None:
        return "NUMBER"
    raw_value = cell.get("rawValue")
    if isinstance(raw_value, dict) and raw_value.get("type") in {
        "date",
        "datetime",
        "time",
        "timedelta",
    }:
        return "DATE"
    return "TEXT"


def _normalized_number_display(
    number: float,
    number_format: object,
    fallback: str,
) -> tuple[str, str, float]:
    """Normalize only source-explicit display scales without semantic guessing."""

    format_text = _text(number_format, limit=80)
    if "%" not in format_text:
        return fallback, "RAW", 1.0
    decimal_match = re.search(r"\.([0#]+)%", format_text)
    decimal_places = len(decimal_match.group(1)) if decimal_match else 0
    scaled = number * 100.0
    return f"{scaled:.{decimal_places}f}%", "PERCENT", 100.0


def _bounds(cells: Sequence[dict[str, Any]]) -> tuple[int, int, int, int]:
    return (
        min(int(cell["row"]) for cell in cells),
        min(int(cell["column"]) for cell in cells),
        max(int(cell["row"]) for cell in cells),
        max(int(cell["column"]) for cell in cells),
    )


def _frequency_value_from_text(value: object) -> float | None:
    text = _text(value, limit=80).casefold()
    match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*(k)?hz\s*", text)
    if match is None:
        return None
    number = float(match.group(1))
    if match.group(2):
        number *= 1000.0
    return number if math.isfinite(number) and number > 0 else None


def _increasing_frequency_axis(
    points: Sequence[tuple[int, float, str]],
    *,
    minimum_points: int,
    require_wide_span: bool,
) -> bool:
    if len(points) < minimum_points:
        return False
    ordered = sorted(points, key=lambda point: point[0])
    values = [point[1] for point in ordered]
    if len(set(values)) < minimum_points:
        return False
    increasing = sum(
        right > left for left, right in zip(values, values[1:])
    )
    if increasing / max(1, len(values) - 1) < 0.9:
        return False
    minimum = min(values)
    maximum = max(values)
    if maximum < 500:
        return False
    return not require_wide_span or maximum / minimum >= 10


def _frequency_axes(
    cells: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Find explicit horizontal Hz headers or conservative numeric Hz columns."""

    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    by_column: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
        by_column[int(cell["column"])].append(cell)

    axes: list[dict[str, Any]] = []
    for row, row_cells in sorted(by_row.items()):
        points = [
            (int(cell["column"]), frequency, str(cell["coordinate"]))
            for cell in row_cells
            if (frequency := _frequency_value_from_text(_cell_display(cell)))
            is not None
        ]
        if not _increasing_frequency_axis(
            points,
            minimum_points=4,
            require_wide_span=False,
        ):
            continue
        ordered = sorted(points, key=lambda point: point[0])
        axes.append(
            {
                "orientation": "ROW",
                "row": row,
                "range": f"{ordered[0][2]}:{ordered[-1][2]}",
                "pointCount": len(ordered),
                "minimumHz": min(point[1] for point in ordered),
                "maximumHz": max(point[1] for point in ordered),
                "unitEvidence": "EXPLICIT_HZ",
            }
        )

    for column, column_cells in sorted(by_column.items()):
        points_by_row: dict[int, tuple[int, float, str]] = {}
        for cell in column_cells:
            number = _cell_number(cell)
            if number is None or number <= 0:
                continue
            row = int(cell["row"])
            points_by_row.setdefault(
                row,
                (row, number, str(cell["coordinate"])),
            )
        points = list(points_by_row.values())
        if not _increasing_frequency_axis(
            points,
            minimum_points=12,
            require_wide_span=True,
        ):
            continue
        ordered = sorted(points, key=lambda point: point[0])
        axes.append(
            {
                "orientation": "COLUMN",
                "column": _column_label(column),
                "range": f"{ordered[0][2]}:{ordered[-1][2]}",
                "pointCount": len(ordered),
                "minimumHz": min(point[1] for point in ordered),
                "maximumHz": max(point[1] for point in ordered),
                "unitEvidence": "FREQUENCY_SEQUENCE",
            }
        )
    return axes[:8]


def _raw_frequency_metric_families(
    cells: Sequence[dict[str, Any]],
    *,
    sheet_title: str,
) -> list[str]:
    """Return only source-explicit SPL/THD/IMP raw-data metric families."""

    cell_texts = [
        _cell_display(cell)
        for cell in cells
        if _cell_number(cell) is None and _cell_display(cell)
    ]
    evidence = [sheet_title, *cell_texts]
    folded = [re.sub(r"\s+", " ", value).strip().casefold() for value in evidence]
    joined = " | ".join(folded)
    sheet_folded = re.sub(r"\s+", " ", sheet_title).strip().casefold()
    sheet_is_raw_data = "raw data" in sheet_folded

    families: list[str] = []
    patterns = {
        "SPL": r"(?<![a-z0-9])spl\s+(?:raw\s+)?data(?![a-z0-9])",
        "THD": r"(?<![a-z0-9])thd\s+(?:raw\s+)?data(?![a-z0-9])",
        "IMP": (
            r"(?<![a-z0-9])imp(?:edance)?\s+(?:raw\s+)?data"
            r"(?![a-z0-9])"
        ),
    }
    for family, pattern in patterns.items():
        if re.search(pattern, joined, flags=re.IGNORECASE):
            families.append(family)

    if re.search(r"frequency\s+response\s*\[[^\]]*db\s*spl[^\]]*\]", joined):
        if "SPL" not in families:
            families.append("SPL")

    if sheet_is_raw_data:
        raw_sheet_patterns = {
            "SPL": r"^spl(?:[_\s-]*freq\d*)?$",
            "THD": r"^thd(?:[_\s-]*freq\d*)?$",
            "IMP": r"^imp(?:edance)?(?:[_\s-]*freq\d*)?$",
        }
        for family, pattern in raw_sheet_patterns.items():
            if family not in families and any(
                re.search(pattern, value, flags=re.IGNORECASE)
                for value in folded
            ):
                families.append(family)
    return families


def _raw_frequency_response_signature(
    cells: Sequence[dict[str, Any]],
    *,
    sheet_title: str,
) -> dict[str, Any] | None:
    families = _raw_frequency_metric_families(cells, sheet_title=sheet_title)
    if not families:
        return None
    axes = _frequency_axes(cells)
    if not axes:
        return None
    return {"metricFamilies": families, "frequencyAxes": axes}


def _split_contiguous_rows(
    cells: Sequence[dict[str, Any]],
) -> list[list[dict[str, Any]]]:
    """Split coherent blocks without fragmenting spaced measurement rows.

    Two or more fully blank rows always form a boundary.  One blank row forms
    a boundary only when the following row looks like a new title/header.
    """

    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    blocks: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    previous_row: int | None = None
    for row in sorted(by_row):
        row_cells = by_row[row]
        text_values = [
            _cell_display(cell)
            for cell in row_cells
            if _cell_number(cell) is None and _cell_display(cell)
        ]
        joined = " ".join(text_values).lower()
        starts_named_block = (
            not any(_cell_number(cell) is not None for cell in row_cells)
            and bool(text_values)
            and len(text_values) <= 4
            and (
                len(joined) >= 20
                or any(
                    term in joined
                    for term in (
                        "title",
                        "report",
                        "result",
                        "test",
                        "check",
                        "analysis",
                        "conclusion",
                        "결과",
                        "시험",
                        "분석",
                        "결론",
                    )
                )
            )
        )
        numbered_section_title = bool(
            text_values
            and len(text_values) <= 2
            and re.match(
                r"^\s*(?:[ivx]+\.|\d+(?:\.\d+)*\.)\s*",
                joined,
                flags=re.IGNORECASE,
            )
        )
        internal_named_boundary = bool(
            current
            and not any(_cell_number(cell) is not None for cell in row_cells)
            and (
                numbered_section_title
                or (
                    len(text_values) == 1
                    and any(
                        term in joined
                        for term in (
                            "result check",
                            "result checking",
                            "function ng",
                            "measurement result",
                            "test result",
                        )
                    )
                )
            )
        )
        header_after_blank = bool(
            previous_row is not None
            and row == previous_row + 2
            and not any(_cell_number(cell) is not None for cell in row_cells)
            and bool(text_values)
            and (starts_named_block or len(text_values) >= 4)
        )
        if previous_row is not None and (
            row > previous_row + 2
            or header_after_blank
            or internal_named_boundary
        ):
            blocks.append(current)
            current = []
        current.extend(sorted(row_cells, key=lambda item: int(item["column"])))
        previous_row = row
    if current:
        blocks.append(current)
    return blocks


def _is_table_block(cells: Sequence[dict[str, Any]]) -> bool:
    rows = {int(cell["row"]) for cell in cells}
    columns = {int(cell["column"]) for cell in cells}
    numeric_count = sum(_cell_number(cell) is not None for cell in cells)
    return (
        (len(rows) >= 2 and len(columns) >= 2)
        or (numeric_count >= 1 and len(columns) >= 2)
    )


def _is_source_metric_label(cell: dict[str, Any]) -> bool:
    if _cell_number(cell) is not None or cell.get("formula"):
        return False
    value = re.sub(r"\s+", " ", _cell_display(cell)).strip().casefold()
    if not value or value.startswith("="):
        return False
    return bool(
        re.fullmatch(
            r"(?:f[0o]|fo|spl average|average spl)",
            value,
        )
    )


def _has_numeric_or_formula(cells: Sequence[dict[str, Any]]) -> bool:
    return any(
        _cell_number(cell) is not None
        or bool(cell.get("formula"))
        or _cell_display(cell).startswith("=")
        for cell in cells
    )


def _split_inline_metric_row(
    cells: Sequence[dict[str, Any]],
) -> list[list[dict[str, Any]]]:
    """Split one physical row when it contains adjacent named metrics."""

    rows = {int(cell["row"]) for cell in cells}
    if len(rows) != 1:
        return [list(cells)]
    ordered = sorted(cells, key=lambda item: int(item["column"]))
    anchors = [cell for cell in ordered if _is_source_metric_label(cell)]
    if len(anchors) < 2:
        return [list(cells)]
    segments: list[list[dict[str, Any]]] = []
    for index, anchor in enumerate(anchors):
        start = int(anchor["column"])
        end = (
            int(anchors[index + 1]["column"]) - 1
            if index + 1 < len(anchors)
            else max(int(cell["column"]) for cell in ordered)
        )
        segment = [
            cell
            for cell in ordered
            if start <= int(cell["column"]) <= end
        ]
        if not _has_numeric_or_formula(segment):
            return [list(cells)]
        segments.append(segment)
    return segments


def _split_summary_metric_row(
    cells: Sequence[dict[str, Any]],
) -> list[list[dict[str, Any]]]:
    """Separate a top Fo/F0 summary row from a following response matrix."""

    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    rows = sorted(by_row)
    if len(rows) < 4:
        return [list(cells)]
    first_row = rows[0]
    for row in rows[1:3]:
        row_cells = by_row[row]
        metric_labels = [
            cell for cell in row_cells if _is_source_metric_label(cell)
        ]
        if len(metric_labels) != 1:
            continue
        label = re.sub(
            r"\s+", " ", _cell_display(metric_labels[0])
        ).strip().casefold()
        if label not in {"f0", "fo"}:
            continue
        numeric_count = sum(
            _cell_number(cell) is not None for cell in row_cells
        )
        later_numeric_rows = sum(
            any(_cell_number(cell) is not None for cell in by_row[later_row])
            for later_row in rows
            if later_row > row
        )
        if numeric_count < 3 or later_numeric_rows < 2:
            continue
        header_cells = [
            cell for header_row in rows if first_row <= header_row < row
            for cell in by_row[header_row]
        ]
        summary = [*header_cells, *row_cells]
        response = [
            *header_cells,
            *[
                cell
                for later_row in rows
                if later_row > row
                for cell in by_row[later_row]
            ],
        ]
        if _is_table_block(summary) and _is_table_block(response):
            return [summary, response]
    return [list(cells)]


def _semantic_subblocks(
    cells: Sequence[dict[str, Any]],
) -> list[list[dict[str, Any]]]:
    result: list[list[dict[str, Any]]] = []
    for summary_block in _split_summary_metric_row(cells):
        result.extend(_split_inline_metric_row(summary_block))
    return result


def _selected_preview_rows(
    cells: Sequence[dict[str, Any]],
    max_preview_rows: int,
) -> list[int]:
    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    rows = sorted(by_row)
    numeric_rows = [
        row
        for row in rows
        if any(_cell_number(cell) is not None for cell in by_row[row])
    ]
    selected: list[int] = []
    for row in [*rows[:5], *numeric_rows[:4], *rows[-2:]]:
        if row not in selected:
            selected.append(row)
        if len(selected) >= max_preview_rows:
            break
    return sorted(selected)


def _preview_rows(
    cells: Sequence[dict[str, Any]],
    *,
    max_preview_rows: int,
    max_preview_columns: int,
) -> list[dict[str, Any]]:
    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    rows: list[dict[str, Any]] = []
    for row in _selected_preview_rows(cells, max_preview_rows):
        source_cells = sorted(by_row[row], key=lambda item: int(item["column"]))
        compact_cells: list[dict[str, Any]] = []
        for cell in source_cells[:max_preview_columns]:
            item = {
                "coordinate": str(cell["coordinate"]),
                "value": _cell_display(cell),
                "kind": _cell_kind(cell),
            }
            number_format = _text(cell.get("numberFormat"), limit=80)
            if number_format and number_format != "General":
                item["numberFormat"] = number_format
            if cell.get("mergeRange"):
                item["mergeRange"] = str(cell["mergeRange"])
            compact_cells.append(item)
        rows.append(
            {
                "rowId": f"row_{row}",
                "row": row,
                "cells": compact_cells,
                "omittedCellCount": max(0, len(source_cells) - len(compact_cells)),
            }
        )
    return rows


def _header_texts_for_column(
    cells: Sequence[dict[str, Any]],
    column: int,
    first_numeric_row: int,
) -> list[str]:
    values: list[str] = []
    for cell in sorted(cells, key=lambda item: (int(item["row"]), int(item["column"]))):
        if int(cell["row"]) > first_numeric_row:
            continue
        if _cell_number(cell) is not None:
            continue
        applies = int(cell["column"]) == column
        merge_range = str(cell.get("mergeRange") or "")
        match = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", merge_range)
        if match:
            def column_number(label: str) -> int:
                result = 0
                for character in label:
                    result = result * 26 + ord(character) - ord("A") + 1
                return result

            applies = column_number(match.group(1)) <= column <= column_number(
                match.group(3)
            )
        value = _cell_display(cell)
        if applies and value and value not in values:
            values.append(value)
    return values[-4:]


def _numeric_column_role(header_texts: Sequence[str]) -> str:
    labels = [
        re.sub(r"\s+", " ", str(value)).strip().casefold()
        for value in header_texts
        if str(value).strip()
    ]
    label = labels[-1] if labels else ""
    if re.fullmatch(r"(?:min|minimum)", label):
        return "AGGREGATE_MIN"
    if re.fullmatch(r"(?:max|maximum)", label):
        return "AGGREGATE_MAX"
    if re.fullmatch(r"(?:avg|average|mean)", label):
        return "AGGREGATE_AVERAGE"
    identifier_pattern = (
        r"(?:date|no\.?|number|position|posistion|possition|q'?ty|qty|input|spec|"
        r"result|status|judg(?:e)?ment|ok|pass|fail)"
    )
    if any(re.fullmatch(identifier_pattern, value) for value in labels):
        return "IDENTIFIER_OR_BASIS"
    return "MEASURE_VALUE"


def _numeric_columns(
    cells: Sequence[dict[str, Any]],
    *,
    table_id: str,
    max_value_samples: int,
) -> list[dict[str, Any]]:
    by_column: dict[int, list[tuple[dict[str, Any], float]]] = defaultdict(list)
    numeric_rows: list[int] = []
    for cell in cells:
        number = _cell_number(cell)
        if number is None:
            continue
        by_column[int(cell["column"])].append((cell, number))
        numeric_rows.append(int(cell["row"]))
    if not numeric_rows:
        return []
    first_numeric_row = min(numeric_rows)
    result: list[dict[str, Any]] = []
    for column in sorted(by_column):
        entries = by_column[column]
        numbers = [number for _, number in entries]
        header_texts = _header_texts_for_column(
            cells,
            column,
            first_numeric_row,
        )
        is_rate_column = any(
            re.search(r"(?:\brate\b|percent|percentage|%)", text, re.IGNORECASE)
            for text in header_texts
        )
        formats: list[str] = []
        samples: list[dict[str, Any]] = []
        for cell, number in entries:
            number_format = _text(cell.get("numberFormat"), limit=80)
            if number_format and number_format not in formats:
                formats.append(number_format)
            if len(samples) < max_value_samples:
                display = _cell_display(cell)
                normalized_display, display_scale, scale_factor = (
                    _normalized_number_display(
                        number,
                        cell.get("numberFormat"),
                        display,
                    )
                )
                if (
                    display_scale == "RAW"
                    and is_rate_column
                    and "/" in str(cell.get("formula") or "")
                    and abs(number) <= 1
                ):
                    normalized_display = f"{number * 100.0:.2f}%"
                    display_scale = "PERCENT_FROM_RATE_FORMULA"
                sample = {
                    "coordinate": str(cell["coordinate"]),
                    "rawNumber": number,
                }
                if display_scale.startswith("PERCENT"):
                    sample["normalizedDisplay"] = normalized_display
                    sample["displayScale"] = display_scale
                elif display and display != _text(number):
                    sample["sourceDisplay"] = display
                samples.append(sample)
        result.append(
            {
                "columnId": f"{table_id}_col_{_column_label(column)}",
                "column": _column_label(column),
                "headerTexts": header_texts,
                "columnRole": _numeric_column_role(header_texts),
                "numericCount": len(entries),
                "min": min(numbers),
                "max": max(numbers),
                "average": math.fsum(numbers) / len(numbers),
                "displaySamples": samples,
                "numberFormats": formats[:5],
                "sourceRange": _range(
                    min(int(cell["row"]) for cell, _ in entries),
                    column,
                    max(int(cell["row"]) for cell, _ in entries),
                    column,
                ),
            }
        )
    return result


def _compact_numeric_inventory(
    numeric_columns: Sequence[dict[str, Any]],
    *,
    table_id: str,
    detailed_column_limit: int = 48,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Collapse very wide raw matrices into bounded semantic axis groups."""

    columns = list(numeric_columns)
    if len(columns) <= detailed_column_limit:
        return columns, []

    def column_number(label: str) -> int:
        result = 0
        for character in label:
            result = result * 26 + ord(character) - ord("A") + 1
        return result

    groups: list[list[dict[str, Any]]] = []
    pending: list[dict[str, Any]] = []
    previous_number: int | None = None
    previous_role = ""
    for column in columns:
        number = column_number(str(column["column"]))
        role = str(column["columnRole"])
        if pending and (
            number != (previous_number or 0) + 1 or role != previous_role
        ):
            groups.append(pending)
            pending = []
        pending.append(column)
        previous_number = number
        previous_role = role
    if pending:
        groups.append(pending)

    series: list[dict[str, Any]] = []
    for index, group in enumerate(groups, start=1):
        total_count = sum(int(column["numericCount"]) for column in group)
        weighted_sum = math.fsum(
            float(column["average"]) * int(column["numericCount"])
            for column in group
        )
        sample_columns = [*group[:6], *group[-3:]]
        header_samples: list[dict[str, Any]] = []
        sample_values: list[dict[str, Any]] = []
        for column in sample_columns:
            header_item = {
                "column": column["column"],
                "headerTexts": column["headerTexts"],
            }
            if header_item not in header_samples:
                header_samples.append(header_item)
            for sample in column["displaySamples"][:1]:
                if sample not in sample_values:
                    sample_values.append(sample)
        series.append(
            {
                "seriesId": f"{table_id}_series_{index}",
                "columnRole": group[0]["columnRole"],
                "columnRange": (
                    f"{group[0]['column']}:{group[-1]['column']}"
                ),
                "numericColumnCount": len(group),
                "numericCellCount": total_count,
                "min": min(float(column["min"]) for column in group),
                "max": max(float(column["max"]) for column in group),
                "average": weighted_sum / total_count,
                "headerSamples": header_samples,
                "valueSamples": sample_values[:6],
            }
        )

    mandatory = [
        column
        for column in columns
        if str(column["columnRole"]) != "MEASURE_VALUE"
    ]
    selected = [*mandatory, *columns[:16], *columns[-8:]]
    selected_ids: set[str] = set()
    detailed: list[dict[str, Any]] = []
    for column in columns:
        column_id = str(column["columnId"])
        if column in selected and column_id not in selected_ids:
            selected_ids.add(column_id)
            detailed.append(column)
    return detailed, series


def _aggregate_checks(
    cells: Sequence[dict[str, Any]],
    numeric_columns: Sequence[dict[str, Any]],
    *,
    table_id: str,
) -> list[dict[str, Any]]:
    """Verify or derive Min/Max/Average from the captured raw matrix."""

    def column_number(label: str) -> int:
        result = 0
        for character in label:
            result = result * 26 + ord(character) - ord("A") + 1
        return result

    role_by_column = {
        column_number(str(column["column"])): str(column["columnRole"])
        for column in numeric_columns
    }
    metadata_by_column = {
        column_number(str(column["column"])): column for column in numeric_columns
    }
    aggregate_roles = {
        "AGGREGATE_MIN",
        "AGGREGATE_MAX",
        "AGGREGATE_AVERAGE",
    }
    aggregate_headers_by_row: defaultdict[
        int, list[tuple[int, str]]
    ] = defaultdict(list)
    for cell in cells:
        if _cell_number(cell) is not None:
            continue
        role = _numeric_column_role([_cell_display(cell)])
        if role not in aggregate_roles:
            continue
        column = int(cell["column"])
        role_by_column[column] = role
        aggregate_headers_by_row[int(cell["row"])].append((column, role))

    aggregate_groups: list[dict[str, Any]] = []

    def append_header_groups(
        header_row: int,
        entries: Sequence[tuple[int, str]],
    ) -> None:
        pending: dict[str, int] = {}
        for column, role in sorted(entries):
            if role in pending:
                pending = {}
            pending[role] = column
            if set(pending) == aggregate_roles:
                aggregate_groups.append(
                    {
                        "headerRow": header_row,
                        "columns": dict(pending),
                    }
                )
                pending = {}

    for header_row, entries in sorted(aggregate_headers_by_row.items()):
        append_header_groups(header_row, entries)

    if not aggregate_groups:
        return []

    cell_by_position = {
        (int(cell["row"]), int(cell["column"])): cell for cell in cells
    }
    max_row = max(int(cell["row"]) for cell in cells)
    checks: list[dict[str, Any]] = []

    def normalized_header_family(column: int, header_row: int) -> str:
        labels = _header_texts_for_column(cells, column, header_row)
        for value in reversed(labels):
            label = re.sub(r"\s+", " ", str(value)).strip().casefold()
            if not label or re.fullmatch(
                r"(?:min|minimum|max|maximum|avg|average|mean)", label
            ):
                continue
            if re.fullmatch(
                r"(?:(?:samples?|position|posistion|pos|no\.?)"
                r"[\s#-]*\d+|#\s*\d+|no\.?\s+samples?)",
                label,
            ):
                return "sample"
            label = re.sub(
                r"(?:sample|position|posistion|pos|no\.?)[\s#-]*\d+",
                "sample",
                label,
            )
            label = re.sub(r"#\s*\d+", "", label)
            label = re.sub(r"\b(?:samples?|no\.?|number)\b", "", label)
            label = re.sub(r"[^0-9a-z가-힣]+", " ", label).strip()
            if label:
                return label
        return ""

    def is_basis_column(column: int, header_row: int) -> bool:
        labels = _header_texts_for_column(cells, column, header_row)
        leaf_label = str(labels[-1]).casefold() if labels else ""
        return bool(
            re.search(
                r"(?:\btotal\s*ng\b|\bng\s*total\b|\bsample\s*no\.?\b|"
                r"\bq'?ty\b|\bqty\b|\binput\b|\bdate\b|\btype\b|"
                r"\bline\b|\bmachine\b|\bcavity\b|\bposition\b|"
                r"\bposistion\b|\bpossition\b|^\s*sample\s*$|"
                r"\bspec(?:ification)?\b|"
                r"\btime\b|\btemperature\b|\bpressure\b|\bcondition\b)",
                leaf_label,
            )
        )

    def raw_column_runs(header_row: int) -> list[list[int]]:
        candidates = [
            column
            for column, role in sorted(role_by_column.items())
            if role == "MEASURE_VALUE" and not is_basis_column(column, header_row)
        ]
        runs: list[list[int]] = []
        for column in candidates:
            family = normalized_header_family(column, header_row)
            if runs:
                previous = runs[-1][-1]
                previous_family = normalized_header_family(previous, header_row)
                same_family = (
                    family == previous_family
                    or not family
                    or not previous_family
                )
                if column != previous + 1 or not same_family:
                    runs.append([])
            else:
                runs.append([])
            runs[-1].append(column)
        return runs

    def choose_raw_columns(group: dict[str, Any]) -> list[int]:
        header_row = int(group["headerRow"])
        group_columns = list(group["columns"].values())
        group_families = {
            normalized_header_family(column, header_row)
            for column in group_columns
        } - {""}
        runs = raw_column_runs(header_row)
        if not runs:
            return []
        matching_runs = [
            run
            for run in runs
            if group_families
            and normalized_header_family(run[0], header_row) in group_families
        ]
        group_min = min(group_columns)
        group_max = max(group_columns)

        def score(run: list[int]) -> tuple[int, int]:
            distance = min(
                abs(group_min - run[-1]),
                abs(run[0] - group_max),
            )
            return distance, -len(run)

        if matching_runs:
            return min(matching_runs, key=score)
        adjacent_runs: list[list[int]] = []
        for run in runs:
            distance = score(run)[0]
            if distance == 1:
                adjacent_runs.append(run)
                continue
            if distance != 2:
                continue
            between_column = (
                run[-1] + 1 if run[-1] < group_min else group_max + 1
            )
            if is_basis_column(between_column, header_row):
                adjacent_runs.append(run)
        return min(adjacent_runs, key=score) if adjacent_runs else []

    def calculate(values: Sequence[float]) -> dict[str, float]:
        return {
            "min": min(values),
            "max": max(values),
            "average": math.fsum(values) / len(values),
        }

    def matches(
        calculated: dict[str, float],
        explicit: dict[str, float],
    ) -> bool:
        if set(explicit) != {"min", "max", "average"}:
            return False
        return all(
            math.isclose(
                calculated[name],
                explicit[name],
                rel_tol=1e-9,
                abs_tol=1e-6,
            )
            for name in ("min", "max", "average")
        )

    for group_index, group in enumerate(aggregate_groups, start=1):
        raw_columns = choose_raw_columns(group)
        if not raw_columns:
            continue
        aggregate_columns = {
            int(column): str(role)
            for role, column in group["columns"].items()
        }
        header_row = int(group["headerRow"])
        later_overlapping_headers = [
            int(other["headerRow"])
            for other in aggregate_groups
            if int(other["headerRow"]) > header_row
            and set(other["columns"].values()) & set(aggregate_columns)
        ]
        group_max_row = (
            min(later_overlapping_headers) - 1
            if later_overlapping_headers
            else max_row
        )
        rows_with_aggregates = sorted(
            {
                row
                for row, column in cell_by_position
                if column in aggregate_columns
                and header_row < row <= group_max_row
            }
        )
        for row_index, row in enumerate(rows_with_aggregates):
            explicit_cells: dict[str, dict[str, Any]] = {}
            explicit_values: dict[str, float] = {}
            for column, role in aggregate_columns.items():
                cell = cell_by_position.get((row, column))
                number = _cell_number(cell or {})
                name = role.removeprefix("AGGREGATE_").casefold()
                if number is not None:
                    explicit_values[name] = number
                if cell is not None:
                    explicit_cells[name] = {
                        "coordinate": str(cell["coordinate"]),
                        "rawNumber": number,
                        "formula": cell.get("formula"),
                    }
            explicit_complete = set(explicit_values) == {
                "min",
                "max",
                "average",
            }

            same_row_values = [
                number
                for column in raw_columns
                if (
                    number := _cell_number(
                        cell_by_position.get((row, column), {})
                    )
                )
                is not None
            ]
            mode = "ROW"
            raw_values = same_row_values
            raw_min_row = row
            raw_max_row = row
            calculated = calculate(raw_values) if len(raw_values) >= 2 else None
            if (
                explicit_complete
                and (calculated is None or not matches(calculated, explicit_values))
            ):
                block_values: list[float] = []
                block_rows: list[int] = []
                block_max_row = (
                    rows_with_aggregates[row_index + 1] - 1
                    if row_index + 1 < len(rows_with_aggregates)
                    else group_max_row
                )
                for block_row in range(row, block_max_row + 1):
                    for column in raw_columns:
                        number = _cell_number(
                            cell_by_position.get((block_row, column), {})
                        )
                        if number is not None:
                            block_values.append(number)
                            block_rows.append(block_row)
                if len(block_values) >= 2:
                    block_calculated = calculate(block_values)
                    if matches(block_calculated, explicit_values):
                        mode = "BLOCK"
                        raw_values = block_values
                        raw_min_row = min(block_rows)
                        raw_max_row = max(block_rows)
                        calculated = block_calculated
            if calculated is None:
                continue
            checks.append(
                {
                    "checkId": (
                        f"{table_id}_aggregate_{group_index}_{row}"
                    ),
                    "aggregateHeaderRow": header_row,
                    "mode": mode,
                    "rawRange": _range(
                        raw_min_row,
                        min(raw_columns),
                        raw_max_row,
                        max(raw_columns),
                    ),
                    "rawCount": len(raw_values),
                    "explicitCells": explicit_cells,
                    "explicit": explicit_values,
                    "calculated": calculated,
                    "status": (
                        "CALCULATED_NO_SOURCE_VALUE"
                        if not explicit_complete
                        else (
                            "MATCH"
                            if matches(calculated, explicit_values)
                            else "MISMATCH"
                        )
                    ),
                }
            )
    return checks


def _row_labels(
    cells: Sequence[dict[str, Any]],
    *,
    table_id: str,
    limit: int = 24,
) -> list[dict[str, Any]]:
    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    result: list[dict[str, Any]] = []
    for row in sorted(by_row):
        labels = [
            _cell_display(cell)
            for cell in sorted(by_row[row], key=lambda item: int(item["column"]))
            if _cell_number(cell) is None and _cell_display(cell)
        ][:3]
        if not labels:
            continue
        result.append(
            {
                "rowId": f"{table_id}_row_{row}",
                "row": row,
                "labels": labels,
            }
        )
        if len(result) >= limit:
            break
    return result


def _title_candidates(cells: Sequence[dict[str, Any]]) -> list[str]:
    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    numeric_rows = [
        row
        for row, row_cells in by_row.items()
        if any(_cell_number(cell) is not None for cell in row_cells)
    ]
    first_numeric_row = min(numeric_rows) if numeric_rows else None
    title_rows = [
        row
        for row in sorted(by_row)
        if first_numeric_row is None or row < first_numeric_row
    ][:4]
    candidates: list[str] = []
    for row in title_rows:
        row_cells = sorted(by_row[row], key=lambda item: int(item["column"]))
        text_values = [
            _cell_display(cell)
            for cell in row_cells
            if _cell_number(cell) is None and _cell_display(cell)
        ]
        if text_values:
            candidate = " | ".join(text_values[:4])
            if candidate not in candidates:
                candidates.append(candidate)
    return candidates


def _table_payload(
    cells: Sequence[dict[str, Any]],
    *,
    revision_uid: str,
    sheet_index: int,
    sheet_title: str,
    section_index: int,
    block_index: object,
    max_preview_rows: int,
    max_preview_columns: int,
    max_value_samples: int,
) -> dict[str, Any]:
    min_row, min_column, max_row, max_column = _bounds(cells)
    table_id = _stable_id(
        "table",
        revision_uid,
        sheet_index,
        section_index,
        block_index,
        min_row,
        min_column,
        max_row,
        max_column,
    )
    all_numeric_columns = _numeric_columns(
        cells,
        table_id=table_id,
        max_value_samples=max_value_samples,
    )
    aggregate_checks = _aggregate_checks(
        cells,
        all_numeric_columns,
        table_id=table_id,
    )
    numeric_columns, numeric_series = _compact_numeric_inventory(
        all_numeric_columns,
        table_id=table_id,
    )
    metric_labels = [
        _cell_display(cell)
        for cell in sorted(
            cells,
            key=lambda item: (int(item["row"]), int(item["column"])),
        )
        if _is_source_metric_label(cell)
    ]
    metric_axis_refs = [
        str(series["seriesId"])
        for series in numeric_series
        if not str(series.get("columnRole") or "").startswith("AGGREGATE_")
    ] or [
        str(column["columnId"])
        for column in numeric_columns
        if not str(column.get("columnRole") or "").startswith("AGGREGATE_")
    ]
    unique_metric_labels = list(dict.fromkeys(metric_labels))
    metric_hint_eligible = (
        len({int(cell["row"]) for cell in cells}) <= 2
        and len(unique_metric_labels) == 1
    )
    metric_hints = [
        {
            "name": label,
            "axisRefs": metric_axis_refs,
        }
        for label in unique_metric_labels
        if metric_axis_refs and metric_hint_eligible
    ]
    return {
        "tableId": table_id,
        "sheetIndex": sheet_index,
        "sheet": sheet_title,
        "range": _range(min_row, min_column, max_row, max_column),
        "bounds": {
            "minRow": min_row,
            "minColumn": min_column,
            "maxRow": max_row,
            "maxColumn": max_column,
        },
        "sourceCellCount": len(cells),
        "numericCellCount": sum(
            int(column["numericCount"]) for column in all_numeric_columns
        ),
        "numericColumnCount": len(all_numeric_columns),
        "titleCandidates": _title_candidates(cells),
        "previewRows": _preview_rows(
            cells,
            max_preview_rows=max_preview_rows,
            max_preview_columns=max_preview_columns,
        ),
        "rowLabels": _row_labels(cells, table_id=table_id),
        "numericColumns": numeric_columns,
        "numericSeries": numeric_series,
        "metricHints": metric_hints,
        "aggregateChecks": aggregate_checks,
        "nearbyTextIds": [],
    }


def _raw_frequency_response_exclusion_payload(
    cells: Sequence[dict[str, Any]],
    *,
    revision_uid: str,
    sheet_index: int,
    sheet_title: str,
    section_index: int,
    block_index: object,
    signature: dict[str, Any],
) -> dict[str, Any]:
    min_row, min_column, max_row, max_column = _bounds(cells)
    source_table_id = _stable_id(
        "table",
        revision_uid,
        sheet_index,
        section_index,
        block_index,
        min_row,
        min_column,
        max_row,
        max_column,
    )
    return {
        "exclusionId": _stable_id(
            "raw_frequency_exclusion",
            source_table_id,
            "RAW_FREQUENCY_RESPONSE_DATA",
        ),
        "sourceTableId": source_table_id,
        "reason": "RAW_FREQUENCY_RESPONSE_DATA",
        "codeOwner": "DETERMINISTIC",
        "sheetIndex": sheet_index,
        "sheet": sheet_title,
        "range": _range(min_row, min_column, max_row, max_column),
        "bounds": {
            "minRow": min_row,
            "minColumn": min_column,
            "maxRow": max_row,
            "maxColumn": max_column,
        },
        "sourceCellCount": len(cells),
        "numericCellCount": sum(
            _cell_number(cell) is not None for cell in cells
        ),
        "metricFamilies": list(signature["metricFamilies"]),
        "frequencyAxes": list(signature["frequencyAxes"]),
        "sourceStorage": {
            "schemaVersion": "semantic-source-packet-v1",
            "coordinatesPreserved": True,
            "valuesPreserved": True,
        },
    }


def _workbook_learning_exclusion_payload(
    *,
    raw_frequency_response_exclusions: Sequence[dict[str, Any]],
    tables: Sequence[dict[str, Any]],
    text_blocks: Sequence[dict[str, Any]],
    captured_primary_cell_count: int,
) -> dict[str, Any]:
    """Exclude an entire workbook when it contains acoustic RAW curves.

    The source packet remains lossless, but no table or nearby text from the
    workbook is exposed to workbook-level semantic AI or recipe learning.
    """

    metric_families = sorted(
        {
            str(family)
            for exclusion in raw_frequency_response_exclusions
            for family in exclusion.get("metricFamilies") or []
            if str(family) in {"SPL", "THD", "IMP"}
        }
    )
    return {
        "excluded": True,
        "reason": "WORKBOOK_CONTAINS_SPL_THD_IMP_RAW_FREQUENCY_DATA",
        "codeOwner": "DETERMINISTIC",
        "metricFamilies": metric_families,
        "triggerTableCount": len(raw_frequency_response_exclusions),
        "excludedNonRawTableCount": len(tables),
        "excludedTextBlockCount": len(text_blocks),
        "excludedCapturedPrimaryCellCount": int(
            captured_primary_cell_count
        ),
        "sourceStorage": {
            "schemaVersion": "semantic-source-packet-v1",
            "coordinatesPreserved": True,
            "valuesPreserved": True,
        },
    }


def _text_block_payload(
    cells: Sequence[dict[str, Any]],
    *,
    revision_uid: str,
    sheet_index: int,
    sheet_title: str,
    section_index: int,
    block_index: int,
) -> dict[str, Any]:
    min_row, min_column, max_row, max_column = _bounds(cells)
    by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        by_row[int(cell["row"])].append(cell)
    lines = []
    for row in sorted(by_row):
        line = " | ".join(
            value
            for value in (
                _cell_display(cell)
                for cell in sorted(
                    by_row[row],
                    key=lambda item: int(item["column"]),
                )
            )
            if value
        )
        if line:
            lines.append({"row": row, "text": _text(line, limit=500)})
    return {
        "textId": _stable_id(
            "text",
            revision_uid,
            sheet_index,
            section_index,
            block_index,
            min_row,
            min_column,
            max_row,
            max_column,
        ),
        "sheetIndex": sheet_index,
        "sheet": sheet_title,
        "range": _range(min_row, min_column, max_row, max_column),
        "bounds": {
            "minRow": min_row,
            "minColumn": min_column,
            "maxRow": max_row,
            "maxColumn": max_column,
        },
        "lines": lines[:30],
        "omittedLineCount": max(0, len(lines) - 30),
    }


def build_table_first_request(
    packet_set: dict[str, Any],
    *,
    max_preview_rows: int = 12,
    max_preview_columns: int = 16,
    max_value_samples: int = 3,
    term_dictionary_adapter: TermDictionaryAdapter | None = None,
    term_dictionary_path: str | Path | None = None,
) -> dict[str, Any]:
    """Build one bounded workbook request from a lossless source packet."""

    if packet_set.get("schemaVersion") != "semantic-source-packet-v1":
        raise TableFirstError("Expected semantic-source-packet-v1 input.")
    for name, value in (
        ("max_preview_rows", max_preview_rows),
        ("max_preview_columns", max_preview_columns),
        ("max_value_samples", max_value_samples),
    ):
        if int(value) < 1:
            raise ValueError(f"{name} must be at least 1.")
    if term_dictionary_adapter is not None and term_dictionary_path is not None:
        raise ValueError(
            "term_dictionary_adapter and term_dictionary_path are mutually exclusive."
        )
    term_dictionary = term_dictionary_adapter or load_term_dictionary_adapter(
        term_dictionary_path
    )
    term_dictionary_snapshot = term_dictionary.snapshot()
    inventory = packet_set.get("inventory")
    if not isinstance(inventory, dict):
        raise TableFirstError("Source packet has no inventory.")
    revision = inventory.get("sourceRevision")
    if not isinstance(revision, dict):
        raise TableFirstError("Source packet has no source revision.")
    revision_uid = _text(revision.get("revisionUid"))
    content_sha256 = _text(revision.get("contentSha256"))
    if not revision_uid or not content_sha256:
        raise TableFirstError("Source revision identity is incomplete.")

    source_chunks = list(packet_set.get("chunks") or [])
    formula_count = sum(
        1
        for chunk in source_chunks
        for cell in chunk.get("cells") or []
        if isinstance(cell, dict) and cell.get("formula")
    )
    formula_derivation: dict[str, Any] = {
        "status": "NOT_NEEDED",
        "evaluatorVersion": FORMULA_EVALUATOR_VERSION,
        "formulaCount": formula_count,
        "numericCount": 0,
        "nonNumericCount": 0,
        "errorCount": 0,
    }
    analysis_chunks = source_chunks
    if formula_count > 5000:
        formula_derivation["status"] = "SKIPPED_FORMULA_BUDGET"
    elif formula_count:
        try:
            overlay = derive_formula_overlay(
                source_chunks,
                expected_revision_uid=revision_uid,
                expected_content_sha256=content_sha256,
                tolerate_unsupported=True,
            )
        except (FormulaDerivationError, AssertionError) as exc:
            formula_derivation["status"] = "SKIPPED_UNSUPPORTED"
            formula_derivation["reason"] = _text(exc, limit=300)
        else:
            numeric_count = int(overlay["numericCount"])
            non_numeric_count = int(overlay["nonNumericCount"])
            error_count = int(overlay["errorCount"])
            apply_overlay = error_count == 0 or numeric_count >= 2
            if apply_overlay:
                analysis_chunks = apply_formula_overlay_to_chunks(
                    source_chunks,
                    overlay,
                    validate=False,
                )
            error_samples = [
                str(error)[:200]
                for error in list(overlay.get("errorsByCode") or {})[:3]
            ]
            formula_derivation.update(
                {
                    "status": (
                        (
                            "DERIVED"
                            if numeric_count
                            else "CLASSIFIED_NON_NUMERIC"
                        )
                        if error_count == 0
                        else (
                            "PARTIALLY_DERIVED"
                            if apply_overlay
                            else "SKIPPED_UNSUPPORTED"
                        )
                    ),
                    "numericCount": numeric_count,
                    "nonNumericCount": non_numeric_count,
                    "appliedNumericCount": numeric_count if apply_overlay else 0,
                    "errorCount": error_count,
                    "overlaySha256": str(overlay["overlaySha256"]),
                    "errorSamples": error_samples,
                }
            )

    sections: dict[
        tuple[int, str, int],
        dict[str, dict[str, Any]],
    ] = {}
    seen_cell_keys: set[str] = set()
    for chunk in analysis_chunks:
        if not isinstance(chunk, dict):
            continue
        sheet = chunk.get("sheet") or {}
        key = (
            int(sheet.get("sheetIndex") or 0),
            str(sheet.get("title") or ""),
            int(chunk.get("sectionIndex") or 0),
        )
        section_cells = sections.setdefault(key, {})
        for cell in chunk.get("cells") or []:
            if not isinstance(cell, dict) or not cell.get("primary", True):
                continue
            source_key = str(cell.get("sourceCellKey") or "")
            if not source_key:
                raise TableFirstError("A source cell has no sourceCellKey.")
            if source_key in seen_cell_keys:
                raise TableFirstError(
                    f"Duplicate primary source cell in packet: {source_key}"
                )
            seen_cell_keys.add(source_key)
            section_cells[source_key] = cell

    tables: list[dict[str, Any]] = []
    text_blocks: list[dict[str, Any]] = []
    raw_frequency_response_exclusions: list[dict[str, Any]] = []
    for (sheet_index, sheet_title, section_index), keyed_cells in sorted(
        sections.items()
    ):
        blocks = _split_contiguous_rows(list(keyed_cells.values()))
        for block_index, block in enumerate(blocks, start=1):
            if not _is_table_block(block):
                text_blocks.append(
                    _text_block_payload(
                        block,
                        revision_uid=revision_uid,
                        sheet_index=sheet_index,
                        sheet_title=sheet_title,
                        section_index=section_index,
                        block_index=block_index,
                    )
                )
                continue
            subblocks = _semantic_subblocks(block)
            for subblock_index, subblock in enumerate(subblocks, start=1):
                payload_block_index = (
                    block_index
                    if len(subblocks) == 1
                    else f"{block_index}.{subblock_index}"
                )
                raw_frequency_signature = _raw_frequency_response_signature(
                    subblock,
                    sheet_title=sheet_title,
                )
                if raw_frequency_signature is not None:
                    raw_frequency_response_exclusions.append(
                        _raw_frequency_response_exclusion_payload(
                            subblock,
                            revision_uid=revision_uid,
                            sheet_index=sheet_index,
                            sheet_title=sheet_title,
                            section_index=section_index,
                            block_index=payload_block_index,
                            signature=raw_frequency_signature,
                        )
                    )
                    continue
                tables.append(
                    _table_payload(
                        subblock,
                        revision_uid=revision_uid,
                        sheet_index=sheet_index,
                        sheet_title=sheet_title,
                        section_index=section_index,
                        block_index=payload_block_index,
                        max_preview_rows=max_preview_rows,
                        max_preview_columns=max_preview_columns,
                        max_value_samples=max_value_samples,
                    )
                )

    for table in tables:
        table_bounds = table["bounds"]
        table["nearbyTextIds"] = [
            text_block["textId"]
            for text_block in text_blocks
            if text_block["sheetIndex"] == table["sheetIndex"]
            and (
                0
                <= int(table_bounds["minRow"])
                - int(text_block["bounds"]["maxRow"])
                <= 3
                or 0
                <= int(text_block["bounds"]["minRow"])
                - int(table_bounds["maxRow"])
                <= 3
            )
        ]

    workbook_learning_exclusion: dict[str, Any] | None = None
    if raw_frequency_response_exclusions:
        workbook_learning_exclusion = (
            _workbook_learning_exclusion_payload(
                raw_frequency_response_exclusions=(
                    raw_frequency_response_exclusions
                ),
                tables=tables,
                text_blocks=text_blocks,
                captured_primary_cell_count=len(seen_cell_keys),
            )
        )
        tables = []
        text_blocks = []

    request_id = _stable_id(
        "table_request",
        REQUEST_SCHEMA_VERSION,
        BUILDER_VERSION,
        revision_uid,
        content_sha256,
        max_preview_rows,
        max_preview_columns,
        max_value_samples,
        len(tables),
        len(text_blocks),
        len(raw_frequency_response_exclusions),
        term_dictionary_snapshot["adapterVersion"],
        term_dictionary_snapshot["status"],
        term_dictionary_snapshot["contentSha256"],
    )
    result = {
        "schemaVersion": REQUEST_SCHEMA_VERSION,
        "builderVersion": BUILDER_VERSION,
        "requestId": request_id,
        "source": {
            "revisionUid": revision_uid,
            "contentSha256": content_sha256,
            "fileName": _text(revision.get("fileName"), limit=500),
            "sourcePath": _text(revision.get("sourcePath"), limit=1000),
        },
        "workbook": {
            "status": str((inventory.get("workbook") or {}).get("status") or ""),
            "learningStatus": (
                "EXCLUDED_RAW_FREQUENCY_RESPONSE_WORKBOOK"
                if workbook_learning_exclusion is not None
                else "ELIGIBLE"
            ),
            "sheetCount": int(
                (inventory.get("workbook") or {}).get("sheetCount") or 0
            ),
            "tableCount": len(tables),
            "rawFrequencyResponseExclusionCount": len(
                raw_frequency_response_exclusions
            ),
            "textBlockCount": len(text_blocks),
            "capturedPrimaryCellCount": len(seen_cell_keys),
            "coverageStatus": str(
                (inventory.get("coverage") or {}).get("status") or ""
            ),
        },
        "policy": {
            "aiCallBudget": (
                0 if workbook_learning_exclusion is not None else 1
            ),
            "workbookSemanticLearningEnabled": (
                workbook_learning_exclusion is None
            ),
            "valuesAreCodeOwned": True,
            "statisticsAreCodeOwned": True,
            "aggregateChecksAreCodeOwned": True,
            "evidenceIsCodeOwned": True,
            "termDictionaryRulesAreCodeOwned": True,
            "defaultVerificationStatus": "NEEDS_REVIEW",
        },
        "formulaDerivation": formula_derivation,
        "codeOwnedTermDictionary": term_dictionary_snapshot,
        "codeOwnedExclusions": {
            "rawFrequencyResponseTables": raw_frequency_response_exclusions,
            "workbookLearningExclusion": workbook_learning_exclusion,
        },
        "limits": {
            "maxPreviewRows": max_preview_rows,
            "maxPreviewColumns": max_preview_columns,
            "maxValueSamples": max_value_samples,
        },
        "tables": tables,
        "textBlocks": text_blocks,
    }
    result["requestBytes"] = 0
    for _ in range(4):
        request_bytes = (
            len(
                json.dumps(
                    result,
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                ).encode("utf-8")
            )
            + 1
        )
        if result["requestBytes"] == request_bytes:
            break
        result["requestBytes"] = request_bytes
    return result


def table_first_output_schema() -> dict[str, Any]:
    text_array = {"type": "array", "items": {"type": "string"}}
    group = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "label": {"type": "string"},
            "role": {"type": "string", "enum": list(GROUP_ROLES)},
            "basis": {"type": "string"},
        },
        "required": ["label", "role", "basis"],
    }
    metric = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "name": {"type": "string"},
            "unit": {"type": "string"},
            "axisRefs": text_array,
        },
        "required": ["name", "unit", "axisRefs"],
    }
    relation = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "leftGroup": {"type": "string"},
            "rightGroup": {"type": "string"},
            "basis": {"type": "string"},
        },
        "required": ["leftGroup", "rightGroup", "basis"],
    }
    table = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "tableId": {"type": "string"},
            "title": {"type": "string"},
            "type": {"type": "string", "enum": list(TABLE_TYPES)},
            "studyGroup": {"type": "string"},
            "groups": {"type": "array", "items": group},
            "metrics": {"type": "array", "items": metric},
            "comparisonRelations": {"type": "array", "items": relation},
            "textLinks": text_array,
            "relatedTableIds": text_array,
            "confidence": {
                "type": "string",
                "enum": list(CONFIDENCE_LEVELS),
            },
            "limitations": text_array,
        },
        "required": [
            "tableId",
            "title",
            "type",
            "studyGroup",
            "groups",
            "metrics",
            "comparisonRelations",
            "textLinks",
            "relatedTableIds",
            "confidence",
            "limitations",
        ],
    }
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": ANALYSIS_SCHEMA_VERSION,
            },
            "promptVersion": {
                "type": "string",
                "const": PROMPT_VERSION,
            },
            "requestId": {"type": "string"},
            "revisionUid": {"type": "string"},
            "status": {
                "type": "string",
                "enum": ["ANALYZED", "NEEDS_REVIEW", "NO_TABLES"],
            },
            "workbookSummary": {"type": "string"},
            "tables": {"type": "array", "items": table},
            "notes": text_array,
        },
        "required": [
            "schemaVersion",
            "promptVersion",
            "requestId",
            "revisionUid",
            "status",
            "workbookSummary",
            "tables",
            "notes",
        ],
    }


def _semantic_prompt_projection(
    request: dict[str, Any],
) -> dict[str, Any]:
    """Remove code-owned numeric detail while preserving semantic evidence."""

    prompt_request = copy.deepcopy(request)
    prompt_request.pop("codeOwnedExclusions", None)
    prompt_request.pop("codeOwnedTermDictionary", None)
    for table in prompt_request.get("tables") or []:
        role_by_column = {
            str(column.get("column") or ""): str(
                column.get("columnRole") or ""
            )
            for column in table.get("numericColumns") or []
        }
        labels_by_row = {
            int(row.get("row") or 0): {
                str(label) for label in row.get("labels") or []
            }
            for row in table.get("rowLabels") or []
        }
        aggregate_checks = list(table.pop("aggregateChecks", []) or [])
        if aggregate_checks:
            status_counts: dict[str, int] = {}
            for check in aggregate_checks:
                status = str(check.get("status") or "UNKNOWN")
                status_counts[status] = status_counts.get(status, 0) + 1
            table["aggregateCheckSummary"] = {
                "count": len(aggregate_checks),
                "statusCounts": dict(sorted(status_counts.items())),
            }
        for column in table.get("numericColumns") or []:
            column.pop("average", None)
            column.pop("max", None)
            column.pop("min", None)
            column.pop("numericCount", None)
            column.pop("sourceRange", None)
            if column.get("numberFormats") == ["General"]:
                column.pop("numberFormats", None)
            samples = list(column.get("displaySamples") or [])
            role = str(column.get("columnRole") or "")
            if role == "MEASURE_VALUE" or role.startswith("AGGREGATE_"):
                samples = [
                    sample
                    for sample in samples
                    if sample.get("normalizedDisplay") not in (None, "")
                    or str(sample.get("displayScale") or "").startswith(
                        "PERCENT"
                    )
                ][:1]
            for sample in samples:
                if any(
                    sample.get(key) not in (None, "")
                    for key in ("sourceDisplay", "normalizedDisplay")
                ):
                    sample.pop("rawNumber", None)
            if samples:
                column["displaySamples"] = samples
            else:
                column.pop("displaySamples", None)
        for series in table.get("numericSeries") or []:
            series.pop("average", None)
            series.pop("max", None)
            series.pop("min", None)
            series.pop("valueSamples", None)
        compact_preview_rows: list[dict[str, Any]] = []
        for row in table.get("previewRows") or []:
            row_number = int(row.get("row") or 0)
            row_labels = labels_by_row.get(row_number, set())
            cells: list[dict[str, Any]] = []
            removed_count = 0
            for cell in row.get("cells") or []:
                coordinate = str(cell.get("coordinate") or "")
                match = re.match(r"[A-Z]+", coordinate)
                role = role_by_column.get(match.group(0) if match else "", "")
                is_code_owned_number = str(cell.get("kind") or "") in {
                    "NUMBER",
                    "FORMULA",
                } and (
                    role == "MEASURE_VALUE" or role.startswith("AGGREGATE_")
                )
                is_repeated_row_text = (
                    str(cell.get("kind") or "") == "TEXT"
                    and str(cell.get("value")) in row_labels
                )
                if is_code_owned_number or is_repeated_row_text:
                    removed_count += 1
                    continue
                cells.append(cell)
            if not cells:
                continue
            row["cells"] = cells
            row["omittedCellCount"] = int(
                row.get("omittedCellCount") or 0
            ) + removed_count
            compact_preview_rows.append(row)
        table["previewRows"] = compact_preview_rows
    return prompt_request


def build_table_first_prompt(request: dict[str, Any]) -> str:
    if request.get("schemaVersion") != REQUEST_SCHEMA_VERSION:
        raise TableFirstError("Expected table-first-request-v1 request.")
    prompt_request = _semantic_prompt_projection(request)
    return (
        "You classify the tables and nearby text of one captured Excel workbook.\n"
        "This is a single workbook-level semantic pass. Use only IDs and wording "
        "present in the request. Return every input table exactly once and in the "
        "same order. Some input tables may contain templateOccurrences: those are "
        "code-owned repeated copies of the representative table. Do not return "
        "the occurrence IDs; code expands the representative result afterward. "
        "Do not reproduce, correct, aggregate, or calculate numeric "
        "values: captured values, percentages, Min/Max/Average, and exact evidence "
        "are owned by deterministic code.\n"
        "For each table identify its title, whether it is COMPARISON, DESCRIPTIVE, "
        "SUPPORTING, or TEXT, the studyGroup it belongs to, source-authored group "
        "labels, metrics, and what is compared. Preserve original group wording. "
        "COMPARISON requires two or more source-authored experimental conditions "
        "and an explicit contrast. Jig positions, nozzle numbers, sample numbers, "
        "replicate numbers, dates, and measurement points are axes/strata, not "
        "comparison groups. A table with those axes but no experimental contrast "
        "is DESCRIPTIVE. SUPPORTING is only a test plan, setting, specification, "
        "or context table; a numeric result summary is DESCRIPTIVE. "
        "Do not call Normal a CONTROL; use REFERENCE only when its wording supports "
        "that role. Use UNASSESSED when a role is unclear. Similar tables with the "
        "same purpose, groups, metrics, units, and important context may share one "
        "studyGroup. Otherwise keep them separate and use relatedTableIds when useful. "
        "Ambiguity belongs in limitations and LOW/MEDIUM confidence; it never removes "
        "another usable table. axisRefs may only use that table's columnId or rowId. "
        "For a very wide raw matrix, numericSeries is the compact axis and its "
        "seriesId is preferred over enumerating sampled numericColumns. "
        "numericColumns whose columnRole begins AGGREGATE_ and "
        "aggregateCheckSummary are deterministic code results. Never emit Min, "
        "Max, Average, Avg, or Mean as "
        "a metric and never use AGGREGATE_ columnIds in metric axisRefs. Use the "
        "underlying MEASURE_VALUE columns for the real metric. A percent-formatted "
        "sample's normalizedDisplay is authoritative for display scale. "
        "formulaDerivation contains only deterministic restricted-grammar results; "
        "they may identify metrics but AI must not recalculate them. "
        "textLinks may only use supplied textId values. relatedTableIds may only use "
        "other supplied tableId values. No result is VERIFIED here.\n\n"
        "REQUEST_JSON:\n"
        + json.dumps(prompt_request, ensure_ascii=False, separators=(",", ":"))
    )


def _normalized_template_text(value: object) -> str:
    text = str(value or "").casefold()
    text = re.sub(
        r"\b\d{4}-\d{1,2}-\d{1,2}(?:t\d{1,2}:\d{2}:\d{2})?\b",
        "<date>",
        text,
    )
    text = re.sub(
        r"(?<![a-z])[-+]?\d+(?:\.\d+)?%?(?![a-z])",
        "<n>",
        text,
    )
    return re.sub(r"\s+", " ", text).strip()


def _table_template_signature(table: dict[str, Any]) -> tuple[Any, ...] | None:
    titles = tuple(
        _normalized_template_text(value)
        for value in (table.get("titleCandidates") or [])[:3]
    )
    row_labels = tuple(
        tuple(
            _normalized_template_text(value)
            for value in (row.get("labels") or [])
        )
        for row in table.get("rowLabels") or []
        if isinstance(row, dict)
    )
    if not any(titles) and not any(any(row) for row in row_labels):
        return None
    return (
        bool(int(table.get("numericCellCount") or 0)),
        titles,
        row_labels,
    )


def _compact_repeated_table_templates(
    request: dict[str, Any],
    *,
    minimum_occurrences: int = 3,
) -> tuple[dict[str, Any], dict[str, list[dict[str, Any]]]]:
    """Keep one semantic representative for repeated workbook templates."""

    tables = list(request.get("tables") or [])
    if len(tables) < minimum_occurrences:
        return request, {}
    grouped: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    unique: list[dict[str, Any]] = []
    for table in tables:
        signature = _table_template_signature(table)
        if signature is None:
            unique.append(table)
        else:
            grouped[signature].append(table)

    families: dict[str, list[dict[str, Any]]] = {}
    for members in grouped.values():
        if len(members) < minimum_occurrences:
            unique.extend(members)
            continue
        representative_id = str(members[0]["tableId"])
        families[representative_id] = members
        unique.append(members[0])
    if not families:
        return request, {}

    original_order = {
        str(table["tableId"]): index for index, table in enumerate(tables)
    }
    unique.sort(key=lambda table: original_order[str(table["tableId"])])
    prompt_request = copy.deepcopy(request)
    prompt_tables: list[dict[str, Any]] = []
    for table in unique:
        representative = copy.deepcopy(table)
        representative_id = str(representative["tableId"])
        members = families.get(representative_id)
        if members:
            representative["templateOccurrenceCount"] = len(members)
            representative["templateOccurrences"] = [
                {
                    "tableId": member["tableId"],
                    "sheet": member["sheet"],
                    "range": member["range"],
                    "titleCandidates": list(
                        (member.get("titleCandidates") or [])[:2]
                    ),
                }
                for member in members
            ]
        prompt_tables.append(representative)
    prompt_request["tables"] = prompt_tables
    prompt_request["workbook"]["sourceTableCount"] = len(tables)
    prompt_request["workbook"]["tableCount"] = len(prompt_tables)
    prompt_request["workbook"]["repeatedTemplateCount"] = len(families)
    prompt_request["workbook"]["repeatedOccurrenceCount"] = sum(
        len(members) - 1 for members in families.values()
    )
    return prompt_request, families


def table_first_prompt_stats(request: dict[str, Any]) -> dict[str, int]:
    """Measure the exact one-call prompt after template and semantic compaction."""

    prompt_request, families = _compact_repeated_table_templates(request)
    prompt_bytes = len(build_table_first_prompt(prompt_request).encode("utf-8"))
    return {
        "promptBytes": prompt_bytes,
        "sourceTableCount": len(request.get("tables") or []),
        "promptTableCount": len(prompt_request.get("tables") or []),
        "repeatedTemplateFamilyCount": len(families),
        "repeatedOccurrenceCount": sum(
            len(members) - 1 for members in families.values()
        ),
    }


def _table_axis_id_map(
    representative: dict[str, Any],
    member: dict[str, Any],
) -> dict[str, str]:
    result: dict[str, str] = {}
    for field, id_field in (
        ("rowLabels", "rowId"),
        ("numericColumns", "columnId"),
        ("numericSeries", "seriesId"),
    ):
        source_items = [
            item
            for item in representative.get(field) or []
            if isinstance(item, dict) and item.get(id_field)
        ]
        target_items = [
            item
            for item in member.get(field) or []
            if isinstance(item, dict) and item.get(id_field)
        ]
        for source_item, target_item in zip(
            source_items,
            target_items,
        ):
            result[str(source_item[id_field])] = str(target_item[id_field])
    return result


def _expand_repeated_table_analysis(
    analysis: dict[str, Any],
    *,
    request: dict[str, Any],
    prompt_request: dict[str, Any],
    families: dict[str, list[dict[str, Any]]],
) -> dict[str, Any]:
    if not families:
        return analysis
    prompt_tables = {
        str(table["tableId"]): table
        for table in prompt_request.get("tables") or []
    }
    analyzed_tables = {
        str(table["tableId"]): table for table in analysis.get("tables") or []
    }
    member_analysis: dict[str, dict[str, Any]] = {}
    for representative_id, members in families.items():
        representative_result = analyzed_tables[representative_id]
        representative_source = prompt_tables[representative_id]
        for member in members:
            member_id = str(member["tableId"])
            if member_id == representative_id:
                member_analysis[member_id] = representative_result
                continue
            clone = copy.deepcopy(representative_result)
            clone["tableId"] = member_id
            if member.get("titleCandidates"):
                clone["title"] = str(member["titleCandidates"][0])
            clone["textLinks"] = list(member.get("nearbyTextIds") or [])
            axis_id_map = _table_axis_id_map(
                representative_source,
                member,
            )
            for metric in clone.get("metrics") or []:
                metric["axisRefs"] = [
                    axis_id_map[axis_ref]
                    for axis_ref in metric.get("axisRefs") or []
                    if axis_ref in axis_id_map
                ]
            member_analysis[member_id] = clone

    expanded = copy.deepcopy(analysis)
    expanded_tables: list[dict[str, Any]] = []
    for source_table in request.get("tables") or []:
        table_id = str(source_table["tableId"])
        table_result = member_analysis.get(table_id)
        if table_result is None:
            table_result = analyzed_tables[table_id]
        expanded_tables.append(table_result)
    expanded["tables"] = expanded_tables
    return expanded


def _identity_group_label(label: str) -> bool:
    normalized = re.sub(r"\s+", " ", label).strip().casefold()
    return bool(
        re.fullmatch(r"#?\d+(?:\.\d+)?", normalized)
        or re.fullmatch(
            r"(?:position|posistion|pos|nozzle|sample|specimen|replicate)"
            r"\s*#?\d+",
            normalized,
        )
    )


def _is_error_fragment_source_table(table: dict[str, Any]) -> bool:
    bounds = table.get("bounds") or {}
    row_count = (
        int(bounds.get("maxRow") or 0)
        - int(bounds.get("minRow") or 0)
        + 1
    )
    column_count = (
        int(bounds.get("maxColumn") or 0)
        - int(bounds.get("minColumn") or 0)
        + 1
    )
    labels = [
        str(label).strip()
        for row in table.get("rowLabels") or []
        if isinstance(row, dict)
        for label in row.get("labels") or []
        if str(label).strip()
    ]
    return bool(
        int(table.get("numericCellCount") or 0)
        and row_count <= 2
        and column_count <= 3
        and not table.get("titleCandidates")
        and labels
        and all(
            re.fullmatch(
                r"#(?:REF|DIV/0|N/A|VALUE|NAME\?|NUM|NULL)!?",
                label,
                flags=re.IGNORECASE,
            )
            for label in labels
        )
    )


def _metric_name_key(
    value: object,
    *,
    term_dictionary: TermDictionaryAdapter | None = None,
) -> str:
    if term_dictionary is not None:
        return term_dictionary.semantic_key(value)
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def _column_number(label: object) -> int:
    result = 0
    for character in str(label or "").strip().upper():
        if not "A" <= character <= "Z":
            return 0
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _source_ng_rate_is_ppm(table: dict[str, Any]) -> bool:
    columns = {
        _column_number(column.get("column")): column
        for column in table.get("numericColumns") or []
        if isinstance(column, dict) and _column_number(column.get("column"))
    }
    verified = 0
    for rate_number, rate_column in columns.items():
        headers = {
            re.sub(r"\s+", " ", str(header)).strip().casefold()
            for header in rate_column.get("headerTexts") or []
        }
        if "ng rate" not in headers:
            continue
        ng_column = columns.get(rate_number - 1)
        input_column = columns.get(rate_number - 2)
        if ng_column is None or input_column is None:
            continue
        ng_headers = {
            re.sub(r"\s+", " ", str(header)).strip().casefold()
            for header in ng_column.get("headerTexts") or []
        }
        input_headers = {
            re.sub(r"\s+", " ", str(header)).strip().casefold()
            for header in input_column.get("headerTexts") or []
        }
        if "ng" not in ng_headers or "input" not in input_headers:
            continue

        def samples_by_row(column: dict[str, Any]) -> dict[int, float]:
            result: dict[int, float] = {}
            for sample in column.get("displaySamples") or []:
                match = re.fullmatch(
                    r"[A-Z]+(\d+)", str(sample.get("coordinate") or "")
                )
                value = sample.get("rawNumber")
                if (
                    match
                    and isinstance(value, (int, float))
                    and not isinstance(value, bool)
                ):
                    result[int(match.group(1))] = float(value)
            return result

        rates = samples_by_row(rate_column)
        ng_values = samples_by_row(ng_column)
        inputs = samples_by_row(input_column)
        for row in sorted(set(rates) & set(ng_values) & set(inputs)):
            if inputs[row] == 0:
                continue
            expected = ng_values[row] / inputs[row] * 1_000_000.0
            if not math.isclose(
                rates[row], expected, rel_tol=1e-9, abs_tol=1e-6
            ):
                return False
            verified += 1
    return verified >= 2


def _is_reliability_raw_data(
    table: dict[str, Any],
    *,
    source_file_name: object,
    analysis_title: object,
) -> bool:
    if "reliability" not in str(source_file_name or "").casefold():
        return False
    source_text = " ".join(
        [
            str(table.get("sheet") or ""),
            *(str(value) for value in table.get("titleCandidates") or []),
            str(analysis_title or ""),
        ]
    ).casefold()
    return "raw data" in source_text


def normalize_table_first_analysis(
    result: dict[str, Any],
    *,
    request: dict[str, Any],
) -> dict[str, Any]:
    """Apply cheap guards that remove known non-semantic AI overreach."""

    normalized = copy.deepcopy(result)
    term_dictionary = TermDictionaryAdapter.from_snapshot(
        request.get("codeOwnedTermDictionary")
    )
    request_tables = {
        str(table["tableId"]): table for table in request.get("tables") or []
    }
    formula_status = str(
        (request.get("formulaDerivation") or {}).get("status") or ""
    )
    partial_formula_limitation = (
        "Formula derivation was partial: supported formulas were derived, "
        "while unsupported formulas remain unresolved."
    )
    for table in normalized.get("tables") or []:
        if not isinstance(table, dict):
            continue
        source_table = request_tables.get(str(table.get("tableId") or ""))
        if source_table is None:
            continue
        groups: list[dict[str, Any]] = []
        seen_group_labels: set[str] = set()
        for group in table.get("groups") or []:
            if not isinstance(group, dict):
                continue
            label = str(group.get("label") or "")
            normalized_label = term_dictionary.semantic_key(label)
            if (
                _identity_group_label(label)
                or normalized_label in seen_group_labels
            ):
                continue
            seen_group_labels.add(normalized_label)
            groups.append(group)
        table["groups"] = groups
        group_labels = {str(group.get("label") or "") for group in groups}
        table["comparisonRelations"] = [
            relation
            for relation in table.get("comparisonRelations") or []
            if isinstance(relation, dict)
            and str(relation.get("leftGroup") or "") in group_labels
            and str(relation.get("rightGroup") or "") in group_labels
            and str(relation.get("leftGroup") or "")
            != str(relation.get("rightGroup") or "")
        ]

        aggregate_axis_refs = {
            str(column["columnId"])
            for column in source_table.get("numericColumns") or []
            if str(column.get("columnRole") or "").startswith("AGGREGATE_")
        } | {
            str(series["seriesId"])
            for series in source_table.get("numericSeries") or []
            if str(series.get("columnRole") or "").startswith("AGGREGATE_")
        }
        metrics: list[dict[str, Any]] = []
        for metric in table.get("metrics") or []:
            if not isinstance(metric, dict):
                continue
            if re.fullmatch(
                r"\s*(?:min|minimum|max|maximum|avg|average|mean)\s*",
                str(metric.get("name") or ""),
                flags=re.IGNORECASE,
            ):
                continue
            metric["axisRefs"] = [
                axis_ref
                for axis_ref in metric.get("axisRefs") or []
                if str(axis_ref) not in aggregate_axis_refs
            ]
            metrics.append(metric)
        metric_hints = {
            _metric_name_key(
                hint.get("name"),
                term_dictionary=term_dictionary,
            ): hint
            for hint in source_table.get("metricHints") or []
            if isinstance(hint, dict)
            and str(hint.get("name") or "").strip()
        }
        existing_metric_names: set[str] = set()
        deduplicated_metrics: list[dict[str, Any]] = []
        metric_index_by_name: dict[str, int] = {}
        for metric in metrics:
            normalized_name = _metric_name_key(
                metric.get("name"),
                term_dictionary=term_dictionary,
            )
            hint = metric_hints.get(normalized_name)
            if metric_hints and hint is None:
                continue
            if hint is not None:
                metric["name"] = re.sub(
                    r"\s+", " ", str(hint.get("name") or "")
                ).strip()
                metric["axisRefs"] = list(hint.get("axisRefs") or [])
            previous_index = metric_index_by_name.get(normalized_name)
            if previous_index is not None:
                previous = deduplicated_metrics[previous_index]
                if not str(previous.get("unit") or "").strip() and str(
                    metric.get("unit") or ""
                ).strip():
                    deduplicated_metrics[previous_index] = metric
                continue
            metric_index_by_name[normalized_name] = len(deduplicated_metrics)
            deduplicated_metrics.append(metric)
            existing_metric_names.add(normalized_name)
        metrics = deduplicated_metrics
        for normalized_name, hint in metric_hints.items():
            if normalized_name in existing_metric_names:
                continue
            name = re.sub(r"\s+", " ", str(hint.get("name") or "")).strip()
            metrics.append(
                {
                    "name": name,
                    "unit": "",
                    "axisRefs": list(hint.get("axisRefs") or []),
                }
            )
            existing_metric_names.add(normalized_name)
        table["metrics"] = metrics

        if _source_ng_rate_is_ppm(source_table):
            for metric in table["metrics"]:
                if _metric_name_key(
                    metric.get("name"),
                    term_dictionary=term_dictionary,
                ) == "ng rate":
                    metric["unit"] = "PPM"
            table["limitations"] = [
                str(limitation)
                for limitation in table.get("limitations") or []
                if not re.search(
                    r"ng rate.*(?:scale|non-percent|unit)|"
                    r"(?:scale|non-percent|unit).*ng rate",
                    str(limitation),
                    flags=re.IGNORECASE,
                )
            ]
            if table.get("confidence") == "LOW":
                table["confidence"] = "MEDIUM"

        if formula_status == "PARTIALLY_DERIVED":
            limitations: list[str] = []
            replaced_formula_limitation = False
            for limitation in table.get("limitations") or []:
                value = str(limitation)
                if re.search(
                    r"formula derivation was skipped|formula cells are not recalculated",
                    value,
                    flags=re.IGNORECASE,
                ):
                    if not replaced_formula_limitation:
                        limitations.append(partial_formula_limitation)
                        replaced_formula_limitation = True
                    continue
                limitations.append(value)
            table["limitations"] = limitations

        if _is_error_fragment_source_table(source_table):
            table["type"] = "TEXT"
            table["groups"] = []
            table["metrics"] = []
            table["comparisonRelations"] = []
            table["confidence"] = "HIGH"
            table["limitations"] = [
                "Spreadsheet error fragment excluded from semantic result analysis."
            ]
            continue

        if (
            table.get("confidence") == "LOW"
            and _is_reliability_raw_data(
                source_table,
                source_file_name=(request.get("source") or {}).get("fileName"),
                analysis_title=table.get("title"),
            )
        ):
            table["confidence"] = "MEDIUM"

        if table.get("type") != "COMPARISON":
            table["comparisonRelations"] = []
        elif not table["comparisonRelations"]:
            table["type"] = "DESCRIPTIVE"
        source_title = " ".join(
            str(value) for value in source_table.get("titleCandidates") or []
        ).casefold()
        if (
            table.get("type") == "SUPPORTING"
            and table.get("metrics")
            and any(term in source_title for term in ("result", "결과"))
        ):
            table["type"] = "DESCRIPTIVE"
    return normalized


def _require_string(value: object, path: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value.strip()):
        raise TableFirstError(f"{path} must be a non-empty string.")
    return value


def _require_string_list(value: object, path: str) -> list[str]:
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        raise TableFirstError(f"{path} must be a string list.")
    return value


def validate_table_first_analysis(
    result: dict[str, Any],
    *,
    request: dict[str, Any],
) -> dict[str, Any]:
    """Validate identities and all request-bound references in an AI response."""

    if not isinstance(result, dict):
        raise TableFirstError("Analysis must be a JSON object.")
    if result.get("schemaVersion") != ANALYSIS_SCHEMA_VERSION:
        raise TableFirstError("Invalid analysis schemaVersion.")
    if result.get("promptVersion") != PROMPT_VERSION:
        raise TableFirstError("Invalid analysis promptVersion.")
    if result.get("requestId") != request.get("requestId"):
        raise TableFirstError("Analysis requestId does not match.")
    if result.get("revisionUid") != (request.get("source") or {}).get("revisionUid"):
        raise TableFirstError("Analysis revisionUid does not match.")
    if result.get("status") not in {"ANALYZED", "NEEDS_REVIEW", "NO_TABLES"}:
        raise TableFirstError("Invalid analysis status.")
    _require_string(result.get("workbookSummary"), "workbookSummary", allow_empty=True)
    _require_string_list(result.get("notes"), "notes")

    request_tables = request.get("tables") or []
    response_tables = result.get("tables")
    if not isinstance(response_tables, list):
        raise TableFirstError("tables must be a list.")
    expected_ids = [str(table["tableId"]) for table in request_tables]
    actual_ids = [
        _require_string(table.get("tableId"), f"tables[{index}].tableId")
        if isinstance(table, dict)
        else ""
        for index, table in enumerate(response_tables)
    ]
    if actual_ids != expected_ids:
        raise TableFirstError(
            "Analysis must return every input table exactly once in input order."
        )
    if not expected_ids and result["status"] != "NO_TABLES":
        raise TableFirstError("An empty request must have NO_TABLES status.")
    if expected_ids and result["status"] == "NO_TABLES":
        raise TableFirstError("NO_TABLES is invalid when input tables exist.")

    text_ids = {str(item["textId"]) for item in request.get("textBlocks") or []}
    table_ids = set(expected_ids)
    for index, (table, source_table) in enumerate(
        zip(response_tables, request_tables, strict=True)
    ):
        prefix = f"tables[{index}]"
        _require_string(table.get("title"), f"{prefix}.title", allow_empty=True)
        if table.get("type") not in TABLE_TYPES:
            raise TableFirstError(f"{prefix}.type is invalid.")
        _require_string(table.get("studyGroup"), f"{prefix}.studyGroup")
        if table.get("confidence") not in CONFIDENCE_LEVELS:
            raise TableFirstError(f"{prefix}.confidence is invalid.")
        for field in ("textLinks", "relatedTableIds", "limitations"):
            _require_string_list(table.get(field), f"{prefix}.{field}")
        unknown_text = sorted(set(table["textLinks"]) - text_ids)
        if unknown_text:
            raise TableFirstError(
                f"{prefix}.textLinks contains unknown ids: {', '.join(unknown_text)}"
            )
        unknown_related = sorted(set(table["relatedTableIds"]) - table_ids)
        if unknown_related or table["tableId"] in table["relatedTableIds"]:
            raise TableFirstError(f"{prefix}.relatedTableIds is invalid.")

        groups = table.get("groups")
        if not isinstance(groups, list):
            raise TableFirstError(f"{prefix}.groups must be a list.")
        for group_index, group in enumerate(groups):
            group_path = f"{prefix}.groups[{group_index}]"
            if not isinstance(group, dict):
                raise TableFirstError(f"{group_path} must be an object.")
            _require_string(group.get("label"), f"{group_path}.label")
            if group.get("role") not in GROUP_ROLES:
                raise TableFirstError(f"{group_path}.role is invalid.")
            _require_string(group.get("basis"), f"{group_path}.basis", allow_empty=True)

        allowed_axis_refs = {
            str(column["columnId"])
            for column in source_table.get("numericColumns") or []
            if not str(column.get("columnRole") or "").startswith(
                "AGGREGATE_"
            )
        } | {
            str(row["rowId"])
            for row in source_table.get("rowLabels") or []
        } | {
            str(series["seriesId"])
            for series in source_table.get("numericSeries") or []
        }
        metrics = table.get("metrics")
        if not isinstance(metrics, list):
            raise TableFirstError(f"{prefix}.metrics must be a list.")
        for metric_index, metric in enumerate(metrics):
            metric_path = f"{prefix}.metrics[{metric_index}]"
            if not isinstance(metric, dict):
                raise TableFirstError(f"{metric_path} must be an object.")
            _require_string(metric.get("name"), f"{metric_path}.name")
            _require_string(metric.get("unit"), f"{metric_path}.unit", allow_empty=True)
            axis_refs = _require_string_list(
                metric.get("axisRefs"),
                f"{metric_path}.axisRefs",
            )
            unknown_refs = sorted(set(axis_refs) - allowed_axis_refs)
            if unknown_refs:
                raise TableFirstError(
                    f"{metric_path}.axisRefs contains unknown ids: "
                    + ", ".join(unknown_refs)
                )

        relations = table.get("comparisonRelations")
        if not isinstance(relations, list):
            raise TableFirstError(f"{prefix}.comparisonRelations must be a list.")
        if table.get("type") == "COMPARISON" and not relations:
            raise TableFirstError(
                f"{prefix} COMPARISON requires a comparison relation."
            )
        if table.get("type") != "COMPARISON" and relations:
            raise TableFirstError(
                f"{prefix} non-comparison table cannot contain relations."
            )
        group_labels = {str(group["label"]) for group in groups}
        for relation_index, relation in enumerate(relations):
            relation_path = f"{prefix}.comparisonRelations[{relation_index}]"
            if not isinstance(relation, dict):
                raise TableFirstError(f"{relation_path} must be an object.")
            left = _require_string(
                relation.get("leftGroup"),
                f"{relation_path}.leftGroup",
            )
            right = _require_string(
                relation.get("rightGroup"),
                f"{relation_path}.rightGroup",
            )
            _require_string(
                relation.get("basis"),
                f"{relation_path}.basis",
                allow_empty=True,
            )
            if left not in group_labels or right not in group_labels or left == right:
                raise TableFirstError(
                    f"{relation_path} must reference two distinct declared groups."
                )
    return result


RunCommand = Callable[..., subprocess.CompletedProcess[str]]


def _codex_command(command: Sequence[str] | None) -> list[str]:
    if command:
        return list(command)
    executable = shutil.which("codex.cmd" if os.name == "nt" else "codex")
    if not executable:
        raise TableFirstError("Codex CLI executable was not found on PATH.")
    return [executable]


def run_codex_table_first_analysis(
    *,
    request: dict[str, Any],
    output_path: str | Path,
    model: str | None = None,
    reasoning_effort: str | None = "low",
    codex_command: Sequence[str] | None = None,
    timeout_seconds: int = 600,
    run_command: RunCommand = subprocess.run,
) -> dict[str, Any]:
    """Run exactly one read-only Codex call for one workbook request."""

    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    prompt_request, repeated_families = _compact_repeated_table_templates(
        request
    )
    prompt = build_table_first_prompt(prompt_request)
    with tempfile.TemporaryDirectory(prefix="table-first-analysis-") as temp_dir:
        schema_path = Path(temp_dir) / "table-first.schema.json"
        last_message_path = Path(temp_dir) / "last-message.json"
        schema_path.write_text(
            json.dumps(table_first_output_schema(), ensure_ascii=False),
            encoding="utf-8",
        )
        command = [
            *_codex_command(codex_command),
            "exec",
            "--ephemeral",
            "--sandbox",
            "read-only",
            "--output-schema",
            str(schema_path),
            "--output-last-message",
            str(last_message_path),
        ]
        if reasoning_effort:
            command.extend(["-c", f'model_reasoning_effort="{reasoning_effort}"'])
        if model:
            command.extend(["--model", model])
        command.append("-")
        completed = run_command(
            command,
            input=prompt,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            timeout=timeout_seconds,
            check=False,
        )
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout or "").strip()
            raise TableFirstError(
                "Codex table-first analysis failed with exit code "
                f"{completed.returncode}: {detail[-2000:]}"
            )
        if not last_message_path.is_file():
            raise TableFirstError("Codex did not produce an output message.")
        try:
            result = json.loads(last_message_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise TableFirstError("Codex output is not valid JSON.") from exc
    guarded = normalize_table_first_analysis(
        result,
        request=prompt_request,
    )
    prompt_validated = validate_table_first_analysis(
        guarded,
        request=prompt_request,
    )
    expanded = _expand_repeated_table_analysis(
        prompt_validated,
        request=request,
        prompt_request=prompt_request,
        families=repeated_families,
    )
    normalized = normalize_table_first_analysis(expanded, request=request)
    validated = validate_table_first_analysis(normalized, request=request)
    target.write_text(
        json.dumps(validated, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return validated


def project_table_first_analysis(
    request: dict[str, Any],
    analysis: dict[str, Any],
) -> dict[str, Any]:
    """Create a deterministic, non-approved study/evidence projection."""

    validated = validate_table_first_analysis(analysis, request=request)
    source_tables = {
        str(table["tableId"]): table for table in request.get("tables") or []
    }
    studies: dict[str, dict[str, Any]] = {}
    for table_analysis in validated["tables"]:
        table_id = str(table_analysis["tableId"])
        source_table = source_tables[table_id]
        study_group = str(table_analysis["studyGroup"])
        study = studies.setdefault(
            study_group,
            {
                "studyGroup": study_group,
                "verificationStatus": "NEEDS_REVIEW",
                "titles": [],
                "tableTypes": [],
                "groups": [],
                "metrics": [],
                "comparisonRelations": [],
                "evidence": [],
                "deterministicNumericFacts": [],
                "deterministicNumericSeries": [],
                "deterministicAggregateChecks": [],
                "limitations": [],
            },
        )
        for field, values in (
            ("titles", [table_analysis["title"]]),
            ("tableTypes", [table_analysis["type"]]),
            ("groups", table_analysis["groups"]),
            ("metrics", table_analysis["metrics"]),
            ("comparisonRelations", table_analysis["comparisonRelations"]),
            ("limitations", table_analysis["limitations"]),
        ):
            for value in values:
                if value not in study[field]:
                    study[field].append(value)
        study["evidence"].append(
            {
                "tableId": table_id,
                "sheet": source_table["sheet"],
                "range": source_table["range"],
            }
        )
        for column in source_table.get("numericColumns") or []:
            study["deterministicNumericFacts"].append(
                {
                    "tableId": table_id,
                    "columnId": column["columnId"],
                    "sourceRange": column["sourceRange"],
                    "columnRole": column["columnRole"],
                    "numericCount": column["numericCount"],
                    "min": column["min"],
                    "max": column["max"],
                    "average": column["average"],
                    "calculationAuthority": "CODE_FROM_CAPTURED_RAW_VALUES",
                }
            )
        study["deterministicNumericSeries"].extend(
            source_table.get("numericSeries") or []
        )
        study["deterministicAggregateChecks"].extend(
            source_table.get("aggregateChecks") or []
        )
        if table_analysis["confidence"] != "HIGH":
            limitation = (
                f"{table_id} semantic confidence is "
                f"{table_analysis['confidence']}."
            )
            if limitation not in study["limitations"]:
                study["limitations"].append(limitation)
    return {
        "schemaVersion": PROJECTION_SCHEMA_VERSION,
        "source": request["source"],
        "requestId": request["requestId"],
        "analysisStatus": validated["status"],
        "verificationStatus": (
            "NEEDS_REVIEW" if validated["tables"] else "NO_TABLES"
        ),
        "queryEligibility": "NOT_ELIGIBLE_UNTIL_CANONICAL_REVIEW",
        "studies": list(studies.values()),
        "textBlocks": request.get("textBlocks") or [],
    }


def table_first_json_bytes(value: dict[str, Any]) -> bytes:
    return (
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n"
    ).encode("utf-8")


__all__ = [
    "ANALYSIS_SCHEMA_VERSION",
    "BUILDER_VERSION",
    "PROMPT_VERSION",
    "PROJECTION_SCHEMA_VERSION",
    "REQUEST_SCHEMA_VERSION",
    "TableFirstError",
    "build_table_first_prompt",
    "build_table_first_request",
    "normalize_table_first_analysis",
    "project_table_first_analysis",
    "run_codex_table_first_analysis",
    "table_first_prompt_stats",
    "table_first_json_bytes",
    "table_first_output_schema",
    "validate_table_first_analysis",
]
