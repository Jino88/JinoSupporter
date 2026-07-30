from __future__ import annotations

import json
import math
import re
import sqlite3
import unicodedata
from collections import Counter
from collections.abc import Callable
from typing import Any

from inference_data_ai_formula_derivation import FormulaOverlayLookup
from inference_data_ai_schema import (
    find_concept_id,
    normalize_key_part,
    public_id,
    record_schema_candidate,
    resolve_unit_id,
    stable_uid,
)
from inference_data_ai_study_contract import validate_study_manifest


class AnalysisQuarantineError(ValueError):
    """Raised when a canonical analysis is not safe to quarantine."""


_QUARANTINE_VALIDATOR_NAME = "canonical-analysis-quarantine"
_QUARANTINE_VALIDATOR_VERSION = "1"


def _json(value: object, fallback: object) -> str:
    return json.dumps(fallback if value is None else value, ensure_ascii=False, separators=(",", ":"))


def _column_number(label: str) -> int:
    value = 0
    for char in label.upper():
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def _column_label(number: int) -> str:
    result = ""
    value = number
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def parse_a1_range(address: str) -> tuple[int, int, int, int]:
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,4})\$?([1-9]\d*)(?::\$?([A-Za-z]{1,4})\$?([1-9]\d*))?",
        address.strip(),
    )
    if not match:
        raise ValueError(f"invalid A1 range: {address}")
    start_col = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_col = _column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_row < start_row or end_col < start_col:
        raise ValueError(f"reversed A1 range: {address}")
    return start_row, start_col, end_row, end_col


def resolve_manifest_revision(conn: sqlite3.Connection, source: dict[str, Any]) -> sqlite3.Row:
    dataset = str(source["dataset"]).strip()
    source_path = str(source["sourcePath"]).strip()
    revision_uid = str(source.get("revisionUid") or "").strip()
    content_sha256 = str(source.get("contentSha256") or "").strip().lower()
    fingerprint = str(source.get("fingerprint") or "").strip()
    sql = """
        SELECT sr.*, sd.dataset, sd.source_path, sd.document_id
        FROM source_revisions sr
        JOIN source_documents sd ON sd.document_id=sr.document_id
        WHERE sd.dataset=? AND sd.source_path=?
    """
    params: list[object] = [dataset, source_path]
    if revision_uid:
        sql += " AND sr.revision_uid=?"
        params.append(revision_uid)
    elif content_sha256:
        sql += " AND LOWER(sr.content_sha256)=?"
        params.append(content_sha256)
    else:
        sql += " AND sr.source_fingerprint=?"
        params.append(fingerprint)
    sql += " ORDER BY sr.is_current DESC, sr.revision_id DESC LIMIT 1"
    row = conn.execute(sql, params).fetchone()
    if not row:
        raise ValueError("manifest source revision is not present in the source database")
    return row


def make_database_evidence_checker(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
) -> Callable[[dict[str, Any]], None]:
    legacy_workbook_id = revision["legacy_workbook_id"]

    def check(item: dict[str, Any]) -> None:
        sheet_name = str(item["sheet"])
        start_row, start_col, end_row, end_col = parse_a1_range(str(item["range"]))
        if legacy_workbook_id is not None:
            sheet = conn.execute(
                """
                SELECT used_top, used_left, used_bottom, used_right
                FROM worksheets
                WHERE workbook_id=? AND sheet_name=?
                LIMIT 1
                """,
                (int(legacy_workbook_id), sheet_name),
            ).fetchone()
            if not sheet:
                raise ValueError(f"evidence sheet is not present in source revision: {sheet_name}")
            if (
                start_row < int(sheet["used_top"])
                or start_col < int(sheet["used_left"])
                or end_row > int(sheet["used_bottom"])
                or end_col > int(sheet["used_right"])
            ):
                raise ValueError(f"evidence range is outside source bounds: {sheet_name}!{item['range']}")
            expected = (end_row - start_row + 1) * (end_col - start_col + 1)
            actual = int(
                conn.execute(
                    """
                    SELECT COUNT(*)
                    FROM grid_sheet_cells
                    WHERE workbook_id=? AND sheet_name=?
                      AND row_number BETWEEN ? AND ?
                      AND col_number BETWEEN ? AND ?
                    """,
                    (
                        int(legacy_workbook_id),
                        sheet_name,
                        start_row,
                        end_row,
                        start_col,
                        end_col,
                    ),
                ).fetchone()[0]
            )
            if actual != expected:
                raise ValueError(
                    f"evidence range is not fully represented in source grid: "
                    f"{sheet_name}!{item['range']} ({actual}/{expected})"
                )
            return

        capture_v2_revision_id = revision["capture_v2_revision_id"]
        if capture_v2_revision_id is not None and _table_exists(conn, "capture_v2_sheets"):
            sheet = conn.execute(
                """
                SELECT sheet_id, used_bounds_json, content_bounds_json
                FROM capture_v2_sheets
                WHERE revision_id=? AND title=?
                LIMIT 1
                """,
                (int(capture_v2_revision_id), sheet_name),
            ).fetchone()
            if not sheet:
                raise ValueError(f"evidence sheet is not present in Capture v2: {sheet_name}")
            bounds_text = sheet["content_bounds_json"] or sheet["used_bounds_json"]
            if not bounds_text:
                raise ValueError(f"evidence sheet has no captured cells: {sheet_name}")
            bounds = json.loads(str(bounds_text))
            if (
                start_row < int(bounds["minRow"])
                or start_col < int(bounds["minColumn"])
                or end_row > int(bounds["maxRow"])
                or end_col > int(bounds["maxColumn"])
            ):
                raise ValueError(f"evidence range is outside Capture v2 bounds: {sheet_name}!{item['range']}")
            represented = int(
                conn.execute(
                    """
                    SELECT COUNT(*)
                    FROM capture_v2_cells
                    WHERE sheet_id=?
                      AND row_index BETWEEN ? AND ?
                      AND column_index BETWEEN ? AND ?
                    """,
                    (
                        int(sheet["sheet_id"]),
                        start_row,
                        end_row,
                        start_col,
                        end_col,
                    ),
                ).fetchone()[0]
            )
            if represented == 0:
                raise ValueError(f"evidence range contains no captured source cells: {sheet_name}!{item['range']}")
            return
        raise ValueError("source revision has no cell-grid adapter for evidence validation")

    return check


def _table_exists(conn: sqlite3.Connection, table: str) -> bool:
    return (
        conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1", (table,)).fetchone()
        is not None
    )


def _portable_number(value: object) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None
    if isinstance(value, str) and re.fullmatch(
        r"\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?\s*",
        value,
    ):
        try:
            number = float(value)
        except ValueError:
            return None
        return number if math.isfinite(number) else None
    if isinstance(value, dict) and str(value.get("type") or "") in {
        "decimal",
        "float",
    }:
        try:
            number = float(value.get("value"))
        except (TypeError, ValueError):
            return None
        return number if math.isfinite(number) else None
    return None


def _portable_axis_number(value: object, unit: object) -> float | None:
    number = _portable_number(value)
    if number is not None:
        return number
    if not isinstance(value, str):
        return None
    unit_text = str(unit or "").strip()
    suffix = rf"\s*{re.escape(unit_text)}" if unit_text else r"\s*"
    match = re.fullmatch(
        rf"\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?){suffix}\s*",
        value,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    try:
        result = float(match.group(1))
    except ValueError:
        return None
    return result if math.isfinite(result) else None


def _portable_identity(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        text = value.strip()
        return text or None
    if isinstance(value, (int, float)):
        number = _portable_number(value)
        if number is None:
            return None
        return format(number, ".15g")
    if isinstance(value, dict):
        if "value" in value:
            text = str(value["value"]).strip()
            return text or None
        if value.get("type") == "timedelta" and "seconds" in value:
            return str(value["seconds"])
    return None


def _formula_lookup_for_revision(
    revision: sqlite3.Row,
    formula_overlay: dict[str, Any] | FormulaOverlayLookup | None,
) -> FormulaOverlayLookup | None:
    """Bind one derived overlay to the exact immutable Capture revision."""

    if formula_overlay is None:
        return None
    if isinstance(formula_overlay, FormulaOverlayLookup):
        source = formula_overlay.overlay.get("source", {})
        if (
            str(source.get("revisionUid") or "")
            != str(revision["revision_uid"])
            or str(source.get("contentSha256") or "").lower()
            != str(revision["content_sha256"]).lower()
        ):
            raise ValueError(
                "Formula overlay does not match the canonical source revision"
            )
        return formula_overlay
    return FormulaOverlayLookup(
        formula_overlay,
        revision_uid=str(revision["revision_uid"]),
        content_sha256=str(revision["content_sha256"]),
    )


def _derived_formula_source_json(
    formula_lookup: FormulaOverlayLookup | None,
    *,
    sheet_name: str,
    coordinate: object,
    formula_text: str,
) -> str | None:
    """Return provenance-bearing JSON, never a fabricated Capture cache."""

    if formula_lookup is None or not formula_text:
        return None
    entry = formula_lookup.entry(
        sheet=sheet_name,
        coordinate=str(coordinate),
        formula=formula_text,
    )
    if entry is None:
        return None
    if entry["status"] == "NUMERIC":
        value: object = entry["numericValue"]
        value_type = "float"
    elif entry["status"] == "ERROR" and entry["error"] == "#DIV/0!":
        value = "#DIV/0!"
        value_type = "formulaError"
    else:
        raise ValueError(
            f"Unsupported deterministic formula result at "
            f"{sheet_name}!{coordinate}"
        )
    return json.dumps(
        {
            "type": value_type,
            "value": value,
            "derivation": {
                "schemaVersion": formula_lookup.overlay["schemaVersion"],
                "evaluatorVersion": entry["evaluatorVersion"],
                "sourceCellKey": entry["sourceCellKey"],
                "formula": entry["formula"],
                "provenanceSha256": entry["provenanceSha256"],
                "overlaySha256": formula_lookup.overlay_sha256,
            },
        },
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def _measurement_series_evidence(
    series: dict[str, Any],
) -> list[dict[str, Any]]:
    return [
        {
            "sheet": series["sheet"],
            "range": series["headerRange"],
            "role": "MEASUREMENT_HEADER",
        },
        {
            "sheet": series["sheet"],
            "range": series["valueRange"],
            "role": "MEASUREMENT_VALUES",
        },
        {
            "sheet": series["sheet"],
            "range": series["rowIdentityRange"],
            "role": "ROW_IDENTITY",
        },
    ]


def _capture_cells_for_range(
    conn: sqlite3.Connection,
    *,
    capture_revision_id: int,
    sheet_name: str,
    address: str,
    path: str,
    resolve_merged_label: bool = False,
) -> dict[tuple[int, int], sqlite3.Row]:
    start_row, start_col, end_row, end_col = parse_a1_range(address)
    sheet = conn.execute(
        """
        SELECT sheet_id
        FROM capture_v2_sheets
        WHERE revision_id=? AND title=?
        LIMIT 1
        """,
        (capture_revision_id, sheet_name),
    ).fetchone()
    if not sheet:
        raise ValueError(
            f"{path} sheet is not present in current Capture v2: {sheet_name}"
        )
    rows = conn.execute(
        """
        SELECT
            row_index, column_index, coordinate, raw_value_json,
            formula_text, cached_value_json, display_value_json,
            number_format, merge_range, merge_role
        FROM capture_v2_cells
        WHERE sheet_id=?
          AND row_index BETWEEN ? AND ?
          AND column_index BETWEEN ? AND ?
        ORDER BY row_index, column_index
        """,
        (
            int(sheet["sheet_id"]),
            start_row,
            end_row,
            start_col,
            end_col,
        ),
    ).fetchall()
    by_coordinate = {
        (int(row["row_index"]), int(row["column_index"])): row
        for row in rows
    }
    for row_index in range(start_row, end_row + 1):
        for column_index in range(start_col, end_col + 1):
            key = (row_index, column_index)
            if key not in by_coordinate:
                coordinate = f"{_column_label(column_index)}{row_index}"
                raise ValueError(
                    f"{path} is missing captured source cell "
                    f"{sheet_name}!{coordinate}"
                )
            cell = by_coordinate[key]
            if not resolve_merged_label:
                continue
            formula_text = str(cell["formula_text"] or "")
            source_json = (
                cell["cached_value_json"]
                if formula_text
                else cell["raw_value_json"]
            )
            if (
                source_json is not None
                or str(cell["merge_role"] or "").casefold() != "covered"
                or not str(cell["merge_range"] or "").strip()
            ):
                continue
            anchor = conn.execute(
                """
                SELECT
                    c.row_index, c.column_index, c.coordinate,
                    c.raw_value_json, c.formula_text,
                    c.cached_value_json, c.display_value_json,
                    c.number_format, c.merge_range, c.merge_role
                FROM capture_v2_merged_ranges mr
                JOIN capture_v2_cells c
                  ON c.sheet_id=mr.sheet_id
                 AND c.coordinate=mr.anchor_coordinate
                WHERE mr.sheet_id=?
                  AND mr.address=?
                  AND c.merge_range=mr.address
                  AND c.merge_role='anchor'
                LIMIT 1
                """,
                (
                    int(sheet["sheet_id"]),
                    str(cell["merge_range"]),
                ),
            ).fetchone()
            if anchor is None:
                raise ValueError(
                    f"{path} covered merged source cell "
                    f"{cell['coordinate']} has no exact same-sheet anchor"
                )
            anchor_formula = str(anchor["formula_text"] or "")
            anchor_source_json = (
                anchor["cached_value_json"]
                if anchor_formula
                else anchor["raw_value_json"]
            )
            if anchor_source_json is None:
                raise ValueError(
                    f"{path} merged anchor {anchor['coordinate']} has no "
                    "usable label value"
                )
            by_coordinate[key] = anchor
    return by_coordinate


def _custom_number_format_identity(
    value: object,
    number_format: object,
) -> str | None:
    """Render source-authored literal labels around a numeric identity.

    This intentionally supports only the conservative identity-label subset
    used by source workbooks: quoted or escaped literals surrounding one
    numeric placeholder run. Ordinary numeric/date/percent formats remain on
    the existing raw-value path.
    """

    number = _portable_number(value)
    format_text = str(number_format or "")
    if number is None or '"' not in format_text:
        return None
    section = format_text.split(";", 1)[0]
    if not re.search(r"[0#?]", section):
        return None

    rendered: list[str] = []
    literal_found = False
    placeholder_written = False
    index = 0
    while index < len(section):
        char = section[index]
        if char == '"':
            end = index + 1
            literal: list[str] = []
            while end < len(section):
                if section[end] == '"':
                    break
                literal.append(section[end])
                end += 1
            if end >= len(section):
                return None
            literal_text = "".join(literal)
            rendered.append(literal_text)
            literal_found = literal_found or bool(literal_text)
            index = end + 1
            continue
        if char == "\\" and index + 1 < len(section):
            rendered.append(section[index + 1])
            literal_found = True
            index += 2
            continue
        if char in {"_", "*"} and index + 1 < len(section):
            index += 2
            continue
        if char == "[":
            end = section.find("]", index + 1)
            if end < 0:
                return None
            index = end + 1
            continue
        if char in "0#?":
            while index < len(section) and section[index] in "0#?.,": 
                index += 1
            if not placeholder_written:
                rendered.append(_portable_identity(number) or str(number))
                placeholder_written = True
            continue
        if char.isspace():
            rendered.append(char)
        index += 1

    if not literal_found or not placeholder_written:
        return None
    result = "".join(rendered).strip()
    return result or None


def _capture_cell_payload(
    cell: sqlite3.Row,
    *,
    path: str,
    numeric: bool,
    numeric_unit: object = "",
    sheet_name: str = "",
    formula_lookup: FormulaOverlayLookup | None = None,
) -> tuple[object, str]:
    formula_text = str(cell["formula_text"] or "")
    source_json = (
        cell["cached_value_json"]
        if formula_text
        else cell["raw_value_json"]
    )
    if formula_text and source_json is None:
        source_json = _derived_formula_source_json(
            formula_lookup,
            sheet_name=sheet_name,
            coordinate=cell["coordinate"],
            formula_text=formula_text,
        )
    if source_json is None:
        raise ValueError(
            f"{path} source cell {cell['coordinate']} has no "
            f"{'cached formula value' if formula_text else 'value'}"
        )
    try:
        parsed = json.loads(str(source_json))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"{path} source cell {cell['coordinate']} has invalid Capture v2 JSON"
        ) from exc
    if numeric:
        number = _portable_number(parsed)
        if number is None:
            raise ValueError(
                f"{path} source cell {cell['coordinate']} must be numeric"
            )
        if (
            "%" in str(cell["number_format"] or "")
            and str(numeric_unit or "").strip().casefold()
            in {"%", "percent", "percentage", "pct"}
        ):
            number *= 100.0
        return number, str(source_json)
    display_json = cell["display_value_json"]
    if display_json is not None:
        try:
            display_value = json.loads(str(display_json))
        except json.JSONDecodeError:
            display_value = parsed
    else:
        display_value = parsed
    formatted_identity = _custom_number_format_identity(
        parsed,
        cell["number_format"],
    )
    if formatted_identity is not None:
        display_value = formatted_identity
    identity = _portable_identity(display_value)
    if identity is None:
        raise ValueError(
            f"{path} source cell {cell['coordinate']} has no usable identity"
        )
    return identity, str(source_json)


def _expand_measurement_series(
    conn: sqlite3.Connection,
    *,
    revision: sqlite3.Row,
    series: dict[str, Any],
    series_uid: str,
    path: str,
    formula_lookup: FormulaOverlayLookup | None = None,
) -> list[dict[str, Any]]:
    if not int(revision["is_current"]):
        raise ValueError(
            f"{path} requires the current canonical source revision"
        )
    capture_revision_id = revision["capture_v2_revision_id"]
    if capture_revision_id is None or not _table_exists(
        conn,
        "capture_v2_cells",
    ):
        raise ValueError(f"{path} requires current Capture v2 cells")

    sheet_name = str(series["sheet"])
    header_range = parse_a1_range(str(series["headerRange"]))
    value_range = parse_a1_range(str(series["valueRange"]))
    identity_range = parse_a1_range(str(series["rowIdentityRange"]))
    header_cells = _capture_cells_for_range(
        conn,
        capture_revision_id=int(capture_revision_id),
        sheet_name=sheet_name,
        address=str(series["headerRange"]),
        path=f"{path}.headerRange",
        resolve_merged_label=True,
    )
    value_cells = _capture_cells_for_range(
        conn,
        capture_revision_id=int(capture_revision_id),
        sheet_name=sheet_name,
        address=str(series["valueRange"]),
        path=f"{path}.valueRange",
    )
    identity_cells = _capture_cells_for_range(
        conn,
        capture_revision_id=int(capture_revision_id),
        sheet_name=sheet_name,
        address=str(series["rowIdentityRange"]),
        path=f"{path}.rowIdentityRange",
        resolve_merged_label=True,
    )

    header_row, header_start_col, _, _ = header_range
    value_start_row, value_start_col, value_end_row, value_end_col = (
        value_range
    )
    identity_start_row, identity_col, _, _ = identity_range
    headers: list[tuple[str, float | None, str]] = []
    for column_index in range(
        header_start_col,
        value_end_col - value_start_col + header_start_col + 1,
    ):
        header_cell = header_cells[(header_row, column_index)]
        header_value, header_source_json = _capture_cell_payload(
            header_cell,
            path=f"{path}.headerRange",
            numeric=False,
            sheet_name=sheet_name,
            formula_lookup=formula_lookup,
        )
        parsed_header = json.loads(str(header_source_json))
        headers.append(
            (
                str(header_value),
                _portable_axis_number(
                    parsed_header,
                    series.get("axisUnit"),
                ),
                str(header_cell["coordinate"]),
            )
        )
    header_coordinates = [item[2] for item in headers]
    if len(header_coordinates) != len(set(header_coordinates)):
        raise ValueError(
            f"{path}.headerRange has multiple logical header cells that "
            "resolve to the same merged anchor; cite distinct lower-level "
            "header identities instead of fabricating duplicate replicate "
            "or axis keys"
        )

    identities: list[tuple[str, float | None, str]] = []
    for row_offset, row_index in enumerate(
        range(
            identity_start_row,
            value_end_row - value_start_row + identity_start_row + 1,
        )
    ):
        identity_cell = identity_cells[(row_index, identity_col)]
        identity_value, identity_source_json = _capture_cell_payload(
            identity_cell,
            path=f"{path}.rowIdentityRange",
            numeric=False,
            sheet_name=sheet_name,
            formula_lookup=formula_lookup,
        )
        parsed_identity = json.loads(str(identity_source_json))
        identities.append(
            (
                str(identity_value),
                _portable_axis_number(
                    parsed_identity,
                    series.get("axisUnit"),
                ),
                str(identity_cell["coordinate"]),
            )
        )

    points: list[dict[str, Any]] = []
    axis_source = str(series["axisSource"]).upper()
    series_role = str(series.get("seriesRole") or "RAW").upper()
    aggregate_replicate_coordinates: set[tuple[int, int]] = set()
    for aggregate_range in series.get("aggregateReplicateRanges", []):
        (
            aggregate_start_row,
            aggregate_start_column,
            aggregate_end_row,
            aggregate_end_column,
        ) = parse_a1_range(str(aggregate_range))
        for aggregate_row in range(
            aggregate_start_row,
            aggregate_end_row + 1,
        ):
            for aggregate_column in range(
                aggregate_start_column,
                aggregate_end_column + 1,
            ):
                aggregate_replicate_coordinates.add(
                    (aggregate_row, aggregate_column)
                )
    for row_ordinal, row_index in enumerate(
        range(value_start_row, value_end_row + 1),
        start=1,
    ):
        for column_ordinal, column_index in enumerate(
            range(value_start_col, value_end_col + 1),
            start=1,
        ):
            value_cell = value_cells[(row_index, column_index)]
            value_number, source_value_json = _capture_cell_payload(
                value_cell,
                path=f"{path}.valueRange",
                numeric=True,
                numeric_unit=series.get("valueUnit"),
                sheet_name=sheet_name,
                formula_lookup=formula_lookup,
            )
            header_label, header_value, header_coordinate = headers[
                column_ordinal - 1
            ]
            (
                identity_label,
                identity_value,
                identity_coordinate,
            ) = identities[row_ordinal - 1]
            if axis_source == "HEADER":
                axis_label = header_label
                axis_value = header_value
                axis_coordinate = header_coordinate
                replicate_key = identity_label
                replicate_coordinate = identity_coordinate
            else:
                axis_label = identity_label
                axis_value = identity_value
                axis_coordinate = identity_coordinate
                replicate_key = header_label
                replicate_coordinate = header_coordinate
            replicate_role = "RAW"
            if series_role == "AGGREGATE":
                replicate_role = "AGGREGATE"
            elif (
                (
                    (
                        identity_start_row + row_ordinal - 1,
                        identity_col,
                    )
                    if axis_source == "HEADER"
                    else (
                        header_row,
                        header_start_col + column_ordinal - 1,
                    )
                )
                in aggregate_replicate_coordinates
            ):
                replicate_role = "AGGREGATE"
            coordinate = str(value_cell["coordinate"])
            points.append(
                {
                    "pointUid": stable_uid(
                        "measurement-point",
                        series_uid,
                        coordinate,
                    ),
                    "rowOrdinal": row_ordinal,
                    "columnOrdinal": column_ordinal,
                    "axisLabel": axis_label,
                    "axisValue": axis_value,
                    "replicateKey": replicate_key,
                    "replicateRole": replicate_role,
                    "axisSourceCoordinate": axis_coordinate,
                    "replicateSourceCoordinate": replicate_coordinate,
                    "sourceRowIndex": row_index,
                    "sourceColumnIndex": column_index,
                    "sourceCoordinate": coordinate,
                    "sourceValueJson": source_value_json,
                    "sourceFormulaText": str(value_cell["formula_text"] or ""),
                    "valueNumber": value_number,
                }
            )
    return points


_AVERAGE_REPLICATE_TOKEN_PATTERN = re.compile(r"[0-9A-Za-z가-힣]+")
_AVERAGE_REPLICATE_TOKENS = {"avg", "average", "mean", "평균"}


def _is_average_replicate_key(value: object) -> bool:
    tokens = {
        token.casefold()
        for token in _AVERAGE_REPLICATE_TOKEN_PATTERN.findall(str(value or ""))
    }
    return bool(tokens & _AVERAGE_REPLICATE_TOKENS)


def _validate_average_aggregate_points(
    points: list[dict[str, Any]],
    *,
    path: str,
) -> None:
    """Fail closed when a source-labeled average excludes its raw members."""

    points_by_axis: dict[str, list[dict[str, Any]]] = {}
    for point in points:
        axis_coordinate = str(point["axisSourceCoordinate"])
        points_by_axis.setdefault(axis_coordinate, []).append(point)

    for axis_points in points_by_axis.values():
        raw_values = [
            float(point["valueNumber"])
            for point in axis_points
            if point["replicateRole"] == "RAW"
        ]
        for point in axis_points:
            if point["replicateRole"] != "AGGREGATE" or not _is_average_replicate_key(
                point["replicateKey"]
            ):
                continue
            if not raw_values:
                raise ValueError(
                    f"{path} average aggregate source cell "
                    f"{point['sourceCoordinate']} has no RAW members"
                )
            expected = math.fsum(raw_values) / len(raw_values)
            actual = float(point["valueNumber"])
            if not math.isclose(
                actual,
                expected,
                rel_tol=1e-6,
                abs_tol=1e-9,
            ):
                raise ValueError(
                    f"{path} average aggregate source cell "
                    f"{point['sourceCoordinate']} value {actual} does not "
                    f"equal the arithmetic mean {expected} of its "
                    f"{len(raw_values)} RAW members"
                )


def _normalized_axis_identity(point: dict[str, Any]) -> tuple[str, object]:
    axis_value = point.get("axisValue")
    if axis_value is not None:
        number = float(axis_value)
        if math.isfinite(number):
            return ("NUMBER", format(number, ".15g"))
    label = re.sub(
        r"\s+",
        " ",
        unicodedata.normalize(
            "NFKC",
            str(point.get("axisLabel") or ""),
        ).strip().casefold(),
    )
    return ("LABEL", label)


def _points_by_axis_coordinate(
    points: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = {}
    for point in points:
        result.setdefault(
            str(point["axisSourceCoordinate"]),
            [],
        ).append(point)
    return result


def _validate_standalone_average_series(
    expanded_series: dict[
        str,
        tuple[dict[str, Any], list[dict[str, Any]]],
    ],
    *,
    study_path: str,
) -> None:
    """Validate source-authored standalone averages without copying raw data."""

    for series_key, (series, aggregate_points) in expanded_series.items():
        if str(series.get("seriesRole") or "RAW").upper() != "AGGREGATE":
            continue
        path = f"{study_path}.measurementSeries[{series_key!r}]"
        if (
            str(series.get("aggregationFunction") or "").upper()
            != "AVERAGE"
        ):
            raise ValueError(
                f"{path}.aggregationFunction must be AVERAGE"
            )
        if not aggregate_points:
            raise ValueError(f"{path} has no aggregate points")
        if any(
            point["replicateRole"] != "AGGREGATE"
            for point in aggregate_points
        ):
            raise ValueError(f"{path} must contain only AGGREGATE points")

        aggregate_axes = _points_by_axis_coordinate(aggregate_points)
        for coordinate, points in aggregate_axes.items():
            if len(points) != 1:
                raise ValueError(
                    f"{path} has ambiguous aggregate axis {coordinate}"
                )

        source_lookups: list[
            tuple[
                str,
                dict[str, Any],
                dict[str, list[dict[str, Any]]],
                dict[tuple[str, object], list[str]],
            ]
        ] = []
        for source_key_value in series.get("aggregateOfSeries", []):
            source_key = str(source_key_value)
            source_entry = expanded_series.get(source_key)
            if source_entry is None:
                raise ValueError(
                    f"{path}.aggregateOfSeries references unknown "
                    f"measurementSeries {source_key}"
                )
            source_series, source_points = source_entry
            raw_points = [
                point
                for point in source_points
                if point["replicateRole"] == "RAW"
            ]
            if not raw_points:
                raise ValueError(
                    f"{path} source measurementSeries {source_key} "
                    "has no RAW points"
                )
            source_axes = _points_by_axis_coordinate(raw_points)
            normalized_axes: dict[
                tuple[str, object],
                list[str],
            ] = {}
            for source_coordinate, axis_points in source_axes.items():
                identities = {
                    _normalized_axis_identity(point)
                    for point in axis_points
                }
                if len(identities) != 1:
                    raise ValueError(
                        f"{path} source measurementSeries {source_key} "
                        f"has ambiguous axis {source_coordinate}"
                    )
                normalized_axes.setdefault(
                    next(iter(identities)),
                    [],
                ).append(source_coordinate)
            source_lookups.append(
                (
                    source_key,
                    source_series,
                    source_axes,
                    normalized_axes,
                )
            )

        matched_source_axes: list[set[str]] = [
            set() for _item in source_lookups
        ]
        for aggregate_coordinate, aggregate_axis_points in (
            aggregate_axes.items()
        ):
            aggregate_point = aggregate_axis_points[0]
            raw_values: list[float] = []
            for source_index, (
                source_key,
                source_series,
                source_axes,
                normalized_axes,
            ) in enumerate(source_lookups):
                matched_coordinate: str | None = None
                if (
                    str(source_series["sheet"]) == str(series["sheet"])
                    and aggregate_coordinate in source_axes
                ):
                    matched_coordinate = aggregate_coordinate
                else:
                    identity = _normalized_axis_identity(aggregate_point)
                    candidates = normalized_axes.get(identity, [])
                    if not candidates:
                        raise ValueError(
                            f"{path} aggregate axis "
                            f"{aggregate_point['axisLabel']!r} is missing "
                            f"from source measurementSeries {source_key}"
                        )
                    if len(candidates) != 1:
                        raise ValueError(
                            f"{path} aggregate axis "
                            f"{aggregate_point['axisLabel']!r} is ambiguous "
                            f"in source measurementSeries {source_key}"
                        )
                    matched_coordinate = candidates[0]
                if (
                    matched_coordinate
                    in matched_source_axes[source_index]
                ):
                    raise ValueError(
                        f"{path} aggregate axes do not map one-to-one to "
                        f"source measurementSeries {source_key}"
                    )
                matched_source_axes[source_index].add(matched_coordinate)
                raw_values.extend(
                    float(point["valueNumber"])
                    for point in source_axes[matched_coordinate]
                )
            if not raw_values:
                raise ValueError(
                    f"{path} aggregate source cell "
                    f"{aggregate_point['sourceCoordinate']} has no RAW "
                    "members"
                )
            expected = math.fsum(raw_values) / len(raw_values)
            actual = float(aggregate_point["valueNumber"])
            if not math.isclose(
                actual,
                expected,
                rel_tol=1e-6,
                abs_tol=1e-9,
            ):
                raise ValueError(
                    f"{path} aggregate source cell "
                    f"{aggregate_point['sourceCoordinate']} value {actual} "
                    f"does not equal the arithmetic mean {expected} of its "
                    f"{len(raw_values)} referenced RAW points"
                )


def _numeric_cells_from_capture_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    evidence_items: list[dict[str, Any]],
    *,
    formula_lookup: FormulaOverlayLookup | None = None,
) -> list[tuple[float, bool]]:
    capture_revision_id = revision["capture_v2_revision_id"]
    if capture_revision_id is None or not _table_exists(conn, "capture_v2_sheets"):
        return []
    numbers: list[tuple[float, bool]] = []
    for item in evidence_items:
        start_row, start_col, end_row, end_col = parse_a1_range(str(item["range"]))
        rows = conn.execute(
            """
            SELECT
                c.raw_value_json, c.formula_text, c.cached_value_json,
                c.display_value_json, c.number_format, c.coordinate
            FROM capture_v2_sheets s
            JOIN capture_v2_cells c ON c.sheet_id=s.sheet_id
            WHERE s.revision_id=? AND s.title=?
              AND c.row_index BETWEEN ? AND ?
              AND c.column_index BETWEEN ? AND ?
            """,
            (
                int(capture_revision_id),
                str(item["sheet"]),
                start_row,
                end_row,
                start_col,
                end_col,
            ),
        ).fetchall()
        for row in rows:
            # Formula text itself is never numeric evidence. A formula contributes
            # only its Capture cache or an independently checksum-validated,
            # provenance-bearing deterministic overlay value.
            formula_text = str(row["formula_text"] or "")
            source_json = (
                row["cached_value_json"]
                if formula_text
                else row["raw_value_json"]
            )
            if formula_text and source_json is None:
                source_json = _derived_formula_source_json(
                    formula_lookup,
                    sheet_name=str(item["sheet"]),
                    coordinate=row["coordinate"],
                    formula_text=formula_text,
                )
            if source_json is None:
                continue
            try:
                parsed = json.loads(str(source_json))
            except json.JSONDecodeError:
                continue
            number = _portable_number(parsed)
            if number is None:
                continue
            numbers.append(
                (
                    number,
                    "%" in str(row["number_format"] or ""),
                )
            )
    return numbers


_EXPLICIT_COUNT_RATIO_PATTERN = re.compile(
    r"^\s*(?P<numerator>[0-9]+)\s*/\s*"
    r"(?P<denominator>[1-9][0-9]*)\s*(?:pcs?|ea)\s*$",
    flags=re.IGNORECASE,
)


def _explicit_count_ratio_numbers(
    value: object,
) -> tuple[float, float] | None:
    """Parse only a whole-cell, count-unit-qualified numerator/denominator."""

    if not isinstance(value, str):
        return None
    match = _EXPLICIT_COUNT_RATIO_PATTERN.fullmatch(value)
    if match is None:
        return None
    numerator = int(match.group("numerator"))
    denominator = int(match.group("denominator"))
    if numerator > denominator:
        return None
    return float(numerator), float(denominator)


def _count_ratio_numbers_from_capture_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    evidence_items: list[dict[str, Any]],
) -> list[float]:
    capture_revision_id = revision["capture_v2_revision_id"]
    if (
        capture_revision_id is None
        or not _table_exists(conn, "capture_v2_sheets")
    ):
        return []
    numbers: list[float] = []
    for item in evidence_items:
        start_row, start_col, end_row, end_col = parse_a1_range(
            str(item["range"])
        )
        rows = conn.execute(
            """
            SELECT
                c.raw_value_json, c.formula_text, c.cached_value_json,
                c.display_value_json
            FROM capture_v2_sheets s
            JOIN capture_v2_cells c ON c.sheet_id=s.sheet_id
            WHERE s.revision_id=? AND s.title=?
              AND c.row_index BETWEEN ? AND ?
              AND c.column_index BETWEEN ? AND ?
            """,
            (
                int(capture_revision_id),
                str(item["sheet"]),
                start_row,
                end_row,
                start_col,
                end_col,
            ),
        ).fetchall()
        for row in rows:
            formula_text = str(row["formula_text"] or "")
            source_json = (
                row["cached_value_json"]
                if formula_text
                else row["raw_value_json"]
            )
            # A displayed formula result is not numeric evidence when Capture
            # has no cached result for that formula.
            if formula_text and source_json is None:
                continue
            ratios: set[tuple[float, float]] = set()
            for candidate_json in (
                source_json,
                row["display_value_json"],
            ):
                if candidate_json is None:
                    continue
                try:
                    candidate = json.loads(str(candidate_json))
                except json.JSONDecodeError:
                    continue
                ratio = _explicit_count_ratio_numbers(candidate)
                if ratio is not None:
                    ratios.add(ratio)
            for numerator, denominator in sorted(ratios):
                numbers.extend((numerator, denominator))
    return numbers


def _numbers_from_capture_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    evidence_items: list[dict[str, Any]],
    *,
    formula_lookup: FormulaOverlayLookup | None = None,
) -> list[float]:
    numbers: list[float] = []
    for number, is_percent_format in _numeric_cells_from_capture_evidence(
        conn,
        revision,
        evidence_items,
        formula_lookup=formula_lookup,
    ):
        numbers.append(number)
        if is_percent_format:
            numbers.append(number * 100.0)
    numbers.extend(
        _count_ratio_numbers_from_capture_evidence(
            conn,
            revision,
            evidence_items,
        )
    )
    return numbers


def _human_percent_numbers_from_capture_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    evidence_items: list[dict[str, Any]],
    *,
    formula_lookup: FormulaOverlayLookup | None = None,
) -> list[float]:
    """Return percent-formatted cells only on their canonical human scale."""

    return [
        number * 100.0
        for number, is_percent_format in _numeric_cells_from_capture_evidence(
            conn,
            revision,
            evidence_items,
            formula_lookup=formula_lookup,
        )
        if is_percent_format
    ]


def _labeled_percent_numbers_from_capture_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    evidence_items: list[dict[str, Any]],
    *,
    outcome_label: str,
) -> list[float]:
    """Return percentages explicitly paired with the exact outcome label.

    A compound source cell can contain multiple labeled percentages, so an
    unqualified numeric-token search would permit the wrong metric.  This
    helper accepts only a number immediately following the outcome's exact
    source label.
    """

    capture_revision_id = revision["capture_v2_revision_id"]
    normalized_label = " ".join(str(outcome_label or "").split())
    if (
        capture_revision_id is None
        or not normalized_label
        or not _table_exists(conn, "capture_v2_sheets")
    ):
        return []
    label_aliases = [normalized_label]
    source_style_label = re.sub(
        r"\s+(?:percentage|percent|pct|rate)\s*$",
        "",
        normalized_label,
        flags=re.IGNORECASE,
    ).strip()
    if source_style_label and source_style_label not in label_aliases:
        label_aliases.append(source_style_label)
    label_pattern = "(?:" + "|".join(
        r"\s+".join(re.escape(part) for part in alias.split())
        for alias in label_aliases
    ) + ")"
    labeled_percent = re.compile(
        rf"(?<![A-Za-z0-9_]){label_pattern}(?![A-Za-z0-9_])"
        r"\s*[:=]?\s*"
        r"(?P<number>[+-]?(?:\d+(?:,\d{3})*|\d*)"
        r"(?:\.\d+)?(?:[Ee][+-]?\d+)?)\s*%",
        re.IGNORECASE,
    )
    numbers: list[float] = []
    for item in evidence_items:
        start_row, start_col, end_row, end_col = parse_a1_range(
            str(item["range"])
        )
        rows = conn.execute(
            """
            SELECT
                c.raw_value_json, c.formula_text, c.cached_value_json,
                c.display_value_json
            FROM capture_v2_sheets s
            JOIN capture_v2_cells c ON c.sheet_id=s.sheet_id
            WHERE s.revision_id=? AND s.title=?
              AND c.row_index BETWEEN ? AND ?
              AND c.column_index BETWEEN ? AND ?
            """,
            (
                int(capture_revision_id),
                str(item["sheet"]),
                start_row,
                end_row,
                start_col,
                end_col,
            ),
        ).fetchall()
        for row in rows:
            formula_text = str(row["formula_text"] or "")
            source_json = (
                row["cached_value_json"]
                if formula_text
                else row["raw_value_json"]
            )
            if formula_text and source_json is None:
                continue
            for candidate_json in (
                source_json,
                row["display_value_json"],
            ):
                if candidate_json is None:
                    continue
                try:
                    candidate = json.loads(str(candidate_json))
                except json.JSONDecodeError:
                    continue
                if not isinstance(candidate, str):
                    continue
                for match in labeled_percent.finditer(candidate):
                    try:
                        number = float(
                            match.group("number").replace(",", "")
                        )
                    except ValueError:
                        continue
                    if math.isfinite(number) and number not in numbers:
                        numbers.append(number)
    return numbers


_CONCLUSION_MARKERS = frozenset(
    {
        "conclusion",
        "decision",
        "finding",
        "follow standard",
        "judgment",
        "recommendation",
        "result",
        "can use",
        "test more",
        "verdict",
        "결과",
        "결론",
        "판정",
    }
)
_WHOLE_CELL_QUANTITY_PATTERN = re.compile(
    r"^([+-]?(?:\d+(?:\.\d+)?|\.\d+))\s*([A-Za-z%°µμ]+)$"
)
_ORDINAL_TOKEN_PATTERN = re.compile(
    r"^\d+(?:st|nd|rd|th)$",
    flags=re.IGNORECASE,
)
_PLAIN_NUMBER_PATTERN = re.compile(
    r"^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$"
)
_NORMAL_ARM_PATTERN = re.compile(
    r"^normal(?:\s*\([^()]*\))?$",
    flags=re.IGNORECASE,
)
_CONTROL_SOURCE_PATTERN = re.compile(
    r"(?<![A-Za-z])control(?![A-Za-z])|대조(?:군)?",
    flags=re.IGNORECASE,
)
_REFERENCE_SOURCE_PATTERN = re.compile(
    r"(?<![A-Za-z])(?:normal|reference|standard|spec(?:ification)?)"
    r"(?![A-Za-z])|기준|표준|사양",
    flags=re.IGNORECASE,
)
_REFERENCE_REPLICATE_PATTERN = re.compile(
    r"^(?:normal|reference|standard|spec(?:ification)?)"
    r"\s*#\s*(\d+)$",
    flags=re.IGNORECASE,
)


def _normalized_source_text(value: object) -> str:
    return re.sub(
        r"\s+",
        " ",
        unicodedata.normalize(
            "NFKC",
            str(value or ""),
        ).strip().casefold(),
    )


def _has_ordered_grouped_reference_identities(
    values: list[str],
) -> bool:
    if len(values) < 2:
        return False
    replicate_numbers: list[int] = []
    for value in values:
        match = _REFERENCE_REPLICATE_PATTERN.fullmatch(str(value).strip())
        if match is None:
            return False
        replicate_numbers.append(int(match.group(1)))
    return (
        len(replicate_numbers) == len(set(replicate_numbers))
        and all(
            left < right
            for left, right in zip(
                replicate_numbers,
                replicate_numbers[1:],
            )
        )
    )


def _capture_text_values_for_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    item: dict[str, Any],
) -> list[str]:
    capture_revision_id = revision["capture_v2_revision_id"]
    if (
        capture_revision_id is None
        or not _table_exists(conn, "capture_v2_sheets")
    ):
        return []
    start_row, start_col, end_row, end_col = parse_a1_range(
        str(item["range"])
    )
    rows = conn.execute(
        """
        SELECT
            c.raw_value_json, c.formula_text, c.cached_value_json,
            c.display_value_json
        FROM capture_v2_sheets s
        JOIN capture_v2_cells c ON c.sheet_id=s.sheet_id
        WHERE s.revision_id=? AND s.title=?
          AND c.row_index BETWEEN ? AND ?
          AND c.column_index BETWEEN ? AND ?
        ORDER BY c.row_index, c.column_index
        """,
        (
            int(capture_revision_id),
            str(item["sheet"]),
            start_row,
            end_row,
            start_col,
            end_col,
        ),
    ).fetchall()
    values: list[str] = []
    for row in rows:
        formula_text = str(row["formula_text"] or "")
        source_json = (
            row["cached_value_json"]
            if formula_text
            else row["raw_value_json"]
        )
        if formula_text and source_json is None:
            continue
        for candidate_json in (
            source_json,
            row["display_value_json"],
        ):
            if candidate_json is None:
                continue
            try:
                candidate = json.loads(str(candidate_json))
            except json.JSONDecodeError:
                continue
            if isinstance(candidate, str) and candidate.strip():
                values.append(candidate.strip())
    return list(dict.fromkeys(values))


def _captured_texts_for_evidence_items(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    evidence_items: list[dict[str, Any]],
) -> list[str]:
    values: list[str] = []
    for item in evidence_items:
        values.extend(
            _capture_text_values_for_evidence(
                conn,
                revision,
                item,
            )
        )
    return list(dict.fromkeys(values))


def _whole_cell_quantity(
    conn: sqlite3.Connection,
    value: object,
) -> tuple[float, str, int] | None:
    quantity_syntax = _whole_cell_quantity_syntax(value)
    if quantity_syntax is None:
        return None
    number, unit_text = quantity_syntax
    unit_id = resolve_unit_id(conn, unit_text)
    if unit_id is None:
        return None
    return number, unit_text, unit_id


def _whole_cell_quantity_syntax(
    value: object,
) -> tuple[float, str] | None:
    """Parse strict numeric+unit syntax without consulting the unit catalog."""

    text = str(value or "").strip()
    if _ORDINAL_TOKEN_PATTERN.fullmatch(text):
        return None
    match = _WHOLE_CELL_QUANTITY_PATTERN.fullmatch(text)
    if match is None:
        return None
    unit_text = match.group(2)
    return float(match.group(1)), unit_text


def validate_factor_and_arm_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    manifest: dict[str, Any],
) -> None:
    """Validate source-grounded factor quantities and Arm reference roles."""

    for study_index, study in enumerate(manifest.get("studies", [])):
        factors = {
            str(factor.get("key") or ""): factor
            for factor in study.get("factors", [])
        }
        for arm_index, arm in enumerate(study.get("arms", [])):
            arm_path = f"studies[{study_index}].arms[{arm_index}]"
            role = str(arm.get("role") or "OTHER").strip().upper()
            label = str(arm.get("label") or "").strip()
            condition = str(arm.get("condition") or "").strip()
            arm_texts = _captured_texts_for_evidence_items(
                conn,
                revision,
                arm.get("evidence", []),
            )
            normalized_arm_texts = {
                _normalized_source_text(value)
                for value in arm_texts
                if _normalized_source_text(value)
            }
            normalized_label = _normalized_source_text(label)
            label_is_cited = normalized_label in normalized_arm_texts

            if _NORMAL_ARM_PATTERN.fullmatch(label):
                if not label_is_cited:
                    raise ValueError(
                        f"{arm_path} Normal label must exactly match a "
                        "captured cell in Arm evidence"
                    )
                if role != "REFERENCE":
                    raise ValueError(
                        f"{arm_path} literal source Normal maps to REFERENCE, "
                        "never CONTROL by itself"
                    )
            if role == "CONTROL":
                control_identity_values = {
                    _normalized_source_text(value)
                    for value in (label, condition)
                    if str(value or "").strip()
                }
                supported_control = any(
                    _CONTROL_SOURCE_PATTERN.search(source_text)
                    and _normalized_source_text(source_text)
                    in control_identity_values
                    for source_text in arm_texts
                )
                if not supported_control:
                    raise ValueError(
                        f"{arm_path}.role CONTROL requires directly cited "
                        "captured explicit Control wording matching the Arm "
                        "label or condition"
                    )
            if role == "REFERENCE":
                reference_identity_values = {
                    _normalized_source_text(value)
                    for value in (label, condition)
                    if str(value or "").strip()
                }
                supported_reference = any(
                    _REFERENCE_SOURCE_PATTERN.search(source_text)
                    and _normalized_source_text(source_text)
                    in reference_identity_values
                    for source_text in arm_texts
                )
                grouped_reference = (
                    any(
                        _REFERENCE_SOURCE_PATTERN.search(identity)
                        for identity in (label, condition)
                    )
                    and _has_ordered_grouped_reference_identities(arm_texts)
                )
                if not supported_reference and not grouped_reference:
                    raise ValueError(
                        f"{arm_path}.role REFERENCE requires directly cited "
                        "captured full Normal, Reference, Standard, Spec, or "
                        "equivalent reference wording matching the Arm label "
                        "or condition, or at least two exact ordered distinct "
                        "full reference #N identity cells for a descriptive "
                        "grouped Arm; a bare abbreviation such as ST or mixed "
                        "Test/Normal evidence is not reference semantics"
                    )

            for factor_value_index, factor_value in enumerate(
                arm.get("factorValues", [])
            ):
                factor_key = str(factor_value.get("factor") or "")
                factor = factors.get(factor_key)
                if factor is None:
                    continue
                factor_value_path = (
                    f"{arm_path}.factorValues[{factor_value_index}]"
                )
                cited_texts = list(arm_texts)
                cited_texts.extend(
                    _captured_texts_for_evidence_items(
                        conn,
                        revision,
                        factor.get("evidence", []),
                    )
                )
                quantity_cells = [
                    (source_text, quantity)
                    for source_text in dict.fromkeys(cited_texts)
                    if (
                        quantity := _whole_cell_quantity(
                            conn,
                            source_text,
                        )
                    )
                    is not None
                ]
                quantity_syntax_cells = [
                    (source_text, quantity)
                    for source_text in dict.fromkeys(cited_texts)
                    if (
                        quantity := _whole_cell_quantity_syntax(source_text)
                    )
                    is not None
                ]
                value_text = str(factor_value.get("value") or "").strip()
                normalized_value_text = _normalized_source_text(value_text)
                normalized_cited_texts = {
                    _normalized_source_text(source_text)
                    for source_text in cited_texts
                    if _normalized_source_text(source_text)
                }
                if (
                    (
                        _ORDINAL_TOKEN_PATTERN.fullmatch(value_text)
                        or normalized_value_text == "total"
                    )
                    and normalized_value_text not in normalized_cited_texts
                    and any(
                        normalized_value_text in source_text
                        for source_text in normalized_cited_texts
                    )
                ):
                    raise ValueError(
                        f"{factor_value_path}.value must not isolate an "
                        "ordinal or Total token from composite evidence; "
                        "preserve the exact whole-cell compound value"
                    )
                declared_quantity_syntax = _whole_cell_quantity_syntax(
                    value_text
                )
                declared_quantity = _whole_cell_quantity(
                    conn,
                    value_text,
                )
                exact_syntax_matches = [
                    item
                    for item in quantity_syntax_cells
                    if item[0].strip() == value_text
                ]
                exact_matches = [
                    item
                    for item in quantity_cells
                    if item[0].strip() == value_text
                ]
                if (
                    declared_quantity_syntax is not None
                    and not exact_syntax_matches
                ):
                    raise ValueError(
                        f"{factor_value_path}.value must preserve an exact "
                        "whole-cell quantity from its Arm/factor evidence"
                    )
                if (
                    declared_quantity is None
                    and _PLAIN_NUMBER_PATTERN.fullmatch(value_text)
                ):
                    numeric_value = float(value_text)
                    numeric_matches = [
                        item
                        for item in quantity_cells
                        if math.isclose(
                            item[1][0],
                            numeric_value,
                            rel_tol=1e-9,
                            abs_tol=1e-12,
                        )
                    ]
                    if len(numeric_matches) == 1:
                        raise ValueError(
                            f"{factor_value_path}.value must preserve exact "
                            f"source text {numeric_matches[0][0]!r}"
                        )
                if not exact_matches:
                    continue
                source_text, (
                    source_number,
                    source_unit,
                    source_unit_id,
                ) = exact_matches[0]
                value_number = factor_value.get("valueNumber")
                if (
                    isinstance(value_number, bool)
                    or not isinstance(value_number, (int, float))
                    or not math.isfinite(float(value_number))
                    or not math.isclose(
                        float(value_number),
                        source_number,
                        rel_tol=1e-9,
                        abs_tol=1e-12,
                    )
                ):
                    raise ValueError(
                        f"{factor_value_path}.valueNumber must equal "
                        f"{source_number:g} from captured whole-cell "
                        f"{source_text!r}"
                    )
                declared_unit = str(
                    factor_value.get("unit") or ""
                ).strip()
                declared_unit_id = resolve_unit_id(conn, declared_unit)
                if (
                    declared_unit_id is None
                    or declared_unit_id != source_unit_id
                ):
                    raise ValueError(
                        f"{factor_value_path}.unit must resolve to captured "
                        f"unit {source_unit!r}"
                    )


_SCALAR_QUANTITATIVE_FIELDS = (
    "valueNumber",
    "numerator",
    "denominator",
    "ratePpm",
    "min",
    "max",
    "average",
    "sampleSize",
)


def _comparison_unit_identity(
    conn: sqlite3.Connection,
    value: object,
) -> tuple[str, object]:
    text = str(value or "").strip()
    if not text:
        return ("EMPTY", "")
    unit_id = resolve_unit_id(conn, text)
    if unit_id is not None:
        return ("UNIT", unit_id)
    return ("LITERAL", _normalized_source_text(text))


def _raw_series_alignment_signature(
    conn: sqlite3.Connection,
    series: dict[str, Any],
    points: list[dict[str, Any]],
) -> tuple[object, ...]:
    axis_source = str(series["axisSource"]).upper()
    if axis_source == "HEADER":
        axis_points = sorted(
            (
                point
                for point in points
                if int(point["rowOrdinal"]) == 1
            ),
            key=lambda point: int(point["columnOrdinal"]),
        )
    else:
        axis_points = sorted(
            (
                point
                for point in points
                if int(point["columnOrdinal"]) == 1
            ),
            key=lambda point: int(point["rowOrdinal"]),
        )
    row_count = max(
        (int(point["rowOrdinal"]) for point in points),
        default=0,
    )
    column_count = max(
        (int(point["columnOrdinal"]) for point in points),
        default=0,
    )
    return (
        _comparison_unit_identity(conn, series.get("valueUnit")),
        _comparison_unit_identity(conn, series.get("axisUnit")),
        axis_source,
        tuple(
            _normalized_axis_identity(point)
            for point in axis_points
        ),
        row_count,
        column_count,
        _normalized_source_text(series.get("stratumKey")),
    )


def _scalar_representation_is_compatible(
    left: list[dict[str, Any]],
    right: list[dict[str, Any]],
) -> bool:
    left_fields = {
        field
        for observation in left
        for field in _SCALAR_QUANTITATIVE_FIELDS
        if observation.get(field) not in (None, "")
    }
    right_fields = {
        field
        for observation in right
        for field in _SCALAR_QUANTITATIVE_FIELDS
        if observation.get(field) not in (None, "")
    }
    if left_fields or right_fields:
        return bool(left_fields & right_fields)
    left_qualitative = any(
        str(observation.get("valueText") or "").strip()
        for observation in left
    )
    right_qualitative = any(
        str(observation.get("valueText") or "").strip()
        for observation in right
    )
    return left_qualitative and right_qualitative


def validate_comparison_representation_alignment(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    manifest: dict[str, Any],
    *,
    formula_overlay: (
        dict[str, Any] | FormulaOverlayLookup | None
    ) = None,
) -> None:
    """Require compared Arms to represent shared Outcomes compatibly."""

    source = manifest.get("source", {})
    formula_lookup = _formula_lookup_for_revision(revision, formula_overlay)
    for study_index, study in enumerate(manifest.get("studies", [])):
        study_path = f"studies[{study_index}]"
        observations_by_outcome_arm: dict[
            tuple[str, str],
            list[dict[str, Any]],
        ] = {}
        for outcome in study.get("outcomes", []):
            outcome_key = str(outcome.get("key") or "")
            for observation in outcome.get("observations", []):
                observations_by_outcome_arm.setdefault(
                    (
                        outcome_key,
                        str(observation.get("arm") or ""),
                    ),
                    [],
                ).append(observation)

        raw_series_by_outcome_arm: dict[
            tuple[str, str],
            list[tuple[dict[str, Any], list[dict[str, Any]]]],
        ] = {}
        for series_index, series in enumerate(
            study.get("measurementSeries", [])
        ):
            if str(series.get("seriesRole") or "RAW").upper() != "RAW":
                continue
            series_uid = stable_uid(
                "comparison-alignment-series",
                source.get("revisionUid"),
                study.get("key"),
                series.get("key"),
            )
            points = _expand_measurement_series(
                conn,
                revision=revision,
                series=series,
                series_uid=series_uid,
                path=(
                    f"{study_path}.measurementSeries"
                    f"[{series_index}]"
                ),
                formula_lookup=formula_lookup,
            )
            raw_series_by_outcome_arm.setdefault(
                (
                    str(series.get("outcome") or ""),
                    str(series.get("arm") or ""),
                ),
                [],
            ).append((series, points))

        outcome_keys = {
            str(outcome.get("key") or "")
            for outcome in study.get("outcomes", [])
        }
        for comparison_index, comparison in enumerate(
            study.get("comparisons", [])
        ):
            comparison_path = (
                f"{study_path}.comparisons[{comparison_index}]"
            )
            left_arm = str(comparison.get("comparedArm") or "")
            right_arm = str(comparison.get("controlArm") or "")
            shared_outcomes = [
                outcome_key
                for outcome_key in outcome_keys
                if (
                    observations_by_outcome_arm.get(
                        (outcome_key, left_arm)
                    )
                    or raw_series_by_outcome_arm.get(
                        (outcome_key, left_arm)
                    )
                )
                and (
                    observations_by_outcome_arm.get(
                        (outcome_key, right_arm)
                    )
                    or raw_series_by_outcome_arm.get(
                        (outcome_key, right_arm)
                    )
                )
            ]
            if not shared_outcomes:
                raise ValueError(
                    f"{comparison_path} requires at least one shared "
                    "Outcome represented by both Arms; omit the invalid "
                    "Comparison, preserve its Arms/Outcomes/series, and add "
                    "a limitation"
                )
            for outcome_key in shared_outcomes:
                left_raw = raw_series_by_outcome_arm.get(
                    (outcome_key, left_arm),
                    [],
                )
                right_raw = raw_series_by_outcome_arm.get(
                    (outcome_key, right_arm),
                    [],
                )
                left_scalar = observations_by_outcome_arm.get(
                    (outcome_key, left_arm),
                    [],
                )
                right_scalar = observations_by_outcome_arm.get(
                    (outcome_key, right_arm),
                    [],
                )
                if bool(left_raw) != bool(right_raw):
                    raise ValueError(
                        f"{comparison_path} shared Outcome {outcome_key!r} "
                        "has incompatible RAW measurementSeries versus "
                        "scalar/summary representation; omit the invalid "
                        "Comparison, preserve its Arms/Outcomes/series, and "
                        "add a limitation"
                    )
                if left_raw and right_raw:
                    left_signatures = Counter(
                        _raw_series_alignment_signature(
                            conn,
                            series,
                            points,
                        )
                        for series, points in left_raw
                    )
                    right_signatures = Counter(
                        _raw_series_alignment_signature(
                            conn,
                            series,
                            points,
                        )
                        for series, points in right_raw
                    )
                    if left_signatures != right_signatures:
                        raise ValueError(
                            f"{comparison_path} shared Outcome "
                            f"{outcome_key!r} RAW representations require "
                            "compatible value units and aligned ordered axis "
                            "identity, shape, and stratum; omit the invalid "
                            "Comparison, preserve its Arms/Outcomes/series, "
                            "and add a limitation"
                        )
                    continue
                if not _scalar_representation_is_compatible(
                    left_scalar,
                    right_scalar,
                ):
                    raise ValueError(
                        f"{comparison_path} shared Outcome {outcome_key!r} "
                        "scalar representations require a compatible shared "
                        "quantitative field or qualitative-only values on "
                        "both Arms; omit the invalid Comparison, preserve "
                        "its Arms/Outcomes/series, and add a limitation"
                    )


def _source_conclusion_is_supported(
    conclusion: dict[str, Any],
    cited_texts: list[str],
    *,
    source_text: str,
) -> bool:
    normalized_source = _normalized_source_text(source_text)
    normalized_claim = _normalized_source_text(conclusion.get("text"))
    normalized_cited_ordered = list(
        dict.fromkeys(
            _normalized_source_text(value)
            for value in cited_texts
            if _normalized_source_text(value)
        )
    )
    normalized_cited = set(normalized_cited_ordered)
    ordered_source_candidates = set(normalized_cited_ordered)
    for start in range(len(normalized_cited_ordered)):
        for end in range(start + 2, len(normalized_cited_ordered) + 1):
            ordered_cells = normalized_cited_ordered[start:end]
            for separator in (" ", "; ", " ; "):
                ordered_source_candidates.add(
                    separator.join(ordered_cells)
                )
    if (
        not normalized_source
        or normalized_source not in ordered_source_candidates
        or not normalized_claim
        or (
            normalized_source not in normalized_claim
            and normalized_claim not in normalized_source
        )
    ):
        return False
    source_tokens = re.findall(
        r"[^\W\d_]+",
        normalized_source,
        flags=re.UNICODE,
    )
    if not source_tokens:
        return False
    if set(source_tokens) <= _CONCLUSION_MARKERS:
        return False
    cited_has_marker = any(
        any(marker in value for marker in _CONCLUSION_MARKERS)
        for value in normalized_cited
    )
    alphabetic_count = sum(
        1 for char in normalized_source if char.isalpha()
    )
    return bool(
        cited_has_marker
        or (len(source_tokens) >= 3 and alphabetic_count >= 6)
    )


def validate_conclusion_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    manifest: dict[str, Any],
) -> None:
    """Prove SOURCE_CONCLUSION claims from exact captured narrative text."""

    for study_index, study in enumerate(manifest.get("studies", [])):
        for conclusion_index, conclusion in enumerate(
            study.get("conclusions", [])
        ):
            if (
                str(conclusion.get("claimType") or "").upper()
                != "SOURCE_CONCLUSION"
            ):
                continue
            supported = False
            for item in conclusion.get("evidence", []):
                source_text = str(item.get("sourceText") or "").strip()
                if not source_text:
                    continue
                cited_texts = _capture_text_values_for_evidence(
                    conn,
                    revision,
                    item,
                )
                if _source_conclusion_is_supported(
                    conclusion,
                    cited_texts,
                    source_text=source_text,
                ):
                    supported = True
                    break
            if not supported:
                raise ValueError(
                    f"studies[{study_index}].conclusions"
                    f"[{conclusion_index}] SOURCE_CONCLUSION requires "
                    "directly cited captured narrative decision/conclusion "
                    "text matching evidence.sourceText and supporting the "
                    "claim; classify an AI synthesis as "
                    "AI_DERIVED_DESCRIPTIVE or omit it"
                )


def _approximately_contains(values: list[float], expected: float) -> bool:
    tolerance = max(1e-9, abs(expected) * 1e-6)
    return any(abs(value - expected) <= tolerance for value in values)


def _normalized_observation_values(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    outcome: dict[str, Any],
    observation: dict[str, Any],
) -> tuple[dict[str, Any], list[str]]:
    """Store percent-formatted scalar observations in displayed percent units.

    Excel stores a displayed 3.6% as approximately 0.036.  Canonical outcome
    unit "%" means the human percentage scale, matching measurement-series
    normalization and later percentage-point effect calculations.  AI drafts
    may supply either the raw or displayed scale, so normalize only when a
    directly cited Capture v2 cell proves the raw percent-formatted value.
    """

    unit = str(outcome.get("unit") or "").strip().casefold()
    if unit not in {"%", "percent", "percentage", "pct"}:
        return observation, []
    numeric_cells = _numeric_cells_from_capture_evidence(
        conn,
        revision,
        observation.get("evidence", []),
    )
    percent_raw_values = [
        number
        for number, is_percent_format in numeric_cells
        if is_percent_format
    ]
    if not percent_raw_values:
        return observation, []

    normalized = dict(observation)
    changed_fields: list[str] = []
    for field in ("valueNumber", "min", "max", "average"):
        value = observation.get(field)
        if value in (None, ""):
            continue
        number = float(value)
        raw_matches = [
            raw
            for raw in percent_raw_values
            if _approximately_contains([raw], number)
        ]
        displayed_matches = [
            raw * 100.0
            for raw in percent_raw_values
            if _approximately_contains([raw * 100.0], number)
        ]
        if raw_matches and not displayed_matches:
            normalized[field] = number * 100.0
            changed_fields.append(field)
    return normalized, changed_fields


def validate_numeric_observation_evidence(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    manifest: dict[str, Any],
    *,
    formula_overlay: (
        dict[str, Any] | FormulaOverlayLookup | None
    ) = None,
) -> None:
    """Validate observation claims and every dense measurement-series cell.

    This complements the coordinate checker. It intentionally does not infer
    values from labels or layout. Formula text remains unusable unless the caller
    supplies a checksum-validated restricted-grammar derivation overlay for the
    exact source revision. PPM may also be supported by an explicitly cited
    numerator/denominator pair. Measurement series are expanded read-only here as
    a pre-import gate so blank, malformed, or error cells are rejected during
    semantic drafting rather than after the import stage starts.
    """

    if revision["capture_v2_revision_id"] is None:
        return
    formula_lookup = _formula_lookup_for_revision(revision, formula_overlay)
    for study_index, study in enumerate(manifest.get("studies", [])):
        expanded_series: dict[
            str,
            tuple[dict[str, Any], list[dict[str, Any]]],
        ] = {}
        for series_index, series in enumerate(
            study.get("measurementSeries", [])
        ):
            points = _expand_measurement_series(
                conn,
                revision=revision,
                series=series,
                series_uid=(
                    f"validation-study-{study_index}-series-{series_index}"
                ),
                path=(
                    f"studies[{study_index}]"
                    f".measurementSeries[{series_index}]"
                ),
                formula_lookup=formula_lookup,
            )
            series_key = str(series["key"])
            expanded_series[series_key] = (series, points)
            if str(series.get("seriesRole") or "RAW").upper() == "RAW":
                _validate_average_aggregate_points(
                    points,
                    path=(
                        f"studies[{study_index}]"
                        f".measurementSeries[{series_index}]"
                    ),
                )
        _validate_standalone_average_series(
            expanded_series,
            study_path=f"studies[{study_index}]",
        )
        for outcome_index, outcome in enumerate(study.get("outcomes", [])):
            for observation_index, observation in enumerate(outcome.get("observations", [])):
                path = (
                    f"studies[{study_index}].outcomes[{outcome_index}]"
                    f".observations[{observation_index}]"
                )
                evidence_items = observation.get("evidence", [])
                source_numbers = _numbers_from_capture_evidence(
                    conn,
                    revision,
                    evidence_items,
                    formula_lookup=formula_lookup,
                )
                outcome_unit = str(
                    outcome.get("unit") or ""
                ).strip().casefold()
                human_percent_source_numbers = (
                    _human_percent_numbers_from_capture_evidence(
                        conn,
                        revision,
                        evidence_items,
                        formula_lookup=formula_lookup,
                    )
                    if outcome_unit
                    in {"%", "percent", "percentage", "pct"}
                    else []
                )
                labeled_percent_source_numbers = (
                    _labeled_percent_numbers_from_capture_evidence(
                        conn,
                        revision,
                        evidence_items,
                        outcome_label=str(
                            outcome.get("originalLabel") or ""
                        ),
                    )
                    if outcome_unit
                    in {"%", "percent", "percentage", "pct"}
                    else []
                )
                claims: list[tuple[str, float]] = []
                for field in (
                    "valueNumber",
                    "numerator",
                    "denominator",
                    "ratePpm",
                    "min",
                    "max",
                    "average",
                ):
                    value = observation.get(field)
                    if value in (None, ""):
                        continue
                    try:
                        number = float(value)
                    except (TypeError, ValueError) as exc:
                        raise ValueError(f"{path}.{field} must be numeric") from exc
                    if not math.isfinite(number):
                        raise ValueError(f"{path}.{field} must be finite")
                    claims.append((field, number))
                numerator = observation.get("numerator")
                denominator = observation.get("denominator")
                derived_ppm = None
                if numerator not in (None, "") and denominator not in (None, ""):
                    denominator_number = float(denominator)
                    if denominator_number:
                        derived_ppm = float(numerator) / denominator_number * 1_000_000.0
                for field, expected in claims:
                    allowed_source_numbers = source_numbers
                    if (
                        (
                            human_percent_source_numbers
                            or labeled_percent_source_numbers
                        )
                        and field
                        in {"valueNumber", "min", "max", "average"}
                    ):
                        allowed_source_numbers = [
                            *human_percent_source_numbers,
                            *labeled_percent_source_numbers,
                        ]
                    if _approximately_contains(
                        allowed_source_numbers,
                        expected,
                    ):
                        continue
                    if (
                        field == "ratePpm"
                        and derived_ppm is not None
                        and abs(derived_ppm - expected)
                        <= max(1e-9, abs(expected) * 1e-6)
                    ):
                        continue
                    raise ValueError(
                        f"{path}.{field}={expected:g} is not present in its cited Capture v2 cells"
                    )


def unsupported_rate_pair_observation_paths(
    conn: sqlite3.Connection,
    revision: sqlite3.Row,
    manifest: dict[str, Any],
    *,
    formula_overlay: (
        dict[str, Any] | FormulaOverlayLookup | None
    ) = None,
) -> list[tuple[int, int, int]]:
    """Return every observation whose declared count pair lacks evidence."""

    if revision["capture_v2_revision_id"] is None:
        return []
    formula_lookup = _formula_lookup_for_revision(
        revision,
        formula_overlay,
    )
    unsupported: list[tuple[int, int, int]] = []
    for study_index, study in enumerate(manifest.get("studies", [])):
        for outcome_index, outcome in enumerate(
            study.get("outcomes", [])
        ):
            for observation_index, observation in enumerate(
                outcome.get("observations", [])
            ):
                numerator = observation.get("numerator")
                denominator = observation.get("denominator")
                if numerator in (None, "") and denominator in (None, ""):
                    continue
                if numerator in (None, "") or denominator in (None, ""):
                    # The canonical contract validator owns incomplete pairs.
                    continue
                try:
                    numerator_number = float(numerator)
                    denominator_number = float(denominator)
                except (TypeError, ValueError):
                    # The canonical contract validator owns non-numeric pairs.
                    continue
                if not (
                    math.isfinite(numerator_number)
                    and math.isfinite(denominator_number)
                ):
                    continue
                source_numbers = _numbers_from_capture_evidence(
                    conn,
                    revision,
                    observation.get("evidence", []),
                    formula_lookup=formula_lookup,
                )
                if (
                    _approximately_contains(
                        source_numbers,
                        numerator_number,
                    )
                    and _approximately_contains(
                        source_numbers,
                        denominator_number,
                    )
                ):
                    continue
                unsupported.append(
                    (
                        study_index,
                        outcome_index,
                        observation_index,
                    )
                )
    return unsupported


def _concept_id(
    conn: sqlite3.Connection,
    item: dict[str, Any],
    *,
    default_kind: str,
    original_value: object,
    now_iso: Callable[[], str],
) -> int | None:
    concept = item.get("concept")
    if isinstance(concept, dict):
        concept_uid = str(concept.get("uid") or concept.get("conceptUid") or "").strip()
        if concept_uid:
            row = conn.execute("SELECT concept_id FROM knowledge_concepts WHERE concept_uid=?", (concept_uid,)).fetchone()
            if row:
                return int(row[0])
        kind = str(concept.get("kind") or default_kind).strip().upper()
        canonical_name = str(concept.get("canonicalName") or "").strip()
    else:
        kind = default_kind
        canonical_name = str(item.get("canonicalName") or "").strip()
    match = find_concept_id(conn, canonical_name or original_value, kind)
    if match is not None:
        return match
    record_schema_candidate(
        conn,
        candidate_kind=f"CONCEPT:{kind}",
        original_value=original_value,
        suggested_canonical_name=canonical_name,
        now_iso=now_iso,
    )
    return None


def _unit_id(
    conn: sqlite3.Connection,
    unit: object,
    *,
    now_iso: Callable[[], str],
) -> int | None:
    result = resolve_unit_id(conn, unit)
    if result is None and str(unit or "").strip():
        record_schema_candidate(
            conn,
            candidate_kind="UNIT",
            original_value=unit,
            suggested_canonical_name=unit,
            now_iso=now_iso,
        )
    return result


def _store_evidence(
    conn: sqlite3.Connection,
    *,
    revision: sqlite3.Row,
    entity_type: str,
    entity_uid: str,
    evidence_items: list[dict[str, Any]],
    now_iso: Callable[[], str],
) -> int:
    count = 0
    for item in evidence_items:
        start_row, start_col, end_row, end_col = parse_a1_range(str(item["range"]))
        role = str(item.get("role") or "SOURCE").strip().upper()
        evidence_uid = stable_uid(
            "evidence",
            revision["revision_uid"],
            entity_type,
            entity_uid,
            item["sheet"],
            item["range"],
            role,
        )
        conn.execute(
            """
            INSERT INTO evidence_items(
                evidence_uid, public_evidence_id, revision_id, legacy_evidence_id,
                evidence_kind, sheet_name, start_row, start_col, end_row, end_col,
                range_address, evidence_role, source_text, note, content_sha256,
                verification_status, created_at
            ) VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', 'VERIFIED', ?)
            ON CONFLICT(evidence_uid) DO UPDATE SET
                source_text=excluded.source_text,
                note=excluded.note,
                verification_status='VERIFIED'
            """,
            (
                evidence_uid,
                public_id("EVD", evidence_uid),
                int(revision["revision_id"]),
                str(item.get("kind") or "CELL_RANGE").upper(),
                str(item["sheet"]),
                start_row,
                start_col,
                end_row,
                end_col,
                str(item["range"]),
                role,
                str(item.get("sourceText") or ""),
                str(item.get("note") or ""),
                now_iso(),
            ),
        )
        evidence_id = int(conn.execute("SELECT evidence_id FROM evidence_items WHERE evidence_uid=?", (evidence_uid,)).fetchone()[0])
        conn.execute(
            """
            INSERT OR IGNORE INTO entity_evidence_links(
                entity_type, entity_uid, evidence_id, evidence_role, claim_scope
            ) VALUES (?, ?, ?, ?, '')
            """,
            (entity_type, entity_uid, evidence_id, role),
        )
        count += 1
    return count


_IMPORT_VALIDATOR_NAME = "canonical-study-import"
_IMPORT_VALIDATOR_VERSION = "1"


def _manifest_entity_uids(
    manifest: dict[str, Any],
    analysis_uid: str,
) -> dict[str, set[str]]:
    desired: dict[str, set[str]] = {
        "WORKBOOK_ANALYSIS": {analysis_uid},
        "STUDY": set(),
        "CONTEXT": set(),
        "FACTOR": set(),
        "ARM": set(),
        "OUTCOME": set(),
        "OBSERVATION": set(),
        "MEASUREMENT_SERIES": set(),
        "MEASUREMENT_POINT": set(),
        "COMPARISON": set(),
        "EFFECT": set(),
        "CLAIM": set(),
    }
    for study in manifest["studies"]:
        study_uid = stable_uid("study", analysis_uid, study["key"])
        desired["STUDY"].add(study_uid)
        arm_uids = {
            str(arm["key"]): stable_uid("arm", study_uid, arm["key"])
            for arm in study.get("arms", [])
        }
        outcome_uids = {
            str(outcome["key"]): stable_uid("outcome", study_uid, outcome["key"])
            for outcome in study.get("outcomes", [])
        }
        desired["CONTEXT"].update(
            stable_uid("context", study_uid, context["key"])
            for context in study.get("contexts", [])
        )
        desired["FACTOR"].update(
            stable_uid("factor", study_uid, factor["key"])
            for factor in study.get("factors", [])
        )
        desired["ARM"].update(arm_uids.values())
        desired["OUTCOME"].update(outcome_uids.values())
        for outcome in study.get("outcomes", []):
            outcome_uid = outcome_uids[str(outcome["key"])]
            desired["OBSERVATION"].update(
                stable_uid(
                    "observation",
                    outcome_uid,
                    arm_uids[str(observation["arm"])],
                    observation["key"],
                )
                for observation in outcome.get("observations", [])
            )
        for series in study.get("measurementSeries", []):
            series_uid = stable_uid(
                "measurement-series",
                study_uid,
                series["key"],
            )
            desired["MEASUREMENT_SERIES"].add(series_uid)
            (
                start_row,
                start_col,
                end_row,
                end_col,
            ) = parse_a1_range(str(series["valueRange"]))
            desired["MEASUREMENT_POINT"].update(
                stable_uid(
                    "measurement-point",
                    series_uid,
                    f"{_column_label(column_index)}{row_index}",
                )
                for row_index in range(start_row, end_row + 1)
                for column_index in range(start_col, end_col + 1)
            )
        for comparison in study.get("comparisons", []):
            comparison_uid = stable_uid("comparison", study_uid, comparison["key"])
            desired["COMPARISON"].add(comparison_uid)
            desired["EFFECT"].update(
                stable_uid(
                    "effect",
                    comparison_uid,
                    effect["outcome"],
                    effect["effectType"],
                    effect.get("formulaVersion") or "canonical-v1",
                )
                for effect in comparison.get("effects", [])
            )
        desired["CLAIM"].update(
            stable_uid("claim", study_uid, conclusion["key"])
            for conclusion in study.get("conclusions", [])
        )
    return desired


def _analysis_entity_uids(
    conn: sqlite3.Connection,
    *,
    workbook_analysis_id: int,
    analysis_uid: str,
) -> dict[str, set[str]]:
    queries = {
        "STUDY": """
            SELECT s.study_uid
            FROM knowledge_studies s
            WHERE s.workbook_analysis_id=?
        """,
        "CONTEXT": """
            SELECT c.context_uid
            FROM knowledge_study_contexts c
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "FACTOR": """
            SELECT f.factor_uid
            FROM knowledge_factors f
            JOIN knowledge_studies s ON s.study_id=f.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "ARM": """
            SELECT a.arm_uid
            FROM knowledge_arms a
            JOIN knowledge_studies s ON s.study_id=a.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "OUTCOME": """
            SELECT o.outcome_uid
            FROM knowledge_outcomes o
            JOIN knowledge_studies s ON s.study_id=o.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "OBSERVATION": """
            SELECT v.observation_uid
            FROM knowledge_observations v
            JOIN knowledge_outcomes o ON o.outcome_id=v.outcome_id
            JOIN knowledge_studies s ON s.study_id=o.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "MEASUREMENT_SERIES": """
            SELECT ms.series_uid
            FROM knowledge_measurement_series ms
            JOIN knowledge_studies s ON s.study_id=ms.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "MEASUREMENT_POINT": """
            SELECT mp.point_uid
            FROM knowledge_measurement_points mp
            JOIN knowledge_measurement_series ms
              ON ms.series_id=mp.series_id
            JOIN knowledge_studies s ON s.study_id=ms.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "COMPARISON": """
            SELECT c.comparison_uid
            FROM knowledge_comparisons c
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "EFFECT": """
            SELECT e.effect_uid
            FROM knowledge_effects e
            JOIN knowledge_comparisons c ON c.comparison_id=e.comparison_id
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
        """,
        "CLAIM": """
            SELECT c.claim_uid
            FROM knowledge_claims c
            WHERE c.workbook_analysis_id=?
        """,
    }
    result = {"WORKBOOK_ANALYSIS": {analysis_uid}}
    for entity_type, sql in queries.items():
        result[entity_type] = {
            str(row[0])
            for row in conn.execute(sql, (workbook_analysis_id,))
        }
    return result


_QUARANTINE_CHILD_TARGETS = (
    (
        "EFFECT",
        "knowledge_effects",
        "effect_uid",
        """
        comparison_id IN (
            SELECT c.comparison_id
            FROM knowledge_comparisons c
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
        )
        """,
        True,
    ),
    (
        "COMPARISON",
        "knowledge_comparisons",
        "comparison_uid",
        """
        study_id IN (
            SELECT study_id FROM knowledge_studies
            WHERE workbook_analysis_id=?
        )
        """,
        True,
    ),
    (
        "MEASUREMENT_POINT",
        "knowledge_measurement_points",
        "point_uid",
        """
        series_id IN (
            SELECT ms.series_id
            FROM knowledge_measurement_series ms
            JOIN knowledge_studies s ON s.study_id=ms.study_id
            WHERE s.workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "MEASUREMENT_SERIES",
        "knowledge_measurement_series",
        "series_uid",
        """
        study_id IN (
            SELECT study_id FROM knowledge_studies
            WHERE workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "OBSERVATION",
        "knowledge_observations",
        "observation_uid",
        """
        outcome_id IN (
            SELECT o.outcome_id
            FROM knowledge_outcomes o
            JOIN knowledge_studies s ON s.study_id=o.study_id
            WHERE s.workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "CLAIM",
        "knowledge_claims",
        "claim_uid",
        "workbook_analysis_id=?",
        False,
    ),
    (
        "OUTCOME",
        "knowledge_outcomes",
        "outcome_uid",
        """
        study_id IN (
            SELECT study_id FROM knowledge_studies
            WHERE workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "ARM",
        "knowledge_arms",
        "arm_uid",
        """
        study_id IN (
            SELECT study_id FROM knowledge_studies
            WHERE workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "FACTOR",
        "knowledge_factors",
        "factor_uid",
        """
        study_id IN (
            SELECT study_id FROM knowledge_studies
            WHERE workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "CONTEXT",
        "knowledge_study_contexts",
        "context_uid",
        """
        study_id IN (
            SELECT study_id FROM knowledge_studies
            WHERE workbook_analysis_id=?
        )
        """,
        False,
    ),
    (
        "STUDY",
        "knowledge_studies",
        "study_uid",
        "workbook_analysis_id=?",
        False,
    ),
)


def quarantine_canonical_analysis(
    conn: sqlite3.Connection,
    *,
    public_analysis_id: str,
    reason: str,
    now_iso: Callable[[], str],
) -> dict[str, Any]:
    """Atomically hide one unverified current canonical analysis from queries."""

    analysis_id = str(public_analysis_id or "").strip()
    quarantine_reason = str(reason or "").strip()
    if not analysis_id:
        raise AnalysisQuarantineError("publicAnalysisId is required")
    if not quarantine_reason:
        raise AnalysisQuarantineError("quarantine reason must not be empty")

    conn.execute("SAVEPOINT canonical_analysis_quarantine")
    try:
        row = conn.execute(
            """
            SELECT
                wa.workbook_analysis_id, wa.analysis_uid,
                wa.public_analysis_id, wa.analysis_status,
                wa.verification_status, wa.analyzer_name,
                wa.legacy_analysis_report_id, r.is_current,
                d.lifecycle_status
            FROM workbook_analyses wa
            JOIN source_revisions r ON r.revision_id=wa.revision_id
            JOIN source_documents d ON d.document_id=wa.document_id
            WHERE wa.public_analysis_id=?
            """,
            (analysis_id,),
        ).fetchone()
        if row is None:
            raise AnalysisQuarantineError(
                f"canonical analysis not found: {analysis_id}"
            )
        if (
            str(row["analyzer_name"]) != "canonical-study-import"
            or row["legacy_analysis_report_id"] is not None
        ):
            raise AnalysisQuarantineError(
                f"analysis is not a canonical study import: {analysis_id}"
            )
        if (
            int(row["is_current"]) != 1
            or str(row["lifecycle_status"]).upper() != "ACTIVE"
        ):
            raise AnalysisQuarantineError(
                f"analysis is not current and active: {analysis_id}"
            )
        analysis_status = str(row["analysis_status"]).upper()
        verification_status = str(row["verification_status"]).upper()
        if verification_status == "VERIFIED":
            raise AnalysisQuarantineError(
                f"refusing to quarantine VERIFIED analysis: {analysis_id}"
            )
        if (
            verification_status == "STALE"
            or analysis_status == "STALE"
        ):
            raise AnalysisQuarantineError(
                f"analysis is already quarantined: {analysis_id}"
            )

        workbook_analysis_id = int(row["workbook_analysis_id"])
        analysis_uid = str(row["analysis_uid"])
        entity_uids = _analysis_entity_uids(
            conn,
            workbook_analysis_id=workbook_analysis_id,
            analysis_uid=analysis_uid,
        )
        protected_entities = {
            (entity_type, entity_uid)
            for entity_type, values in entity_uids.items()
            for entity_uid in values
        }
        decision = next(
            (
                decision_row
                for decision_row in conn.execute(
                    """
                    SELECT entity_type, entity_uid
                    FROM review_decisions
                    ORDER BY review_decision_id
                    """
                )
                if (
                    str(decision_row["entity_type"]).upper(),
                    str(decision_row["entity_uid"]),
                )
                in protected_entities
            ),
            None,
        )
        if decision is not None:
            raise AnalysisQuarantineError(
                "refusing to quarantine analysis with review decision on "
                f"{str(decision['entity_type']).upper()}:"
                f"{decision['entity_uid']}"
            )

        for (
            entity_type,
            table,
            uid_column,
            predicate,
            _clear_aggregation,
        ) in _QUARANTINE_CHILD_TARGETS:
            verified_child = conn.execute(
                f"""
                SELECT {uid_column}
                FROM {table}
                WHERE {predicate}
                  AND verification_status='VERIFIED'
                LIMIT 1
                """,
                (workbook_analysis_id,),
            ).fetchone()
            if verified_child is not None:
                raise AnalysisQuarantineError(
                    "refusing to quarantine analysis with VERIFIED child "
                    f"{entity_type}:{verified_child[0]}"
                )

        updated_children: dict[str, int] = {}
        for (
            entity_type,
            table,
            _uid_column,
            predicate,
            clear_aggregation,
        ) in _QUARANTINE_CHILD_TARGETS:
            assignments = ["verification_status='STALE'"]
            if clear_aggregation:
                assignments.append("aggregation_eligible=0")
            cursor = conn.execute(
                f"""
                UPDATE {table}
                SET {", ".join(assignments)}
                WHERE {predicate}
                """,
                (workbook_analysis_id,),
            )
            updated_children[entity_type] = int(cursor.rowcount)

        quarantined_at = now_iso()
        conn.execute(
            """
            UPDATE workbook_analyses
            SET analysis_status='STALE',
                verification_status='STALE',
                updated_at=?
            WHERE workbook_analysis_id=?
            """,
            (quarantined_at, workbook_analysis_id),
        )
        issue_uid = stable_uid(
            "validation-issue",
            _QUARANTINE_VALIDATOR_NAME,
            analysis_uid,
            quarantined_at,
            quarantine_reason,
        )
        conn.execute(
            """
            INSERT INTO validation_issues(
                issue_uid, entity_type, entity_uid, issue_code, severity,
                message, details_json, status, validator_name,
                validator_version, created_at
            ) VALUES (?, 'WORKBOOK_ANALYSIS', ?, ?, 'BLOCKING', ?, ?,
                      'OPEN', ?, ?, ?)
            """,
            (
                issue_uid,
                analysis_uid,
                "MANUAL_QUARANTINE_"
                + issue_uid[-12:].upper(),
                quarantine_reason,
                _json(
                    {
                        "publicAnalysisId": analysis_id,
                        "reason": quarantine_reason,
                    },
                    {},
                ),
                _QUARANTINE_VALIDATOR_NAME,
                _QUARANTINE_VALIDATOR_VERSION,
                quarantined_at,
            ),
        )
        result = {
            "publicAnalysisId": analysis_id,
            "analysisUid": analysis_uid,
            "reason": quarantine_reason,
            "quarantinedAt": quarantined_at,
            "analysisStatus": "STALE",
            "verificationStatus": "STALE",
            "updatedChildren": updated_children,
            "preservedEvidence": True,
        }
    except Exception:
        conn.execute(
            "ROLLBACK TO SAVEPOINT canonical_analysis_quarantine"
        )
        conn.execute("RELEASE SAVEPOINT canonical_analysis_quarantine")
        raise
    conn.execute("RELEASE SAVEPOINT canonical_analysis_quarantine")
    return result


def _synchronize_analysis_children(
    conn: sqlite3.Connection,
    *,
    workbook_analysis_id: int,
    analysis_uid: str,
    desired: dict[str, set[str]],
) -> None:
    existing = _analysis_entity_uids(
        conn,
        workbook_analysis_id=workbook_analysis_id,
        analysis_uid=analysis_uid,
    )

    # Importer-generated evidence links are an authoritative projection of the
    # current manifest. Preserve independently authored links on entities that
    # remain current; links for entities about to be deleted cannot remain.
    # Evidence rows themselves are retained because the same source evidence may
    # be linked elsewhere and has its own stable public identity.
    for entity_type, entity_uids in existing.items():
        for entity_uid in entity_uids:
            if entity_uid not in desired[entity_type]:
                conn.execute(
                    "DELETE FROM entity_evidence_links WHERE entity_type=? AND entity_uid=?",
                    (entity_type, entity_uid),
                )
            else:
                linked_evidence = conn.execute(
                    """
                    SELECT
                        e.evidence_id, e.evidence_uid, e.sheet_name,
                        e.range_address, e.evidence_role, r.revision_uid
                    FROM entity_evidence_links l
                    JOIN evidence_items e ON e.evidence_id=l.evidence_id
                    JOIN source_revisions r ON r.revision_id=e.revision_id
                    WHERE l.entity_type=? AND l.entity_uid=?
                    """,
                    (entity_type, entity_uid),
                ).fetchall()
                for evidence in linked_evidence:
                    expected_uid = stable_uid(
                        "evidence",
                        evidence["revision_uid"],
                        entity_type,
                        entity_uid,
                        evidence["sheet_name"],
                        evidence["range_address"],
                        evidence["evidence_role"],
                    )
                    if str(evidence["evidence_uid"]) != expected_uid:
                        continue
                    conn.execute(
                        """
                        DELETE FROM entity_evidence_links
                        WHERE entity_type=? AND entity_uid=? AND evidence_id=?
                        """,
                        (
                            entity_type,
                            entity_uid,
                            int(evidence["evidence_id"]),
                        ),
                    )
            conn.execute(
                """
                DELETE FROM validation_issues
                WHERE entity_type=? AND entity_uid=? AND validator_name=?
                """,
                (entity_type, entity_uid, _IMPORT_VALIDATOR_NAME),
            )

    # The join table has no manifest identity of its own, so rebuild it exactly
    # from the current arms' factorValues during the following upsert pass.
    conn.execute(
        """
        DELETE FROM knowledge_arm_factor_values
        WHERE arm_id IN (
            SELECT a.arm_id
            FROM knowledge_arms a
            JOIN knowledge_studies s ON s.study_id=a.study_id
            WHERE s.workbook_analysis_id=?
        )
        """,
        (workbook_analysis_id,),
    )

    # Delete only rows owned by this analysis and absent from the new manifest.
    # Dependency order keeps behavior correct even when foreign keys are disabled
    # by a caller, while normal callers still benefit from cascade enforcement.
    delete_order = (
        ("EFFECT", "knowledge_effects", "effect_uid"),
        ("COMPARISON", "knowledge_comparisons", "comparison_uid"),
        (
            "MEASUREMENT_POINT",
            "knowledge_measurement_points",
            "point_uid",
        ),
        (
            "MEASUREMENT_SERIES",
            "knowledge_measurement_series",
            "series_uid",
        ),
        ("OBSERVATION", "knowledge_observations", "observation_uid"),
        ("CLAIM", "knowledge_claims", "claim_uid"),
        ("OUTCOME", "knowledge_outcomes", "outcome_uid"),
        ("ARM", "knowledge_arms", "arm_uid"),
        ("FACTOR", "knowledge_factors", "factor_uid"),
        ("CONTEXT", "knowledge_study_contexts", "context_uid"),
        ("STUDY", "knowledge_studies", "study_uid"),
    )
    for entity_type, table, uid_column in delete_order:
        stale_uids = existing[entity_type] - desired[entity_type]
        for entity_uid in stale_uids:
            conn.execute(
                f"DELETE FROM {table} WHERE {uid_column}=?",
                (entity_uid,),
            )


def _supersede_other_canonical_analyses(
    conn: sqlite3.Connection,
    *,
    revision_id: int,
    current_analysis_uid: str,
    now_iso: Callable[[], str],
) -> dict[str, int]:
    result = {"deleted": 0, "preservedStale": 0}
    rows = conn.execute(
        """
        SELECT
            workbook_analysis_id, analysis_uid, verification_status
        FROM workbook_analyses
        WHERE revision_id=?
          AND analysis_uid<>?
          AND legacy_analysis_report_id IS NULL
          AND analyzer_name='canonical-study-import'
        ORDER BY workbook_analysis_id
        """,
        (revision_id, current_analysis_uid),
    ).fetchall()
    for row in rows:
        workbook_analysis_id = int(row["workbook_analysis_id"])
        analysis_uid = str(row["analysis_uid"])
        existing = _analysis_entity_uids(
            conn,
            workbook_analysis_id=workbook_analysis_id,
            analysis_uid=analysis_uid,
        )
        protected = str(row["verification_status"]).upper() == "VERIFIED"
        if not protected:
            for entity_type, entity_uids in existing.items():
                if any(
                    conn.execute(
                        """
                        SELECT 1
                        FROM review_decisions
                        WHERE entity_type=? AND entity_uid=?
                        LIMIT 1
                        """,
                        (entity_type, entity_uid),
                    ).fetchone()
                    is not None
                    for entity_uid in entity_uids
                ):
                    protected = True
                    break
        if not protected:
            protected = (
                conn.execute(
                    """
                    SELECT 1
                    FROM knowledge_studies s
                    WHERE s.workbook_analysis_id=?
                      AND s.verification_status='VERIFIED'
                    UNION ALL
                    SELECT 1
                    FROM knowledge_comparisons c
                    JOIN knowledge_studies s ON s.study_id=c.study_id
                    WHERE s.workbook_analysis_id=?
                      AND c.verification_status='VERIFIED'
                    UNION ALL
                    SELECT 1
                    FROM knowledge_effects e
                    JOIN knowledge_comparisons c
                      ON c.comparison_id=e.comparison_id
                    JOIN knowledge_studies s ON s.study_id=c.study_id
                    WHERE s.workbook_analysis_id=?
                      AND e.verification_status='VERIFIED'
                    LIMIT 1
                    """,
                    (
                        workbook_analysis_id,
                        workbook_analysis_id,
                        workbook_analysis_id,
                    ),
                ).fetchone()
                is not None
            )
        if protected:
            conn.execute(
                """
                UPDATE workbook_analyses
                SET analysis_status='STALE',
                    verification_status='STALE',
                    updated_at=?
                WHERE workbook_analysis_id=?
                """,
                (now_iso(), workbook_analysis_id),
            )
            result["preservedStale"] += 1
            continue

        empty_projection = {
            entity_type: set()
            for entity_type in existing
        }
        _synchronize_analysis_children(
            conn,
            workbook_analysis_id=workbook_analysis_id,
            analysis_uid=analysis_uid,
            desired=empty_projection,
        )
        conn.execute(
            """
            DELETE FROM workbook_analyses
            WHERE workbook_analysis_id=?
            """,
            (workbook_analysis_id,),
        )
        result["deleted"] += 1
    return result


def _record_validation_issue(
    conn: sqlite3.Connection,
    *,
    entity_type: str,
    entity_uid: str,
    issue_code: str,
    severity: str,
    message: str,
    now_iso: Callable[[], str],
) -> None:
    issue_uid = stable_uid(
        "validation-issue",
        _IMPORT_VALIDATOR_NAME,
        entity_type,
        entity_uid,
        issue_code,
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO validation_issues(
            issue_uid, entity_type, entity_uid, issue_code, severity, message,
            details_json, status, validator_name, validator_version, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, '{}', 'OPEN', ?, ?, ?)
        """,
        (
            issue_uid,
            entity_type,
            entity_uid,
            issue_code,
            severity,
            message,
            _IMPORT_VALIDATOR_NAME,
            _IMPORT_VALIDATOR_VERSION,
            now_iso(),
        ),
    )


def _record_analysis_validation_issues(
    conn: sqlite3.Connection,
    *,
    workbook_analysis_id: int,
    now_iso: Callable[[], str],
) -> None:
    status_queries = (
        (
            "WORKBOOK_ANALYSIS",
            "SELECT analysis_uid, verification_status FROM workbook_analyses WHERE workbook_analysis_id=?",
        ),
        (
            "STUDY",
            """
            SELECT s.study_uid, s.verification_status
            FROM knowledge_studies s
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "CONTEXT",
            """
            SELECT c.context_uid, c.verification_status
            FROM knowledge_study_contexts c
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "FACTOR",
            """
            SELECT f.factor_uid, f.verification_status
            FROM knowledge_factors f
            JOIN knowledge_studies s ON s.study_id=f.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "ARM",
            """
            SELECT a.arm_uid, a.verification_status
            FROM knowledge_arms a
            JOIN knowledge_studies s ON s.study_id=a.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "OUTCOME",
            """
            SELECT o.outcome_uid, o.verification_status
            FROM knowledge_outcomes o
            JOIN knowledge_studies s ON s.study_id=o.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "OBSERVATION",
            """
            SELECT v.observation_uid, v.verification_status
            FROM knowledge_observations v
            JOIN knowledge_outcomes o ON o.outcome_id=v.outcome_id
            JOIN knowledge_studies s ON s.study_id=o.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "MEASUREMENT_SERIES",
            """
            SELECT ms.series_uid, ms.verification_status
            FROM knowledge_measurement_series ms
            JOIN knowledge_studies s ON s.study_id=ms.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "COMPARISON",
            """
            SELECT c.comparison_uid,
                   CASE
                       WHEN c.validity_status IN ('NEEDS_REVIEW','EXCLUDED')
                       THEN c.validity_status
                       ELSE c.verification_status
                   END
            FROM knowledge_comparisons c
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "EFFECT",
            """
            SELECT e.effect_uid, e.verification_status
            FROM knowledge_effects e
            JOIN knowledge_comparisons c ON c.comparison_id=e.comparison_id
            JOIN knowledge_studies s ON s.study_id=c.study_id
            WHERE s.workbook_analysis_id=?
            """,
        ),
        (
            "CLAIM",
            """
            SELECT c.claim_uid, c.verification_status
            FROM knowledge_claims c
            WHERE c.workbook_analysis_id=?
            """,
        ),
    )
    for entity_type, sql in status_queries:
        for entity_uid, status in conn.execute(sql, (workbook_analysis_id,)):
            normalized_status = str(status or "").upper()
            if normalized_status == "NEEDS_REVIEW":
                _record_validation_issue(
                    conn,
                    entity_type=entity_type,
                    entity_uid=str(entity_uid),
                    issue_code="NEEDS_REVIEW",
                    severity="WARNING",
                    message="The imported entity requires human review before answer eligibility.",
                    now_iso=now_iso,
                )
            elif normalized_status == "EXCLUDED":
                _record_validation_issue(
                    conn,
                    entity_type=entity_type,
                    entity_uid=str(entity_uid),
                    issue_code="EXCLUDED",
                    severity="INFO",
                    message="The imported entity is explicitly excluded from answer eligibility.",
                    now_iso=now_iso,
                )

    for study_uid, confounding_status, comparison_count in conn.execute(
        """
        SELECT s.study_uid, s.confounding_status, COUNT(c.comparison_id)
        FROM knowledge_studies s
        LEFT JOIN knowledge_comparisons c ON c.study_id=s.study_id
        WHERE s.workbook_analysis_id=?
        GROUP BY s.study_id
        """,
        (workbook_analysis_id,),
    ):
        if int(comparison_count) == 0:
            _record_validation_issue(
                conn,
                entity_type="STUDY",
                entity_uid=str(study_uid),
                issue_code="NO_COMPARISON",
                severity="WARNING",
                message="No explicit control/comparison record is available; results are descriptive only.",
                now_iso=now_iso,
            )
        _record_confounding_issue(
            conn,
            entity_type="STUDY",
            entity_uid=str(study_uid),
            confounding_status=str(confounding_status),
            now_iso=now_iso,
        )

    for comparison_uid, confounding_status in conn.execute(
        """
        SELECT c.comparison_uid, c.confounding_status
        FROM knowledge_comparisons c
        JOIN knowledge_studies s ON s.study_id=c.study_id
        WHERE s.workbook_analysis_id=?
        """,
        (workbook_analysis_id,),
    ):
        _record_confounding_issue(
            conn,
            entity_type="COMPARISON",
            entity_uid=str(comparison_uid),
            confounding_status=str(confounding_status),
            now_iso=now_iso,
        )


def _record_confounding_issue(
    conn: sqlite3.Connection,
    *,
    entity_type: str,
    entity_uid: str,
    confounding_status: str,
    now_iso: Callable[[], str],
) -> None:
    normalized = confounding_status.upper()
    issue = {
        "POSSIBLE": (
            "CONFOUNDING_POSSIBLE",
            "WARNING",
            "Possible confounding prevents unqualified aggregation.",
        ),
        "CONFOUNDED": (
            "CONFOUNDED",
            "ERROR",
            "The entity is confounded and is not eligible for aggregation.",
        ),
        "UNASSESSED": (
            "CONFOUNDING_UNASSESSED",
            "WARNING",
            "Confounding has not been assessed.",
        ),
    }.get(normalized)
    if issue is None:
        return
    issue_code, severity, message = issue
    _record_validation_issue(
        conn,
        entity_type=entity_type,
        entity_uid=entity_uid,
        issue_code=issue_code,
        severity=severity,
        message=message,
        now_iso=now_iso,
    )


def import_study_manifest(
    conn: sqlite3.Connection,
    data: dict[str, Any],
    *,
    now_iso: Callable[[], str],
    formula_overlay: dict[str, Any] | None = None,
    source_claims_prevalidated: bool = False,
) -> dict[str, Any]:
    """Import a canonical manifest after validating its schema and claims.

    ``source_claims_prevalidated`` is reserved for the incremental workflow
    when the exact in-memory manifest has already passed the same database-
    backed source-claim validators against its resolved revision. Direct
    callers retain the fail-closed validation path by default.
    """
    revision = resolve_manifest_revision(conn, data["source"])
    formula_lookup = _formula_lookup_for_revision(revision, formula_overlay)
    if source_claims_prevalidated:
        manifest = validate_study_manifest(data)
    else:
        checker = make_database_evidence_checker(conn, revision)
        manifest = validate_study_manifest(data, evidence_checker=checker)
        validate_numeric_observation_evidence(
            conn,
            revision,
            manifest,
            formula_overlay=formula_lookup,
        )
        validate_factor_and_arm_evidence(conn, revision, manifest)
        validate_comparison_representation_alignment(
            conn,
            revision,
            manifest,
            formula_overlay=formula_lookup,
        )
        validate_conclusion_evidence(conn, revision, manifest)
    source = manifest["source"]
    analysis = manifest["workbookAnalysis"]
    analysis_uid = stable_uid("analysis", source["dataset"], source["sourcePath"], analysis["key"])
    desired_entity_uids = _manifest_entity_uids(manifest, analysis_uid)
    verification_status = str(analysis.get("verificationStatus") or "NEEDS_REVIEW").upper()

    conn.execute("SAVEPOINT canonical_study_import")
    try:
        superseded = _supersede_other_canonical_analyses(
            conn,
            revision_id=int(revision["revision_id"]),
            current_analysis_uid=analysis_uid,
            now_iso=now_iso,
        )
        conn.execute(
            """
            INSERT INTO workbook_analyses(
                analysis_uid, public_analysis_id, document_id, revision_id,
                legacy_analysis_report_id, analysis_key, title, analysis_type,
                purpose, scope_text, analysis_status, verification_status,
                decision_text, consolidated_summary, limitations_json,
                analyzer_name, analyzer_version, created_at, updated_at
            ) VALUES (?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                      'canonical-study-import', '1', ?, ?)
            ON CONFLICT(analysis_uid) DO UPDATE SET
                revision_id=excluded.revision_id,
                title=excluded.title,
                analysis_type=excluded.analysis_type,
                purpose=excluded.purpose,
                scope_text=excluded.scope_text,
                analysis_status=excluded.analysis_status,
                verification_status=excluded.verification_status,
                decision_text=excluded.decision_text,
                consolidated_summary=excluded.consolidated_summary,
                limitations_json=excluded.limitations_json,
                analyzer_name=excluded.analyzer_name,
                analyzer_version=excluded.analyzer_version,
                updated_at=excluded.updated_at
            """,
            (
                analysis_uid,
                public_id("ANALYSIS", analysis_uid),
                int(revision["document_id"]),
                int(revision["revision_id"]),
                str(analysis["key"]),
                str(analysis["title"]),
                str(analysis.get("type") or ""),
                str(analysis.get("purpose") or ""),
                str(analysis.get("scope") or ""),
                str(analysis.get("status") or verification_status),
                verification_status,
                str(analysis.get("decision") or ""),
                str(analysis.get("summary") or ""),
                _json(analysis.get("limitations"), []),
                now_iso(),
                now_iso(),
            ),
        )
        workbook_analysis_id = int(
            conn.execute("SELECT workbook_analysis_id FROM workbook_analyses WHERE analysis_uid=?", (analysis_uid,)).fetchone()[0]
        )
        conn.execute(
            """
            UPDATE validation_issues
            SET status='RESOLVED', resolved_at=?
            WHERE entity_type='WORKBOOK_ANALYSIS'
              AND entity_uid=?
              AND validator_name=?
              AND status='OPEN'
            """,
            (now_iso(), analysis_uid, _QUARANTINE_VALIDATOR_NAME),
        )
        _synchronize_analysis_children(
            conn,
            workbook_analysis_id=workbook_analysis_id,
            analysis_uid=analysis_uid,
            desired=desired_entity_uids,
        )
        evidence_count = _store_evidence(
            conn,
            revision=revision,
            entity_type="WORKBOOK_ANALYSIS",
            entity_uid=analysis_uid,
            evidence_items=analysis.get("evidence", []),
            now_iso=now_iso,
        )

        counts = {
            "studies": 0,
            "contexts": 0,
            "factors": 0,
            "arms": 0,
            "outcomes": 0,
            "observations": 0,
            "measurementSeries": 0,
            "measurementPoints": 0,
            "comparisons": 0,
            "effects": 0,
            "claims": 0,
        }
        for study in manifest["studies"]:
            study_uid = stable_uid("study", analysis_uid, study["key"])
            conn.execute(
                """
                INSERT INTO knowledge_studies(
                    study_uid, public_data_id, workbook_analysis_id,
                    legacy_review_item_id, study_key, title, purpose, hypothesis,
                    objective, design_type, comparison_basis, analysis_status,
                    verification_status, comparability_status, confounding_status,
                    decision_text, summary_text, limitations_json, created_at, updated_at
                ) VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(study_uid) DO UPDATE SET
                    title=excluded.title,
                    purpose=excluded.purpose,
                    hypothesis=excluded.hypothesis,
                    objective=excluded.objective,
                    design_type=excluded.design_type,
                    comparison_basis=excluded.comparison_basis,
                    analysis_status=excluded.analysis_status,
                    verification_status=excluded.verification_status,
                    comparability_status=excluded.comparability_status,
                    confounding_status=excluded.confounding_status,
                    decision_text=excluded.decision_text,
                    summary_text=excluded.summary_text,
                    limitations_json=excluded.limitations_json,
                    updated_at=excluded.updated_at
                """,
                (
                    study_uid,
                    public_id("DATA", study_uid),
                    workbook_analysis_id,
                    str(study["key"]),
                    str(study["title"]),
                    str(study.get("purpose") or ""),
                    str(study.get("hypothesis") or ""),
                    str(study.get("objective") or ""),
                    str(study["designType"]),
                    str(study.get("comparisonBasis") or ""),
                    str(study.get("status") or study["verificationStatus"]),
                    str(study["verificationStatus"]).upper(),
                    str(study["comparabilityStatus"]).upper(),
                    str(study["confoundingStatus"]).upper(),
                    str(study.get("decision") or ""),
                    str(study.get("summary") or ""),
                    _json(study.get("limitations"), []),
                    now_iso(),
                    now_iso(),
                ),
            )
            study_id = int(conn.execute("SELECT study_id FROM knowledge_studies WHERE study_uid=?", (study_uid,)).fetchone()[0])
            counts["studies"] += 1
            evidence_count += _store_evidence(
                conn,
                revision=revision,
                entity_type="STUDY",
                entity_uid=study_uid,
                evidence_items=study.get("evidence", []),
                now_iso=now_iso,
            )

            factor_ids: dict[str, int] = {}
            factor_uids: dict[str, str] = {}
            for context in study.get("contexts", []):
                context_uid = stable_uid("context", study_uid, context["key"])
                concept_id = _concept_id(
                    conn,
                    context,
                    default_kind=str(context["kind"]).upper(),
                    original_value=context["originalValue"],
                    now_iso=now_iso,
                )
                unit_id = _unit_id(conn, context.get("unit"), now_iso=now_iso)
                conn.execute(
                    """
                    INSERT INTO knowledge_study_contexts(
                        context_uid, study_id, context_kind, concept_id,
                        original_value, normalized_value, value_number, unit_id,
                        start_value, end_value, attributes_json, verification_status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(context_uid) DO UPDATE SET
                        concept_id=excluded.concept_id,
                        original_value=excluded.original_value,
                        normalized_value=excluded.normalized_value,
                        value_number=excluded.value_number,
                        unit_id=excluded.unit_id,
                        start_value=excluded.start_value,
                        end_value=excluded.end_value,
                        attributes_json=excluded.attributes_json,
                        verification_status=excluded.verification_status
                    """,
                    (
                        context_uid,
                        study_id,
                        str(context["kind"]).upper(),
                        concept_id,
                        str(context["originalValue"]),
                        str(context.get("normalizedValue") or normalize_key_part(context["originalValue"])),
                        context.get("valueNumber"),
                        unit_id,
                        str(context.get("startValue") or ""),
                        str(context.get("endValue") or ""),
                        _json(context.get("attributes"), {}),
                        str(context.get("verificationStatus") or study["verificationStatus"]).upper(),
                    ),
                )
                counts["contexts"] += 1
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="CONTEXT",
                    entity_uid=context_uid,
                    evidence_items=context.get("evidence", []),
                    now_iso=now_iso,
                )

            for factor in study.get("factors", []):
                factor_uid = stable_uid("factor", study_uid, factor["key"])
                concept_id = _concept_id(
                    conn,
                    factor,
                    default_kind="CHANGED_FACTOR",
                    original_value=factor["originalLabel"],
                    now_iso=now_iso,
                )
                conn.execute(
                    """
                    INSERT INTO knowledge_factors(
                        factor_uid, study_id, concept_id, factor_key, factor_domain,
                        original_label, baseline_condition, changed_condition,
                        change_direction, isolation_status, verification_status,
                        attributes_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(factor_uid) DO UPDATE SET
                        concept_id=excluded.concept_id,
                        factor_domain=excluded.factor_domain,
                        original_label=excluded.original_label,
                        baseline_condition=excluded.baseline_condition,
                        changed_condition=excluded.changed_condition,
                        change_direction=excluded.change_direction,
                        isolation_status=excluded.isolation_status,
                        verification_status=excluded.verification_status,
                        attributes_json=excluded.attributes_json
                    """,
                    (
                        factor_uid,
                        study_id,
                        concept_id,
                        str(factor["key"]),
                        str(factor.get("domain") or ""),
                        str(factor["originalLabel"]),
                        str(factor["baselineCondition"]),
                        str(factor["changedCondition"]),
                        str(factor.get("changeDirection") or ""),
                        str(factor["isolationStatus"]).upper(),
                        str(factor.get("verificationStatus") or study["verificationStatus"]).upper(),
                        _json(factor.get("attributes"), {}),
                    ),
                )
                factor_id = int(conn.execute("SELECT factor_id FROM knowledge_factors WHERE factor_uid=?", (factor_uid,)).fetchone()[0])
                factor_ids[str(factor["key"])] = factor_id
                factor_uids[str(factor["key"])] = factor_uid
                counts["factors"] += 1
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="FACTOR",
                    entity_uid=factor_uid,
                    evidence_items=factor.get("evidence", []),
                    now_iso=now_iso,
                )

            arm_ids: dict[str, int] = {}
            arm_uids: dict[str, str] = {}
            for arm in study.get("arms", []):
                arm_uid = stable_uid("arm", study_uid, arm["key"])
                conn.execute(
                    """
                    INSERT INTO knowledge_arms(
                        arm_uid, study_id, legacy_cohort_id, arm_key, arm_role,
                        label, condition_text, sample_size, sample_basis,
                        matching_basis, attributes_json, verification_status
                    ) VALUES (?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(arm_uid) DO UPDATE SET
                        arm_role=excluded.arm_role,
                        label=excluded.label,
                        condition_text=excluded.condition_text,
                        sample_size=excluded.sample_size,
                        sample_basis=excluded.sample_basis,
                        matching_basis=excluded.matching_basis,
                        attributes_json=excluded.attributes_json,
                        verification_status=excluded.verification_status
                    """,
                    (
                        arm_uid,
                        study_id,
                        str(arm["key"]),
                        str(arm["role"]).upper(),
                        str(arm["label"]),
                        str(arm.get("condition") or ""),
                        arm.get("sampleSize"),
                        str(arm.get("sampleBasis") or ""),
                        str(arm.get("matchingBasis") or ""),
                        _json(arm.get("attributes"), {}),
                        str(arm.get("verificationStatus") or study["verificationStatus"]).upper(),
                    ),
                )
                arm_id = int(conn.execute("SELECT arm_id FROM knowledge_arms WHERE arm_uid=?", (arm_uid,)).fetchone()[0])
                arm_ids[str(arm["key"])] = arm_id
                arm_uids[str(arm["key"])] = arm_uid
                counts["arms"] += 1
                for factor_value in arm.get("factorValues", []):
                    conn.execute(
                        """
                        INSERT INTO knowledge_arm_factor_values(
                            arm_id, factor_id, original_value, value_number,
                            unit_id, is_baseline, held_constant
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(arm_id, factor_id) DO UPDATE SET
                            original_value=excluded.original_value,
                            value_number=excluded.value_number,
                            unit_id=excluded.unit_id,
                            is_baseline=excluded.is_baseline,
                            held_constant=excluded.held_constant
                        """,
                        (
                            arm_id,
                            factor_ids[str(factor_value["factor"])],
                            str(factor_value.get("value") or ""),
                            factor_value.get("valueNumber"),
                            _unit_id(conn, factor_value.get("unit"), now_iso=now_iso),
                            1 if factor_value.get("isBaseline") else 0,
                            1 if factor_value.get("heldConstant") else 0,
                        ),
                    )
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="ARM",
                    entity_uid=arm_uid,
                    evidence_items=arm.get("evidence", []),
                    now_iso=now_iso,
                )

            outcome_ids: dict[str, int] = {}
            outcome_uids: dict[str, str] = {}
            for outcome in study.get("outcomes", []):
                outcome_uid = stable_uid("outcome", study_uid, outcome["key"])
                concept_id = _concept_id(
                    conn,
                    outcome,
                    default_kind="OUTCOME",
                    original_value=outcome["originalLabel"],
                    now_iso=now_iso,
                )
                unit_id = _unit_id(conn, outcome.get("unit"), now_iso=now_iso)
                conn.execute(
                    """
                    INSERT INTO knowledge_outcomes(
                        outcome_uid, study_id, legacy_metric_id, outcome_key,
                        concept_id, original_label, outcome_domain, metric_type,
                        unit_id, original_unit, denominator_basis,
                        favorable_direction, definition_text, spec_text,
                        verification_status, attributes_json
                    ) VALUES (?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(outcome_uid) DO UPDATE SET
                        concept_id=excluded.concept_id,
                        original_label=excluded.original_label,
                        outcome_domain=excluded.outcome_domain,
                        metric_type=excluded.metric_type,
                        unit_id=excluded.unit_id,
                        original_unit=excluded.original_unit,
                        denominator_basis=excluded.denominator_basis,
                        favorable_direction=excluded.favorable_direction,
                        definition_text=excluded.definition_text,
                        spec_text=excluded.spec_text,
                        verification_status=excluded.verification_status,
                        attributes_json=excluded.attributes_json
                    """,
                    (
                        outcome_uid,
                        study_id,
                        str(outcome["key"]),
                        concept_id,
                        str(outcome["originalLabel"]),
                        str(outcome.get("domain") or ""),
                        str(outcome["metricType"]),
                        unit_id,
                        str(outcome.get("unit") or ""),
                        str(outcome.get("denominatorBasis") or ""),
                        str(outcome.get("favorableDirection") or "UNKNOWN").upper(),
                        str(outcome.get("definition") or ""),
                        str(outcome.get("spec") or ""),
                        str(outcome.get("verificationStatus") or study["verificationStatus"]).upper(),
                        _json(outcome.get("attributes"), {}),
                    ),
                )
                outcome_id = int(conn.execute("SELECT outcome_id FROM knowledge_outcomes WHERE outcome_uid=?", (outcome_uid,)).fetchone()[0])
                outcome_ids[str(outcome["key"])] = outcome_id
                outcome_uids[str(outcome["key"])] = outcome_uid
                counts["outcomes"] += 1
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="OUTCOME",
                    entity_uid=outcome_uid,
                    evidence_items=outcome.get("evidence", []),
                    now_iso=now_iso,
                )
                for observation in outcome.get("observations", []):
                    observation_uid = stable_uid("observation", outcome_uid, arm_uids[str(observation["arm"])], observation["key"])
                    stored_observation, normalized_percent_fields = (
                        _normalized_observation_values(
                            conn,
                            revision,
                            outcome,
                            observation,
                        )
                    )
                    observation_details = dict(
                        stored_observation.get("details") or {}
                    )
                    if normalized_percent_fields:
                        observation_details["sourcePercentScaleApplied"] = (
                            normalized_percent_fields
                        )
                    conn.execute(
                        """
                        INSERT INTO knowledge_observations(
                            observation_uid, outcome_id, arm_id, legacy_metric_value_id,
                            observation_key, stratum_key, replicate_key, observed_at,
                            value_number, value_text, numerator, denominator, rate_ppm,
                            min_value, max_value, average_value, sample_size,
                            result_status, verification_status, details_json
                        ) VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(observation_uid) DO UPDATE SET
                            stratum_key=excluded.stratum_key,
                            replicate_key=excluded.replicate_key,
                            observed_at=excluded.observed_at,
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
                            outcome_id,
                            arm_ids[str(observation["arm"])],
                            str(observation["key"]),
                            str(
                                observation.get("stratumKey")
                                or observation.get("stratum")
                                or ""
                            ),
                            str(
                                observation.get("replicateKey")
                                or observation.get("replicate")
                                or ""
                            ),
                            str(observation.get("observedAt") or ""),
                            stored_observation.get("valueNumber"),
                            str(stored_observation.get("valueText") or ""),
                            stored_observation.get("numerator"),
                            stored_observation.get("denominator"),
                            stored_observation.get("ratePpm"),
                            stored_observation.get("min"),
                            stored_observation.get("max"),
                            stored_observation.get("average"),
                            stored_observation.get("sampleSize")
                            or stored_observation.get("denominator"),
                            str(stored_observation.get("status") or ""),
                            str(
                                stored_observation.get("verificationStatus")
                                or study["verificationStatus"]
                            ).upper(),
                            _json(observation_details, {}),
                        ),
                    )
                    counts["observations"] += 1
                    evidence_count += _store_evidence(
                        conn,
                        revision=revision,
                        entity_type="OBSERVATION",
                        entity_uid=observation_uid,
                        evidence_items=observation.get("evidence", []),
                        now_iso=now_iso,
                    )

            for series_index, series in enumerate(
                study.get("measurementSeries", [])
            ):
                series_uid = stable_uid(
                    "measurement-series",
                    study_uid,
                    series["key"],
                )
                series_path = (
                    f"studies[{counts['studies'] - 1}]"
                    f".measurementSeries[{series_index}]"
                )
                points = _expand_measurement_series(
                    conn,
                    revision=revision,
                    series=series,
                    series_uid=series_uid,
                    path=series_path,
                    formula_lookup=formula_lookup,
                )
                axis_unit = str(series.get("axisUnit") or "")
                value_unit = str(series.get("valueUnit") or "")
                axis_unit_id = _unit_id(
                    conn,
                    axis_unit,
                    now_iso=now_iso,
                )
                value_unit_id = _unit_id(
                    conn,
                    value_unit,
                    now_iso=now_iso,
                )
                verification = str(
                    series.get("verificationStatus")
                    or "NEEDS_REVIEW"
                ).upper()
                series_details = series.get("details")
                if not isinstance(series_details, dict):
                    series_details = {}
                series_details = {
                    **series_details,
                    "seriesRole": str(
                        series.get("seriesRole") or "RAW"
                    ).upper(),
                    "aggregationFunction": str(
                        series.get("aggregationFunction") or ""
                    ).upper(),
                    "aggregateOfSeries": [
                        str(value)
                        for value in series.get(
                            "aggregateOfSeries",
                            [],
                        )
                    ],
                    "aggregateReplicateRanges": [
                        str(value)
                        for value in series.get(
                            "aggregateReplicateRanges",
                            [],
                        )
                    ],
                }
                conn.execute(
                    """
                    INSERT INTO knowledge_measurement_series(
                        series_uid, public_series_id, study_id, outcome_id,
                        arm_id, series_key, sheet_name, header_range,
                        value_range, row_identity_range, axis_name,
                        axis_source, axis_unit_id, original_axis_unit, value_unit_id,
                        original_value_unit, stratum_key,
                        verification_status, details_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                              ?, ?, ?, ?)
                    ON CONFLICT(series_uid) DO UPDATE SET
                        outcome_id=excluded.outcome_id,
                        arm_id=excluded.arm_id,
                        sheet_name=excluded.sheet_name,
                        header_range=excluded.header_range,
                        value_range=excluded.value_range,
                        row_identity_range=excluded.row_identity_range,
                        axis_name=excluded.axis_name,
                        axis_source=excluded.axis_source,
                        axis_unit_id=excluded.axis_unit_id,
                        original_axis_unit=excluded.original_axis_unit,
                        value_unit_id=excluded.value_unit_id,
                        original_value_unit=excluded.original_value_unit,
                        stratum_key=excluded.stratum_key,
                        verification_status=excluded.verification_status,
                        details_json=excluded.details_json
                    """,
                    (
                        series_uid,
                        public_id("SER", series_uid),
                        study_id,
                        outcome_ids[str(series["outcome"])],
                        arm_ids[str(series["arm"])],
                        str(series["key"]),
                        str(series["sheet"]),
                        str(series["headerRange"]),
                        str(series["valueRange"]),
                        str(series["rowIdentityRange"]),
                        str(series.get("axisLabel") or ""),
                        str(series["axisSource"]).upper(),
                        axis_unit_id,
                        axis_unit,
                        value_unit_id,
                        value_unit,
                        str(series.get("stratumKey") or ""),
                        verification,
                        _json(series_details, {}),
                    ),
                )
                series_id = int(
                    conn.execute(
                        """
                        SELECT series_id
                        FROM knowledge_measurement_series
                        WHERE series_uid=?
                        """,
                        (series_uid,),
                    ).fetchone()[0]
                )
                counts["measurementSeries"] += 1
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="MEASUREMENT_SERIES",
                    entity_uid=series_uid,
                    evidence_items=_measurement_series_evidence(series),
                    now_iso=now_iso,
                )
                # Matrix ranges may shift while retaining the same semantic
                # series key. Rebuild the deterministic point set before
                # inserting so old row/column ordinals cannot collide with
                # newly mapped coordinates. Stable point/public IDs still
                # derive from the series UID and exact source coordinate.
                conn.execute(
                    """
                    DELETE FROM knowledge_measurement_points
                    WHERE series_id=?
                    """,
                    (series_id,),
                )
                for point in points:
                    point_uid = str(point["pointUid"])
                    conn.execute(
                        """
                        INSERT INTO knowledge_measurement_points(
                            point_uid, public_point_id, series_id,
                            row_ordinal, column_ordinal, axis_label,
                            axis_value, axis_unit_id, original_axis_unit,
                            replicate_key, replicate_role, stratum_key,
                            value_number,
                            value_unit_id, original_value_unit,
                            source_revision_id, source_sheet_name,
                            source_row_index, source_column_index,
                            source_coordinate, axis_source_coordinate,
                            replicate_source_coordinate, source_value_json,
                            source_formula_text, verification_status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(point_uid) DO UPDATE SET
                            row_ordinal=excluded.row_ordinal,
                            column_ordinal=excluded.column_ordinal,
                            axis_label=excluded.axis_label,
                            axis_value=excluded.axis_value,
                            axis_unit_id=excluded.axis_unit_id,
                            original_axis_unit=excluded.original_axis_unit,
                            replicate_key=excluded.replicate_key,
                            replicate_role=excluded.replicate_role,
                            stratum_key=excluded.stratum_key,
                            value_number=excluded.value_number,
                            value_unit_id=excluded.value_unit_id,
                            original_value_unit=excluded.original_value_unit,
                            source_revision_id=excluded.source_revision_id,
                            source_sheet_name=excluded.source_sheet_name,
                            source_row_index=excluded.source_row_index,
                            source_column_index=excluded.source_column_index,
                            source_coordinate=excluded.source_coordinate,
                            axis_source_coordinate=
                                excluded.axis_source_coordinate,
                            replicate_source_coordinate=
                                excluded.replicate_source_coordinate,
                            source_value_json=excluded.source_value_json,
                            source_formula_text=excluded.source_formula_text,
                            verification_status=excluded.verification_status
                        """,
                        (
                            point_uid,
                            public_id("MPT", point_uid),
                            series_id,
                            point["rowOrdinal"],
                            point["columnOrdinal"],
                            point["axisLabel"],
                            point["axisValue"],
                            axis_unit_id,
                            axis_unit,
                            point["replicateKey"],
                            point["replicateRole"],
                            str(series.get("stratumKey") or ""),
                            point["valueNumber"],
                            value_unit_id,
                            value_unit,
                            int(revision["revision_id"]),
                            str(series["sheet"]),
                            point["sourceRowIndex"],
                            point["sourceColumnIndex"],
                            point["sourceCoordinate"],
                            point["axisSourceCoordinate"],
                            point["replicateSourceCoordinate"],
                            point["sourceValueJson"],
                            point["sourceFormulaText"],
                            verification,
                        ),
                    )
                    counts["measurementPoints"] += 1

            for comparison in study.get("comparisons", []):
                comparison_uid = stable_uid("comparison", study_uid, comparison["key"])
                conn.execute(
                    """
                    INSERT INTO knowledge_comparisons(
                        comparison_uid, public_comparison_id, study_id,
                        legacy_comparison_id, comparison_key, compared_arm_id,
                        control_arm_id, design_type, matching_basis,
                        validity_status, confounding_status, exclusion_reason,
                        direction, summary_text, aggregation_eligible,
                        verification_status, details_json
                    ) VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(comparison_uid) DO UPDATE SET
                        compared_arm_id=excluded.compared_arm_id,
                        control_arm_id=excluded.control_arm_id,
                        design_type=excluded.design_type,
                        matching_basis=excluded.matching_basis,
                        validity_status=excluded.validity_status,
                        confounding_status=excluded.confounding_status,
                        exclusion_reason=excluded.exclusion_reason,
                        direction=excluded.direction,
                        summary_text=excluded.summary_text,
                        aggregation_eligible=excluded.aggregation_eligible,
                        verification_status=excluded.verification_status,
                        details_json=excluded.details_json
                    """,
                    (
                        comparison_uid,
                        public_id("CMP", comparison_uid),
                        study_id,
                        str(comparison["key"]),
                        arm_ids[str(comparison["comparedArm"])],
                        arm_ids[str(comparison["controlArm"])],
                        str(comparison["designType"]),
                        str(comparison.get("matchingBasis") or ""),
                        str(comparison["validityStatus"]).upper(),
                        str(comparison["confoundingStatus"]).upper(),
                        str(comparison.get("exclusionReason") or ""),
                        str(comparison.get("direction") or ""),
                        str(comparison.get("summary") or ""),
                        1 if comparison.get("aggregationEligible") else 0,
                        str(comparison["verificationStatus"]).upper(),
                        _json(comparison.get("details"), {}),
                    ),
                )
                comparison_id = int(
                    conn.execute("SELECT comparison_id FROM knowledge_comparisons WHERE comparison_uid=?", (comparison_uid,)).fetchone()[0]
                )
                counts["comparisons"] += 1
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="COMPARISON",
                    entity_uid=comparison_uid,
                    evidence_items=comparison.get("evidence", []),
                    now_iso=now_iso,
                )
                for effect in comparison.get("effects", []):
                    effect_uid = stable_uid(
                        "effect",
                        comparison_uid,
                        effect["outcome"],
                        effect["effectType"],
                        effect.get("formulaVersion") or "canonical-v1",
                    )
                    conn.execute(
                        """
                        INSERT INTO knowledge_effects(
                            effect_uid, public_effect_id, comparison_id, outcome_id,
                            effect_type, estimate, unit_id, original_unit, ci_lower,
                            ci_upper, formula_version, calculation_text, direction,
                            aggregation_eligible, verification_status, details_json
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(effect_uid) DO UPDATE SET
                            estimate=excluded.estimate,
                            unit_id=excluded.unit_id,
                            original_unit=excluded.original_unit,
                            ci_lower=excluded.ci_lower,
                            ci_upper=excluded.ci_upper,
                            calculation_text=excluded.calculation_text,
                            direction=excluded.direction,
                            aggregation_eligible=excluded.aggregation_eligible,
                            verification_status=excluded.verification_status,
                            details_json=excluded.details_json
                        """,
                        (
                            effect_uid,
                            public_id("EFF", effect_uid),
                            comparison_id,
                            outcome_ids[str(effect["outcome"])],
                            str(effect["effectType"]),
                            effect.get("estimate"),
                            _unit_id(conn, effect.get("unit"), now_iso=now_iso),
                            str(effect.get("unit") or ""),
                            effect.get("ciLower"),
                            effect.get("ciUpper"),
                            str(effect.get("formulaVersion") or "canonical-v1"),
                            str(effect.get("calculation") or ""),
                            str(effect.get("direction") or ""),
                            1 if comparison.get("aggregationEligible") else 0,
                            str(effect["verificationStatus"]).upper(),
                            _json(effect.get("details"), {}),
                        ),
                    )
                    counts["effects"] += 1
                    evidence_count += _store_evidence(
                        conn,
                        revision=revision,
                        entity_type="EFFECT",
                        entity_uid=effect_uid,
                        evidence_items=effect.get("evidence", []),
                        now_iso=now_iso,
                    )

            for conclusion in study.get("conclusions", []):
                claim_uid = stable_uid("claim", study_uid, conclusion["key"])
                conn.execute(
                    """
                    INSERT INTO knowledge_claims(
                        claim_uid, public_claim_id, workbook_analysis_id, study_id,
                        legacy_conclusion_id, claim_key, claim_type, claim_text,
                        verdict, causal_strength, verification_status, limitations_json
                    ) VALUES (?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(claim_uid) DO UPDATE SET
                        claim_type=excluded.claim_type,
                        claim_text=excluded.claim_text,
                        verdict=excluded.verdict,
                        causal_strength=excluded.causal_strength,
                        verification_status=excluded.verification_status,
                        limitations_json=excluded.limitations_json
                    """,
                    (
                        claim_uid,
                        public_id("CLM", claim_uid),
                        workbook_analysis_id,
                        study_id,
                        str(conclusion["key"]),
                        str(conclusion["claimType"]).upper(),
                        str(conclusion["text"]),
                        str(conclusion.get("verdict") or ""),
                        str(conclusion.get("causalStrength") or "UNSPECIFIED").upper(),
                        (
                            "NEEDS_REVIEW"
                            if str(conclusion["claimType"]).upper()
                            == "AI_DERIVED_DESCRIPTIVE"
                            else str(
                                conclusion.get("verificationStatus")
                                or study["verificationStatus"]
                            ).upper()
                        ),
                        _json(conclusion.get("limitations"), []),
                    ),
                )
                counts["claims"] += 1
                evidence_count += _store_evidence(
                    conn,
                    revision=revision,
                    entity_type="CLAIM",
                    entity_uid=claim_uid,
                    evidence_items=conclusion.get("evidence", []),
                    now_iso=now_iso,
                )
        _record_analysis_validation_issues(
            conn,
            workbook_analysis_id=workbook_analysis_id,
            now_iso=now_iso,
        )
    except Exception:
        conn.execute("ROLLBACK TO SAVEPOINT canonical_study_import")
        conn.execute("RELEASE SAVEPOINT canonical_study_import")
        raise
    conn.execute("RELEASE SAVEPOINT canonical_study_import")
    return {
        "analysisUid": analysis_uid,
        "publicAnalysisId": public_id("ANALYSIS", analysis_uid),
        "workbookAnalysisId": workbook_analysis_id,
        **counts,
        "evidence": evidence_count,
        "supersededAnalyses": superseded["deleted"],
        "preservedStaleAnalyses": superseded["preservedStale"],
    }
