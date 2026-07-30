"""Deterministic source-coordinate coverage for complete Study drafts.

The semantic manifest contract proves that cited claims exist in Capture v2.
This module proves the inverse for candidate-bearing source sections: every
primary numeric result cell must be represented by a measurement series, an
exact quantitative Observation, or an exact source-backed design value, and
every locator-identified source conclusion cell must remain evidence-linked
to preserved conclusion wording. Deterministically identifiable non-result
cells remain inventoried with an explicit exclusion reason.
"""

from __future__ import annotations

from bisect import bisect_left, bisect_right
import copy
import hashlib
import math
import re
import unicodedata
from typing import Any, Iterable, Sequence


CONTENT_COVERAGE_SCHEMA_VERSION = "study-content-coverage-v1"

_A1_PATTERN = re.compile(
    r"\$?([A-Za-z]{1,4})\$?([1-9]\d*)"
    r"(?::\$?([A-Za-z]{1,4})\$?([1-9]\d*))?"
)
_DIRECT_RATIO_FORMULA = re.compile(
    r"^\s*=\s*\+?\s*\$?[A-Za-z]{1,4}\$?[1-9]\d*"
    r"\s*/\s*\$?[A-Za-z]{1,4}\$?[1-9]\d*"
    r"(?:\s*\*\s*(?:100|1000000))?\s*$",
    re.IGNORECASE,
)
_NUMBER_PATTERN = re.compile(
    r"^[\s(]*([+-]?(?:\d+(?:,\d{3})*|\d*)(?:\.\d+)?"
    r"(?:[Ee][+-]?\d+)?)[\s%)]*$"
)
_SEQUENCE_LABEL = re.compile(
    r"^\s*(?:no\.?|number|index|seq(?:uence)?|sample\s*no\.?"
    r"|순번|번호|일련번호|회차|차수)\s*$",
    re.IGNORECASE,
)
_IDENTIFIER_HEADER = re.compile(
    r"^\s*IR\s*$",
    re.IGNORECASE,
)
_INTEGER_IDENTIFIER = re.compile(r"^\d{6,}$")
_AGGREGATE_LABEL = re.compile(
    r"(?:^|[\s_/#()-])(?:avg\.?|average|mean|sum|total|subtotal|min(?:imum)?"
    r"|max(?:imum)?|median|stdev|std\.?|variance|합계|소계|평균|최대|최소"
    r"|중앙값|표준편차)(?:$|[\s_/#()-])",
    re.IGNORECASE,
)
_CONDITION_LABEL = re.compile(
    r"(?:^|[\s_/#()-])(?:condition|setting|setpoint|spec(?:ification)?"
    r"|target|nominal|criteria|criterion|조건|설정|사양|스펙|기준|목표)"
    r"(?:$|[\s_/#()-])",
    re.IGNORECASE,
)
_QUANTITATIVE_FIELDS = (
    "valueNumber",
    "numerator",
    "denominator",
    "ratePpm",
    "min",
    "max",
    "average",
)
_CONCLUSION_ROLE = re.compile(
    r"(?:decision|conclusion|judg(?:e)?ment|verdict|판정|결론|결정)",
    re.IGNORECASE,
)
_CONCLUSION_HEADING = re.compile(
    r"^\s*(?:(?:[IVXLCDM]+|\d+|[A-Z])[\s.\-:)]*)?"
    r"(?:decision|conclusion|judg(?:e)?ment|verdict|판정|결론|결정)"
    r"[\s.:\-]*$",
    re.IGNORECASE,
)
_CONCLUSION_COLUMN_HEADING = re.compile(
    r"^\s*(?:decision|conclusion|judg(?:e)?ment|verdict|"
    r"(?:(?:total|overall)\s+)?result|remarks?)"
    r"[\s.:\-]*$",
    re.IGNORECASE,
)
_MIXED_CONCLUSION_ROLE = re.compile(
    r"(?:condition|measurement|factor|context|setting|input|sample|test|"
    r"outcome|aggregate)",
    re.IGNORECASE,
)
_CATEGORICAL_STATUS = re.compile(
    r"^\s*(?:pass(?:ed)?|fail(?:ed)?|ok|n/?g)\s*[.!]?\s*$",
    re.IGNORECASE,
)
_SEMANTIC_EVIDENCE_ROLES = {
    "OUTCOME_LABEL",
    "FACTOR_LABEL",
    "FACTOR_LEVEL",
    "ARM_LABEL",
    "UNIT_QUANTITY",
    "COUNT_RATIO",
}
_FACTOR_LABEL = re.compile(
    r"(?:assembly\s*method|method|condition|factor|setting|process|"
    r"pressure|temperature|duration|amount|quantity|level|dose)",
    re.IGNORECASE,
)
_MG_FACTOR_LABEL = re.compile(
    r"^\s*(?:c|s)\s*-\s*mg(?:\s*\([^)]*\))?\s*$",
    re.IGNORECASE,
)
_FACTOR_MATRIX_LEAF = re.compile(
    r"^\s*(?:spec(?:ification)?|supplier|vendor|condition|setting|"
    r"level|method|material|source|type)\s*$",
    re.IGNORECASE,
)
_OUTCOME_LABEL = re.compile(
    r"(?:\bresult\b|\boutcome\b|\bresponse\b|function\s*n/?g|"
    r"n/?g\s*rate|\bgauss\b|\bthd\b|\bimp(?:edance)?\b)",
    re.IGNORECASE,
)
_RESULT_SECTION_HEADING = re.compile(
    r"^\s*(?:(?:[IVXLCDM]+|\d+)[\s.\-:)]*)?"
    r"(?:result|outcome)"
    r"(?:\s+(?:check(?:ing)?|test|summary|review)\b.*)?"
    r"[\s.:\-]*$",
    re.IGNORECASE,
)
_ARM_LABEL = re.compile(
    r"^\s*(?:control|baseline|reference|normal|test|before|after|"
    r"method\s+[A-Za-z0-9_-]+)\s*$",
    re.IGNORECASE,
)
_UNIT_QUANTITY = re.compile(
    r"^\s*[+-]?(?:\d+(?:\.\d+)?|\.\d+)\s*"
    r"(?:hz|khz|mhz|db|dbspl|s|ms|min|h|mg|g|kg|"
    r"mm|cm|m|pa|kpa|mpa|%|℃|°c)\s*$",
    re.IGNORECASE,
)
_COUNT_RATIO = re.compile(
    r"^\s*(\d+)\s*/\s*(\d+)\s*(?:pcs?|ea|samples?)?\s*$",
    re.IGNORECASE,
)
_FIELD_LABELS = (
    ("NUMERATOR", re.compile(r"(?:ng|fail(?:ed)?)\s*(?:count|qty|quantity)|numerator", re.I)),
    ("DENOMINATOR", re.compile(r"(?:sample|total)\s*(?:size|count|qty|quantity)|denominator", re.I)),
    ("MIN", re.compile(r"(?:^|\W)min(?:imum)?(?:$|\W)", re.I)),
    ("MAX", re.compile(r"(?:^|\W)max(?:imum)?(?:$|\W)", re.I)),
    ("AVERAGE", re.compile(r"(?:average|mean|avg\.?)", re.I)),
    (
        "BASELINE",
        re.compile(
            r"(?:^|[|])\s*(?:(?:baseline|control)"
            r"(?:\s+(?:value|condition|setting|level|amount))?"
            r"|before\s+(?:value|condition|setting|level|amount))"
            r"\s*(?=$|[|])",
            re.I,
        ),
    ),
    (
        "CHANGED",
        re.compile(
            r"(?:^|[|])\s*(?:changed"
            r"(?:\s+(?:value|condition|setting|level|amount))?"
            r"|after\s+(?:value|condition|setting|level|amount)"
            r"|test\s+(?:value|condition|setting|level|amount))"
            r"\s*(?=$|[|])",
            re.I,
        ),
    ),
)
_STATUS_LEGEND_LABEL = re.compile(
    r"(?:legend|key|meaning|status\s*(?:guide|legend)|instruction|example)",
    re.IGNORECASE,
)


class ContentCoverageError(RuntimeError):
    """Raised when a complete draft omits owned source content."""


def _column_number(label: str) -> int:
    value = 0
    for char in label.upper():
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def _range_bounds(address: object) -> tuple[int, int, int, int]:
    match = _A1_PATTERN.fullmatch(str(address or "").strip())
    if not match:
        raise ContentCoverageError(
            f"Content coverage received an invalid A1 range: {address}"
        )
    start_column = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_column = _column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_row < start_row or end_column < start_column:
        raise ContentCoverageError(
            f"Content coverage received a reversed A1 range: {address}"
        )
    return start_row, start_column, end_row, end_column


def _sheet_title(chunk: dict[str, Any]) -> str:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return str(sheet.get("title") or "")
    return str(sheet or "")


def _source_cell_key(
    chunk: dict[str, Any],
    cell: dict[str, Any],
) -> str:
    key = str(cell.get("sourceCellKey") or "").strip()
    if key:
        return key
    coordinate = str(
        cell.get("coordinate") or cell.get("c") or ""
    ).strip()
    if not coordinate:
        raise ContentCoverageError(
            f"Chunk {chunk.get('chunkId')} has a cell without a coordinate"
        )
    sheet = chunk.get("sheet")
    sheet_index = (
        int(sheet.get("sheetIndex") or 0)
        if isinstance(sheet, dict)
        else int(chunk.get("sheetIndex") or 0)
    )
    return f"{sheet_index}:{_sheet_title(chunk)}:{coordinate.upper()}"


def _portable_scalar(value: object) -> object:
    if isinstance(value, dict):
        value_type = str(value.get("type") or "").lower()
        if value_type in {"date", "datetime", "time", "timedelta"}:
            return None
        if "value" in value:
            return value["value"]
    return value


def _parse_numeric(value: object, *, declared_numeric: bool) -> float | None:
    value = _portable_scalar(value)
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        return numeric if math.isfinite(numeric) else None
    if not declared_numeric or not isinstance(value, str):
        return None
    match = _NUMBER_PATTERN.fullmatch(value)
    if not match or not match.group(1):
        return None
    try:
        numeric = float(match.group(1).replace(",", ""))
    except ValueError:
        return None
    if value.strip().startswith("(") and value.strip().endswith(")"):
        numeric = -numeric
    return numeric if math.isfinite(numeric) else None


def _cell_numeric_value(cell: dict[str, Any]) -> float | None:
    if _is_excel_error_cell(cell):
        return None
    formula = str(cell.get("formula") or "").strip()
    declared_numeric = str(
        cell.get("cachedDataType") if formula else cell.get("dataType")
    ).lower() == "n"
    candidates = (
        (cell.get("cachedValue"), cell.get("displayValue"))
        if formula
        else (cell.get("rawValue"), cell.get("displayValue"))
    )
    for value in candidates:
        numeric = _parse_numeric(
            value,
            declared_numeric=(
                declared_numeric or isinstance(value, str)
            ),
        )
        if numeric is not None:
            return numeric
    return None


def _cell_text(cell: dict[str, Any]) -> str:
    for field in ("displayValue", "rawValue"):
        value = _portable_scalar(cell.get(field))
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _is_excel_error_cell(cell: dict[str, Any]) -> bool:
    formula = str(cell.get("formula") or "").strip()
    data_type = str(
        cell.get("cachedDataType") if formula else cell.get("dataType")
    ).lower()
    if data_type == "e":
        return True
    values = (
        cell.get("cachedValue"),
        cell.get("displayValue"),
        cell.get("rawValue"),
    )
    excel_errors = {
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NULL!",
        "#NUM!",
        "#REF!",
        "#SPILL!",
        "#VALUE!",
    }
    for value in values:
        portable = _portable_scalar(value)
        if str(portable or "").strip().upper() in excel_errors:
            return True
        if (
            isinstance(portable, (int, float))
            and not isinstance(portable, bool)
            and float(portable).is_integer()
        ):
            unsigned = int(portable) & 0xFFFFFFFF
            if (
                unsigned >> 16 == 0x800A
                and unsigned & 0xFFFF
                in {2000, 2007, 2015, 2023, 2029, 2036, 2042}
            ):
                return True
    return any(error in formula.upper() for error in excel_errors)


def _formula_reference_positions(
    formula: str,
) -> list[tuple[int, int]]:
    positions: list[tuple[int, int]] = []
    for match in _A1_PATTERN.finditer(formula):
        start_column = _column_number(match.group(1))
        start_row = int(match.group(2))
        end_column = _column_number(match.group(3) or match.group(1))
        end_row = int(match.group(4) or match.group(2))
        if (
            end_row < start_row
            or end_column < start_column
            or (end_row - start_row + 1)
            * (end_column - start_column + 1)
            > 10_000
        ):
            continue
        positions.extend(
            (row, column)
            for row in range(start_row, end_row + 1)
            for column in range(start_column, end_column + 1)
        )
    return positions


def _formula_reference_endpoints(
    formula: str,
) -> set[tuple[int, int]]:
    """Return referenced cell/range endpoints without expanding rectangles."""

    endpoints: set[tuple[int, int]] = set()
    for match in _A1_PATTERN.finditer(formula):
        endpoints.add(
            (
                int(match.group(2)),
                _column_number(match.group(1)),
            )
        )
        if match.group(3) and match.group(4):
            endpoints.add(
                (
                    int(match.group(4)),
                    _column_number(match.group(3)),
                )
            )
    return endpoints


def _formula_has_nonblank_source_input(
    *,
    sheet: str,
    cell: dict[str, Any],
    source_cells: dict[tuple[str, str], dict[str, Any]],
    visiting: set[tuple[str, str]] | None = None,
) -> bool:
    formula = str(cell.get("formula") or "").strip()
    references = _formula_reference_positions(formula)
    if not references:
        # Do not classify constants or unsupported reference syntax as blank.
        return True
    active = visiting if visiting is not None else set()
    cell_coordinate = _coordinate(cell)
    current = (sheet.casefold(), cell_coordinate)
    if current in active:
        return False
    active.add(current)
    try:
        for row, column in references:
            coordinate = _coordinate_from_position(row, column)
            referenced = source_cells.get(
                (sheet.casefold(), coordinate)
            )
            if referenced is None or _is_excel_error_cell(referenced):
                continue
            referenced_formula = str(
                referenced.get("formula") or ""
            ).strip()
            if referenced_formula:
                if _formula_has_nonblank_source_input(
                    sheet=sheet,
                    cell=referenced,
                    source_cells=source_cells,
                    visiting=active,
                ):
                    return True
                continue
            for field in ("rawValue", "displayValue", "cachedValue"):
                value = _portable_scalar(referenced.get(field))
                if value not in (None, ""):
                    return True
        return False
    finally:
        active.remove(current)


def _hidden_formula_without_source_input(
    *,
    sheet: str,
    cell: dict[str, Any],
    numeric_value: float,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> bool:
    hidden = cell.get("hidden")
    formula = str(cell.get("formula") or "").strip()
    row, column = _cell_position(cell)
    reference_positions = _formula_reference_positions(formula)
    visible_blank_label_placeholder = (
        len(reference_positions) == 1
        and sum(
            1
            for (
                candidate_sheet,
                _candidate_coordinate,
            ), candidate in source_cells.items()
            if candidate_sheet == sheet.casefold()
            and _cell_position(candidate)[1] == column
            and row < _cell_position(candidate)[0] <= row + 3
            and _is_excel_error_cell(candidate)
        )
        >= 2
        and any(
            candidate_sheet == sheet.casefold()
            and _cell_position(candidate)[1] == column
            and row + 4 <= _cell_position(candidate)[0] <= row + 6
            and "AVG" in _cell_text(candidate).upper()
            for (
                candidate_sheet,
                _candidate_coordinate,
            ), candidate in source_cells.items()
        )
    )
    return (
        bool(formula)
        and numeric_value == 0
        and not _formula_has_nonblank_source_input(
            sheet=sheet,
            cell=cell,
            source_cells=source_cells,
        )
        and (
            isinstance(hidden, dict)
            and bool(hidden.get("row"))
            or visible_blank_label_placeholder
        )
    )


def _isolated_axis_tail_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Classify detached terminal axes with blank or error-only values.

    The run must contain at least two adjacent numeric axis cells, be the last
    populated content on the sheet, and start after a gap of at least ten
    completely empty rows. Error-only rows must have one numeric axis cell
    followed only by Excel errors. These constraints prevent ordinary two-row
    results or numeric values next to errors from being discarded.
    """

    positioned_by_sheet: dict[
        str,
        list[tuple[int, int, dict[str, Any]]],
    ] = {}
    for (candidate_sheet, _coordinate_key), cell in source_cells.items():
        cell_row, cell_column = _cell_position(cell)
        has_value = bool(
            _cell_numeric_value(cell) is not None
            or _cell_text(cell)
            or _is_excel_error_cell(cell)
            or str(cell.get("formula") or "").strip()
        )
        if has_value:
            positioned_by_sheet.setdefault(
                candidate_sheet,
                [],
            ).append((cell_row, cell_column, cell))

    isolated: dict[tuple[str, int, int], str] = {}
    for sheet_key, positioned in positioned_by_sheet.items():
        cells_by_row: dict[
            int,
            list[tuple[int, dict[str, Any]]],
        ] = {}
        for cell_row, cell_column, cell in positioned:
            cells_by_row.setdefault(cell_row, []).append(
                (cell_column, cell)
            )
        populated_rows = sorted(cells_by_row)
        if not populated_rows:
            continue
        terminal_row = populated_rows[-1]
        tail_rows: dict[tuple[int, str], list[int]] = {}
        for candidate_row, row_cells in cells_by_row.items():
            numeric_cells = [
                (cell_column, cell)
                for cell_column, cell in row_cells
                if _cell_numeric_value(cell) is not None
            ]
            if len(numeric_cells) != 1:
                continue
            candidate_column = numeric_cells[0][0]
            other_cells = [
                (cell_column, cell)
                for cell_column, cell in row_cells
                if cell_column != candidate_column
            ]
            reason = ""
            if not other_cells:
                reason = "ISOLATED_BLANK_AXIS_TAIL"
            elif all(
                cell_column > candidate_column
                and _is_excel_error_cell(cell)
                for cell_column, cell in other_cells
            ):
                reason = "ERROR_ONLY_AXIS_TAIL"
            if reason:
                tail_rows.setdefault(
                    (candidate_column, reason),
                    [],
                ).append(candidate_row)
        for (
            candidate_column,
            reason,
        ), candidate_rows in tail_rows.items():
            ordered_rows = sorted(candidate_rows)
            run: list[int] = []
            runs: list[list[int]] = []
            for candidate_row in ordered_rows:
                if run and candidate_row != run[-1] + 1:
                    runs.append(run)
                    run = []
                run.append(candidate_row)
            if run:
                runs.append(run)
            for candidate_run in runs:
                if (
                    len(candidate_run) < 2
                    or candidate_run[-1] != terminal_row
                ):
                    continue
                previous_rows = [
                    populated_row
                    for populated_row in populated_rows
                    if populated_row < candidate_run[0]
                ]
                if not previous_rows:
                    # A staged fragment may contain only the detached tail,
                    # without the earlier matrix rows that proved the gap in
                    # the complete packet. Preserve the same classification
                    # only for a high-row, high-valued, increasing numeric
                    # axis whose paired result cells are all Excel errors.
                    # These constraints keep ordinary two-row result tables
                    # in scope while allowing the exact A175:B176-style
                    # frequency tail to validate independently.
                    if (
                        reason != "ERROR_ONLY_AXIS_TAIL"
                        or candidate_run[0] < 100
                    ):
                        continue
                    axis_values = [
                        _cell_numeric_value(
                            next(
                                cell
                                for cell_column, cell
                                in cells_by_row[candidate_row]
                                if cell_column == candidate_column
                            )
                        )
                        for candidate_row in candidate_run
                    ]
                    if (
                        any(
                            value is None or value < 1000
                            for value in axis_values
                        )
                        or any(
                            later <= earlier
                            for earlier, later in zip(
                                axis_values,
                                axis_values[1:],
                            )
                        )
                    ):
                        continue
                elif candidate_run[0] - previous_rows[-1] < 11:
                    continue
                for candidate_row in candidate_run:
                    isolated[
                        (
                            sheet_key,
                            candidate_row,
                            candidate_column,
                        )
                    ] = reason
    return isolated


def _is_date_format(number_format: object) -> bool:
    text = str(number_format or "").lower()
    if not text or text == "general":
        return False
    text = re.sub(r'"[^"]*"', "", text)
    text = re.sub(r"\\.", "", text)
    text = re.sub(r"\[(?:black|blue|cyan|green|magenta|red|white|yellow)\]", "", text)
    return bool(
        re.search(r"y{1,4}", text)
        or re.search(r"d{1,4}", text)
        or any(token in text for token in ("년", "월", "일"))
    )


def _hidden_error_grid_companion_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Exclude raw helpers trapped inside a broken hidden formula grid.

    Historical workbooks sometimes retain a fully hidden legacy table whose
    formulas now evaluate to Excel COM error variants. A raw zero/count placed
    between those broken formulas is an implementation input to that unusable
    grid, not a trustworthy standalone result. Visible values and hidden raw
    values without an adjacent hidden error formula remain required.
    """

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    excluded: dict[tuple[str, int, int], str] = {}
    for (sheet, row, column), cell in source_by_position.items():
        hidden = cell.get("hidden")
        if (
            _cell_numeric_value(cell) is None
            or str(cell.get("formula") or "").strip()
            or not isinstance(hidden, dict)
            or not bool(hidden.get("column"))
        ):
            continue
        adjacent_errors = [
            candidate
            for candidate_column in range(
                max(1, column - 2),
                column + 3,
            )
            if candidate_column != column
            and (
                candidate := source_by_position.get(
                    (sheet, row, candidate_column)
                )
            )
            is not None
            and str(candidate.get("formula") or "").strip()
            and isinstance(candidate.get("hidden"), dict)
            and bool(candidate["hidden"].get("column"))
            and _is_excel_error_cell(candidate)
        ]
        if adjacent_errors:
            excluded[(sheet, row, column)] = (
                "HIDDEN_ERROR_GRID_INPUT"
            )
    return excluded


def _formula_label_input_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Identify numeric inputs rendered into a nearby formula label."""

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    excluded: dict[tuple[str, int, int], str] = {}
    for (sheet, row, column), cell in source_by_position.items():
        numeric = _cell_numeric_value(cell)
        if numeric is None:
            continue
        for label_row in range(row + 1, row + 3):
            label = source_by_position.get(
                (sheet, label_row, column)
            )
            label_formula = str(
                (label or {}).get("formula") or ""
            ).strip()
            if (
                label is None
                or not label_formula
                or (row, column)
                not in _formula_reference_endpoints(
                    label_formula
                )
            ):
                continue
            if _DIRECT_RATIO_FORMULA.fullmatch(label_formula):
                continue
            label_text = _cell_text(label)
            numeric_text = (
                str(int(numeric))
                if float(numeric).is_integer()
                else str(numeric)
            )
            if (
                re.search(
                    rf"(?<!\d){re.escape(numeric_text)}(?!\d)",
                    label_text,
                )
                and re.search(r"[A-Za-z%]", label_text)
            ):
                excluded[(sheet, row, column)] = (
                    "FORMULA_LABEL_INPUT"
                )
                break
        if (sheet, row, column) in excluded:
            continue
        for label_column in range(column + 1, column + 3):
            label = source_by_position.get(
                (sheet, row, label_column)
            )
            label_formula = str(
                (label or {}).get("formula") or ""
            ).strip()
            if (
                label is None
                or re.search(
                    r"\bINDEX\s*\(",
                    label_formula,
                    re.IGNORECASE,
                )
                is None
                or (row, column)
                not in _formula_reference_endpoints(label_formula)
                or not re.search(r"[A-Za-z#]", _cell_text(label))
            ):
                continue
            excluded[(sheet, row, column)] = (
                "FORMULA_LABEL_INPUT"
            )
            break
    return excluded


def _formula_lookup_index_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Identify cached MATCH results consumed only as INDEX selectors."""

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    index_reference_counts: dict[tuple[str, int, int], int] = {}
    for (
        candidate_sheet,
        _candidate_row,
        _candidate_column,
    ), candidate in source_by_position.items():
        candidate_formula = str(candidate.get("formula") or "")
        if re.search(
            r"\bINDEX\s*\(",
            candidate_formula,
            re.IGNORECASE,
        ) is None:
            continue
        for referenced_row, referenced_column in set(
            _formula_reference_endpoints(candidate_formula)
        ):
            identity = (
                candidate_sheet,
                referenced_row,
                referenced_column,
            )
            index_reference_counts[identity] = (
                index_reference_counts.get(identity, 0) + 1
            )
    excluded: dict[tuple[str, int, int], str] = {}
    for (sheet, row, column), cell in source_by_position.items():
        formula = str(cell.get("formula") or "")
        if (
            _cell_numeric_value(cell) is None
            or re.search(r"\bMATCH\s*\(", formula, re.IGNORECASE)
            is None
        ):
            continue
        index_consumers = index_reference_counts.get(
            (sheet, row, column),
            0,
        )
        if index_consumers >= 2:
            excluded[(sheet, row, column)] = (
                "FORMULA_LOOKUP_INDEX"
            )
    return excluded


def _formula_layout_sequence_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Identify unlabeled chart-layout sequences extended by formulas."""

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    numeric_by_column: dict[
        tuple[str, int],
        list[tuple[int, dict[str, Any]]],
    ] = {}
    for (sheet, row, column), cell in source_by_position.items():
        if _cell_numeric_value(cell) is not None:
            numeric_by_column.setdefault(
                (sheet, column),
                [],
            ).append((row, cell))
    excluded: dict[tuple[str, int, int], str] = {}
    for (sheet, column), values in numeric_by_column.items():
        if column > 2:
            continue
        values.sort(key=lambda item: item[0])
        groups: list[list[tuple[int, dict[str, Any]]]] = []
        current: list[tuple[int, dict[str, Any]]] = []
        for item in values:
            if current and item[0] != current[-1][0] + 1:
                groups.append(current)
                current = []
            current.append(item)
        if current:
            groups.append(current)
        for group in groups:
            if len(group) < 5:
                continue
            recurrence_deltas: list[float] = []
            column_label = _coordinate_from_position(
                1,
                column,
            )[:-1]
            for row, cell in group:
                formula = str(cell.get("formula") or "").strip()
                match = re.fullmatch(
                    rf"=\s*\$?{column_label}\$?"
                    rf"{row - 1}\s*\+\s*"
                    r"([+-]?(?:\d+(?:\.\d+)?|\.\d+))",
                    formula,
                    re.IGNORECASE,
                )
                if match is not None:
                    recurrence_deltas.append(float(match.group(1)))
            if (
                len(recurrence_deltas) < 2
                or len(set(recurrence_deltas)) != 1
                or recurrence_deltas[0] <= 0
            ):
                continue
            nearby_labels = sum(
                1
                for row, _cell in group
                if any(
                    (
                        candidate := source_by_position.get(
                            (sheet, row, candidate_column)
                        )
                    )
                    is not None
                    and bool(_cell_text(candidate))
                    for candidate_column in range(
                        column + 1,
                        column + 4,
                    )
                )
            )
            if nearby_labels < 3:
                continue
            for row, _cell in group:
                excluded[(sheet, row, column)] = (
                    "FORMULA_LAYOUT_SEQUENCE"
                )
    return excluded


def _duplicate_identifier_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Find repeated row identifiers while preserving their first value.

    Historical result tables sometimes repeat one numeric IR identifier for
    each arm and for the total row.  One source-backed context remains
    queryable; later identical cells are duplicate row identity, not repeated
    measurements.  The exact IR header plus a numeric result to the right is
    required so ordinary repeated result values remain in scope.
    """

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    excluded: dict[tuple[str, int, int], str] = {}
    for (
        sheet,
        header_row,
        header_column,
    ), header_cell in source_by_position.items():
        if not _IDENTIFIER_HEADER.fullmatch(_cell_text(header_cell)):
            continue
        occurrences: dict[str, list[int]] = {}
        for row in range(header_row + 1, header_row + 31):
            cell = source_by_position.get(
                (sheet, row, header_column)
            )
            if cell is None:
                continue
            text = _cell_text(cell).strip()
            if not text:
                numeric = _cell_numeric_value(cell)
                if (
                    numeric is not None
                    and math.isfinite(numeric)
                    and numeric.is_integer()
                ):
                    text = str(int(numeric))
            if not _INTEGER_IDENTIFIER.fullmatch(text):
                continue
            has_result_to_right = any(
                (
                    candidate := source_by_position.get(
                        (sheet, row, column)
                    )
                )
                is not None
                and _cell_numeric_value(candidate) is not None
                for column in range(
                    header_column + 1,
                    header_column + 11,
                )
            )
            if has_result_to_right:
                occurrences.setdefault(text, []).append(row)
        for rows in occurrences.values():
            if len(rows) < 2 or any(
                later - earlier > 4
                for earlier, later in zip(rows, rows[1:])
            ):
                continue
            for row in rows[1:]:
                excluded[
                    (sheet, row, header_column)
                ] = "DUPLICATE_IDENTIFIER"
    return excluded


def _horizontal_replicate_identifier_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Find 1..N column headers backed by numeric replicate rows below."""

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    numeric_rows: dict[
        tuple[str, int],
        list[tuple[int, int]],
    ] = {}
    for (sheet, row, column), cell in source_by_position.items():
        numeric = _cell_numeric_value(cell)
        if (
            numeric is None
            or not math.isfinite(numeric)
            or not numeric.is_integer()
        ):
            continue
        numeric_rows.setdefault((sheet, row), []).append(
            (column, int(numeric))
        )
    excluded: dict[tuple[str, int, int], str] = {}
    marker_pattern = re.compile(
        r"(?:^|\\b)(?:avg|average|mean|sample|no\\.?|pcs?|ea|qty)(?:\\b|$)",
        flags=re.IGNORECASE,
    )
    for (sheet, row), values in numeric_rows.items():
        ordered = sorted(values)
        runs: list[list[tuple[int, int]]] = []
        current: list[tuple[int, int]] = []
        for column, value in ordered:
            if (
                current
                and (
                    column != current[-1][0] + 1
                    or value != current[-1][1] + 1
                )
            ):
                runs.append(current)
                current = []
            current.append((column, value))
        if current:
            runs.append(current)
        for run in runs:
            if len(run) < 3 or run[0][1] != 1:
                continue
            start_column = run[0][0]
            end_column = run[-1][0]
            nearby_marker = any(
                (
                    candidate := source_by_position.get(
                        (sheet, row, column)
                    )
                )
                is not None
                and marker_pattern.search(_cell_text(candidate))
                for column in (
                    *range(max(1, start_column - 3), start_column),
                    *range(end_column + 1, end_column + 4),
                )
            )
            if not nearby_marker:
                continue
            populated_columns = {
                column
                for column, _value in run
                if any(
                    (
                        candidate := source_by_position.get(
                            (sheet, data_row, column)
                        )
                    )
                    is not None
                    and _cell_numeric_value(candidate) is not None
                    for data_row in range(row + 1, row + 11)
                )
            }
            formula_labeled_columns = {
                column
                for column, _value in run
                if any(
                    str(
                        candidate.get("formula") or ""
                    ).strip()
                    for label_row in range(row + 1, row + 3)
                    if (
                        candidate
                        := source_by_position.get(
                            (sheet, label_row, column)
                        )
                    )
                    is not None
                )
            }
            if (
                len(populated_columns) < min(3, len(run))
                and len(formula_labeled_columns)
                < min(3, len(run))
            ):
                continue
            for column, _value in run:
                excluded[
                    (sheet, row, column)
                ] = "HORIZONTAL_REPLICATE_IDENTIFIER"
    return excluded


def _merged_structural_numeric_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Identify numeric merged anchors used as row/group/column labels."""

    source_by_position = {
        (sheet, *_cell_position(cell)): cell
        for (sheet, _coordinate_key), cell in source_cells.items()
    }
    vertical_anchors: dict[
        tuple[str, int, int],
        tuple[int, dict[str, Any]],
    ] = {}
    excluded: dict[tuple[str, int, int], str] = {}
    for (sheet, row, column), cell in source_by_position.items():
        if (
            _cell_numeric_value(cell) is None
            or str(cell.get("mergeRole") or "").casefold() != "anchor"
            or not str(cell.get("mergeRange") or "").strip()
        ):
            continue
        bounds = _range_bounds(cell.get("mergeRange"))
        if (
            bounds[0] != row
            or bounds[1] != column
            or bounds[3] != column
            or bounds[2] <= row
        ):
            continue
        vertical_anchors[(sheet, row, column)] = (bounds[2], cell)

    for (sheet, row, column), (end_row, _cell) in (
        vertical_anchors.items()
    ):
        explicit_sequence_header = any(
            _SEQUENCE_LABEL.fullmatch(_cell_text(header).strip())
            for header_row in range(max(1, row - 32), row)
            if (
                header := source_by_position.get(
                    (sheet, header_row, column)
                )
            )
            is not None
        )
        numeric_to_right = any(
            (
                candidate := source_by_position.get(
                    (sheet, member_row, member_column)
                )
            )
            is not None
            and _cell_numeric_value(candidate) is not None
            for member_row in range(row, end_row + 1)
            for member_column in range(column + 1, column + 12)
        )
        if explicit_sequence_header and numeric_to_right:
            excluded[(sheet, row, column)] = "MERGED_ROW_IDENTIFIER"
            continue

        labeled_result_rows = 0
        for member_row in range(row, end_row + 1):
            label = source_by_position.get(
                (sheet, member_row, column + 1)
            )
            if label is None or not _cell_text(label):
                continue
            if any(
                (
                    result := source_by_position.get(
                        (sheet, member_row, result_column)
                    )
                )
                is not None
                and _cell_numeric_value(result) is not None
                for result_column in range(column + 2, column + 5)
            ):
                labeled_result_rows += 1
        if end_row - row + 1 >= 3 and labeled_result_rows >= 2:
            excluded[(sheet, row, column)] = "MERGED_GROUP_IDENTIFIER"

    column_level_groups: dict[
        tuple[str, int, int],
        list[int],
    ] = {}
    for sheet, row, column in vertical_anchors:
        end_row = vertical_anchors[(sheet, row, column)][0]
        column_level_groups.setdefault(
            (sheet, row, end_row),
            [],
        ).append(column)
    for (sheet, row, end_row), columns in column_level_groups.items():
        ordered_columns = sorted(columns)
        if (
            len(ordered_columns) < 2
            or any(
                later != earlier + 1
                for earlier, later in zip(
                    ordered_columns,
                    ordered_columns[1:],
                )
            )
        ):
            continue
        parent_found = False
        for parent_row in range(max(1, row - 3), row):
            for parent_column in range(
                ordered_columns[0],
                ordered_columns[-1] + 1,
            ):
                parent = source_by_position.get(
                    (sheet, parent_row, parent_column)
                )
                if (
                    parent is None
                    or not _cell_text(parent)
                    or str(
                        parent.get("mergeRole") or ""
                    ).casefold()
                    != "anchor"
                    or not str(
                        parent.get("mergeRange") or ""
                    ).strip()
                ):
                    continue
                bounds = _range_bounds(parent.get("mergeRange"))
                if (
                    bounds[0] == bounds[2] == parent_row
                    and bounds[1] <= ordered_columns[0]
                    and bounds[3] >= ordered_columns[-1]
                ):
                    parent_found = True
                    break
            if parent_found:
                break
        if not parent_found:
            continue
        if not all(
            any(
                (
                    value_cell := source_by_position.get(
                        (sheet, value_row, column)
                    )
                )
                is not None
                and _cell_numeric_value(value_cell) is not None
                for value_row in range(end_row + 1, end_row + 13)
            )
            for column in ordered_columns
        ):
            continue
        for column in ordered_columns:
            excluded[(sheet, row, column)] = "MERGED_COLUMN_LEVEL"
    return excluded


def _coordinate(cell: dict[str, Any]) -> str:
    return str(
        cell.get("coordinate") or cell.get("c") or ""
    ).strip().upper()


def _cell_position(cell: dict[str, Any]) -> tuple[int, int]:
    row = cell.get("row")
    column = cell.get("column")
    if row not in (None, "") and column not in (None, ""):
        return int(row), int(column)
    coordinate = _coordinate(cell)
    bounds = _range_bounds(coordinate)
    return bounds[0], bounds[1]


def _sheet_layout_ordinal_positions(
    *,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> dict[tuple[str, int, int], str]:
    """Identify an isolated print-page ordinal above a merged TITLE band."""

    result: dict[tuple[str, int, int], str] = {}
    by_sheet: dict[str, list[dict[str, Any]]] = {}
    for (sheet, _coordinate_key), cell in source_cells.items():
        by_sheet.setdefault(sheet, []).append(cell)
    for sheet, cells in by_sheet.items():
        row_one_cells = [
            cell
            for cell in cells
            if _cell_position(cell)[0] == 1
            and (
                _cell_text(cell)
                or _cell_numeric_value(cell) is not None
            )
        ]
        if len(row_one_cells) != 1:
            continue
        ordinal_cell = row_one_cells[0]
        numeric = _cell_numeric_value(ordinal_cell)
        if (
            numeric is None
            or not math.isfinite(numeric)
            or not numeric.is_integer()
            or not 1 <= numeric <= 99
        ):
            continue
        _ordinal_row, ordinal_column = _cell_position(ordinal_cell)
        for title_cell in cells:
            title_row, _title_column = _cell_position(title_cell)
            if (
                title_row != 2
                or " ".join(
                    _cell_text(title_cell).upper().split()
                )
                != "TITLE"
                or not str(
                    title_cell.get("mergeRange") or ""
                ).strip()
            ):
                continue
            (
                start_row,
                start_column,
                end_row,
                end_column,
            ) = _range_bounds(title_cell["mergeRange"])
            if (
                start_row == 2
                and start_column <= ordinal_column <= end_column
                and end_row >= 2
            ):
                result[(sheet, 1, ordinal_column)] = (
                    "SHEET_LAYOUT_ORDINAL"
                )
                break
    return result


def _nearby_labels(
    *,
    sheet: str,
    row: int,
    column: int,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> tuple[str, str]:
    row_labels: list[tuple[int, str]] = []
    column_labels: list[tuple[int, str]] = []
    sheet_key = sheet.casefold()
    same_column_rows = sorted(
        {
        candidate_row
        for (
            candidate_sheet,
            _candidate_coordinate,
        ), candidate_cell in source_cells.items()
        for candidate_row, candidate_column in [
            _cell_position(candidate_cell)
        ]
        if candidate_sheet == sheet_key
        and candidate_column == column
        and candidate_row <= row
        }
    )
    for (candidate_sheet, _coordinate_key), cell in source_cells.items():
        if candidate_sheet != sheet_key:
            continue
        text = _cell_text(cell)
        if not text:
            continue
        candidate_row, candidate_column = _cell_position(cell)
        if candidate_row == row and candidate_column < column:
            row_labels.append((column - candidate_column, text))
        path_rows = [
            value
            for value in same_column_rows
            if candidate_row <= value <= row
        ]
        continuously_populated = bool(
            len(path_rows) >= 2
            and all(
                later - earlier <= 2
                for earlier, later in zip(
                    path_rows,
                    path_rows[1:],
                )
            )
        )
        if (
            candidate_column == column
            and candidate_row < row
            and (
                row - candidate_row <= 2
                or continuously_populated
            )
        ):
            column_labels.append((row - candidate_row, text))
    row_labels.sort(key=lambda item: item[0])
    column_labels.sort(key=lambda item: item[0])
    return (
        " | ".join(value for _distance, value in row_labels[:3]),
        " | ".join(
            value for _distance, value in column_labels[:3]
        ),
    )


def _has_conclusion_column_heading_above(
    *,
    sheet: str,
    row: int,
    column: int,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> bool:
    sheet_key = sheet.casefold()
    return any(
        candidate_sheet == sheet_key
        and candidate_column == column
        and 0 < row - candidate_row <= 20
        and _CONCLUSION_COLUMN_HEADING.fullmatch(
            _cell_text(candidate_cell)
        )
        for (
            candidate_sheet,
            _candidate_coordinate,
        ), candidate_cell in source_cells.items()
        for candidate_row, candidate_column in [
            _cell_position(candidate_cell)
        ]
    )


def _numeric_column_size(
    *,
    sheet: str,
    column: int,
    numeric_positions: set[tuple[str, int, int]],
) -> int:
    sheet_key = sheet.casefold()
    return sum(
        1
        for candidate_sheet, _candidate_row, candidate_column
        in numeric_positions
        if candidate_sheet == sheet_key
        and candidate_column == column
    )


def _field_role(row_labels: str, column_labels: str) -> str:
    components = [
        component.strip()
        for labels in (column_labels, row_labels)
        for component in labels.split("|")
        if component.strip()
    ]
    for component in components:
        for role, pattern in _FIELD_LABELS:
            if (
                role in {"MIN", "MAX", "AVERAGE"}
                and len(re.findall(r"\w+", component)) > 4
            ):
                continue
            if pattern.search(component):
                return role
    return ""


def _numeric_source_role(
    *,
    sheet: str,
    row: int,
    column: int,
    row_labels: str,
    column_labels: str,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> str:
    row_components = [
        component.strip()
        for component in row_labels.split("|")
        if component.strip()
    ]
    column_components = [
        component.strip()
        for component in column_labels.split("|")
        if component.strip()
    ]
    # Only the nearest row/leaf-column label may declare an aggregate.  A
    # farther merged parent such as ``Total -> Position 1/2/3`` names a
    # measurement family; it must not turn each raw position into a summary.
    nearest_labels = " | ".join(
        components[0]
        for components in (row_components, column_components)
        if components
    )
    if _AGGREGATE_LABEL.search(nearest_labels):
        return "AGGREGATE"
    labels = f"{row_labels} | {column_labels}"
    if _CONDITION_LABEL.search(labels) or _FACTOR_LABEL.search(labels):
        return "FACTOR"
    # A vertical result column also has numeric cells below every nonterminal
    # result row.  That fact alone cannot make the current cell an axis.
    # Numeric axes are authorized later from the exact paired series geometry.
    return "RESULT"


def _semantic_source_roles(
    *,
    primary_cells: Sequence[
        tuple[dict[str, Any], dict[str, Any], str, int, int]
    ],
    primary_by_coordinate: dict[
        tuple[str, str],
        tuple[dict[str, Any], dict[str, Any], str, int, int],
    ],
    source_cells: dict[tuple[str, str], dict[str, Any]],
    locator_results: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    role_by_key: dict[str, dict[str, Any]] = {}
    source_by_position = {
        (
            candidate_sheet,
            *_cell_position(candidate_cell),
        ): candidate_cell
        for (
            candidate_sheet,
            _candidate_coordinate,
        ), candidate_cell in source_cells.items()
    }

    def looks_like_label(text: str) -> bool:
        return (
            len(re.findall(r"\S+", text)) <= 5
            and re.search(r"[.!?]\s*$", text) is None
        )

    def has_result_below(
        *,
        sheet: str,
        row: int,
        column: int,
    ) -> bool:
        panel_hits = 0
        for primary in primary_cells:
            _chunk, cell, candidate_sheet, candidate_row, candidate_column = (
                primary
            )
            if (
                candidate_sheet.casefold() != sheet.casefold()
                or not 0 < candidate_row - row <= 5
                or not column <= candidate_column <= column + 24
            ):
                continue
            text = _cell_text(cell)
            if (
                _cell_numeric_value(cell) is not None
                or _UNIT_QUANTITY.fullmatch(text)
                or _COUNT_RATIO.fullmatch(text)
                or _CATEGORICAL_STATUS.fullmatch(text)
            ):
                if candidate_column == column:
                    return True
                panel_hits += 1
        return panel_hits >= 2

    def is_merged_parent_with_leaf_grid(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
    ) -> bool:
        """Identify a horizontal group header above leaf headers/values."""

        _chunk, cell, sheet, row, column = primary
        if str(cell.get("mergeRole") or "").casefold() != "anchor":
            return False
        merge_range = str(cell.get("mergeRange") or "").strip()
        if not merge_range:
            return False
        start_row, start_column, end_row, end_column = (
            _range_bounds(merge_range)
        )
        if (
            start_row != end_row
            or end_column - start_column < 1
            or row != start_row
            or column != start_column
        ):
            return False
        leaf_texts: list[str] = []
        sheet_key = sheet.casefold()
        leaf_row = end_row + 1
        value_row = leaf_row + 1
        for leaf_column in range(start_column, end_column + 1):
            leaf_cell = source_by_position.get(
                (sheet_key, leaf_row, leaf_column)
            )
            value_cell = source_by_position.get(
                (sheet_key, value_row, leaf_column)
            )
            if leaf_cell is None or value_cell is None:
                continue
            leaf_text = _cell_text(leaf_cell)
            if (
                not leaf_text
                or _cell_numeric_value(leaf_cell) is not None
                or not (
                    _cell_text(value_cell)
                    or _cell_numeric_value(value_cell) is not None
                )
            ):
                continue
            leaf_texts.append(
                _normalized_narrative(leaf_text)
            )
        return (
            len(leaf_texts) >= 2
            and len(set(leaf_texts)) >= 2
        )

    def is_merged_vertical_table_header(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
    ) -> bool:
        """Recognize vertically merged field headings in a wide table."""

        _chunk, cell, sheet, row, column = primary
        if (
            str(cell.get("mergeRole") or "").casefold() != "anchor"
            or not str(cell.get("mergeRange") or "").strip()
        ):
            return False
        bounds = _range_bounds(cell.get("mergeRange"))
        if (
            bounds[0] != row
            or bounds[1] != bounds[3]
            or bounds[2] <= row
            or bounds[1] != column
        ):
            return False
        sheet_key = sheet.casefold()
        peer_headers = 0
        for (
            candidate_sheet,
            candidate_row,
            _candidate_column,
        ), candidate in source_by_position.items():
            if (
                candidate_sheet != sheet_key
                or candidate_row != row
                or str(
                    candidate.get("mergeRole") or ""
                ).casefold()
                != "anchor"
                or not str(
                    candidate.get("mergeRange") or ""
                ).strip()
            ):
                continue
            candidate_bounds = _range_bounds(
                candidate.get("mergeRange")
            )
            if (
                candidate_bounds[0] == row
                and candidate_bounds[1] == candidate_bounds[3]
                and candidate_bounds[2] == bounds[2]
            ):
                peer_headers += 1
        populated_below = sum(
            1
            for data_row in range(bounds[2] + 1, bounds[2] + 11)
            if (
                candidate := source_by_position.get(
                    (sheet_key, data_row, column)
                )
            )
            is not None
            and bool(
                _cell_text(candidate)
                or _cell_numeric_value(candidate) is not None
            )
        )
        return peer_headers >= 3 and populated_below >= 2

    def is_structural_factor_leaf_header(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
    ) -> bool:
        """Recognize factor leaves only inside a captured merged matrix.

        Bare ``Spec`` or ``Supplier`` text is too ambiguous to make a factor
        requirement.  A horizontal named parent, at least two distinct leaf
        headers, and at least two populated data rows are all required.  This
        captures layouts such as ``S-MG -> Spec/Supplier`` without treating
        unrelated prose or result-grid leaves as factors.
        """

        _chunk, cell, sheet, row, column = primary
        if not _FACTOR_MATRIX_LEAF.fullmatch(_cell_text(cell)):
            return False
        sheet_key = sheet.casefold()
        parent_bounds: tuple[int, int, int, int] | None = None
        for (
            candidate_sheet,
            candidate_row,
            candidate_column,
        ), parent_cell in source_by_position.items():
            if (
                candidate_sheet != sheet_key
                or candidate_row != row - 1
                or str(
                    parent_cell.get("mergeRole") or ""
                ).casefold()
                != "anchor"
                or not str(
                    parent_cell.get("mergeRange") or ""
                ).strip()
                or not _cell_text(parent_cell)
            ):
                continue
            bounds = _range_bounds(parent_cell.get("mergeRange"))
            if (
                bounds[0] == bounds[2] == candidate_row
                and bounds[3] - bounds[1] >= 1
                and bounds[1] <= column <= bounds[3]
            ):
                parent_bounds = bounds
                break
        if parent_bounds is None:
            return False
        leaf_texts = [
            _normalized_narrative(_cell_text(leaf_cell))
            for leaf_column in range(
                parent_bounds[1],
                parent_bounds[3] + 1,
            )
            if (
                leaf_cell := source_by_position.get(
                    (sheet_key, row, leaf_column)
                )
            )
            is not None
            and _FACTOR_MATRIX_LEAF.fullmatch(
                _cell_text(leaf_cell)
            )
        ]
        populated_level_rows = {
            level_row
            for level_row in range(row + 1, row + 11)
            if any(
                (
                    level_cell := source_by_position.get(
                        (sheet_key, level_row, level_column)
                    )
                )
                is not None
                and bool(
                    _cell_text(level_cell)
                    or _cell_numeric_value(level_cell) is not None
                )
                for level_column in range(
                    parent_bounds[1],
                    parent_bounds[3] + 1,
                )
            )
        }
        return (
            len(leaf_texts) >= 2
            and len(set(leaf_texts)) >= 2
            and len(populated_level_rows) >= 2
        )

    def is_typed_mg_matrix_factor_header(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
    ) -> bool:
        """Recognize C-MG/S-MG design columns beside an exact Type header."""

        _chunk, cell, sheet, row, column = primary
        text = _cell_text(cell)
        is_mg_header = _MG_FACTOR_LABEL.fullmatch(text) is not None
        is_type_header = text.strip().casefold() == "type"
        if not is_mg_header and not is_type_header:
            return False
        sheet_key = sheet.casefold()
        type_columns = [
            candidate_column
            for (
                candidate_sheet,
                candidate_row,
                candidate_column,
            ), candidate_cell in source_by_position.items()
            if candidate_sheet == sheet_key
            and candidate_row == row
            and _cell_text(candidate_cell).strip().casefold() == "type"
        ]
        if len(type_columns) != 1:
            return False
        type_column = type_columns[0]
        mg_columns = [
            candidate_column
            for (
                candidate_sheet,
                candidate_row,
                candidate_column,
            ), candidate_cell in source_by_position.items()
            if candidate_sheet == sheet_key
            and candidate_row == row
            and _MG_FACTOR_LABEL.fullmatch(_cell_text(candidate_cell))
            is not None
        ]
        if not mg_columns:
            return False
        populated_rows = [
            candidate_row
            for candidate_row in range(row + 1, row + 11)
            if (
                type_cell := source_by_position.get(
                    (sheet_key, candidate_row, type_column)
                )
            )
            is not None
            and _cell_numeric_value(type_cell) is not None
            and (
                level_cell := source_by_position.get(
                    (
                        sheet_key,
                        candidate_row,
                        (
                            mg_columns[0]
                            if is_type_header
                            else column
                        ),
                    )
                )
            )
            is not None
            and bool(
                _cell_text(level_cell)
                or _cell_numeric_value(level_cell) is not None
            )
        ]
        return len(populated_rows) >= 2

    def is_empty_min_max_summary_placeholder(
        *,
        sheet: str,
        row: int,
        column: int,
        text: str,
    ) -> bool:
        """Exclude a blank summary row's literal ``Normal`` placeholder."""

        if text.strip().casefold() != "normal":
            return False
        sheet_key = sheet.casefold()
        if any(
            candidate_sheet == sheet_key
            and candidate_row == row
            and candidate_column != column
            and bool(
                _cell_text(candidate_cell)
                or _cell_numeric_value(candidate_cell) is not None
            )
            for (
                candidate_sheet,
                candidate_row,
                candidate_column,
            ), candidate_cell in source_by_position.items()
        ):
            return False
        previous_labels = [
            _cell_text(
                source_by_position.get(
                    (sheet_key, previous_row, column),
                    {},
                )
            ).strip().casefold()
            for previous_row in (row - 2, row - 1)
        ]
        return previous_labels == ["min", "max"]

    def is_generic_arm_column_header(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
    ) -> bool:
        """Exclude a merged ``TEST`` heading above concrete TEST 1/2/... arms."""

        _chunk, cell, sheet, row, column = primary
        if _cell_text(cell).strip().casefold() != "test":
            return False
        if str(cell.get("mergeRole") or "").casefold() != "anchor":
            return False
        merge_range = str(cell.get("mergeRange") or "").strip()
        if not merge_range:
            return False
        start_row, start_column, end_row, end_column = _range_bounds(
            merge_range
        )
        if (
            start_row != end_row
            or end_column <= start_column
            or row != start_row
            or column != start_column
        ):
            return False
        sheet_key = sheet.casefold()
        concrete_labels = {
            _normalized_narrative(_cell_text(candidate))
            for child_row in range(row + 1, row + 6)
            if (
                candidate := source_by_position.get(
                    (sheet_key, child_row, start_column)
                )
            )
            is not None
            and re.fullmatch(
                r"test\s*\d+\b.*",
                _cell_text(candidate).strip(),
                re.IGNORECASE,
            )
        }
        return len(concrete_labels) >= 2

    def is_merged_type_table_header(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
    ) -> bool:
        """Recognize merged ``Type`` headings above a concrete type block."""

        _chunk, cell, sheet, row, column = primary
        if _cell_text(cell).strip().casefold() != "type":
            return False
        if str(cell.get("mergeRole") or "").casefold() != "anchor":
            return False
        merge_range = str(cell.get("mergeRange") or "").strip()
        if not merge_range:
            return False
        start_row, start_column, end_row, end_column = _range_bounds(
            merge_range
        )
        if (
            start_row != end_row
            or end_column <= start_column
            or row != start_row
            or column != start_column
        ):
            return False
        child = source_by_position.get(
            (sheet.casefold(), row + 1, start_column)
        )
        if child is None or not _cell_text(child):
            return False
        child_merge = str(child.get("mergeRange") or "").strip()
        if not child_merge:
            return False
        child_start_row, child_start_column, child_end_row, child_end_column = (
            _range_bounds(child_merge)
        )
        return (
            child_start_row == row + 1
            and child_start_column == start_column
            and child_end_column == end_column
            and child_end_row >= child_start_row
        )

    def add(
        primary: tuple[
            dict[str, Any],
            dict[str, Any],
            str,
            int,
            int,
        ],
        role: str,
    ) -> None:
        chunk, cell, sheet, row, column = primary
        source_text = _cell_text(cell)
        if not source_text:
            return
        key = _source_cell_key(chunk, cell)
        item = role_by_key.get(key)
        if item is None:
            item = {
                "sourceCellKey": key,
                "chunkId": str(chunk.get("chunkId") or ""),
                "sheet": sheet,
                "coordinate": _coordinate(cell),
                "row": row,
                "column": column,
                "sourceText": source_text,
                "semanticRoles": [],
            }
            role_by_key[key] = item
        if role not in item["semanticRoles"]:
            item["semanticRoles"].append(role)

    for locator in locator_results:
        if not isinstance(locator, dict):
            continue
        for candidate in locator.get("candidates", []):
            if not isinstance(candidate, dict):
                continue
            for evidence in candidate.get("evidence", []):
                if not isinstance(evidence, dict):
                    continue
                role = str(evidence.get("role") or "").upper()
                if role not in _SEMANTIC_EVIDENCE_ROLES:
                    continue
                sheet = str(evidence.get("sheet") or "")
                bounds = _range_bounds(evidence.get("range"))
                for (
                    candidate_sheet,
                    _coordinate_key,
                ), primary in primary_by_coordinate.items():
                    row, column = primary[3], primary[4]
                    if (
                        candidate_sheet == sheet.casefold()
                        and bounds[0] <= row <= bounds[2]
                        and bounds[1] <= column <= bounds[3]
                    ):
                        if (
                            role == "FACTOR_LABEL"
                            and is_merged_parent_with_leaf_grid(primary)
                            or is_merged_vertical_table_header(primary)
                        ):
                            continue
                        add(primary, role)

    factor_headers: list[tuple[str, int, int]] = [
        (
            str(item["sheet"]).casefold(),
            int(item["row"]),
            int(item["column"]),
        )
        for item in role_by_key.values()
        if "FACTOR_LABEL" in item["semanticRoles"]
    ]
    for primary in primary_cells:
        _chunk, cell, sheet, row, column = primary
        text = _cell_text(cell)
        if not text:
            continue
        if is_merged_vertical_table_header(primary):
            continue
        if _COUNT_RATIO.fullmatch(text):
            add(primary, "COUNT_RATIO")
            continue
        if _UNIT_QUANTITY.fullmatch(text):
            add(primary, "UNIT_QUANTITY")
            continue
        if _ARM_LABEL.fullmatch(text):
            if is_empty_min_max_summary_placeholder(
                sheet=sheet,
                row=row,
                column=column,
                text=text,
            ) or is_generic_arm_column_header(primary):
                continue
            add(primary, "ARM_LABEL")
            continue
        structural_factor_leaf = (
            is_structural_factor_leaf_header(primary)
        )
        typed_mg_matrix_factor = (
            is_typed_mg_matrix_factor_header(primary)
        )
        if (
            (
                _FACTOR_LABEL.search(text)
                or structural_factor_leaf
                or typed_mg_matrix_factor
            )
            and looks_like_label(text)
            and not is_merged_parent_with_leaf_grid(primary)
            and not (
                _OUTCOME_LABEL.search(text)
                and has_result_below(
                    sheet=sheet,
                    row=row,
                    column=column,
                )
            )
        ):
            # Structural leaf headings establish the scope of the factor
            # levels below, but their generic words (for example ``Spec`` and
            # ``Supplier``) are not standalone canonical factor labels.
            if not structural_factor_leaf:
                add(primary, "FACTOR_LABEL")
            factor_header = (sheet.casefold(), row, column)
            if factor_header not in factor_headers:
                factor_headers.append(factor_header)
            continue
        if (
            _OUTCOME_LABEL.search(text)
            and looks_like_label(text)
            and not _RESULT_SECTION_HEADING.fullmatch(text)
            and has_result_below(
                sheet=sheet,
                row=row,
                column=column,
            )
        ):
            add(primary, "OUTCOME_LABEL")
            continue

    for primary in primary_cells:
        _chunk, cell, sheet, row, column = primary
        text = _cell_text(cell)
        if not text:
            continue
        factor_level_source = False
        for candidate_sheet, candidate_row, candidate_column in (
            factor_headers
        ):
            if (
                candidate_sheet != sheet.casefold()
                or candidate_column != column
                or not 0 < row - candidate_row <= 20
            ):
                continue
            populated_rows = sorted(
                {
                    level_row
                    for level_primary in primary_cells
                    for (
                        _level_chunk,
                        level_cell,
                        level_sheet,
                        level_row,
                        level_column,
                    ) in [level_primary]
                    if level_sheet.casefold() == candidate_sheet
                    and level_column == column
                    and candidate_row <= level_row <= row
                    and _cell_text(level_cell)
                }
            )
            if (
                populated_rows
                and populated_rows[0] == candidate_row
                and populated_rows[-1] == row
                and all(
                    later - earlier <= 2
                    for earlier, later in zip(
                        populated_rows,
                        populated_rows[1:],
                    )
                )
            ):
                factor_level_source = True
                break
        if (
            looks_like_label(text)
            and factor_level_source
            and not is_merged_type_table_header(primary)
            and (
                any(
                    candidate_sheet == sheet.casefold()
                    and candidate_row == row
                    and candidate_column != column
                    and (
                        _cell_numeric_value(candidate) is not None
                        or _CATEGORICAL_STATUS.fullmatch(
                            _cell_text(candidate)
                        )
                    )
                    for (
                        candidate_sheet,
                        candidate_row,
                        candidate_column,
                    ), candidate in source_by_position.items()
                )
                or not any(
                    candidate_sheet == sheet.casefold()
                    and candidate_row == row
                    and candidate_column != column
                    and _is_excel_error_cell(candidate)
                    for (
                        candidate_sheet,
                        candidate_row,
                        candidate_column,
                    ), candidate in source_by_position.items()
                )
            )
        ):
            add(primary, "FACTOR_LEVEL")
            key = _source_cell_key(primary[0], cell)
            roles = role_by_key[key]["semanticRoles"]
            if "ARM_LABEL" in roles and "FACTOR_LEVEL" in roles:
                roles.remove("ARM_LABEL")
    return list(role_by_key.values())


def _is_actual_status_cell(
    *,
    sheet: str,
    row: int,
    column: int,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> bool:
    sheet_key = sheet.casefold()
    same_row = [
        cell
        for (candidate_sheet, _coordinate_key), cell
        in source_cells.items()
        if candidate_sheet == sheet_key
        and _cell_position(cell)[0] == row
    ]
    if any(
        _STATUS_LEGEND_LABEL.search(_cell_text(cell))
        for cell in same_row
        if _cell_position(cell)[1] != column
    ):
        return False
    status_count = sum(
        bool(_CATEGORICAL_STATUS.fullmatch(_cell_text(cell)))
        for cell in same_row
    )
    current_cell = source_cells.get(
        (sheet_key, _coordinate_from_position(row, column))
    )
    if (
        current_cell is not None
        and str(current_cell.get("mergeRole") or "").casefold()
        == "anchor"
        and str(current_cell.get("mergeRange") or "").strip()
    ):
        _start_row, _start_column, merged_end_row, _end_column = (
            _range_bounds(current_cell.get("mergeRange"))
        )
        peer_text_count = sum(
            bool(_cell_text(cell))
            and _cell_numeric_value(cell) is None
            for cell in same_row
            if _cell_position(cell)[1] != column
        )
        peer_numeric_count = sum(
            _cell_numeric_value(cell) is not None
            for cell in same_row
            if _cell_position(cell)[1] != column
        )
        populated_below_merge = any(
            candidate_sheet == sheet_key
            and _cell_position(cell)[1] == column
            and merged_end_row < _cell_position(cell)[0]
            <= merged_end_row + 5
            and not _is_excel_error_cell(cell)
            and bool(
                _cell_text(cell)
                or _cell_numeric_value(cell) is not None
            )
            for (candidate_sheet, _coordinate_key), cell
            in source_cells.items()
        )
        # Sparse result tables sometimes use literal ``OK`` as a vertically
        # merged column heading while the data column itself is entirely
        # blank.  Several peer text headings, no peer numeric result, and no
        # populated cell below the merge are required to classify it as a
        # header; ordinary merged PASS/FAIL observations remain in scope.
        if (
            merged_end_row > row
            and peer_text_count >= 3
            and peer_numeric_count == 0
            and not populated_below_merge
        ):
            return False
    has_numeric_below = any(
        candidate_sheet == sheet_key
        and _cell_position(cell)[1] == column
        and row < _cell_position(cell)[0] <= row + 5
        and _cell_numeric_value(cell) is not None
        for (candidate_sheet, _coordinate_key), cell
        in source_cells.items()
    )
    if has_numeric_below:
        return False
    non_status_row_identity = any(
        _cell_position(cell)[1] != column
        and bool(_cell_text(cell) or _cell_numeric_value(cell) is not None)
        and not _CATEGORICAL_STATUS.fullmatch(_cell_text(cell))
        for cell in same_row
    )
    if status_count >= 2 and not non_status_row_identity:
        return False
    vertical_status_count = sum(
        1
        for (candidate_sheet, _coordinate_key), cell
        in source_cells.items()
        if candidate_sheet == sheet_key
        and _cell_position(cell)[1] == column
        and _CATEGORICAL_STATUS.fullmatch(_cell_text(cell))
    )
    if non_status_row_identity or vertical_status_count >= 2:
        return True
    return row > 1


def _coordinate_from_position(row: int, column: int) -> str:
    label = ""
    value = column
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        label = chr(ord("A") + remainder) + label
    return f"{label}{row}"


def _has_explicit_sequence_header_above(
    *,
    sheet: str,
    row: int,
    column: int,
    source_cells: dict[tuple[str, str], dict[str, Any]],
) -> bool:
    """Find the nearest same-column text header across merged data rows.

    Merged row records commonly appear at rows 17, 19, 23 beneath ``No``.
    Requiring adjacent populated rows therefore misclassifies later sequence
    values as measurements.  The nearest textual cell must itself be an
    explicit sequence label and stay within a bounded table window, so an old
    label cannot cross a distant empty section.
    """

    candidates = [
        (candidate_row, _cell_text(cell))
        for (candidate_sheet, _coordinate_key), cell
        in source_cells.items()
        for candidate_row, candidate_column in [_cell_position(cell)]
        if candidate_sheet == sheet.casefold()
        and candidate_column == column
        and candidate_row < row
        and _cell_numeric_value(cell) is None
        and _cell_text(cell)
    ]
    if not candidates:
        return False
    header_row, header_text = max(candidates, key=lambda item: item[0])
    if (
        row - header_row > 32
        or _SEQUENCE_LABEL.fullmatch(header_text.strip()) is None
    ):
        return False
    if row - header_row <= 2:
        return True
    numeric_records = sorted(
        (
            candidate_row,
            numeric_value,
        )
        for (candidate_sheet, _coordinate_key), cell
        in source_cells.items()
        for candidate_row, candidate_column in [_cell_position(cell)]
        if candidate_sheet == sheet.casefold()
        and candidate_column == column
        and header_row < candidate_row <= header_row + 32
        and (
            numeric_value := _cell_numeric_value(cell)
        )
        is not None
    )
    runs: list[list[tuple[int, float]]] = []
    current_run: list[tuple[int, float]] = []
    for candidate in numeric_records:
        if current_run and (
            candidate[0] - current_run[-1][0] > 5
            or candidate[1] <= current_run[-1][1]
        ):
            runs.append(current_run)
            current_run = []
        current_run.append(candidate)
    if current_run:
        runs.append(current_run)
    target_run = next(
        (
            values
            for values in runs
            if any(candidate_row == row for candidate_row, _value in values)
        ),
        [],
    )
    return (
        len(target_run) >= 2
        and all(value > 0 for _candidate_row, value in target_run)
        and all(
            later > earlier
            for earlier, later in zip(
                [value for _candidate_row, value in target_run],
                [value for _candidate_row, value in target_run][1:],
            )
        )
    )


def build_content_coverage_inventory(
    *,
    chunks: Sequence[dict[str, Any]],
    locator_results: Sequence[dict[str, Any]],
    expected_source_cell_keys: Sequence[str] | None = None,
) -> dict[str, Any]:
    """Inventory every owned numeric and candidate-conclusion cell."""

    primary_cells: list[
        tuple[dict[str, Any], dict[str, Any], str, int, int]
    ] = []
    owned_keys: list[str] = []
    seen_keys: set[str] = set()
    source_cells: dict[tuple[str, str], dict[str, Any]] = {}
    for chunk in chunks:
        sheet = _sheet_title(chunk)
        if not sheet:
            raise ContentCoverageError(
                f"Chunk {chunk.get('chunkId')} lacks a sheet title"
            )
        for collection in ("cells", "contextCells"):
            for cell in chunk.get(collection, []):
                coordinate = _coordinate(cell)
                if coordinate:
                    source_cells.setdefault(
                        (sheet.casefold(), coordinate),
                        cell,
                    )
        for cell in chunk.get("cells", []):
            source_key = _source_cell_key(chunk, cell)
            if source_key in seen_keys:
                raise ContentCoverageError(
                    "Content coverage found duplicate primary ownership "
                    f"for {source_key}"
                )
            seen_keys.add(source_key)
            owned_keys.append(source_key)
            row, column = _cell_position(cell)
            primary_cells.append((chunk, cell, sheet, row, column))

    if (
        expected_source_cell_keys is not None
        and owned_keys != list(expected_source_cell_keys)
    ):
        raise ContentCoverageError(
            "Content coverage source-cell ownership is not exact or "
            "source ordered"
        )

    numeric_positions = {
        (sheet.casefold(), row, column)
        for _chunk, cell, sheet, row, column in primary_cells
        if _cell_numeric_value(cell) is not None
    }
    isolated_axis_tail_positions = (
        _isolated_axis_tail_positions(
            source_cells=source_cells,
        )
    )
    duplicate_identifier_positions = (
        _duplicate_identifier_positions(
            source_cells=source_cells,
        )
    )
    horizontal_replicate_positions = (
        _horizontal_replicate_identifier_positions(
            source_cells=source_cells,
        )
    )
    hidden_error_grid_positions = (
        _hidden_error_grid_companion_positions(
            source_cells=source_cells,
        )
    )
    formula_label_input_positions = (
        _formula_label_input_positions(
            source_cells=source_cells,
        )
    )
    formula_lookup_index_positions = (
        _formula_lookup_index_positions(
            source_cells=source_cells,
        )
    )
    formula_layout_sequence_positions = (
        _formula_layout_sequence_positions(
            source_cells=source_cells,
        )
    )
    merged_structural_positions = (
        _merged_structural_numeric_positions(
            source_cells=source_cells,
        )
    )
    sheet_layout_ordinal_positions = (
        _sheet_layout_ordinal_positions(
            source_cells=source_cells,
        )
    )
    numeric_cells: list[dict[str, Any]] = []
    excluded_cells: list[dict[str, Any]] = []
    required_cells: list[dict[str, Any]] = []
    for chunk, cell, sheet, row, column in primary_cells:
        numeric_value = _cell_numeric_value(cell)
        if numeric_value is None:
            continue
        coordinate = _coordinate(cell)
        row_labels, column_labels = _nearby_labels(
            sheet=sheet,
            row=row,
            column=column,
            source_cells=source_cells,
        )
        exclusion = ""
        if _is_date_format(cell.get("numberFormat")):
            exclusion = "DATE_FORMAT"
        elif _hidden_formula_without_source_input(
            sheet=sheet,
            cell=cell,
            numeric_value=numeric_value,
            source_cells=source_cells,
        ):
            exclusion = "HIDDEN_FORMULA_WITHOUT_SOURCE_INPUT"
        elif (
            sheet.casefold(),
            row,
            column,
        ) in isolated_axis_tail_positions:
            exclusion = isolated_axis_tail_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in duplicate_identifier_positions:
            exclusion = duplicate_identifier_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in horizontal_replicate_positions:
            exclusion = horizontal_replicate_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in hidden_error_grid_positions:
            exclusion = hidden_error_grid_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in formula_label_input_positions:
            exclusion = formula_label_input_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in formula_lookup_index_positions:
            exclusion = formula_lookup_index_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in formula_layout_sequence_positions:
            exclusion = formula_layout_sequence_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in merged_structural_positions:
            exclusion = merged_structural_positions[
                (sheet.casefold(), row, column)
            ]
        elif (
            sheet.casefold(),
            row,
            column,
        ) in sheet_layout_ordinal_positions:
            exclusion = sheet_layout_ordinal_positions[
                (sheet.casefold(), row, column)
            ]
        elif _has_explicit_sequence_header_above(
            sheet=sheet,
            row=row,
            column=column,
            source_cells=source_cells,
        ):
            exclusion = "SEQUENCE_LABEL"
        elif any(
            _SEQUENCE_LABEL.fullmatch(component.strip())
            for labels in (row_labels, column_labels)
            for component in labels.split("|")[:1]
            if component.strip()
        ):
            exclusion = "SEQUENCE_LABEL"
        item = {
            "sourceCellKey": _source_cell_key(chunk, cell),
            "chunkId": str(chunk.get("chunkId") or ""),
            "sheet": sheet,
            "coordinate": coordinate,
            "row": row,
            "column": column,
            "numericValue": numeric_value,
            "numberFormat": str(cell.get("numberFormat") or ""),
            "sourceRole": _numeric_source_role(
                sheet=sheet,
                row=row,
                column=column,
                row_labels=row_labels,
                column_labels=column_labels,
                source_cells=source_cells,
            ),
            "fieldRole": _field_role(
                row_labels,
                column_labels,
            ),
            "classification": (
                "EXCLUDED_NON_RESULT" if exclusion else "REQUIRED_RESULT"
            ),
            "exclusionReason": exclusion,
        }
        numeric_cells.append(item)
        if exclusion:
            excluded_cells.append(item)
        else:
            required_cells.append(item)
    primary_by_coordinate: dict[
        tuple[str, str],
        tuple[dict[str, Any], dict[str, Any], str, int, int],
    ] = {
        (sheet.casefold(), _coordinate(cell)): (
            chunk,
            cell,
            sheet,
            row,
            column,
        )
        for chunk, cell, sheet, row, column in primary_cells
        if _coordinate(cell)
    }
    semantic_label_cells = _semantic_source_roles(
        primary_cells=primary_cells,
        primary_by_coordinate=primary_by_coordinate,
        source_cells=source_cells,
        locator_results=locator_results,
    )
    categorical_status_cells: list[dict[str, Any]] = []
    unresolved_formula_cells: list[dict[str, Any]] = []
    for chunk, cell, sheet, row, column in primary_cells:
        formula = str(cell.get("formula") or "").strip()
        cached_value = _portable_scalar(cell.get("cachedValue"))
        display_value = _portable_scalar(cell.get("displayValue"))
        if (
            formula
            and cached_value in (None, "")
            and display_value in (None, "")
        ):
            unresolved_formula_cells.append(
                {
                    "sourceCellKey": _source_cell_key(chunk, cell),
                    "chunkId": str(chunk.get("chunkId") or ""),
                    "sheet": sheet,
                    "coordinate": _coordinate(cell),
                    "row": row,
                    "column": column,
                    "formula": formula,
                }
            )
        source_text = _cell_text(cell)
        if (
            not source_text
            or not _CATEGORICAL_STATUS.fullmatch(source_text)
        ):
            continue
        if not _is_actual_status_cell(
            sheet=sheet,
            row=row,
            column=column,
            source_cells=source_cells,
        ):
            continue
        categorical_status_cells.append(
            {
                "sourceCellKey": _source_cell_key(chunk, cell),
                "chunkId": str(chunk.get("chunkId") or ""),
                "sheet": sheet,
                "coordinate": _coordinate(cell),
                "row": row,
                "column": column,
                "sourceText": source_text,
            }
        )

    narrative_cells: list[dict[str, Any]] = []
    narrative_by_key: dict[str, dict[str, Any]] = {}
    for locator in locator_results:
        if not isinstance(locator, dict):
            continue
        for candidate in locator.get("candidates", []):
            if not isinstance(candidate, dict):
                continue
            for evidence in candidate.get("evidence", []):
                role = str(
                    evidence.get("role")
                    if isinstance(evidence, dict)
                    else ""
                )
                if (
                    not isinstance(evidence, dict)
                    or not _CONCLUSION_ROLE.search(role)
                ):
                    continue
                mixed_conclusion_region = bool(
                    _MIXED_CONCLUSION_ROLE.search(role)
                )
                sheet = str(evidence.get("sheet") or "")
                start_row, start_column, end_row, end_column = (
                    _range_bounds(evidence.get("range"))
                )
                evidence_text_cells: list[
                    tuple[
                        dict[str, Any],
                        dict[str, Any],
                        str,
                        int,
                        int,
                    ]
                ] = []
                for (
                    candidate_sheet,
                    _coordinate_key,
                ), primary in primary_by_coordinate.items():
                    (
                        _chunk,
                        cell,
                        _sheet,
                        row,
                        column,
                    ) = primary
                    if (
                        candidate_sheet == sheet.casefold()
                        and start_row <= row <= end_row
                        and start_column <= column <= end_column
                        and _cell_text(cell)
                        and _cell_numeric_value(cell) is None
                    ):
                        evidence_text_cells.append(primary)
                if not evidence_text_cells:
                    continue
                evidence_text_cells.sort(
                    key=lambda item: (item[3], item[4])
                )
                heading_rows = [
                    item[3]
                    for item in evidence_text_cells
                    if _CONCLUSION_HEADING.fullmatch(
                        _cell_text(item[1])
                    )
                ]
                for chunk, cell, cell_sheet, row, column in (
                    evidence_text_cells
                ):
                    source_text = _cell_text(cell)
                    if (
                        mixed_conclusion_region
                        and not _has_conclusion_column_heading_above(
                            sheet=cell_sheet,
                            row=row,
                            column=column,
                            source_cells=source_cells,
                        )
                    ):
                        continue
                    if (
                        _CONCLUSION_HEADING.fullmatch(source_text)
                        or _CATEGORICAL_STATUS.fullmatch(source_text)
                        or row in heading_rows
                    ):
                        continue
                    source_key = _source_cell_key(chunk, cell)
                    item = narrative_by_key.get(source_key)
                    if item is None:
                        item = {
                            "sourceCellKey": source_key,
                            "chunkId": str(chunk.get("chunkId") or ""),
                            "sheet": cell_sheet,
                            "coordinate": _coordinate(cell),
                            "row": row,
                            "column": column,
                            "sourceText": source_text,
                            "locatorRoles": [],
                            "candidateKeys": [],
                        }
                        narrative_by_key[source_key] = item
                        narrative_cells.append(item)
                    candidate_key = str(candidate.get("key") or "")
                    if role not in item["locatorRoles"]:
                        item["locatorRoles"].append(role)
                    if (
                        candidate_key
                        and candidate_key not in item["candidateKeys"]
                    ):
                        item["candidateKeys"].append(candidate_key)

    # Independently discover a conclusion heading followed by adjacent
    # narrative so a locator-wide NO_CANDIDATE result cannot hide it.
    ordered_primary = sorted(
        primary_cells,
        key=lambda item: (
            item[2].casefold(),
            item[3],
            item[4],
        ),
    )
    for (
        _heading_chunk,
        heading_cell,
        heading_sheet,
        heading_row,
        _heading_column,
    ) in ordered_primary:
        if not _CONCLUSION_HEADING.fullmatch(
            _cell_text(heading_cell)
        ):
            continue
        for chunk, cell, sheet, row, column in ordered_primary:
            if (
                sheet.casefold() != heading_sheet.casefold()
                or row != heading_row + 1
            ):
                continue
            source_text = _cell_text(cell)
            if (
                len(source_text) < 12
                or _CATEGORICAL_STATUS.fullmatch(source_text)
                or _CONCLUSION_HEADING.fullmatch(source_text)
            ):
                continue
            source_key = _source_cell_key(chunk, cell)
            item = narrative_by_key.get(source_key)
            if item is None:
                item = {
                    "sourceCellKey": source_key,
                    "chunkId": str(chunk.get("chunkId") or ""),
                    "sheet": sheet,
                    "coordinate": _coordinate(cell),
                    "row": row,
                    "column": column,
                    "sourceText": source_text,
                    "locatorRoles": [
                        "DETERMINISTIC_CONCLUSION_HEADING"
                    ],
                    "candidateKeys": [],
                }
                narrative_by_key[source_key] = item
                narrative_cells.append(item)

    return {
        "schemaVersion": CONTENT_COVERAGE_SCHEMA_VERSION,
        "ownedSourceCellKeys": owned_keys,
        "ownedSourceCellCount": len(owned_keys),
        "numericCells": numeric_cells,
        "numericCellCount": len(numeric_cells),
        "requiredCells": required_cells,
        "requiredCellCount": len(required_cells),
        "excludedCells": excluded_cells,
        "excludedCellCount": len(excluded_cells),
        "unresolvedFormulaCells": unresolved_formula_cells,
        "unresolvedFormulaCellCount": len(
            unresolved_formula_cells
        ),
        "semanticLabelCells": semantic_label_cells,
        "semanticLabelCellCount": len(semantic_label_cells),
        "categoricalStatusCells": categorical_status_cells,
        "categoricalStatusCellCount": len(
            categorical_status_cells
        ),
        "narrativeConclusionCells": narrative_cells,
        "narrativeConclusionCellCount": len(narrative_cells),
    }


class _RangeCellIndex:
    """Index sparse worksheet cells by sheet and row for repeated range reads."""

    def __init__(
        self,
        cells_by_coordinate: dict[
            tuple[str, str],
            dict[str, Any],
        ],
    ) -> None:
        cells_by_sheet_row: dict[
            str,
            dict[int, list[dict[str, Any]]],
        ] = {}
        for cell in cells_by_coordinate.values():
            sheet = str(cell["sheet"]).casefold()
            row = int(cell["row"])
            cells_by_sheet_row.setdefault(sheet, {}).setdefault(
                row,
                [],
            ).append(cell)
        self._cells_by_sheet_row = cells_by_sheet_row
        self._rows_by_sheet = {
            sheet: sorted(rows)
            for sheet, rows in cells_by_sheet_row.items()
        }

    def coordinates_in_range(
        self,
        *,
        sheet: str,
        bounds: tuple[int, int, int, int],
    ) -> list[dict[str, Any]]:
        sheet_key = sheet.casefold()
        rows = self._rows_by_sheet.get(sheet_key, [])
        start_row, start_column, end_row, end_column = bounds
        start_index = bisect_left(rows, start_row)
        end_index = bisect_right(rows, end_row)
        return [
            cell
            for row in rows[start_index:end_index]
            for cell in self._cells_by_sheet_row[sheet_key][row]
            if start_column <= int(cell["column"]) <= end_column
        ]


def _coordinates_in_range(
    *,
    sheet: str,
    address: object,
    cells_by_coordinate: (
        dict[tuple[str, str], dict[str, Any]]
        | _RangeCellIndex
    ),
) -> list[dict[str, Any]]:
    bounds = _range_bounds(address)
    if isinstance(cells_by_coordinate, _RangeCellIndex):
        return cells_by_coordinate.coordinates_in_range(
            sheet=sheet,
            bounds=bounds,
        )
    start_row, start_column, end_row, end_column = bounds
    sheet_key = sheet.casefold()
    return [
        cell
        for (candidate_sheet, _coordinate_key), cell
        in cells_by_coordinate.items()
        if candidate_sheet == sheet_key
        and start_row <= int(cell["row"]) <= end_row
        and start_column <= int(cell["column"]) <= end_column
    ]


def _claim_values(observation: dict[str, Any]) -> list[float]:
    values: list[float] = []
    for field in _QUANTITATIVE_FIELDS:
        numeric = _parse_numeric(
            observation.get(field),
            declared_numeric=True,
        )
        if numeric is None:
            continue
        values.append(numeric)
    return values


def _claim_field_values(
    item: dict[str, Any],
    fields: Sequence[str],
) -> list[tuple[str, float]]:
    result: list[tuple[str, float]] = []
    for field in fields:
        numeric = _parse_numeric(
            item.get(field),
            declared_numeric=True,
        )
        if numeric is not None:
            result.append((field, numeric))
    return result


def _strict_numeric_values(
    item: dict[str, Any],
    fields: Sequence[str],
) -> list[float]:
    values: list[float] = []
    for field in fields:
        numeric = _parse_numeric(
            item.get(field),
            declared_numeric=True,
        )
        if numeric is None:
            continue
        values.append(numeric)
    return values


def _claim_matches_cell(claim: float, cell: dict[str, Any]) -> bool:
    candidates = [float(cell["numericValue"])]
    number_format = str(cell.get("numberFormat") or "")
    if "%" in number_format:
        candidates.append(candidates[0] * 100.0)
    return any(
        math.isclose(claim, candidate, rel_tol=1e-9, abs_tol=1e-9)
        for candidate in candidates
    )


def _evidence_items(value: object) -> Iterable[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    return (item for item in value if isinstance(item, dict))


def _normalized_narrative(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return re.sub(r"\s+", " ", text).strip()


def _exact_label_or_component(
    canonical_label: object,
    source_component: str,
) -> bool:
    label = _normalized_narrative(canonical_label)
    if label == source_component:
        return True
    component_tokens = re.findall(r"[\w]+", source_component)
    label_tokens = re.findall(r"[\w]+", label)
    if (
        not component_tokens
        or source_component in {"result", "outcome", "response"}
    ):
        return False
    width = len(component_tokens)
    return any(
        label_tokens[index:index + width] == component_tokens
        for index in range(len(label_tokens) - width + 1)
    )


def _cover_exact_numeric_evidence(
    *,
    claims: Sequence[float],
    claim_fields: Sequence[str] | None = None,
    evidence_items: Iterable[dict[str, Any]],
    cells_by_coordinate: (
        dict[tuple[str, str], dict[str, Any]]
        | _RangeCellIndex
    ),
    covered: dict[str, str],
    reason: str,
    used_sources: set[str] | None = None,
    binding_errors: list[str] | None = None,
    required_claim_count: int | None = None,
) -> set[str]:
    """Injectively match exact source cells to quantitative claim slots."""

    if not claims:
        return set()
    normalized_fields = list(claim_fields or [""] * len(claims))
    if len(normalized_fields) != len(claims):
        raise ContentCoverageError(
            "Numeric claim fields do not align with claim values"
        )
    used = used_sources if used_sources is not None else set()

    def field_matches(
        claim_index: int,
        cell: dict[str, Any],
    ) -> bool:
        cell_role = str(cell.get("fieldRole") or "")
        claim_role = {
            "numerator": "NUMERATOR",
            "denominator": "DENOMINATOR",
            "sampleSize": "DENOMINATOR",
            "min": "MIN",
            "max": "MAX",
            "average": "AVERAGE",
            "baselineCondition": "BASELINE",
            "changedCondition": "CHANGED",
        }.get(normalized_fields[claim_index], "")
        # A row can legitimately be named "Changed condition" while every
        # metric cell in that row is still a generic observation value.
        # Enforce field binding only when the canonical claim itself declares
        # a field-specific role.
        if not claim_role or not cell_role:
            return True
        return claim_role == cell_role

    candidate_cells: dict[str, dict[str, Any]] = {}
    value_candidate_sources: set[str] = set()
    value_candidate_claims: set[float] = set()
    for evidence in evidence_items:
        for cell in _coordinates_in_range(
                sheet=str(evidence.get("sheet") or ""),
                address=evidence.get("range"),
                cells_by_coordinate=cells_by_coordinate,
        ):
            source_key = str(cell["sourceCellKey"])
            if any(
                _claim_matches_cell(claim, cell)
                for claim in claims
            ):
                value_candidate_sources.add(source_key)
                value_candidate_claims.update(
                    claim
                    for claim in claims
                    if _claim_matches_cell(claim, cell)
                )
            if source_key in used:
                continue
            if any(
                _claim_matches_cell(claim, cell)
                and field_matches(claim_index, cell)
                for claim_index, claim in enumerate(claims)
            ):
                candidate_cells.setdefault(
                    source_key,
                    cell,
                )

    ordered_cells = sorted(
        candidate_cells.values(),
        key=lambda cell: (
            str(cell["sourceCellKey"]) in covered,
            str(cell["sheet"]).casefold(),
            int(cell["row"]),
            int(cell["column"]),
            str(cell["sourceCellKey"]),
        ),
    )
    claim_to_source: dict[int, str] = {}

    def assign(
        source_key: str,
        cell: dict[str, Any],
        seen_claims: set[int],
    ) -> bool:
        for claim_index, claim in enumerate(claims):
            if (
                claim_index in seen_claims
                or not _claim_matches_cell(claim, cell)
                or not field_matches(claim_index, cell)
            ):
                continue
            seen_claims.add(claim_index)
            previous_source = claim_to_source.get(claim_index)
            if previous_source is None:
                claim_to_source[claim_index] = source_key
                return True
            previous_cell = candidate_cells[previous_source]
            if assign(
                previous_source,
                previous_cell,
                seen_claims,
            ):
                claim_to_source[claim_index] = source_key
                return True
        return False

    matched_sources: set[str] = set()
    for cell in ordered_cells:
        source_key = str(cell["sourceCellKey"])
        if assign(source_key, cell, set()):
            matched_sources = set(claim_to_source.values())
    for source_key in matched_sources:
        covered[source_key] = reason
    used.update(matched_sources)
    minimum_required = (
        len(set(claims))
        if required_claim_count is None
        else required_claim_count
    )
    minimum_required = min(
        minimum_required,
        len(value_candidate_claims),
    )
    if (
        binding_errors is not None
        and value_candidate_sources
        and len(matched_sources) < minimum_required
    ):
        binding_errors.append(
            f"{reason} bound {len(matched_sources)} of "
            f"{minimum_required} required field-specific claim slot(s)"
        )
    return matched_sources


def _cover_exact_categorical_evidence(
    *,
    claim: object,
    evidence_items: Iterable[dict[str, Any]],
    cells_by_coordinate: (
        dict[tuple[str, str], dict[str, Any]]
        | _RangeCellIndex
    ),
    covered: dict[str, str],
) -> None:
    normalized_claim = _normalized_narrative(claim)
    if not normalized_claim:
        return
    candidates: dict[str, dict[str, Any]] = {}
    for evidence in evidence_items:
        for cell in _coordinates_in_range(
            sheet=str(evidence.get("sheet") or ""),
            address=evidence.get("range"),
            cells_by_coordinate=cells_by_coordinate,
        ):
            if (
                _normalized_narrative(cell.get("sourceText"))
                == normalized_claim
            ):
                candidates.setdefault(
                    str(cell["sourceCellKey"]),
                    cell,
                )
    if not candidates:
        return
    for cell in candidates.values():
        covered[str(cell["sourceCellKey"])] = (
            "CATEGORICAL_OBSERVATION"
        )


def _evidence_covers_semantic_cell(
    evidence_items: Iterable[dict[str, Any]],
    cell: dict[str, Any],
) -> bool:
    for evidence in evidence_items:
        if (
            str(evidence.get("sheet") or "").casefold()
            != str(cell.get("sheet") or "").casefold()
        ):
            continue
        row = int(cell["row"])
        column = int(cell["column"])
        bounds = _range_bounds(evidence.get("range"))
        if (
            bounds[0] <= row <= bounds[2]
            and bounds[1] <= column <= bounds[3]
        ):
            return True
    return False


def _semantic_manifest_coverage(
    *,
    manifest: dict[str, Any],
    semantic_cells: Sequence[dict[str, Any]],
) -> dict[str, str]:
    covered: dict[str, str] = {}
    for cell in semantic_cells:
        source_key = str(cell["sourceCellKey"])
        source_text = _normalized_narrative(cell.get("sourceText"))
        roles = set(cell.get("semanticRoles", []))
        ratio_match = _COUNT_RATIO.fullmatch(
            str(cell.get("sourceText") or "")
        )
        for study in manifest.get("studies", []):
            if not isinstance(study, dict):
                continue
            study_evidence = list(
                _evidence_items(study.get("evidence", []))
            )
            if (
                "OUTCOME_LABEL" in roles
                and _exact_label_or_component(
                    study.get("title"),
                    source_text,
                )
                and any(
                    _normalized_narrative(
                        evidence.get("sourceText")
                    )
                    == source_text
                    and _evidence_covers_semantic_cell(
                        [evidence],
                        cell,
                    )
                    for evidence in study_evidence
                )
            ):
                covered[source_key] = "STUDY_TITLE"
            for context in study.get("contexts", []):
                if not isinstance(context, dict):
                    continue
                if (
                    roles.intersection(
                        {
                            "FACTOR_LABEL",
                            "FACTOR_LEVEL",
                            "UNIT_QUANTITY",
                        }
                    )
                    and any(
                        _normalized_narrative(context.get(field))
                        == source_text
                        for field in ("kind", "originalValue")
                    )
                    and _evidence_covers_semantic_cell(
                        _evidence_items(context.get("evidence", [])),
                        cell,
                    )
                ):
                    covered[source_key] = "SEMANTIC_CONTEXT"
            for factor in study.get("factors", []):
                if not isinstance(factor, dict):
                    continue
                if (
                    "FACTOR_LABEL" in roles
                    and _normalized_narrative(
                        factor.get("originalLabel")
                    )
                    == source_text
                    and _evidence_covers_semantic_cell(
                        _evidence_items(factor.get("evidence", [])),
                        cell,
                    )
                ):
                    covered[source_key] = "FACTOR_LABEL"
            for arm in study.get("arms", []):
                if not isinstance(arm, dict):
                    continue
                arm_evidence = list(
                    _evidence_items(arm.get("evidence", []))
                )
                if (
                    "ARM_LABEL" in roles
                    and any(
                        _normalized_narrative(arm.get(field))
                        == source_text
                        for field in ("label", "condition")
                    )
                    and _evidence_covers_semantic_cell(
                        arm_evidence,
                        cell,
                    )
                ):
                    covered[source_key] = "ARM_LABEL"
                elif (
                    "ARM_LABEL" in roles
                    and any(
                        _normalized_narrative(arm.get(field))
                        == source_text
                        for field in ("label", "condition")
                    )
                    and any(
                        (
                            bounds := _range_bounds(
                                evidence.get("range")
                            )
                        )
                        and (
                            bounds[2] - bounds[0] + 1
                        )
                        * (
                            bounds[3] - bounds[1] + 1
                        )
                        <= 256
                        and _exact_label_or_component(
                            evidence.get("sourceText"),
                            source_text,
                        )
                        and _evidence_covers_semantic_cell(
                            [evidence],
                            cell,
                        )
                        for evidence in study_evidence
                    )
                ):
                    covered[source_key] = (
                        "ARM_LABEL_STUDY_EVIDENCE"
                    )
                for factor_value in arm.get("factorValues", []):
                    if not isinstance(factor_value, dict):
                        continue
                    factor_value_evidence = [
                        *_evidence_items(
                            factor_value.get("evidence", [])
                        ),
                        *arm_evidence,
                    ]
                    if (
                        roles.intersection(
                            {"FACTOR_LEVEL", "UNIT_QUANTITY"}
                        )
                        and any(
                            _normalized_narrative(
                                factor_value.get(field)
                            )
                            == source_text
                            for field in (
                                "value",
                                "originalValue",
                            )
                        )
                        and _evidence_covers_semantic_cell(
                            factor_value_evidence,
                            cell,
                        )
                    ):
                        covered[source_key] = (
                            "FACTOR_LEVEL_OR_QUANTITY"
                        )
            for outcome in study.get("outcomes", []):
                if not isinstance(outcome, dict):
                    continue
                outcome_evidence = list(
                    _evidence_items(outcome.get("evidence", []))
                )
                exact_source_label_evidence = any(
                    _normalized_narrative(
                        evidence.get("sourceText")
                    )
                    == source_text
                    and _evidence_covers_semantic_cell(
                        [evidence],
                        cell,
                    )
                    for evidence in outcome_evidence
                )
                if (
                    "OUTCOME_LABEL" in roles
                    and (
                        _exact_label_or_component(
                            outcome.get("originalLabel"),
                            source_text,
                        )
                        or exact_source_label_evidence
                    )
                    and _evidence_covers_semantic_cell(
                        outcome_evidence,
                        cell,
                    )
                ):
                    covered[source_key] = "OUTCOME_LABEL"
                for observation in outcome.get("observations", []):
                    if not isinstance(observation, dict):
                        continue
                    observation_evidence = list(
                        _evidence_items(
                            observation.get("evidence", [])
                        )
                    )
                    if (
                        "FACTOR_LEVEL" in roles
                        and _normalized_narrative(
                            observation.get("stratumKey")
                        )
                        == source_text
                        and _evidence_covers_semantic_cell(
                            observation_evidence,
                            cell,
                        )
                    ):
                        covered[source_key] = (
                            "OBSERVATION_STRATUM_IDENTITY"
                        )
                    if (
                        "UNIT_QUANTITY" in roles
                        and _normalized_narrative(
                            observation.get("valueText")
                        )
                        == source_text
                        and _evidence_covers_semantic_cell(
                            observation_evidence,
                            cell,
                        )
                    ):
                        covered[source_key] = "UNIT_QUANTITY"
                    if (
                        "COUNT_RATIO" in roles
                        and ratio_match is not None
                        and _parse_numeric(
                            observation.get("numerator"),
                            declared_numeric=True,
                        )
                        == float(ratio_match.group(1))
                        and _parse_numeric(
                            observation.get("denominator"),
                            declared_numeric=True,
                        )
                        == float(ratio_match.group(2))
                        and _evidence_covers_semantic_cell(
                            observation_evidence,
                            cell,
                        )
                    ):
                        covered[source_key] = "COUNT_RATIO"
            for series in study.get("measurementSeries", []):
                if (
                    not isinstance(series, dict)
                    or str(series.get("sheet") or "").casefold()
                    != str(cell.get("sheet") or "").casefold()
                ):
                    continue
                header_bounds = _range_bounds(
                    series.get("headerRange")
                )
                value_bounds = _range_bounds(
                    series.get("valueRange")
                )
                identity_bounds = _range_bounds(
                    series.get("rowIdentityRange")
                )
                axis_source = str(
                    series.get("axisSource") or ""
                ).upper()
                horizontal_header_axis = (
                    axis_source == "HEADER"
                    and header_bounds[0] == header_bounds[2]
                    and header_bounds[1] == value_bounds[1]
                    and header_bounds[3] == value_bounds[3]
                    and header_bounds[3] > header_bounds[1]
                    and value_bounds[0] > header_bounds[2]
                )
                vertical_row_identity = (
                    axis_source == "ROW_IDENTITY"
                    and identity_bounds[1] == identity_bounds[3]
                    and identity_bounds[0] == value_bounds[0]
                    and identity_bounds[2] == value_bounds[2]
                    and identity_bounds[2] > identity_bounds[0]
                )
                if (
                    horizontal_header_axis
                    and _evidence_covers_semantic_cell(
                        [
                            {
                                "sheet": series.get("sheet"),
                                "range": series.get("headerRange"),
                            }
                        ],
                        cell,
                    )
                ):
                    covered[source_key] = (
                        "SERIES_HEADER_IDENTITY"
                    )
                if (
                    vertical_row_identity
                    and _evidence_covers_semantic_cell(
                        [
                            {
                                "sheet": series.get("sheet"),
                                "range": series.get(
                                    "rowIdentityRange"
                                ),
                            }
                        ],
                        cell,
                    )
                ):
                    covered[source_key] = (
                        "SERIES_ROW_IDENTITY"
                    )
                if "UNIT_QUANTITY" not in roles:
                    continue
                for field, reason in (
                    ("headerRange", "SERIES_AXIS_QUANTITY"),
                    ("rowIdentityRange", "SERIES_AXIS_QUANTITY"),
                    ("valueRange", "SERIES_VALUE_QUANTITY"),
                ):
                    if _evidence_covers_semantic_cell(
                        [
                            {
                                "sheet": series.get("sheet"),
                                "range": series.get(field),
                            }
                        ],
                        cell,
                    ):
                        covered[source_key] = reason
    return covered


def _observation_has_exact_count_ratio_source(
    *,
    observation: dict[str, Any],
    semantic_cells: Sequence[dict[str, Any]],
) -> bool:
    numerator = _parse_numeric(
        observation.get("numerator"),
        declared_numeric=True,
    )
    denominator = _parse_numeric(
        observation.get("denominator"),
        declared_numeric=True,
    )
    if numerator is None or denominator is None:
        return False
    evidence = list(
        _evidence_items(observation.get("evidence", []))
    )
    for cell in semantic_cells:
        if "COUNT_RATIO" not in set(cell.get("semanticRoles", [])):
            continue
        ratio_match = _COUNT_RATIO.fullmatch(
            str(cell.get("sourceText") or "")
        )
        if (
            ratio_match is not None
            and numerator == float(ratio_match.group(1))
            and denominator == float(ratio_match.group(2))
            and _evidence_covers_semantic_cell(evidence, cell)
        ):
            return True
    return False


def augment_exact_source_conclusions(
    *,
    manifest: dict[str, Any],
    inventory: dict[str, Any],
) -> dict[str, Any]:
    """Attach exact locator-proven decisions omitted as conclusion objects."""

    result = copy.deepcopy(manifest)
    studies = [
        study
        for study in result.get("studies", [])
        if isinstance(study, dict)
    ]

    def evidence_items(value: object) -> Iterable[dict[str, Any]]:
        if isinstance(value, dict):
            direct = value.get("evidence")
            if isinstance(direct, list):
                yield from _evidence_items(direct)
            for key, child in value.items():
                if key != "evidence":
                    yield from evidence_items(child)
        elif isinstance(value, list):
            for child in value:
                yield from evidence_items(child)

    for cell in inventory.get("narrativeConclusionCells", []):
        if not isinstance(cell, dict):
            continue
        source_text = str(cell.get("sourceText") or "").strip()
        source_key = str(cell.get("sourceCellKey") or "")
        sheet = str(cell.get("sheet") or "")
        coordinate = str(cell.get("coordinate") or "").upper()
        if not source_text or not source_key or not sheet or not coordinate:
            continue
        normalized_source_text = _normalized_narrative(source_text)
        already_preserved = any(
            str(conclusion.get("claimType") or "").upper()
            == "SOURCE_CONCLUSION"
            and any(
                str(evidence.get("sheet") or "").casefold()
                == sheet.casefold()
                and str(evidence.get("range") or "").upper()
                == coordinate
                and _normalized_narrative(
                    evidence.get("sourceText")
                )
                == normalized_source_text
                for evidence in _evidence_items(
                    conclusion.get("evidence", [])
                )
            )
            for study in studies
            for conclusion in study.get("conclusions", [])
            if isinstance(conclusion, dict)
        )
        if already_preserved:
            continue

        candidates: list[
            tuple[tuple[int, int, int], dict[str, Any]]
        ] = []
        for study_index, study in enumerate(studies):
            best_score: tuple[int, int, int] | None = None
            for evidence in evidence_items(study):
                if (
                    str(evidence.get("sheet") or "").casefold()
                    != sheet.casefold()
                ):
                    continue
                bounds = _range_bounds(evidence.get("range"))
                row = int(cell["row"])
                column = int(cell["column"])
                if not (
                    bounds[0] <= row <= bounds[2]
                    and bounds[1] <= column <= bounds[3]
                ):
                    continue
                exact_text = (
                    _normalized_narrative(
                        evidence.get("sourceText")
                    )
                    == normalized_source_text
                )
                area = (
                    bounds[2] - bounds[0] + 1
                ) * (
                    bounds[3] - bounds[1] + 1
                )
                score = (
                    0 if exact_text else 1,
                    area,
                    study_index,
                )
                if best_score is None or score < best_score:
                    best_score = score
            if best_score is not None:
                candidates.append((best_score, study))
        if not candidates:
            continue
        _score, target_study = min(
            candidates,
            key=lambda item: item[0],
        )
        conclusions = target_study.setdefault("conclusions", [])
        if not isinstance(conclusions, list):
            conclusions = []
            target_study["conclusions"] = conclusions
        conclusion_digest = hashlib.sha256(
            source_key.encode("utf-8")
        ).hexdigest()[:24]
        conclusions.append(
            {
                "key": f"source_conclusion_{conclusion_digest}",
                "text": source_text,
                "claimType": "SOURCE_CONCLUSION",
                "causalStrength": "DESCRIPTIVE",
                "evidence": [
                    {
                        "sheet": sheet,
                        "range": coordinate,
                        "role": "SOURCE",
                        "sourceText": source_text,
                        "note": "",
                    }
                ],
            }
        )
    return result


def validate_content_manifest_coverage(
    *,
    manifest: dict[str, Any],
    inventory: dict[str, Any],
    require_complete: bool,
) -> dict[str, Any]:
    """Fail when owned quantitative or conclusion content is unrepresented."""

    if (
        inventory.get("schemaVersion")
        != CONTENT_COVERAGE_SCHEMA_VERSION
    ):
        raise ContentCoverageError(
            "Content coverage inventory schema is invalid"
        )
    required_cells = list(inventory.get("requiredCells", []))
    cells_by_coordinate = _RangeCellIndex(
        {
            (
                str(cell["sheet"]).casefold(),
                str(cell["coordinate"]).upper(),
            ): cell
            for cell in required_cells
        }
    )
    covered: dict[str, str] = {}
    used_numeric_sources: set[str] = set()
    binding_errors: list[str] = []
    semantic_cells = list(
        inventory.get("semanticLabelCells", [])
    )

    for study in manifest.get("studies", []):
        if not isinstance(study, dict):
            continue
        study_series = [
            series
            for series in study.get("measurementSeries", [])
            if isinstance(series, dict)
        ]
        header_axis_value_rows: dict[
            tuple[str, tuple[int, int, int, int]],
            dict[int, set[int]],
        ] = {}
        for series in study_series:
            if str(series.get("axisSource") or "").upper() != "HEADER":
                continue
            header_bounds = _range_bounds(series.get("headerRange"))
            value_bounds = _range_bounds(series.get("valueRange"))
            if (
                header_bounds[0] != header_bounds[2]
                or value_bounds[1] != header_bounds[1]
                or value_bounds[3] != header_bounds[3]
                or value_bounds[0] <= header_bounds[2]
            ):
                continue
            sheet = str(series.get("sheet") or "")
            identity = (sheet.casefold(), header_bounds)
            rows_by_column = header_axis_value_rows.setdefault(
                identity,
                {
                    column: set()
                    for column in range(
                        header_bounds[1],
                        header_bounds[3] + 1,
                    )
                },
            )
            for cell in _coordinates_in_range(
                sheet=sheet,
                address=series.get("valueRange"),
                cells_by_coordinate=cells_by_coordinate,
            ):
                column = int(cell["column"])
                if column in rows_by_column:
                    rows_by_column[column].add(int(cell["row"]))
        authorized_header_axes = {
            identity
            for identity, rows_by_column
            in header_axis_value_rows.items()
            if rows_by_column
            and all(
                len(value_rows) >= 2
                for value_rows in rows_by_column.values()
            )
        }
        for series in study_series:
            series_role = str(
                series.get("seriesRole") or "RAW"
            ).upper()
            sheet = str(series.get("sheet") or "")
            value_cells = _coordinates_in_range(
                sheet=sheet,
                address=series.get("valueRange"),
                cells_by_coordinate=cells_by_coordinate,
            )
            for cell in value_cells:
                source_role = str(cell.get("sourceRole") or "RESULT")
                if (
                    series_role == "RAW"
                    and source_role in {"RESULT", "FACTOR"}
                    or series_role == "AGGREGATE"
                    and source_role in {"AGGREGATE", "RESULT"}
                ):
                    covered[str(cell["sourceCellKey"])] = (
                        f"{series_role}_SERIES_VALUE"
                    )
            axis_source = str(
                series.get("axisSource") or ""
            ).upper()
            if axis_source in {"HEADER", "ROW_IDENTITY"}:
                value_bounds = _range_bounds(
                    series.get("valueRange")
                )
                row_identity_bounds = _range_bounds(
                    series.get("rowIdentityRange")
                )
                header_identity = (
                    sheet.casefold(),
                    _range_bounds(series.get("headerRange")),
                )
                populated_value_rows = {
                    int(cell["row"])
                    for cell in value_cells
                }
                paired_multirow_identity = (
                    axis_source == "ROW_IDENTITY"
                    and row_identity_bounds[0] == value_bounds[0]
                    and row_identity_bounds[2] == value_bounds[2]
                    and row_identity_bounds[2] > row_identity_bounds[0]
                    and row_identity_bounds[1] == row_identity_bounds[3]
                    and len(populated_value_rows) >= 2
                    and populated_value_rows
                    == set(
                        range(
                            value_bounds[0],
                            value_bounds[2] + 1,
                        )
                    )
                )
                header_replicate_identity = (
                    axis_source == "HEADER"
                    and header_identity in authorized_header_axes
                    and row_identity_bounds[0]
                    == row_identity_bounds[2]
                    and row_identity_bounds[1]
                    == row_identity_bounds[3]
                    and value_bounds[0] == value_bounds[2]
                    and row_identity_bounds[0] == value_bounds[0]
                    and value_bounds[3] - value_bounds[1] >= 1
                    and header_identity[1][0]
                    == header_identity[1][2]
                    and header_identity[1][3]
                    - header_identity[1][1]
                    == value_bounds[3] - value_bounds[1]
                    and len(value_cells) >= 2
                )
                for axis_field in (
                    "headerRange",
                    "rowIdentityRange",
                ):
                    for cell in _coordinates_in_range(
                        sheet=sheet,
                        address=series.get(axis_field),
                        cells_by_coordinate=cells_by_coordinate,
                    ):
                        if (
                            axis_field == "rowIdentityRange"
                            and paired_multirow_identity
                            or axis_field == "rowIdentityRange"
                            and header_replicate_identity
                            or axis_field == "headerRange"
                            and header_identity in authorized_header_axes
                        ):
                            covered[str(cell["sourceCellKey"])] = (
                                (
                                    f"{series_role}_SERIES_"
                                    "REPLICATE_IDENTITY"
                                    if (
                                        axis_field
                                        == "rowIdentityRange"
                                        and header_replicate_identity
                                    )
                                    else (
                                        f"{series_role}_SERIES_"
                                        f"{axis_field.upper()}"
                                    )
                                )
                            )

        factor_by_key = {
            str(factor.get("key") or ""): factor
            for factor in study.get("factors", [])
            if isinstance(factor, dict)
            and str(factor.get("key") or "")
        }
        for context in study.get("contexts", []):
            if not isinstance(context, dict):
                continue
            context_claims = _claim_field_values(
                context,
                ("valueNumber", "originalValue", "value"),
            )
            _cover_exact_numeric_evidence(
                claims=[value for _field, value in context_claims],
                claim_fields=[field for field, _value in context_claims],
                evidence_items=_evidence_items(
                    context.get("evidence", [])
                ),
                cells_by_coordinate=cells_by_coordinate,
                covered=covered,
                reason="NUMERIC_CONTEXT",
                binding_errors=binding_errors,
            )
        for factor in factor_by_key.values():
            factor_claims = _claim_field_values(
                factor,
                (
                    "valueNumber",
                    "baselineCondition",
                    "changedCondition",
                ),
            )
            _cover_exact_numeric_evidence(
                claims=[value for _field, value in factor_claims],
                claim_fields=[field for field, _value in factor_claims],
                evidence_items=_evidence_items(
                    factor.get("evidence", [])
                ),
                cells_by_coordinate=cells_by_coordinate,
                covered=covered,
                reason="NUMERIC_FACTOR",
                binding_errors=binding_errors,
            )
        for arm in study.get("arms", []):
            if not isinstance(arm, dict):
                continue
            arm_evidence = list(
                _evidence_items(arm.get("evidence", []))
            )
            arm_identity_claims = _claim_field_values(
                arm,
                ("label", "condition"),
            )
            _cover_exact_numeric_evidence(
                claims=[
                    value
                    for _field, value in arm_identity_claims
                ],
                claim_fields=[
                    field
                    for field, _value in arm_identity_claims
                ],
                evidence_items=arm_evidence,
                cells_by_coordinate=cells_by_coordinate,
                covered=covered,
                reason="ARM_NUMERIC_IDENTITY",
                binding_errors=binding_errors,
            )
            arm_claims = _claim_field_values(
                arm,
                ("sampleSize",),
            )
            _cover_exact_numeric_evidence(
                claims=[value for _field, value in arm_claims],
                claim_fields=[field for field, _value in arm_claims],
                evidence_items=arm_evidence,
                cells_by_coordinate=cells_by_coordinate,
                covered=covered,
                reason="ARM_SAMPLE_SIZE",
                binding_errors=binding_errors,
            )
            for factor_value in arm.get("factorValues", []):
                if not isinstance(factor_value, dict):
                    continue
                factor = factor_by_key.get(
                    str(factor_value.get("factor") or "")
                )
                evidence_items = list(
                    _evidence_items(
                        factor_value.get("evidence", [])
                    )
                )
                evidence_items.extend(arm_evidence)
                if factor is not None:
                    evidence_items.extend(
                        _evidence_items(
                            factor.get("evidence", [])
                        )
                    )
                factor_value_claims = _claim_field_values(
                    factor_value,
                    (
                        "valueNumber",
                        "value",
                        "originalValue",
                    ),
                )
                _cover_exact_numeric_evidence(
                    claims=[
                        value
                        for _field, value in factor_value_claims
                    ],
                    claim_fields=[
                        field
                        for field, _value in factor_value_claims
                    ],
                    evidence_items=evidence_items,
                    cells_by_coordinate=cells_by_coordinate,
                    covered=covered,
                    reason="NUMERIC_FACTOR_VALUE",
                    binding_errors=binding_errors,
                )

        for outcome in study.get("outcomes", []):
            if not isinstance(outcome, dict):
                continue
            shareable_sample_size_source = (
                str(outcome.get("metricType") or "")
                .strip()
                .casefold()
                == "sample_size"
            )
            for observation in outcome.get("observations", []):
                if not isinstance(observation, dict):
                    continue
                observation_claims = _claim_field_values(
                    observation,
                    _QUANTITATIVE_FIELDS,
                )
                if _observation_has_exact_count_ratio_source(
                    observation=observation,
                    semantic_cells=semantic_cells,
                ):
                    observation_claims = [
                        (field, value)
                        for field, value in observation_claims
                        if field not in {"numerator", "denominator"}
                    ]
                if not observation_claims:
                    continue
                is_primary_scalar = (
                    len(observation_claims) == 1
                    and observation_claims[0][0] == "valueNumber"
                )
                _cover_exact_numeric_evidence(
                    claims=[
                        value
                        for _field, value in observation_claims
                    ],
                    claim_fields=[
                        field
                        for field, _value in observation_claims
                    ],
                    evidence_items=_evidence_items(
                        observation.get("evidence", [])
                    ),
                    cells_by_coordinate=cells_by_coordinate,
                    covered=covered,
                    reason="QUANTITATIVE_OBSERVATION",
                    used_sources=(
                        used_numeric_sources
                        if (
                            is_primary_scalar
                            and not shareable_sample_size_source
                        )
                        else None
                    ),
                    binding_errors=binding_errors,
                    required_claim_count=len(
                        {
                            value
                            for _field, value in observation_claims
                        }
                    ),
                )

    uncovered = [
        cell
        for cell in required_cells
        if str(cell["sourceCellKey"]) not in covered
    ]
    semantic_covered = _semantic_manifest_coverage(
        manifest=manifest,
        semantic_cells=semantic_cells,
    )
    uncovered_semantic = [
        cell
        for cell in semantic_cells
        if str(cell["sourceCellKey"]) not in semantic_covered
    ]
    categorical_cells = list(
        inventory.get("categoricalStatusCells", [])
    )
    categorical_by_coordinate = _RangeCellIndex(
        {
            (
                str(cell["sheet"]).casefold(),
                str(cell["coordinate"]).upper(),
            ): cell
            for cell in categorical_cells
        }
    )
    categorical_covered: dict[str, str] = {}
    for study in manifest.get("studies", []):
        if not isinstance(study, dict):
            continue
        for outcome in study.get("outcomes", []):
            if not isinstance(outcome, dict):
                continue
            for observation in outcome.get("observations", []):
                if not isinstance(observation, dict):
                    continue
                _cover_exact_categorical_evidence(
                    claim=observation.get("valueText"),
                    evidence_items=_evidence_items(
                        observation.get("evidence", [])
                    ),
                    cells_by_coordinate=(
                        categorical_by_coordinate
                    ),
                    covered=categorical_covered,
                )
    uncovered_categorical = [
        cell
        for cell in categorical_cells
        if str(cell["sourceCellKey"])
        not in categorical_covered
    ]
    unresolved_formula_cells = list(
        inventory.get("unresolvedFormulaCells", [])
    )
    narrative_cells = list(
        inventory.get("narrativeConclusionCells", [])
    )
    narrative_by_coordinate = _RangeCellIndex(
        {
            (
                str(cell["sheet"]).casefold(),
                str(cell["coordinate"]).upper(),
            ): cell
            for cell in narrative_cells
        }
    )
    narrative_covered: dict[str, str] = {}
    for study in manifest.get("studies", []):
        if not isinstance(study, dict):
            continue
        for conclusion in study.get("conclusions", []):
            if (
                not isinstance(conclusion, dict)
                or str(conclusion.get("claimType") or "").upper()
                != "SOURCE_CONCLUSION"
            ):
                continue
            conclusion_text = _normalized_narrative(
                conclusion.get("text")
            )
            for evidence in _evidence_items(
                conclusion.get("evidence", [])
            ):
                evidence_text = _normalized_narrative(
                    evidence.get("sourceText")
                )
                preserved_text = " ".join(
                    value
                    for value in (conclusion_text, evidence_text)
                    if value
                )
                range_cells = _coordinates_in_range(
                    sheet=str(evidence.get("sheet") or ""),
                    address=evidence.get("range"),
                    cells_by_coordinate=narrative_by_coordinate,
                )
                range_cells.sort(
                    key=lambda cell: (
                        int(cell["row"]),
                        int(cell["column"]),
                    )
                )
                search_offset = 0
                for cell in range_cells:
                    source_text = _normalized_narrative(
                        cell.get("sourceText")
                    )
                    source_offset = preserved_text.find(
                        source_text,
                        search_offset,
                    )
                    if source_text and source_offset >= 0:
                        narrative_covered[
                            str(cell["sourceCellKey"])
                        ] = "SOURCE_CONCLUSION"
                        search_offset = source_offset + len(source_text)
    uncovered_narrative = [
        cell
        for cell in narrative_cells
        if str(cell["sourceCellKey"]) not in narrative_covered
    ]
    report = {
        "schemaVersion": CONTENT_COVERAGE_SCHEMA_VERSION,
        "requiredCellCount": len(required_cells),
        "coveredCellCount": len(required_cells) - len(uncovered),
        "excludedCellCount": int(inventory.get("excludedCellCount") or 0),
        "uncoveredCellCount": len(uncovered),
        "uncoveredCells": uncovered,
        "coverageBySourceCellKey": covered,
        "semanticLabelCellCount": len(semantic_cells),
        "coveredSemanticLabelCellCount": (
            len(semantic_cells) - len(uncovered_semantic)
        ),
        "uncoveredSemanticLabelCellCount": len(
            uncovered_semantic
        ),
        "uncoveredSemanticLabelCells": uncovered_semantic,
        "semanticCoverageBySourceCellKey": semantic_covered,
        "bindingErrors": binding_errors,
        "categoricalStatusCellCount": len(categorical_cells),
        "coveredCategoricalStatusCellCount": (
            len(categorical_cells) - len(uncovered_categorical)
        ),
        "uncoveredCategoricalStatusCellCount": len(
            uncovered_categorical
        ),
        "uncoveredCategoricalStatusCells": (
            uncovered_categorical
        ),
        "categoricalCoverageBySourceCellKey": (
            categorical_covered
        ),
        "unresolvedFormulaCellCount": len(
            unresolved_formula_cells
        ),
        "unresolvedFormulaCells": unresolved_formula_cells,
        "narrativeConclusionCellCount": len(narrative_cells),
        "coveredNarrativeConclusionCellCount": (
            len(narrative_cells) - len(uncovered_narrative)
        ),
        "uncoveredNarrativeConclusionCellCount": len(
            uncovered_narrative
        ),
        "uncoveredNarrativeConclusionCells": uncovered_narrative,
        "narrativeCoverageBySourceCellKey": narrative_covered,
    }
    if require_complete and (
        uncovered
        or uncovered_semantic
        or binding_errors
        or uncovered_categorical
        or unresolved_formula_cells
        or uncovered_narrative
    ):
        quantitative_preview = ", ".join(
            f"{cell['sheet']}!{cell['coordinate']}"
            for cell in uncovered[:20]
        )
        narrative_preview = ", ".join(
            f"{cell['sheet']}!{cell['coordinate']}"
            for cell in uncovered_narrative[:20]
        )
        categorical_preview = ", ".join(
            f"{cell['sheet']}!{cell['coordinate']}"
            for cell in uncovered_categorical[:20]
        )
        formula_preview = ", ".join(
            f"{cell['sheet']}!{cell['coordinate']}"
            for cell in unresolved_formula_cells[:20]
        )
        semantic_preview = ", ".join(
            f"{cell['sheet']}!{cell['coordinate']}"
            for cell in uncovered_semantic[:20]
        )
        quantitative_suffix = (
            ""
            if len(uncovered) <= 20
            else f" (+{len(uncovered) - 20} more)"
        )
        narrative_suffix = (
            ""
            if len(uncovered_narrative) <= 20
            else f" (+{len(uncovered_narrative) - 20} more)"
        )
        categorical_suffix = (
            ""
            if len(uncovered_categorical) <= 20
            else f" (+{len(uncovered_categorical) - 20} more)"
        )
        formula_suffix = (
            ""
            if len(unresolved_formula_cells) <= 20
            else f" (+{len(unresolved_formula_cells) - 20} more)"
        )
        semantic_suffix = (
            ""
            if len(uncovered_semantic) <= 20
            else f" (+{len(uncovered_semantic) - 20} more)"
        )
        details: list[str] = []
        if uncovered:
            details.append(
                f"{len(uncovered)} quantitative cell(s): "
                f"{quantitative_preview}{quantitative_suffix}"
            )
        if uncovered_narrative:
            details.append(
                f"{len(uncovered_narrative)} source conclusion cell(s): "
                f"{narrative_preview}{narrative_suffix}"
            )
        if uncovered_semantic:
            details.append(
                f"{len(uncovered_semantic)} semantic label cell(s): "
                f"{semantic_preview}{semantic_suffix}"
            )
        if binding_errors:
            details.append(
                "field/source binding: "
                + " | ".join(binding_errors[:10])
            )
        if uncovered_categorical:
            details.append(
                f"{len(uncovered_categorical)} categorical status cell(s): "
                f"{categorical_preview}{categorical_suffix}"
            )
        if unresolved_formula_cells:
            details.append(
                f"{len(unresolved_formula_cells)} unresolved formula cell(s): "
                f"{formula_preview}{formula_suffix}"
            )
        raise ContentCoverageError(
            "Source content coverage is incomplete; "
            + "; ".join(details)
        )
    return report


__all__ = [
    "CONTENT_COVERAGE_SCHEMA_VERSION",
    "ContentCoverageError",
    "augment_exact_source_conclusions",
    "build_content_coverage_inventory",
    "validate_content_manifest_coverage",
]
