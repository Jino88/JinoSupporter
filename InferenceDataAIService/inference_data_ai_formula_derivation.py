"""Deterministic, fail-closed derivation for a restricted Excel formula grammar.

Capture v2 deliberately stores workbook source facts without evaluating
formulas.  This module keeps that contract intact: it evaluates only a small,
audited grammar from semantic packet cells and returns a separate provenance
overlay.  Callers may project that overlay onto a deep copy for content
analysis, but must never write the derived values back to Capture v2.
"""

from __future__ import annotations

import copy
import hashlib
import json
import math
import re
from dataclasses import dataclass
from typing import Any, Iterable, Sequence


FORMULA_DERIVATION_SCHEMA_VERSION = "deterministic-formula-overlay-v2"
FORMULA_EVALUATOR_VERSION = "restricted-a1-arithmetic-v2"
DERIVED_VALUE_SOURCE = "DETERMINISTIC_FORMULA_DERIVED"

_ALLOWED_FUNCTIONS = {"SUM", "MIN", "MAX", "AVERAGE"}
_A1_PATTERN = re.compile(r"\$?([A-Za-z]{1,4})\$?([1-9]\d*)")
_SAME_SHEET_DIRECT_PATTERN = re.compile(
    r"^\s*=\s*(?P<coordinate>\$?[A-Za-z]{1,4}\$?[1-9]\d*)\s*$"
)
_CROSS_SHEET_DIRECT_PATTERN = re.compile(
    r"^\s*=\s*(?:'(?P<quoted>(?:[^']|'')+)'|"
    r"(?P<plain>[A-Za-z_][A-Za-z0-9_.]*))!\s*"
    r"(?P<coordinate>\$?[A-Za-z]{1,4}\$?[1-9]\d*)\s*$"
)
_CELL_EXPRESSION = r"\$?[A-Za-z]{1,4}\$?[1-9]\d*"
_NUMBER_EXPRESSION = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?"
_TEXT_LITERAL_EXPRESSION = r'"(?:[^"]|"")*"'
_TEXT_CONCATENATION_PATTERN = re.compile(
    rf"^\s*=\s*(?:{_CELL_EXPRESSION}|{_NUMBER_EXPRESSION}|"
    rf"{_TEXT_LITERAL_EXPRESSION})(?:\s*&\s*(?:{_CELL_EXPRESSION}|"
    rf"{_NUMBER_EXPRESSION}|{_TEXT_LITERAL_EXPRESSION}))+\s*$",
    re.IGNORECASE,
)
_TEXT_IF_AND_PATTERN = re.compile(
    rf"^\s*=\s*IF\s*\(\s*AND\s*\(\s*{_CELL_EXPRESSION}\s*"
    rf"(?:<=|>=|<>|=|<|>)\s*{_NUMBER_EXPRESSION}\s*,\s*"
    rf"{_CELL_EXPRESSION}\s*(?:<=|>=|<>|=|<|>)\s*"
    rf"{_NUMBER_EXPRESSION}\s*\)\s*,\s*{_TEXT_LITERAL_EXPRESSION}\s*,\s*"
    rf"{_TEXT_LITERAL_EXPRESSION}\s*\)\s*$",
    re.IGNORECASE,
)
_TOKEN_PATTERN = re.compile(
    r"\s*(?:"
    r"(?P<NUMBER>(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)"
    r"|(?P<CELL>\$?[A-Za-z]{1,4}\$?[1-9]\d*)"
    r"|(?P<IDENT>[A-Za-z_][A-Za-z0-9_.]*)"
    r"|(?P<OP>[+\-*/(),:])"
    r")"
)


class FormulaDerivationError(ValueError):
    """Raised when deterministic derivation cannot be proven safe."""


class _DivisionByZero(Exception):
    pass


class _NonNumericFormula(Exception):
    pass


@dataclass(frozen=True)
class _Token:
    kind: str
    value: str


@dataclass(frozen=True)
class _RangeValue:
    values: tuple[float, ...]


@dataclass
class _CellRecord:
    sheet: str
    sheet_index: int
    coordinate: str
    source_cell_key: str
    cell: dict[str, Any]


def _canonical_json(value: object) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )


def _sha256(value: object) -> str:
    return hashlib.sha256(_canonical_json(value).encode("utf-8")).hexdigest()


def _portable_scalar(value: object) -> object:
    if isinstance(value, dict):
        value_type = str(value.get("type") or "").lower()
        if value_type in {"date", "datetime", "time", "timedelta"}:
            return value
        if "value" in value:
            return value["value"]
    return value


def _finite_number(value: object) -> float | None:
    value = _portable_scalar(value)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    number = float(value)
    if not math.isfinite(number):
        raise FormulaDerivationError(
            "Formula derivation encountered a non-finite source number"
        )
    return 0.0 if number == 0 else number


def _normalize_coordinate(value: object) -> str:
    text = str(value or "").strip()
    match = _A1_PATTERN.fullmatch(text)
    if match is None:
        raise FormulaDerivationError(
            f"Formula derivation received an invalid A1 coordinate: {value}"
        )
    return f"{match.group(1).upper()}{int(match.group(2))}"


def _direct_reference(
    formula: str,
    *,
    owner_sheet: str,
) -> tuple[str, str] | None:
    same_sheet_match = _SAME_SHEET_DIRECT_PATTERN.fullmatch(formula)
    if same_sheet_match is not None:
        return owner_sheet, _normalize_coordinate(
            same_sheet_match.group("coordinate")
        )
    cross_sheet_match = _CROSS_SHEET_DIRECT_PATTERN.fullmatch(formula)
    if cross_sheet_match is None:
        return None
    quoted = cross_sheet_match.group("quoted")
    sheet = (
        quoted.replace("''", "'")
        if quoted is not None
        else str(cross_sheet_match.group("plain"))
    )
    return sheet, _normalize_coordinate(
        cross_sheet_match.group("coordinate")
    )


def _non_numeric_formula_reason(formula: str) -> str | None:
    if _TEXT_CONCATENATION_PATTERN.fullmatch(formula):
        return "TEXT_CONCATENATION"
    if _TEXT_IF_AND_PATTERN.fullmatch(formula):
        return "TEXT_IF_AND"
    return None


def _column_number(label: str) -> int:
    value = 0
    for char in label.upper():
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def _column_label(number: int) -> str:
    chars: list[str] = []
    while number:
        number, remainder = divmod(number - 1, 26)
        chars.append(chr(ord("A") + remainder))
    return "".join(reversed(chars))


def _coordinate_parts(coordinate: str) -> tuple[int, int]:
    match = _A1_PATTERN.fullmatch(coordinate)
    if match is None:
        raise FormulaDerivationError(
            f"Formula derivation received an invalid A1 coordinate: {coordinate}"
        )
    return int(match.group(2)), _column_number(match.group(1))


def _expand_range(start: str, end: str) -> Iterable[str]:
    start_row, start_column = _coordinate_parts(start)
    end_row, end_column = _coordinate_parts(end)
    if end_row < start_row or end_column < start_column:
        raise FormulaDerivationError(
            f"Formula derivation received a reversed range: {start}:{end}"
        )
    for row in range(start_row, end_row + 1):
        for column in range(start_column, end_column + 1):
            yield f"{_column_label(column)}{row}"


def _sheet_title(chunk: dict[str, Any]) -> str:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return str(sheet.get("title") or "")
    return str(sheet or "")


def _sheet_index(chunk: dict[str, Any]) -> int:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return int(sheet.get("sheetIndex") or 0)
    return int(chunk.get("sheetIndex") or 0)


def _source_revision(
    chunks: Sequence[dict[str, Any]],
) -> tuple[str, str]:
    identities: set[tuple[str, str]] = set()
    for chunk in chunks:
        source = chunk.get("sourceRevision")
        if not isinstance(source, dict):
            raise FormulaDerivationError(
                "Every formula-derivation chunk requires sourceRevision"
            )
        revision_uid = str(source.get("revisionUid") or "").strip()
        content_sha256 = str(source.get("contentSha256") or "").strip().lower()
        if not revision_uid or not re.fullmatch(r"[0-9a-f]{64}", content_sha256):
            raise FormulaDerivationError(
                "Formula derivation requires revisionUid and contentSha256"
            )
        identities.add((revision_uid, content_sha256))
    if len(identities) != 1:
        raise FormulaDerivationError(
            "Formula derivation cannot mix source revisions"
        )
    return next(iter(identities))


def _cell_signature(cell: dict[str, Any]) -> str:
    source_fields = {
        key: cell.get(key)
        for key in (
            "coordinate",
            "sourceCellKey",
            "formula",
            "rawValue",
            "cachedValue",
            "displayValue",
            "dataType",
            "cachedDataType",
            "numberFormat",
            "valueSource",
        )
        if key in cell
    }
    if source_fields.get("valueSource") == DERIVED_VALUE_SOURCE:
        source_fields["cachedValue"] = None
        source_fields["displayValue"] = None
        source_fields["cachedDataType"] = None
        source_fields["valueSource"] = "FORMULA_NO_CACHE"
    return _canonical_json(source_fields)


def _cell_records(
    chunks: Sequence[dict[str, Any]],
) -> tuple[
    dict[tuple[str, str], _CellRecord],
    dict[str, _CellRecord],
    set[str],
]:
    by_coordinate: dict[tuple[str, str], _CellRecord] = {}
    derivation_targets: dict[str, _CellRecord] = {}
    signatures: dict[tuple[str, str], str] = {}
    sheets: set[str] = set()

    for chunk in chunks:
        sheet = _sheet_title(chunk)
        sheet_index = _sheet_index(chunk)
        if not sheet or sheet_index <= 0:
            raise FormulaDerivationError(
                "Formula derivation requires a titled, indexed sheet"
            )
        sheets.add(sheet)
        for collection_name in ("cells", "contextCells"):
            cells = chunk.get(collection_name, [])
            if not isinstance(cells, list):
                raise FormulaDerivationError(
                    f"Chunk {chunk.get('chunkId')} {collection_name} must be a list"
                )
            for cell in cells:
                if not isinstance(cell, dict):
                    raise FormulaDerivationError(
                        "Formula derivation received a non-object cell"
                    )
                coordinate = _normalize_coordinate(
                    cell.get("coordinate") or cell.get("c")
                )
                source_cell_key = str(cell.get("sourceCellKey") or "").strip()
                if not source_cell_key:
                    raise FormulaDerivationError(
                        f"{sheet}!{coordinate} has no sourceCellKey"
                    )
                record = _CellRecord(
                    sheet=sheet,
                    sheet_index=sheet_index,
                    coordinate=coordinate,
                    source_cell_key=source_cell_key,
                    cell=cell,
                )
                coordinate_key = (sheet, coordinate)
                signature = _cell_signature(cell)
                prior_signature = signatures.get(coordinate_key)
                if prior_signature is not None and prior_signature != signature:
                    raise FormulaDerivationError(
                        f"Conflicting packet copies for {sheet}!{coordinate}"
                    )
                signatures[coordinate_key] = signature
                by_coordinate.setdefault(coordinate_key, record)

                formula = str(cell.get("formula") or "")
                is_derived_projection = (
                    str(cell.get("valueSource") or "") == DERIVED_VALUE_SOURCE
                )
                has_no_cache = (
                    cell.get("cachedValue") in (None, "")
                    and cell.get("displayValue") in (None, "")
                )
                if (
                    collection_name == "cells"
                    and formula
                    and (has_no_cache or is_derived_projection)
                ):
                    prior = derivation_targets.get(source_cell_key)
                    if prior is not None and (
                        prior.sheet != sheet
                        or prior.coordinate != coordinate
                        or str(prior.cell.get("formula") or "") != formula
                    ):
                        raise FormulaDerivationError(
                            f"Conflicting formula sourceCellKey {source_cell_key}"
                        )
                    derivation_targets[source_cell_key] = record
    return by_coordinate, derivation_targets, sheets


def _tokenize(formula: str) -> list[_Token]:
    if not formula.startswith("="):
        raise FormulaDerivationError(
            f"Restricted formula must start with '=': {formula}"
        )
    expression = formula[1:]
    if not expression.strip():
        raise FormulaDerivationError("Restricted formula cannot be empty")
    if any(char in expression for char in ("!", "[", "]", "#", "'", '"', ";")):
        raise FormulaDerivationError(
            f"External, error, string, or array syntax is unsupported: {formula}"
        )
    tokens: list[_Token] = []
    position = 0
    while position < len(expression):
        match = _TOKEN_PATTERN.match(expression, position)
        if match is None:
            raise FormulaDerivationError(
                f"Unsupported token at offset {position + 1}: {formula}"
            )
        kind = str(match.lastgroup)
        value = match.group(kind)
        if kind == "OP":
            kind = {
                "+": "PLUS",
                "-": "MINUS",
                "*": "STAR",
                "/": "SLASH",
                "(": "LPAREN",
                ")": "RPAREN",
                ",": "COMMA",
                ":": "COLON",
            }[value]
        tokens.append(_Token(kind, value))
        position = match.end()
    tokens.append(_Token("EOF", ""))
    return tokens


class _Parser:
    def __init__(
        self,
        *,
        evaluator: "_Evaluator",
        record: _CellRecord,
        formula: str,
    ) -> None:
        self.evaluator = evaluator
        self.record = record
        self.formula = formula
        self.direct_reference = _direct_reference(
            formula,
            owner_sheet=record.sheet,
        )
        self.tokens = (
            [_Token("EOF", "")]
            if self.direct_reference is not None
            else _tokenize(formula)
        )
        self.position = 0
        self.dependencies: list[dict[str, Any]] = []

    def parse(self) -> float:
        if self.direct_reference is not None:
            sheet, coordinate = self.direct_reference
            value = self.evaluator.reference_value(
                self.record,
                coordinate,
                sheet=sheet,
                range_member=False,
                dependency_sink=self.dependencies,
                allow_nonnumeric_result=True,
            )
            if value is None:
                raise FormulaDerivationError(
                    f"Unexpected ignored direct reference in {self._location()}"
                )
            return self._finite(value)
        result = self._expression()
        self._require("EOF")
        if isinstance(result, _RangeValue):
            raise FormulaDerivationError(
                f"Bare ranges are unsupported in {self._location()}"
            )
        return self._finite(result)

    def _location(self) -> str:
        return f"{self.record.sheet}!{self.record.coordinate} {self.formula}"

    def _peek(self, offset: int = 0) -> _Token:
        return self.tokens[self.position + offset]

    def _accept(self, kind: str) -> _Token | None:
        token = self._peek()
        if token.kind != kind:
            return None
        self.position += 1
        return token

    def _require(self, kind: str) -> _Token:
        token = self._accept(kind)
        if token is None:
            actual = self._peek()
            raise FormulaDerivationError(
                f"Expected {kind}, found {actual.kind} in {self._location()}"
            )
        return token

    def _finite(self, value: float) -> float:
        number = float(value)
        if not math.isfinite(number):
            raise FormulaDerivationError(
                f"Non-finite result in {self._location()}"
            )
        return 0.0 if number == 0 else number

    def _expression(self) -> float:
        value = self._term()
        while self._peek().kind in {"PLUS", "MINUS"}:
            operator = self._peek().kind
            self.position += 1
            right = self._term()
            value = value + right if operator == "PLUS" else value - right
            value = self._finite(value)
        return value

    def _term(self) -> float:
        value = self._unary()
        while self._peek().kind in {"STAR", "SLASH"}:
            operator = self._peek().kind
            self.position += 1
            right = self._unary()
            if operator == "STAR":
                value = self._finite(value * right)
            else:
                if right == 0:
                    raise _DivisionByZero
                value = self._finite(value / right)
        return value

    def _unary(self) -> float:
        if self._accept("PLUS") is not None:
            return self._unary()
        if self._accept("MINUS") is not None:
            return self._finite(-self._unary())
        return self._primary()

    def _primary(self) -> float:
        number = self._accept("NUMBER")
        if number is not None:
            return self._finite(float(number.value))
        cell = self._accept("CELL")
        if cell is not None:
            if self._peek().kind == "COLON":
                raise FormulaDerivationError(
                    f"Ranges are allowed only as function arguments in "
                    f"{self._location()}"
                )
            return self._reference(cell.value, range_member=False)
        identifier = self._accept("IDENT")
        if identifier is not None:
            return self._function(identifier.value)
        if self._accept("LPAREN") is not None:
            value = self._expression()
            self._require("RPAREN")
            return value
        raise FormulaDerivationError(
            f"Expected number, cell, function, or parentheses in "
            f"{self._location()}"
        )

    def _function(self, name: str) -> float:
        function = name.upper()
        if function not in _ALLOWED_FUNCTIONS:
            raise FormulaDerivationError(
                f"Unsupported function {name} in {self._location()}"
            )
        self._require("LPAREN")
        arguments: list[float | _RangeValue] = []
        if self._peek().kind != "RPAREN":
            while True:
                arguments.append(self._function_argument())
                if self._accept("COMMA") is None:
                    break
        self._require("RPAREN")

        values: list[float] = []
        for argument in arguments:
            if isinstance(argument, _RangeValue):
                values.extend(argument.values)
            else:
                values.append(argument)
        if function == "SUM":
            return self._finite(math.fsum(values))
        if function == "AVERAGE":
            if not values:
                raise _DivisionByZero
            return self._finite(math.fsum(values) / len(values))
        if not values:
            return 0.0
        return self._finite(min(values) if function == "MIN" else max(values))

    def _function_argument(self) -> float | _RangeValue:
        if self._peek().kind == "CELL" and self._peek(1).kind == "COLON":
            start = _normalize_coordinate(self._require("CELL").value)
            self._require("COLON")
            end = _normalize_coordinate(self._require("CELL").value)
            values: list[float] = []
            for coordinate in _expand_range(start, end):
                value = self.evaluator.reference_value(
                    self.record,
                    coordinate,
                    sheet=self.record.sheet,
                    range_member=True,
                    dependency_sink=self.dependencies,
                )
                if value is not None:
                    values.append(value)
            return _RangeValue(tuple(values))
        return self._expression()

    def _reference(self, coordinate: str, *, range_member: bool) -> float:
        value = self.evaluator.reference_value(
            self.record,
            _normalize_coordinate(coordinate),
            sheet=self.record.sheet,
            range_member=range_member,
            dependency_sink=self.dependencies,
        )
        if value is None:
            raise FormulaDerivationError(
                f"Unexpected ignored direct reference in {self._location()}"
            )
        return value


class _Evaluator:
    def __init__(
        self,
        *,
        by_coordinate: dict[tuple[str, str], _CellRecord],
        targets: dict[str, _CellRecord],
        sheets: set[str],
        revision_uid: str,
        content_sha256: str,
        tolerate_unsupported: bool,
    ) -> None:
        self.by_coordinate = by_coordinate
        self.targets = targets
        self.sheets = sheets
        self.revision_uid = revision_uid
        self.content_sha256 = content_sha256
        self.tolerate_unsupported = tolerate_unsupported
        self.memo: dict[str, dict[str, Any]] = {}
        self.active: list[str] = []

    def evaluate(self, record: _CellRecord) -> dict[str, Any]:
        memoized = self.memo.get(record.source_cell_key)
        if memoized is not None:
            return memoized
        if record.source_cell_key in self.active:
            cycle_start = self.active.index(record.source_cell_key)
            cycle = self.active[cycle_start:] + [record.source_cell_key]
            raise FormulaDerivationError(
                "Formula dependency cycle detected: " + " -> ".join(cycle)
            )
        formula = str(record.cell.get("formula") or "")
        if not formula:
            raise FormulaDerivationError(
                f"Target {record.sheet}!{record.coordinate} has no formula"
            )
        self.active.append(record.source_cell_key)
        parser: _Parser | None = None
        try:
            try:
                non_numeric_reason = _non_numeric_formula_reason(formula)
                if non_numeric_reason is not None:
                    raise _NonNumericFormula(non_numeric_reason)
                parser = _Parser(evaluator=self, record=record, formula=formula)
                numeric_value = parser.parse()
                status = "NUMERIC"
                error = None
                non_numeric_reason = None
            except _DivisionByZero:
                numeric_value = None
                status = "ERROR"
                error = "#DIV/0!"
                non_numeric_reason = None
            except _NonNumericFormula as exc:
                numeric_value = None
                status = "NON_NUMERIC"
                error = None
                non_numeric_reason = str(exc)
            except FormulaDerivationError as exc:
                if not self.tolerate_unsupported:
                    raise
                numeric_value = None
                status = "UNSUPPORTED"
                error = str(exc)
                non_numeric_reason = None
        finally:
            popped = self.active.pop()
            if popped != record.source_cell_key:
                raise AssertionError("Formula evaluator stack corruption")

        dependencies = parser.dependencies if parser is not None else []
        dependency_keys = list(
            dict.fromkeys(
                str(item["sourceCellKey"])
                for item in dependencies
                if item.get("sourceCellKey")
            )
        )
        entry: dict[str, Any] = {
            "sourceCellKey": record.source_cell_key,
            "sheet": record.sheet,
            "sheetIndex": record.sheet_index,
            "coordinate": record.coordinate,
            "formula": formula,
            "status": status,
            "numericValue": numeric_value,
            "error": error,
            "nonNumericReason": non_numeric_reason,
            "numberFormat": str(record.cell.get("numberFormat") or ""),
            "dependencySourceCellKeys": dependency_keys,
            "dependencySnapshotSha256": _sha256(dependencies),
            "evaluatorVersion": FORMULA_EVALUATOR_VERSION,
        }
        entry["provenanceSha256"] = _sha256(
            {
                "source": {
                    "revisionUid": self.revision_uid,
                    "contentSha256": self.content_sha256,
                },
                **entry,
            }
        )
        self.memo[record.source_cell_key] = entry
        return entry

    def reference_value(
        self,
        owner: _CellRecord,
        coordinate: str,
        *,
        sheet: str,
        range_member: bool,
        dependency_sink: list[dict[str, Any]],
        allow_nonnumeric_result: bool = False,
    ) -> float | None:
        if sheet not in self.sheets:
            raise FormulaDerivationError(
                f"Unknown cross-sheet reference from {owner.sheet}!"
                f"{owner.coordinate} to {sheet}!{coordinate}"
            )
        record = self.by_coordinate.get((sheet, coordinate))
        if record is None:
            dependency_sink.append(
                {
                    "sheet": sheet,
                    "coordinate": coordinate,
                    "sourceCellKey": None,
                    "kind": "ABSENT_BLANK",
                    "value": None,
                }
            )
            return None if range_member else 0.0

        cell = record.cell
        formula = str(cell.get("formula") or "")
        if formula:
            cached_number = _finite_number(cell.get("cachedValue"))
            is_derived_projection = (
                str(cell.get("valueSource") or "") == DERIVED_VALUE_SOURCE
            )
            if cached_number is not None and not is_derived_projection:
                dependency_sink.append(
                    {
                        "sheet": record.sheet,
                        "coordinate": record.coordinate,
                        "sourceCellKey": record.source_cell_key,
                        "kind": "FORMULA_CACHE",
                        "formula": formula,
                        "value": cached_number,
                    }
                )
                return cached_number
            cached_value = _portable_scalar(cell.get("cachedValue"))
            if (
                cached_value not in (None, "")
                and not is_derived_projection
            ):
                if str(cached_value).upper() == "#DIV/0!":
                    raise _DivisionByZero
                if allow_nonnumeric_result:
                    raise _NonNumericFormula("DIRECT_REFERENCE_TO_TEXT")
                raise FormulaDerivationError(
                    f"Unsupported cached formula error at "
                    f"{record.sheet}!{record.coordinate}: {cached_value}"
                )
            result = self.evaluate(record)
            dependency_sink.append(
                {
                    "sheet": record.sheet,
                    "coordinate": record.coordinate,
                    "sourceCellKey": record.source_cell_key,
                    "kind": "FORMULA_DERIVATION",
                    "formula": formula,
                    "status": result["status"],
                    "numericValue": result["numericValue"],
                    "error": result["error"],
                    "provenanceSha256": result["provenanceSha256"],
                }
            )
            if result["status"] != "NUMERIC":
                if result["error"] == "#DIV/0!":
                    raise _DivisionByZero
                if (
                    allow_nonnumeric_result
                    and result["status"] == "NON_NUMERIC"
                ):
                    raise _NonNumericFormula("DIRECT_REFERENCE_TO_TEXT")
                raise FormulaDerivationError(
                    f"Unsupported derived formula error at "
                    f"{record.sheet}!{record.coordinate}: {result['error']}"
                )
            return float(result["numericValue"])

        raw = _portable_scalar(cell.get("rawValue"))
        number = _finite_number(raw)
        dependency = {
            "sheet": record.sheet,
            "coordinate": record.coordinate,
            "sourceCellKey": record.source_cell_key,
            "formula": None,
        }
        if number is not None:
            dependency.update({"kind": "RAW_NUMBER", "value": number})
            dependency_sink.append(dependency)
            return number
        if raw in (None, ""):
            dependency.update({"kind": "BLANK", "value": None})
            dependency_sink.append(dependency)
            return None if range_member else 0.0
        if isinstance(raw, str) and raw.startswith("#"):
            raise FormulaDerivationError(
                f"Source reference error at {record.sheet}!{record.coordinate}: "
                f"{raw}"
            )
        dependency.update(
            {
                "kind": "IGNORED_RANGE_TEXT"
                if range_member
                else "NON_NUMERIC_REFERENCE",
                "value": raw,
            }
        )
        dependency_sink.append(dependency)
        if range_member:
            return None
        if allow_nonnumeric_result:
            raise _NonNumericFormula("DIRECT_REFERENCE_TO_TEXT")
        raise FormulaDerivationError(
            f"Non-numeric direct reference at "
            f"{record.sheet}!{record.coordinate}"
        )


def derive_formula_overlay(
    chunks: Sequence[dict[str, Any]],
    *,
    expected_revision_uid: str | None = None,
    expected_content_sha256: str | None = None,
    tolerate_unsupported: bool = False,
) -> dict[str, Any]:
    """Evaluate missing-cache formulas into a separate deterministic overlay."""

    chunk_list = list(chunks)
    if not chunk_list:
        raise FormulaDerivationError(
            "Formula derivation requires at least one source chunk"
        )
    revision_uid, content_sha256 = _source_revision(chunk_list)
    if (
        expected_revision_uid is not None
        and revision_uid != str(expected_revision_uid)
    ):
        raise FormulaDerivationError(
            "Formula derivation revisionUid does not match the expected source"
        )
    if (
        expected_content_sha256 is not None
        and content_sha256 != str(expected_content_sha256).lower()
    ):
        raise FormulaDerivationError(
            "Formula derivation contentSha256 does not match the expected source"
        )
    by_coordinate, targets, sheets = _cell_records(chunk_list)
    evaluator = _Evaluator(
        by_coordinate=by_coordinate,
        targets=targets,
        sheets=sheets,
        revision_uid=revision_uid,
        content_sha256=content_sha256,
        tolerate_unsupported=tolerate_unsupported,
    )
    values_by_source_cell_key = {
        key: evaluator.evaluate(targets[key])
        for key in sorted(targets)
    }
    numeric_count = sum(
        entry["status"] == "NUMERIC"
        for entry in values_by_source_cell_key.values()
    )
    non_numeric_count = sum(
        entry["status"] == "NON_NUMERIC"
        for entry in values_by_source_cell_key.values()
    )
    error_counts: dict[str, int] = {}
    for entry in values_by_source_cell_key.values():
        error = entry.get("error")
        if error:
            error_counts[str(error)] = error_counts.get(str(error), 0) + 1
    coordinate_lookup = {
        _sheet_coordinate_key(entry["sheet"], entry["coordinate"]): key
        for key, entry in values_by_source_cell_key.items()
    }
    if len(coordinate_lookup) != len(values_by_source_cell_key):
        raise FormulaDerivationError(
            "Formula overlay contains duplicate sheet coordinates"
        )
    overlay: dict[str, Any] = {
        "schemaVersion": FORMULA_DERIVATION_SCHEMA_VERSION,
        "evaluatorVersion": FORMULA_EVALUATOR_VERSION,
        "source": {
            "revisionUid": revision_uid,
            "contentSha256": content_sha256,
        },
        "grammar": {
            "operators": ["unary +", "unary -", "+", "-", "*", "/"],
            "functions": sorted(_ALLOWED_FUNCTIONS),
            "references": (
                "same-sheet A1 cells and rectangular ranges, plus direct "
                "same-workbook cross-sheet A1 cells"
            ),
            "nonNumericClassification": [
                "direct text references",
                "restricted text concatenation",
                "restricted IF(AND(...), text, text)",
            ],
            "tolerateUnsupported": tolerate_unsupported,
        },
        "formulaCount": len(values_by_source_cell_key),
        "numericCount": numeric_count,
        "nonNumericCount": non_numeric_count,
        "errorCount": sum(
            entry["status"] in {"ERROR", "UNSUPPORTED"}
            for entry in values_by_source_cell_key.values()
        ),
        "errorsByCode": dict(sorted(error_counts.items())),
        "valuesBySourceCellKey": values_by_source_cell_key,
        "sourceCellKeysBySheetCoordinate": coordinate_lookup,
    }
    overlay["overlaySha256"] = _sha256(overlay)
    return overlay


def _sheet_coordinate_key(sheet: object, coordinate: object) -> str:
    return f"{str(sheet)}\u001f{_normalize_coordinate(coordinate)}"


def formula_overlay_entry(
    overlay: dict[str, Any],
    *,
    sheet: str,
    coordinate: str,
    formula: str,
    revision_uid: str | None = None,
) -> dict[str, Any] | None:
    """Return an exact formula derivation entry, failing on any mismatch."""

    return FormulaOverlayLookup(
        overlay,
        revision_uid=revision_uid,
    ).entry(
        sheet=sheet,
        coordinate=coordinate,
        formula=formula,
    )


class FormulaOverlayLookup:
    """Checksum-validated reusable lookup for evidence validation loops."""

    def __init__(
        self,
        overlay: dict[str, Any],
        *,
        revision_uid: str | None = None,
        content_sha256: str | None = None,
    ) -> None:
        _validate_overlay_envelope(overlay)
        source = overlay["source"]
        if (
            revision_uid is not None
            and str(source["revisionUid"]) != str(revision_uid)
        ):
            raise FormulaDerivationError(
                "Formula overlay does not match the requested source revision"
            )
        if (
            content_sha256 is not None
            and str(source["contentSha256"]).lower()
            != str(content_sha256).lower()
        ):
            raise FormulaDerivationError(
                "Formula overlay does not match the requested source content"
            )
        self.overlay = overlay

    @property
    def overlay_sha256(self) -> str:
        return str(self.overlay["overlaySha256"])

    def entry(
        self,
        *,
        sheet: str,
        coordinate: str,
        formula: str,
    ) -> dict[str, Any] | None:
        lookup = self.overlay["sourceCellKeysBySheetCoordinate"]
        source_cell_key = lookup.get(
            _sheet_coordinate_key(sheet, coordinate)
        )
        if source_cell_key is None:
            return None
        entry = self.overlay["valuesBySourceCellKey"].get(source_cell_key)
        if not isinstance(entry, dict):
            raise FormulaDerivationError(
                "Formula overlay coordinate index is corrupt"
            )
        if (
            str(entry.get("sheet") or "") != str(sheet)
            or _normalize_coordinate(entry.get("coordinate"))
            != _normalize_coordinate(coordinate)
            or str(entry.get("formula") or "") != str(formula)
        ):
            raise FormulaDerivationError(
                f"Formula overlay source mismatch for {sheet}!{coordinate}"
            )
        return entry


def _validate_overlay_envelope(overlay: dict[str, Any]) -> None:
    if not isinstance(overlay, dict):
        raise FormulaDerivationError("Formula overlay must be an object")
    if overlay.get("schemaVersion") != FORMULA_DERIVATION_SCHEMA_VERSION:
        raise FormulaDerivationError("Unsupported formula overlay schema")
    if overlay.get("evaluatorVersion") != FORMULA_EVALUATOR_VERSION:
        raise FormulaDerivationError("Unsupported formula evaluator version")
    expected_hash = str(overlay.get("overlaySha256") or "")
    unsigned = dict(overlay)
    unsigned.pop("overlaySha256", None)
    if not expected_hash or expected_hash != _sha256(unsigned):
        raise FormulaDerivationError("Formula overlay checksum mismatch")
    if not isinstance(overlay.get("valuesBySourceCellKey"), dict) or not isinstance(
        overlay.get("sourceCellKeysBySheetCoordinate"), dict
    ):
        raise FormulaDerivationError("Formula overlay indexes are missing")


def validate_formula_overlay(
    chunks: Sequence[dict[str, Any]],
    overlay: dict[str, Any],
) -> None:
    """Re-derive the exact overlay and reject stale or altered provenance."""

    _validate_overlay_envelope(overlay)
    source = overlay["source"]
    expected = derive_formula_overlay(
        chunks,
        expected_revision_uid=str(source["revisionUid"]),
        expected_content_sha256=str(source["contentSha256"]),
        tolerate_unsupported=bool(
            (overlay.get("grammar") or {}).get("tolerateUnsupported")
        ),
    )
    if expected != overlay:
        raise FormulaDerivationError(
            "Formula overlay is stale or does not match its source chunks"
        )


def apply_formula_overlay_to_chunks(
    chunks: Sequence[dict[str, Any]],
    overlay: dict[str, Any],
    *,
    validate: bool = True,
) -> list[dict[str, Any]]:
    """Project derived values onto deep-copied chunks without source mutation."""

    if validate:
        validate_formula_overlay(chunks, overlay)
    projected = copy.deepcopy(list(chunks))
    applied_keys: set[str] = set()
    values = overlay["valuesBySourceCellKey"]
    for chunk in projected:
        sheet = _sheet_title(chunk)
        for collection_name in ("cells", "contextCells"):
            for cell in chunk.get(collection_name, []):
                source_cell_key = str(cell.get("sourceCellKey") or "")
                entry = values.get(source_cell_key)
                if entry is None:
                    continue
                coordinate = _normalize_coordinate(
                    cell.get("coordinate") or cell.get("c")
                )
                if (
                    sheet != entry["sheet"]
                    or coordinate != entry["coordinate"]
                    or str(cell.get("formula") or "") != entry["formula"]
                ):
                    raise FormulaDerivationError(
                        f"Formula projection mismatch for {source_cell_key}"
                    )
                if entry["status"] == "NUMERIC":
                    projected_value: object = entry["numericValue"]
                    cached_data_type = "n"
                elif entry["status"] == "ERROR" and entry["error"] == "#DIV/0!":
                    projected_value = "#DIV/0!"
                    cached_data_type = "e"
                elif entry["status"] in {"NON_NUMERIC", "UNSUPPORTED"}:
                    applied_keys.add(source_cell_key)
                    continue
                else:
                    raise FormulaDerivationError(
                        f"Unsupported formula overlay status for {source_cell_key}"
                    )
                cell["cachedValue"] = projected_value
                cell["displayValue"] = projected_value
                cell["cachedDataType"] = cached_data_type
                cell["valueSource"] = DERIVED_VALUE_SOURCE
                cell["formulaDerivation"] = {
                    "schemaVersion": FORMULA_DERIVATION_SCHEMA_VERSION,
                    "evaluatorVersion": FORMULA_EVALUATOR_VERSION,
                    "status": entry["status"],
                    "provenanceSha256": entry["provenanceSha256"],
                    "overlaySha256": overlay["overlaySha256"],
                }
                applied_keys.add(source_cell_key)
    missing = set(values) - applied_keys
    if missing:
        raise FormulaDerivationError(
            "Formula overlay targets are absent from projection chunks: "
            + ", ".join(sorted(missing)[:5])
        )
    return projected


__all__ = [
    "DERIVED_VALUE_SOURCE",
    "FORMULA_DERIVATION_SCHEMA_VERSION",
    "FORMULA_EVALUATOR_VERSION",
    "FormulaDerivationError",
    "FormulaOverlayLookup",
    "apply_formula_overlay_to_chunks",
    "derive_formula_overlay",
    "formula_overlay_entry",
    "validate_formula_overlay",
]
