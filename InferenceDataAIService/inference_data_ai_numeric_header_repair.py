"""Fail-closed deterministic repair for the B09 numeric height header.

The B09 height distribution already preserves each category count and rate as
scalar observations, but its numeric category headers (Sheet1!F6:J6) are not
queryable numeric fields.  This module recognizes only that exact source and
manifest geometry and appends one shared Outcome plus three aligned RAW
measurement series.  It intentionally does not invoke AI or mutate artifacts.
"""

from __future__ import annotations

import copy
import hashlib
import json
import re
from decimal import Decimal, InvalidOperation
from typing import Any, Sequence


NUMERIC_HEADER_REPAIR_SCHEMA_VERSION = (
    "numeric-header-series-repair-target-v1"
)
B09_NUMERIC_HEADER_COVERAGE_ERROR = (
    "ContentCoverageError: Source content coverage is incomplete; "
    "5 quantitative cell(s): Sheet1!F6, Sheet1!G6, Sheet1!H6, "
    "Sheet1!I6, Sheet1!J6"
)

_TITLE = (
    "RESULT HEIGHT CHECK MATERIAL C-MG, S-MG  "
    "( Spec : 0.66~0.70mm )"
)
_STUDY_KEY = "height-category-distribution"
_HEADER_SHEET = "Sheet1"
_HEADER_RANGE = "F6:J6"
_HEADER_COORDINATES = ("F6", "G6", "H6", "I6", "J6")
_HEADER_VALUES = (
    Decimal("0.66"),
    Decimal("0.67"),
    Decimal("0.68"),
    Decimal("0.69"),
    Decimal("0.70"),
)
_FACTOR_KEY = "height-material-type"
_OUTCOME_KEY = "height-category-count-series"
_ERROR_PATTERN = re.compile(
    r"^(?:ContentCoverageError:\s*)?"
    r"Source content coverage is incomplete; "
    r"5 quantitative cell\(s\): "
    r"Sheet1!F6, Sheet1!G6, Sheet1!H6, Sheet1!I6, Sheet1!J6$"
)

_ARM_ROWS = (
    {
        "key": "height-c-mg",
        "label": "C-MG",
        "row": 7,
        "rateRow": 8,
        "seriesKey": "height-cmg-count-series",
    },
    {
        "key": "height-s-mg-long",
        "label": "S-MG Long",
        "row": 9,
        "rateRow": 10,
        "seriesKey": "height-smg-long-count-series",
    },
    {
        "key": "height-s-mg-short",
        "label": "S-MG Short",
        "row": 11,
        "rateRow": 12,
        "seriesKey": "height-smg-short-count-series",
    },
)
_CATEGORY_SUFFIXES = ("066", "067", "068", "069", "070")


class NumericHeaderRepairError(RuntimeError):
    """Raised when an attempted repair is not the exact safe projection."""


def _canonical_sha256(value: object) -> str:
    payload = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _repair_projection(
    manifest: dict[str, Any],
    study_index: int,
) -> dict[str, Any]:
    studies = manifest.get("studies")
    if not isinstance(studies, list) or not 0 <= study_index < len(studies):
        raise NumericHeaderRepairError("repair Study index is invalid")
    source = manifest.get("source")
    study = studies[study_index]
    if not isinstance(source, dict) or not isinstance(study, dict):
        raise NumericHeaderRepairError("repair projection is not canonical")
    return {
        "source": copy.deepcopy(source),
        "study": copy.deepcopy(study),
    }


def _sheet_title(chunk: dict[str, Any]) -> str:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return str(sheet.get("title") or "")
    return str(sheet or chunk.get("sheetTitle") or "")


def _coordinate(cell: dict[str, Any]) -> str:
    return str(cell.get("coordinate") or "").strip().upper()


def _source_cells(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[tuple[str, str], dict[str, Any]]:
    cells: dict[tuple[str, str], dict[str, Any]] = {}
    for chunk in focused_chunks:
        if not isinstance(chunk, dict):
            raise NumericHeaderRepairError("focused chunk is not an object")
        sheet = _sheet_title(chunk)
        if not sheet:
            raise NumericHeaderRepairError("focused chunk has no sheet title")
        for cell in chunk.get("cells", []):
            if not isinstance(cell, dict):
                raise NumericHeaderRepairError("source cell is not an object")
            coordinate = _coordinate(cell)
            if not coordinate:
                raise NumericHeaderRepairError(
                    "source cell has no coordinate"
                )
            key = (sheet.casefold(), coordinate)
            if key in cells:
                raise NumericHeaderRepairError(
                    f"duplicate source coordinate {sheet}!{coordinate}"
                )
            cells[key] = cell
    return cells


def _cell(
    cells: dict[tuple[str, str], dict[str, Any]],
    coordinate: str,
) -> dict[str, Any]:
    cell = cells.get((_HEADER_SHEET.casefold(), coordinate.upper()))
    if cell is None:
        raise NumericHeaderRepairError(
            f"required source cell {_HEADER_SHEET}!{coordinate} is missing"
        )
    return cell


def _decimal(value: object, path: str) -> Decimal:
    if isinstance(value, bool) or value in (None, ""):
        raise NumericHeaderRepairError(f"{path} is not numeric")
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise NumericHeaderRepairError(f"{path} is not numeric") from exc
    if not result.is_finite():
        raise NumericHeaderRepairError(f"{path} is not finite")
    return result


def _raw_decimal(cell: dict[str, Any], path: str) -> Decimal:
    return _decimal(cell.get("rawValue"), path)


def _raw_text(cell: dict[str, Any]) -> str:
    value = cell.get("rawValue")
    if value in (None, ""):
        value = cell.get("displayValue")
    return str(value or "")


def _number_format(cell: dict[str, Any]) -> str:
    return str(cell.get("numberFormat") or "")


def _require_unmerged(cell: dict[str, Any], path: str) -> None:
    if str(cell.get("mergeRange") or ""):
        raise NumericHeaderRepairError(f"{path} must not be merged")
    if str(cell.get("mergeRole") or "none").casefold() not in {
        "",
        "none",
    }:
        raise NumericHeaderRepairError(f"{path} has an invalid merge role")


def _require_merged_anchor(
    cell: dict[str, Any],
    path: str,
    merge_range: str,
) -> None:
    if str(cell.get("mergeRange") or "").upper() != merge_range:
        raise NumericHeaderRepairError(
            f"{path} must anchor merge {merge_range}"
        )
    if str(cell.get("mergeRole") or "").casefold() != "anchor":
        raise NumericHeaderRepairError(f"{path} must be a merge anchor")


def _validate_source_geometry(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Decimal]:
    cells = _source_cells(focused_chunks)
    numeric_values: dict[str, Decimal] = {}
    if _raw_text(_cell(cells, "C4")) != _TITLE:
        raise NumericHeaderRepairError("Sheet1!C4 title is not exact")
    if _raw_text(_cell(cells, "D6")) != "Type":
        raise NumericHeaderRepairError("Sheet1!D6 is not the Type header")
    if _raw_text(_cell(cells, "E6")) != "Q'ty check":
        raise NumericHeaderRepairError(
            "Sheet1!E6 is not the quantity header"
        )

    for coordinate, expected in zip(
        _HEADER_COORDINATES,
        _HEADER_VALUES,
    ):
        cell = _cell(cells, coordinate)
        if _raw_decimal(cell, f"Sheet1!{coordinate}") != expected:
            raise NumericHeaderRepairError(
                f"Sheet1!{coordinate} header value is not exact"
            )
        if _number_format(cell) != "0.00":
            raise NumericHeaderRepairError(
                f"Sheet1!{coordinate} header format is not 0.00"
            )
        _require_unmerged(cell, f"Sheet1!{coordinate}")
        numeric_values[coordinate] = expected

    for arm in _ARM_ROWS:
        row = int(arm["row"])
        rate_row = int(arm["rateRow"])
        label_coordinate = f"D{row}"
        quantity_coordinate = f"E{row}"
        label_cell = _cell(cells, label_coordinate)
        quantity_cell = _cell(cells, quantity_coordinate)
        if _raw_text(label_cell) != arm["label"]:
            raise NumericHeaderRepairError(
                f"Sheet1!{label_coordinate} arm label is not exact"
            )
        _require_merged_anchor(
            label_cell,
            f"Sheet1!{label_coordinate}",
            f"D{row}:D{rate_row}",
        )
        _require_merged_anchor(
            quantity_cell,
            f"Sheet1!{quantity_coordinate}",
            f"E{row}:E{rate_row}",
        )
        quantity = _raw_decimal(
            quantity_cell,
            f"Sheet1!{quantity_coordinate}",
        )
        if quantity <= 0 or quantity != quantity.to_integral_value():
            raise NumericHeaderRepairError(
                f"Sheet1!{quantity_coordinate} quantity is invalid"
            )
        numeric_values[quantity_coordinate] = quantity

        counts: list[Decimal] = []
        for column in "FGHIJ":
            coordinate = f"{column}{row}"
            cell = _cell(cells, coordinate)
            value = _raw_decimal(cell, f"Sheet1!{coordinate}")
            if value < 0 or value != value.to_integral_value():
                raise NumericHeaderRepairError(
                    f"Sheet1!{coordinate} count is invalid"
                )
            if _number_format(cell) != "0":
                raise NumericHeaderRepairError(
                    f"Sheet1!{coordinate} count format is not 0"
                )
            _require_unmerged(cell, f"Sheet1!{coordinate}")
            counts.append(value)
            numeric_values[coordinate] = value
        if sum(counts, Decimal(0)) != quantity:
            raise NumericHeaderRepairError(
                f"Sheet1!F{row}:J{row} counts do not sum to quantity"
            )

        for column, count in zip("FGHIJ", counts):
            coordinate = f"{column}{rate_row}"
            cell = _cell(cells, coordinate)
            rate = _raw_decimal(cell, f"Sheet1!{coordinate}")
            if rate != count / quantity:
                raise NumericHeaderRepairError(
                    f"Sheet1!{coordinate} rate is not count/quantity"
                )
            if _number_format(cell) != "0.0%":
                raise NumericHeaderRepairError(
                    f"Sheet1!{coordinate} rate format is not 0.0%"
                )
            _require_unmerged(cell, f"Sheet1!{coordinate}")
            numeric_values[coordinate] = rate
    return numeric_values


def _unique_by_key(
    values: object,
    path: str,
) -> dict[str, dict[str, Any]]:
    if not isinstance(values, list):
        raise NumericHeaderRepairError(f"{path} is not a list")
    result: dict[str, dict[str, Any]] = {}
    for index, value in enumerate(values):
        if not isinstance(value, dict):
            raise NumericHeaderRepairError(
                f"{path}[{index}] is not an object"
            )
        key = str(value.get("key") or "")
        if not key or key in result:
            raise NumericHeaderRepairError(
                f"{path} has an invalid or duplicate key"
            )
        result[key] = value
    return result


def _has_evidence(
    value: dict[str, Any],
    *,
    coordinate: str,
    source_text: str,
) -> bool:
    evidence = value.get("evidence")
    if not isinstance(evidence, list):
        return False
    return any(
        isinstance(item, dict)
        and str(item.get("sheet") or "") == _HEADER_SHEET
        and str(item.get("range") or "").upper() == coordinate.upper()
        and str(item.get("sourceText") or "") == source_text
        for item in evidence
    )


def _numeric_claim(value: dict[str, Any], field: str, path: str) -> Decimal:
    return _decimal(value.get(field), f"{path}.{field}")


def _validate_manifest_geometry(
    baseline: dict[str, Any],
    numeric_values: dict[str, Decimal],
) -> tuple[int, dict[str, Any]]:
    source = baseline.get("source")
    studies = baseline.get("studies")
    if (
        not isinstance(source, dict)
        or source.get("contentComplete") is not True
        or not isinstance(studies, list)
        or not studies
        or not isinstance(studies[0], dict)
    ):
        raise NumericHeaderRepairError(
            "manifest is not a complete canonical B09 baseline"
        )
    study_index = 0
    study = studies[study_index]
    if (
        str(study.get("key") or "") != _STUDY_KEY
        or str(study.get("title") or "") != _TITLE
    ):
        raise NumericHeaderRepairError(
            "height distribution Study identity is not exact"
        )

    factors = _unique_by_key(
        study.get("factors"),
        "studies[0].factors",
    )
    factor = factors.get(_FACTOR_KEY)
    if (
        factor is None
        or str(factor.get("originalLabel") or "") != "Type"
        or not _has_evidence(
            factor,
            coordinate="D6",
            source_text="Type",
        )
    ):
        raise NumericHeaderRepairError(
            "height material factor is not source-exact"
        )

    arms = _unique_by_key(study.get("arms"), "studies[0].arms")
    outcomes = _unique_by_key(
        study.get("outcomes"),
        "studies[0].outcomes",
    )
    series = _unique_by_key(
        study.get("measurementSeries"),
        "studies[0].measurementSeries",
    )

    input_outcome = outcomes.get("height-input")
    if input_outcome is None:
        raise NumericHeaderRepairError("height input Outcome is missing")
    input_observations = _unique_by_key(
        input_outcome.get("observations"),
        "height-input.observations",
    )

    for arm in _ARM_ROWS:
        row = int(arm["row"])
        rate_row = int(arm["rateRow"])
        arm_value = arms.get(str(arm["key"]))
        if (
            arm_value is None
            or str(arm_value.get("role") or "") != "OTHER"
            or str(arm_value.get("label") or "") != arm["label"]
            or _numeric_claim(
                arm_value,
                "sampleSize",
                f"arms[{arm['key']}]",
            )
            != Decimal("500")
            or not _has_evidence(
                arm_value,
                coordinate=f"D{row}",
                source_text=str(arm["label"]),
            )
            or not _has_evidence(
                arm_value,
                coordinate=f"E{row}",
                source_text="500",
            )
        ):
            raise NumericHeaderRepairError(
                f"arm {arm['key']} is not source-exact"
            )
        input_matches = [
            observation
            for observation in input_observations.values()
            if str(observation.get("arm") or "") == arm["key"]
        ]
        if (
            len(input_matches) != 1
            or _numeric_claim(
                input_matches[0],
                "valueNumber",
                f"height-input[{arm['key']}]",
            )
            != Decimal("500")
            or _numeric_claim(
                input_matches[0],
                "sampleSize",
                f"height-input[{arm['key']}]",
            )
            != numeric_values[f"E{row}"]
            or not _has_evidence(
                input_matches[0],
                coordinate=f"E{row}",
                source_text="500",
            )
        ):
            raise NumericHeaderRepairError(
                f"height input for {arm['key']} is not source-exact"
            )

        for column, suffix, header in zip(
            "FGHIJ",
            _CATEGORY_SUFFIXES,
            _HEADER_VALUES,
        ):
            count_outcome = outcomes.get(f"height-{suffix}-count")
            rate_outcome = outcomes.get(f"height-{suffix}-rate")
            if (
                count_outcome is None
                or rate_outcome is None
                or str(count_outcome.get("metricType") or "")
                != "height_category_count"
                or str(rate_outcome.get("metricType") or "")
                != "height_category_rate"
                or not _has_evidence(
                    count_outcome,
                    coordinate=f"{column}6",
                    source_text=format(header, ".2f"),
                )
                or not _has_evidence(
                    rate_outcome,
                    coordinate=f"{column}6",
                    source_text=format(header, ".2f"),
                )
            ):
                raise NumericHeaderRepairError(
                    f"height category {suffix} Outcomes are not source-exact"
                )
            count_observations = [
                observation
                for observation in _unique_by_key(
                    count_outcome.get("observations"),
                    f"height-{suffix}-count.observations",
                ).values()
                if str(observation.get("arm") or "") == arm["key"]
            ]
            rate_observations = [
                observation
                for observation in _unique_by_key(
                    rate_outcome.get("observations"),
                    f"height-{suffix}-rate.observations",
                ).values()
                if str(observation.get("arm") or "") == arm["key"]
            ]
            if (
                len(count_observations) != 1
                or len(rate_observations) != 1
                or _numeric_claim(
                    count_observations[0],
                    "valueNumber",
                    "count observation",
                )
                != numeric_values[f"{column}{row}"]
                or _numeric_claim(
                    count_observations[0],
                    "numerator",
                    "count observation",
                )
                != numeric_values[f"{column}{row}"]
                or not _has_evidence(
                    count_observations[0],
                    coordinate=f"{column}{row}",
                    source_text=str(
                        int(
                            _numeric_claim(
                                count_observations[0],
                                "valueNumber",
                                "count observation",
                            )
                        )
                    ),
                )
                or not _has_evidence(
                    rate_observations[0],
                    coordinate=f"{column}{rate_row}",
                    source_text=str(
                        rate_observations[0].get("valueText") or ""
                    ),
                )
                or _numeric_claim(
                    count_observations[0],
                    "denominator",
                    "count observation",
                )
                != numeric_values[f"E{row}"]
                or _numeric_claim(
                    count_observations[0],
                    "sampleSize",
                    "count observation",
                )
                != numeric_values[f"E{row}"]
                or _numeric_claim(
                    rate_observations[0],
                    "denominator",
                    "rate observation",
                )
                != numeric_values[f"E{row}"]
                or _numeric_claim(
                    rate_observations[0],
                    "sampleSize",
                    "rate observation",
                )
                != numeric_values[f"E{row}"]
                or _numeric_claim(
                    rate_observations[0],
                    "numerator",
                    "rate observation",
                )
                != _numeric_claim(
                    count_observations[0],
                    "valueNumber",
                    "count observation",
                )
                or _numeric_claim(
                    rate_observations[0],
                    "valueNumber",
                    "rate observation",
                )
                != numeric_values[f"{column}{rate_row}"] * Decimal("100")
            ):
                raise NumericHeaderRepairError(
                    f"height category {suffix} observations are not exact"
                )

    outcome_record, series_records = _repair_records()
    outcome_collision = outcomes.get(_OUTCOME_KEY)
    series_collisions = [
        series.get(str(record["key"]))
        for record in series_records
    ]
    if outcome_collision is None and all(
        collision is None for collision in series_collisions
    ):
        return study_index, study
    if (
        outcome_collision == outcome_record
        and series_collisions == series_records
    ):
        return study_index, study
    raise NumericHeaderRepairError(
        "numeric header repair is partial or conflicts with canonical keys"
    )


def _repair_records() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    outcome = {
        "key": _OUTCOME_KEY,
        "originalLabel": _TITLE,
        "metricType": "height_category_count_by_height",
        "unit": "",
        "favorableDirection": "UNKNOWN",
        "evidence": [
            {
                "sheet": _HEADER_SHEET,
                "range": "C4",
                "role": "SOURCE",
                "sourceText": _TITLE,
                "note": "",
            }
        ],
        "observations": [],
    }
    series: list[dict[str, Any]] = []
    for arm in _ARM_ROWS:
        row = int(arm["row"])
        series.append(
            {
                "key": str(arm["seriesKey"]),
                "seriesRole": "RAW",
                "aggregationFunction": "",
                "aggregateOfSeries": [],
                "outcome": _OUTCOME_KEY,
                "arm": str(arm["key"]),
                "sheet": _HEADER_SHEET,
                "headerRange": _HEADER_RANGE,
                "valueRange": f"F{row}:J{row}",
                "rowIdentityRange": f"D{row}:D{row}",
                "aggregateReplicateRanges": [],
                "axisSource": "HEADER",
                "axisLabel": "Height category",
                "axisUnit": "mm",
                "valueUnit": "",
                "stratumKey": "height-category-count",
                "verificationStatus": "NEEDS_REVIEW",
            }
        )
    return outcome, series


def _append_repair_records(
    manifest: dict[str, Any],
    study_index: int,
) -> dict[str, Any]:
    repaired = copy.deepcopy(manifest)
    study = repaired["studies"][study_index]
    outcome, series = _repair_records()
    outcomes_by_key = _unique_by_key(
        study.get("outcomes"),
        f"studies[{study_index}].outcomes",
    )
    series_by_key = _unique_by_key(
        study.get("measurementSeries"),
        f"studies[{study_index}].measurementSeries",
    )
    outcome_value = outcomes_by_key.get(_OUTCOME_KEY)
    series_values = [
        series_by_key.get(str(record["key"]))
        for record in series
    ]
    if outcome_value is None and all(
        value is None for value in series_values
    ):
        study["outcomes"].append(copy.deepcopy(outcome))
        study["measurementSeries"].extend(copy.deepcopy(series))
        return repaired
    if outcome_value == outcome and series_values == series:
        return repaired
    raise NumericHeaderRepairError(
        "numeric header repair is partial or conflicts with canonical keys"
    )


def numeric_header_series_repair_target(
    validation_error: str,
    baseline: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Return an exact repair target, or ``None`` for any ambiguous input."""

    if _ERROR_PATTERN.fullmatch(str(validation_error or "").strip()) is None:
        return None
    try:
        if not isinstance(baseline, dict):
            raise NumericHeaderRepairError("baseline is not an object")
        numeric_values = _validate_source_geometry(focused_chunks)
        study_index, _study = _validate_manifest_geometry(
            baseline,
            numeric_values,
        )
        outcome, series = _repair_records()
        repaired = _append_repair_records(baseline, study_index)
        return {
            "schemaVersion": NUMERIC_HEADER_REPAIR_SCHEMA_VERSION,
            "studyIndex": study_index,
            "studyKey": _STUDY_KEY,
            "sheet": _HEADER_SHEET,
            "headerRange": _HEADER_RANGE,
            "headerValues": [
                float(value) for value in _HEADER_VALUES
            ],
            "outcome": outcome,
            "measurementSeries": series,
            "baselineProjectionSha256": _canonical_sha256(
                _repair_projection(baseline, study_index)
            ),
            "repairedProjectionSha256": _canonical_sha256(
                _repair_projection(repaired, study_index)
            ),
        }
    except (KeyError, TypeError, ValueError, NumericHeaderRepairError):
        return None


def _validate_target(target: dict[str, Any]) -> int:
    if not isinstance(target, dict):
        raise NumericHeaderRepairError("repair target is not an object")
    outcome, series = _repair_records()
    if (
        target.get("schemaVersion")
        != NUMERIC_HEADER_REPAIR_SCHEMA_VERSION
        or target.get("studyIndex") != 0
        or target.get("studyKey") != _STUDY_KEY
        or target.get("sheet") != _HEADER_SHEET
        or target.get("headerRange") != _HEADER_RANGE
        or target.get("headerValues")
        != [float(value) for value in _HEADER_VALUES]
        or target.get("outcome") != outcome
        or target.get("measurementSeries") != series
    ):
        raise NumericHeaderRepairError("repair target is not exact")
    for field in (
        "baselineProjectionSha256",
        "repairedProjectionSha256",
    ):
        value = str(target.get(field) or "")
        if len(value) != 64 or any(
            char not in "0123456789abcdef" for char in value
        ):
            raise NumericHeaderRepairError(
                f"repair target {field} is invalid"
            )
    return 0


def apply_numeric_header_series_repair(
    baseline: dict[str, Any],
    target: dict[str, Any],
) -> dict[str, Any]:
    """Apply only the target's append-only numeric-header projection."""

    study_index = _validate_target(target)
    current_hash = _canonical_sha256(
        _repair_projection(baseline, study_index)
    )
    baseline_hash = str(target["baselineProjectionSha256"])
    repaired_hash = str(target["repairedProjectionSha256"])
    if current_hash not in {baseline_hash, repaired_hash}:
        raise NumericHeaderRepairError(
            "baseline changed within the protected repair projection"
        )
    repaired = _append_repair_records(baseline, study_index)
    actual_hash = _canonical_sha256(
        _repair_projection(repaired, study_index)
    )
    if actual_hash != repaired_hash:
        raise NumericHeaderRepairError(
            "numeric header repair produced an unexpected projection"
        )
    return repaired


def validate_numeric_header_series_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    target: dict[str, Any],
) -> None:
    """Reject every mutation outside the exact append-only repair."""

    expected = apply_numeric_header_series_repair(baseline, target)
    if repaired != expected:
        raise NumericHeaderRepairError(
            "numeric header repair changed fields outside the exact "
            "Outcome and measurementSeries projection"
        )


__all__ = [
    "B09_NUMERIC_HEADER_COVERAGE_ERROR",
    "NUMERIC_HEADER_REPAIR_SCHEMA_VERSION",
    "NumericHeaderRepairError",
    "apply_numeric_header_series_repair",
    "numeric_header_series_repair_target",
    "validate_numeric_header_series_repair",
]
