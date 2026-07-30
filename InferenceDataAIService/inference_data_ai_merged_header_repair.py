"""Deterministic repair for measurement series that reuse merged headers."""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import Any


_ERROR_PATTERN = re.compile(
    r"(?:^|:\s*)studies\[(?P<study>\d+)\]\.measurementSeries"
    r"\[(?P<series>\d+)\]\.headerRange has multiple logical header "
    r"cells that resolve to the same merged anchor"
)
_A1_PATTERN = re.compile(
    r"^\$?(?P<start_column>[A-Za-z]{1,4})\$?"
    r"(?P<start_row>[1-9]\d*)"
    r"(?:\:\$?(?P<end_column>[A-Za-z]{1,4})\$?"
    r"(?P<end_row>[1-9]\d*))?$"
)


class MergedHeaderRepairError(RuntimeError):
    """Raised when a merged-header repair cannot remain lossless."""


@dataclass(frozen=True)
class MergedHeaderRepairTarget:
    study_index: int
    series_index: int


def _column_number(label: str) -> int:
    result = 0
    for character in label.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _column_label(number: int) -> str:
    if number < 1:
        raise MergedHeaderRepairError("column number must be positive")
    characters: list[str] = []
    while number:
        number, remainder = divmod(number - 1, 26)
        characters.append(chr(ord("A") + remainder))
    return "".join(reversed(characters))


def _range_bounds(address: object) -> tuple[int, int, int, int]:
    match = _A1_PATTERN.fullmatch(str(address or "").strip())
    if match is None:
        raise MergedHeaderRepairError(f"invalid A1 range: {address}")
    start_column = _column_number(match.group("start_column"))
    start_row = int(match.group("start_row"))
    end_column = _column_number(
        match.group("end_column") or match.group("start_column")
    )
    end_row = int(match.group("end_row") or match.group("start_row"))
    if end_column < start_column or end_row < start_row:
        raise MergedHeaderRepairError(f"reversed A1 range: {address}")
    return start_row, start_column, end_row, end_column


def merged_header_series_repair_target(
    validation_error: str,
) -> MergedHeaderRepairTarget | None:
    match = _ERROR_PATTERN.search(str(validation_error or ""))
    if match is None:
        return None
    return MergedHeaderRepairTarget(
        study_index=int(match.group("study")),
        series_index=int(match.group("series")),
    )


def apply_merged_header_series_repair(
    manifest: dict[str, Any],
    target: MergedHeaderRepairTarget,
) -> dict[str, Any]:
    """Split one wide RAW series into source-lossless single-column series."""

    repaired = copy.deepcopy(manifest)
    studies = repaired.get("studies")
    if (
        not isinstance(studies, list)
        or not 0 <= target.study_index < len(studies)
        or not isinstance(studies[target.study_index], dict)
    ):
        raise MergedHeaderRepairError("repair Study index is invalid")
    study = studies[target.study_index]
    series_values = study.get("measurementSeries")
    if (
        not isinstance(series_values, list)
        or not 0 <= target.series_index < len(series_values)
        or not isinstance(series_values[target.series_index], dict)
    ):
        raise MergedHeaderRepairError(
            "repair measurementSeries index is invalid"
        )
    source_series = series_values[target.series_index]
    source_key = str(source_series.get("key") or "").strip()
    if not source_key:
        raise MergedHeaderRepairError("repair series key is empty")
    if (
        str(source_series.get("seriesRole") or "RAW").upper() != "RAW"
        or source_series.get("aggregateOfSeries")
        or str(source_series.get("aggregationFunction") or "").strip()
    ):
        raise MergedHeaderRepairError(
            "only non-aggregate RAW series can be split safely"
        )

    header = _range_bounds(source_series.get("headerRange"))
    values = _range_bounds(source_series.get("valueRange"))
    if (
        header[0] != header[2]
        or header[1] != values[1]
        or header[3] != values[3]
        or header[3] <= header[1]
    ):
        raise MergedHeaderRepairError(
            "header/value ranges are not an aligned multi-column series"
        )

    existing_keys = {
        str(series.get("key") or "")
        for index, series in enumerate(series_values)
        if index != target.series_index and isinstance(series, dict)
    }
    split_series: list[dict[str, Any]] = []
    split_keys: list[str] = []
    for column in range(header[1], header[3] + 1):
        column_label = _column_label(column)
        split_key = f"{source_key}--column-{column_label.casefold()}"
        if split_key in existing_keys or split_key in split_keys:
            raise MergedHeaderRepairError(
                f"split series key collides: {split_key}"
            )
        split = copy.deepcopy(source_series)
        split["key"] = split_key
        split["headerRange"] = f"{column_label}{header[0]}"
        split["valueRange"] = (
            f"{column_label}{values[0]}"
            if values[0] == values[2]
            else f"{column_label}{values[0]}:{column_label}{values[2]}"
        )
        stratum_key = str(split.get("stratumKey") or "").strip()
        split["stratumKey"] = " | ".join(
            value
            for value in (
                stratum_key,
                f"source column {column_label}",
            )
            if value
        )
        split_series.append(split)
        split_keys.append(split_key)

    series_values[target.series_index : target.series_index + 1] = (
        split_series
    )
    for series in series_values:
        if not isinstance(series, dict):
            continue
        aggregate_sources = series.get("aggregateOfSeries")
        if not isinstance(aggregate_sources, list):
            continue
        series["aggregateOfSeries"] = [
            replacement
            for value in aggregate_sources
            for replacement in (
                split_keys if str(value) == source_key else [value]
            )
        ]
    # Once one series proves that this Study used a merged header as a
    # repeated axis, normalize every other aligned RAW multi-column series in
    # the same Study. The representation remains source-lossless and avoids
    # requiring one retry per row of the same table.
    for index, series in enumerate(series_values):
        if (
            not isinstance(series, dict)
            or str(series.get("seriesRole") or "RAW").upper() != "RAW"
            or series.get("aggregateOfSeries")
            or str(series.get("aggregationFunction") or "").strip()
        ):
            continue
        try:
            candidate_header = _range_bounds(series.get("headerRange"))
            candidate_values = _range_bounds(series.get("valueRange"))
        except MergedHeaderRepairError:
            continue
        if (
            candidate_header[0] == candidate_header[2]
            and candidate_header[1] == candidate_values[1]
            and candidate_header[3] == candidate_values[3]
            and candidate_header[3] > candidate_header[1]
        ):
            return apply_merged_header_series_repair(
                repaired,
                MergedHeaderRepairTarget(target.study_index, index),
            )
    return repaired


__all__ = [
    "MergedHeaderRepairError",
    "MergedHeaderRepairTarget",
    "apply_merged_header_series_repair",
    "merged_header_series_repair_target",
]
