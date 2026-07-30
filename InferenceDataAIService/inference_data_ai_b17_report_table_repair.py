"""Deterministic scalar projection for the omitted B17 Report lot table."""

from __future__ import annotations

import copy
import math
import ntpath
from collections.abc import Sequence
from typing import Any


class B17ReportTableRepairError(ValueError):
    """Raised when the exact B17 projection cannot be proven safe."""


B17_REVISION_UID = "capture_revision_f2834f598089a492eadf74c0"
B17_CONTENT_SHA256 = (
    "12db74002323431117c54cbc638eaa3219d25d852e2ab6bac2f5d85500fe8ca4"
)
B17_FILE_NAME = (
    "02. L20S15-07DT Report test new lot CD (3-17)  "
    "( Size 510)date   2025.04.14_clean.xlsx"
)

_REPORT_SHEET = "Report"
_TOP_ROWS = list(range(18, 108, 2))
_STUDY_KEYS = [
    "report-new-lot-cd-checks",
    "vp-cd-assembly-tension-tests",
]
_CONCLUSION_KEYS = [
    "dyne-pen-all-lots-ok",
    "tension-lots-cannot-use",
    "separation-lots-cannot-use",
    "listed-lots-require-second-sample-function-test",
]


def b17_report_table_repair_applicable(
    baseline: dict[str, Any],
    *,
    validation_error: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> bool:
    """Return true only for the exact B17 source-coverage omission."""

    if (
        "Source content coverage is incomplete"
        not in str(validation_error or "")
        or "quantitative cell" not in str(validation_error or "")
    ):
        return False
    try:
        _validate_source_identity(baseline)
        _validate_manifest_geometry(baseline)
        _validate_source_geometry(focused_chunks)
    except (B17ReportTableRepairError, TypeError, ValueError):
        return False
    return True


def apply_b17_report_table_repair(
    baseline: dict[str, Any],
    *,
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Add 45 lot Arms, 12 scalar Outcomes, and two heading citations."""

    _validate_source_identity(baseline)
    _validate_manifest_geometry(baseline)
    cells = _validate_source_geometry(focused_chunks)
    repaired = copy.deepcopy(baseline)

    report_study = repaired["studies"][0]
    tension_conclusion = report_study["conclusions"][1]
    tension_conclusion["evidence"] = [
        _evidence(
            "C109:C110",
            (
                f"{_text(cells['C109'])}; "
                f"{_text(cells['C110'])}"
            ),
        )
    ]
    separation_conclusion = report_study["conclusions"][2]
    separation_conclusion["evidence"] = [
        _evidence(
            "C111:C112",
            (
                f"{_text(cells['C111'])}; "
                f"{_text(cells['C112'])}"
            ),
        )
    ]

    repaired["studies"].insert(
        1,
        _build_report_lot_study(cells),
    )
    return repaired


def _validate_source_identity(baseline: dict[str, Any]) -> None:
    source = baseline.get("source")
    if not isinstance(source, dict) or source.get("contentComplete") is not True:
        raise B17ReportTableRepairError(
            "B17 repair requires a content-complete source"
        )
    identity = (
        str(source.get("revisionUid") or ""),
        str(source.get("contentSha256") or "").lower(),
        ntpath.basename(str(source.get("sourcePath") or "")),
    )
    if identity != (
        B17_REVISION_UID,
        B17_CONTENT_SHA256,
        B17_FILE_NAME,
    ):
        raise B17ReportTableRepairError("B17 source identity is not exact")


def _key_sequence(values: object, path: str) -> list[str]:
    if not isinstance(values, list) or not all(
        isinstance(value, dict) for value in values
    ):
        raise B17ReportTableRepairError(f"{path} is not an object list")
    keys = [str(value.get("key") or "") for value in values]
    if not all(keys) or len(keys) != len(set(keys)):
        raise B17ReportTableRepairError(f"{path} keys are not unique")
    return keys


def _validate_manifest_geometry(baseline: dict[str, Any]) -> None:
    studies = baseline.get("studies")
    if _key_sequence(studies, "studies") != _STUDY_KEYS:
        raise B17ReportTableRepairError("B17 Study geometry is not exact")
    if _key_sequence(
        studies[0].get("conclusions"),
        "B17 conclusions",
    ) != _CONCLUSION_KEYS:
        raise B17ReportTableRepairError(
            "B17 conclusion geometry is not exact"
        )
    if studies[0]["conclusions"][1].get("evidence") != [
        {
            "sheet": "Report",
            "range": "C110",
            "role": "SOURCE",
            "sourceText": (
                " + Lot 18,21,27,30,31,34,41  happen tension "
                "=> Can not use "
            ),
            "note": "",
        }
    ]:
        raise B17ReportTableRepairError(
            "B17 tension conclusion baseline is not exact"
        )
    if studies[0]["conclusions"][2].get("evidence") != [
        {
            "sheet": "Report",
            "range": "C112",
            "role": "SOURCE",
            "sourceText": (
                " + Lot 2,4,7,8,13,17,18,21,24,25,26,27,30,31,"
                "33,34,35, 37,38,40,41,42,44   happen separate "
                "VP/CD  => Can not use "
            ),
            "note": "",
        }
    ]:
        raise B17ReportTableRepairError(
            "B17 separation conclusion baseline is not exact"
        )


def _sheet_title(chunk: dict[str, Any]) -> str:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return str(sheet.get("title") or "")
    return str(sheet or "")


def _source_cells(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for chunk in focused_chunks:
        if not isinstance(chunk, dict):
            raise B17ReportTableRepairError("focused chunk is not an object")
        if _sheet_title(chunk) != _REPORT_SHEET:
            continue
        for cell in chunk.get("cells", []):
            coordinate = str(cell.get("coordinate") or "").upper()
            if not coordinate or coordinate in result:
                raise B17ReportTableRepairError(
                    "Report cells require unique coordinates"
                )
            result[coordinate] = cell
    if not result:
        raise B17ReportTableRepairError("Report source cells are missing")
    return result


def _require(
    cells: dict[str, dict[str, Any]],
    coordinate: str,
    value: object,
    *,
    merge_range: str | None = None,
) -> dict[str, Any]:
    cell = cells.get(coordinate)
    if (
        cell is None
        or cell.get("rawValue") != value
        or cell.get("primary") is not True
        or bool(cell.get("contextOnly"))
        or str(cell.get("mergeRange") or "").upper()
        != str(merge_range or "").upper()
    ):
        raise B17ReportTableRepairError(
            f"B17 source geometry is not exact at Report!{coordinate}"
        )
    return cell


def _number(cell: dict[str, Any]) -> float | int:
    value = cell.get("rawValue")
    if (
        isinstance(value, bool)
        or not isinstance(value, (int, float))
        or not math.isfinite(float(value))
    ):
        raise B17ReportTableRepairError(
            f"Report!{cell.get('coordinate')} is not numeric"
        )
    return value


def _validate_source_geometry(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    cells = _source_cells(focused_chunks)
    headers = {
        "C15": ("Date test", "C15:C17"),
        "D15": ("LOT TEST", "D15:D17"),
        "E15": ("LOT", "E15:E17"),
        "F16": ("Dyne", "F16:F17"),
        "G16": ("Input", "G16:G17"),
        "H16": ("NG / NG rate", "H16:H17"),
        "I16": ("Input", "I16:I17"),
        "J16": ("Tension TEST", "J16:J17"),
        "K16": ("Input", "K16:K17"),
        "L16": ("OK", "L16:L17"),
        "M16": ("VP/CD separate", "M16:M17"),
        "N16": ("Total NG", "N16:N17"),
        "O16": ("NG rate", "O16:O17"),
    }
    for coordinate, (value, merge_range) in headers.items():
        _require(
            cells,
            coordinate,
            value,
            merge_range=merge_range,
        )
    _require(cells, "F18", 40, merge_range="F18:F107")

    lot_values: list[int] = []
    for top_row in _TOP_ROWS:
        lower_row = top_row + 1
        lot_cell = cells.get(f"D{top_row}")
        if (
            lot_cell is None
            or lot_cell.get("primary") is not True
            or str(lot_cell.get("mergeRange") or "").upper()
            != f"D{top_row}:D{lower_row}"
        ):
            raise B17ReportTableRepairError(
                "B17 LOT TEST geometry is not exact at "
                f"Report!D{top_row}"
            )
        lot_number = _number(lot_cell)
        if not float(lot_number).is_integer():
            raise B17ReportTableRepairError(
                f"Report!D{top_row} LOT TEST is not an integer"
            )
        lot_values.append(int(lot_number))
        for column in ("G", "I", "K", "L", "N", "O"):
            cell = cells.get(f"{column}{top_row}")
            if (
                cell is None
                or cell.get("primary") is not True
                or str(cell.get("mergeRange") or "").upper()
                != f"{column}{top_row}:{column}{lower_row}"
            ):
                raise B17ReportTableRepairError(
                    "B17 merged result geometry is not exact at "
                    f"Report!{column}{top_row}"
                )
            _number(cell)
        for column in ("H", "J", "M"):
            for row in (top_row, lower_row):
                cell = cells.get(f"{column}{row}")
                if (
                    cell is None
                    or str(cell.get("mergeRange") or "")
                    or cell.get("primary") is not True
                ):
                    raise B17ReportTableRepairError(
                        "B17 count/rate geometry is not exact at "
                        f"Report!{column}{row}"
                    )
                _number(cell)
    if sorted(lot_values) != list(range(1, 46)):
        raise B17ReportTableRepairError(
            "B17 LOT TEST identities are not exactly 1 through 45"
        )

    _require(
        cells,
        "C109",
        " => Result check tension  :",
        merge_range="C109:O109",
    )
    _require(
        cells,
        "C110",
        (
            " + Lot 18,21,27,30,31,34,41  happen tension "
            "=> Can not use "
        ),
        merge_range="C110:O110",
    )
    _require(
        cells,
        "C111",
        " => Result check separate VP/CD :",
        merge_range="C111:O111",
    )
    _require(
        cells,
        "C112",
        (
            " + Lot 2,4,7,8,13,17,18,21,24,25,26,27,30,31,"
            "33,34,35, 37,38,40,41,42,44   happen separate "
            "VP/CD  => Can not use "
        ),
        merge_range="C112:O112",
    )
    return cells


def _text(cell: dict[str, Any]) -> str:
    value = cell.get("rawValue")
    return str(value)


def _evidence(address: str, source_text: str) -> dict[str, Any]:
    return {
        "sheet": _REPORT_SHEET,
        "range": address,
        "role": "SOURCE",
        "sourceText": source_text,
        "note": "",
    }


def _observation(
    *,
    key: str,
    arm: str,
    value_number: float | int,
    evidence: list[dict[str, Any]],
    numerator: float | int | None = None,
    denominator: float | int | None = None,
) -> dict[str, Any]:
    return {
        "key": key,
        "arm": arm,
        "valueNumber": value_number,
        "valueText": "",
        "numerator": numerator,
        "denominator": denominator,
        "ratePpm": None,
        "min": None,
        "max": None,
        "average": None,
        "sampleSize": None,
        "evidence": evidence,
    }


def _scalar_evidence(
    cells: dict[str, dict[str, Any]],
    coordinate: str,
) -> dict[str, Any]:
    return _evidence(coordinate, _text(cells[coordinate]))


def _rate_observation(
    *,
    cells: dict[str, dict[str, Any]],
    key: str,
    arm: str,
    rate_coordinate: str,
    numerator_coordinate: str,
    denominator_coordinate: str,
) -> dict[str, Any]:
    raw_rate = _number(cells[rate_coordinate])
    numerator = _number(cells[numerator_coordinate])
    denominator = _number(cells[denominator_coordinate])
    return _observation(
        key=key,
        arm=arm,
        value_number=float(raw_rate) * 100.0,
        numerator=numerator,
        denominator=denominator,
        evidence=[
            _scalar_evidence(cells, rate_coordinate),
            _scalar_evidence(cells, numerator_coordinate),
            _scalar_evidence(cells, denominator_coordinate),
        ],
    )


def _outcome(
    *,
    key: str,
    label: str,
    metric_type: str,
    unit: str,
    header_coordinate: str,
    cells: dict[str, dict[str, Any]],
    observations: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "key": key,
        "originalLabel": label,
        "metricType": metric_type,
        "unit": unit,
        "favorableDirection": "UNKNOWN",
        "evidence": [
            _scalar_evidence(cells, header_coordinate),
        ],
        "observations": observations,
    }


def _build_report_lot_study(
    cells: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    arms: list[dict[str, Any]] = []
    observations: dict[str, list[dict[str, Any]]] = {
        key: []
        for key in (
            "material-input",
            "material-ng-count",
            "material-ng-rate",
            "tension-input",
            "tension-ng-count",
            "tension-ng-rate",
            "separation-input",
            "separation-ok-count",
            "separation-ng-count",
            "separation-ng-rate",
            "total-ng-count",
            "total-ng-rate",
        )
    }
    scalar_specs = [
        ("material-input", "G"),
        ("material-ng-count", "H"),
        ("tension-input", "I"),
        ("tension-ng-count", "J"),
        ("separation-input", "K"),
        ("separation-ok-count", "L"),
        ("separation-ng-count", "M"),
        ("total-ng-count", "N"),
    ]
    for top_row in _TOP_ROWS:
        lower_row = top_row + 1
        lot = int(_number(cells[f"D{top_row}"]))
        arm_key = f"lot-test-{lot}"
        lot_coordinate = f"D{top_row}"
        arms.append(
            {
                "key": arm_key,
                "role": "OTHER",
                "label": f"LOT TEST {lot}",
                "condition": f"LOT TEST {lot}; Dyne 40",
                "sampleSize": None,
                "sampleBasis": "",
                "matchingBasis": "",
                "factorValues": [
                    {
                        "factor": "lot-test",
                        "value": str(lot),
                        "valueNumber": lot,
                        "unit": "",
                        "isBaseline": False,
                        "heldConstant": False,
                    },
                    {
                        "factor": "dyne",
                        "value": "40",
                        "valueNumber": 40,
                        "unit": "",
                        "isBaseline": False,
                        "heldConstant": True,
                    },
                ],
                "evidence": [
                    _scalar_evidence(cells, lot_coordinate),
                ],
            }
        )
        for outcome_key, column in scalar_specs:
            coordinate = f"{column}{top_row}"
            observations[outcome_key].append(
                _observation(
                    key=f"{outcome_key}-lot-{lot}",
                    arm=arm_key,
                    value_number=_number(cells[coordinate]),
                    evidence=[_scalar_evidence(cells, coordinate)],
                )
            )
        observations["material-ng-rate"].append(
            _rate_observation(
                cells=cells,
                key=f"material-ng-rate-lot-{lot}",
                arm=arm_key,
                rate_coordinate=f"H{lower_row}",
                numerator_coordinate=f"H{top_row}",
                denominator_coordinate=f"G{top_row}",
            )
        )
        observations["tension-ng-rate"].append(
            _rate_observation(
                cells=cells,
                key=f"tension-ng-rate-lot-{lot}",
                arm=arm_key,
                rate_coordinate=f"J{lower_row}",
                numerator_coordinate=f"J{top_row}",
                denominator_coordinate=f"I{top_row}",
            )
        )
        observations["separation-ng-rate"].append(
            _rate_observation(
                cells=cells,
                key=f"separation-ng-rate-lot-{lot}",
                arm=arm_key,
                rate_coordinate=f"M{lower_row}",
                numerator_coordinate=f"M{top_row}",
                denominator_coordinate=f"K{top_row}",
            )
        )
        observations["total-ng-rate"].append(
            _rate_observation(
                cells=cells,
                key=f"total-ng-rate-lot-{lot}",
                arm=arm_key,
                rate_coordinate=f"O{top_row}",
                numerator_coordinate=f"N{top_row}",
                denominator_coordinate=f"K{top_row}",
            )
        )

    outcome_specs = [
        (
            "material-input",
            "Input",
            "material_input_count",
            "",
            "G16",
        ),
        (
            "material-ng-count",
            "NG / NG rate",
            "material_ng_count",
            "",
            "H16",
        ),
        (
            "material-ng-rate",
            "NG / NG rate",
            "material_ng_rate",
            "%",
            "H16",
        ),
        (
            "tension-input",
            "Input",
            "tension_input_count",
            "",
            "I16",
        ),
        (
            "tension-ng-count",
            "Tension TEST",
            "tension_ng_count",
            "",
            "J16",
        ),
        (
            "tension-ng-rate",
            "Tension TEST",
            "tension_ng_rate",
            "%",
            "J16",
        ),
        (
            "separation-input",
            "Input",
            "separation_input_count",
            "",
            "K16",
        ),
        (
            "separation-ok-count",
            "OK",
            "separation_ok_count",
            "",
            "L16",
        ),
        (
            "separation-ng-count",
            "VP/CD separate",
            "separation_ng_count",
            "",
            "M16",
        ),
        (
            "separation-ng-rate",
            "VP/CD separate",
            "separation_ng_rate",
            "%",
            "M16",
        ),
        (
            "total-ng-count",
            "Total NG",
            "total_ng_count",
            "",
            "N16",
        ),
        (
            "total-ng-rate",
            "NG rate",
            "total_ng_rate",
            "%",
            "O16",
        ),
    ]
    outcomes = [
        _outcome(
            key=key,
            label=label,
            metric_type=metric_type,
            unit=unit,
            header_coordinate=header_coordinate,
            cells=cells,
            observations=observations[key],
        )
        for (
            key,
            label,
            metric_type,
            unit,
            header_coordinate,
        ) in outcome_specs
    ]
    return {
        "key": "report-lot-test-results",
        "title": "Report lot-level material, tension, and VP/CD results",
        "purpose": (
            "Preserve the source lot table as separate scalar Input, count, "
            "rate, OK, and total result fields."
        ),
        "hypothesis": "",
        "objective": "",
        "designType": "Descriptive lot-level result table",
        "comparisonBasis": "",
        "verificationStatus": "NEEDS_REVIEW",
        "comparabilityStatus": "UNASSESSED",
        "confoundingStatus": "UNASSESSED",
        "summary": (
            "Forty-five LOT TEST records are projected without inferring a "
            "relationship to the separate #1–#45 sample matrix."
        ),
        "limitations": [
            "LOT TEST identifiers are record identities, not treatment Arms.",
            (
                "Count and percentage rows remain distinct Outcomes; no "
                "cross-lot comparison is asserted."
            ),
        ],
        "evidence": [
            _evidence("C15:O107", "Lot-level result table"),
        ],
        "contexts": [],
        "factors": [
            {
                "key": "lot-test",
                "originalLabel": "LOT TEST",
                "baselineCondition": "",
                "changedCondition": "",
                "changeDirection": "",
                "isolationStatus": "UNASSESSED",
                "evidence": [
                    _scalar_evidence(cells, "D15"),
                    _evidence(
                        "D18:D107",
                        "LOT TEST identifiers 1 through 45",
                    ),
                ],
            },
            {
                "key": "dyne",
                "originalLabel": "Dyne",
                "baselineCondition": "",
                "changedCondition": "",
                "changeDirection": "",
                "isolationStatus": "UNASSESSED",
                "evidence": [
                    _scalar_evidence(cells, "F16"),
                    _scalar_evidence(cells, "F18"),
                ],
            },
        ],
        "arms": arms,
        "outcomes": outcomes,
        "measurementSeries": [],
        "comparisons": [],
        "conclusions": [],
    }


__all__ = [
    "B17ReportTableRepairError",
    "apply_b17_report_table_repair",
    "b17_report_table_repair_applicable",
]
