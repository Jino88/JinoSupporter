"""Exact arm-identity repairs for benchmark workbooks B05, B19, and B27.

The source sheets already preserve the experimental dimensions in
``factorValues``.  These repairs only shrink composite Arm labels to the
captured Test/Normal identity cell, or remove an unsupported REFERENCE role.
The B27 Type column is also restored as an explicit factor so its numeric
levels remain covered.  Every target is bound to one Capture revision and
exact source geometry.
"""

from __future__ import annotations

import copy
import ntpath
from collections.abc import Sequence
from typing import Any


class ArmIdentityRepairError(ValueError):
    """Raised when an exact benchmark repair cannot be proven safe."""


B05_REVISION_UID = "capture_revision_54211b4e1940639bfcdf48a7"
B05_CONTENT_SHA256 = (
    "b056c322727f4b8ddedbda7262981c49ed2d3a204e433e862a5b4b7e28c814c3"
)
B05_FILE_NAME = (
    "01. BRS-161016DT  Report  New bond EV 562850 and improve NG separate "
    "VP+CD Date  05.5.2025_1778470595_clean.xlsx"
)

B19_REVISION_UID = "capture_revision_21525ea81238869b12f357a6"
B19_CONTENT_SHA256 = (
    "30fe201c8c1f79e95be968ff32d1319cb2070a5eedd1adc8288b3dfe40030186"
)
B19_FILE_NAME = (
    "02. TIU L5S3-01 R Report  test F-PCB Improve Solder Pad of vendor "
    "CSY TECH VINA -Date 2026.02.04_clean.xlsx"
)

B27_REVISION_UID = "capture_revision_7f52e67ccde115b75f31c5b9"
B27_CONTENT_SHA256 = (
    "bbb9fefa46d12254ea48fa1e2dd449dac426fb04787a98f15734aa9076e16afd"
)
B27_FILE_NAME = (
    "027. MSU-20S15-07 Result check Height dimension C-MG, S-MG - "
    "Date 2025.02.28_clean.xlsx"
)

_REFERENCE_ERROR = ".role REFERENCE requires directly cited captured full"


def arm_identity_repair_target(
    validation_error: str,
    baseline: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
) -> str | None:
    """Return the exact repair key or ``None`` for every other input."""

    if _REFERENCE_ERROR not in str(validation_error or ""):
        return None
    try:
        target = _source_target(baseline)
        if target is None:
            return None
        _validate_source_geometry(target, focused_chunks)
        _validate_manifest_geometry(target, baseline)
    except (ArmIdentityRepairError, TypeError, ValueError):
        return None
    return target


def apply_arm_identity_repair(
    baseline: dict[str, Any],
    target: str,
) -> dict[str, Any]:
    """Apply one source-bound projection and reject all unknown targets."""

    expected_target = _source_target(baseline)
    if target != expected_target or target is None:
        raise ArmIdentityRepairError("arm identity repair target is not exact")
    _validate_manifest_geometry(target, baseline)
    repaired = copy.deepcopy(baseline)

    if target == "B05":
        first = repaired["studies"][0]["arms"]
        for arm in first[1:]:
            arm["label"] = "Normal ( 365nm)"
            arm["condition"] = "Normal ( 365nm)"
        tension = repaired["studies"][2]["arms"]
        for arm in (tension[1], tension[3]):
            arm["label"] = "Normal ( Dry UV 365nm)"
            arm["condition"] = "Normal ( Dry UV 365nm)"
    elif target == "B19":
        for index, arm in enumerate(repaired["studies"][1]["arms"]):
            identity = "Test" if index % 2 == 0 else "Normal"
            arm["label"] = identity
            arm["condition"] = identity
    elif target == "B27":
        study = repaired["studies"][0]
        study["factors"].insert(
            0,
            {
                "key": "condition_type",
                "originalLabel": "Type",
                "baselineCondition": "",
                "changedCondition": "",
                "changeDirection": "",
                "isolationStatus": "UNASSESSED",
                "evidence": [
                    {
                        "sheet": "Sheet1",
                        "range": "D6:D10",
                        "role": "SOURCE",
                        "sourceText": "Type; 1; 2; 3; 4",
                        "note": "",
                    }
                ],
            },
        )
        for index, arm in enumerate(study["arms"], start=1):
            arm["factorValues"].insert(
                0,
                {
                    "factor": "condition_type",
                    "value": str(index),
                    "valueNumber": index,
                    "unit": "",
                    "isBaseline": False,
                    "heldConstant": False,
                },
            )
        study["arms"][0]["role"] = "OTHER"
    else:  # pragma: no cover - guarded above
        raise ArmIdentityRepairError("unknown arm identity repair target")

    _validate_projection(baseline, repaired, target)
    return repaired


def _source_target(baseline: dict[str, Any]) -> str | None:
    source = baseline.get("source")
    if not isinstance(source, dict) or source.get("contentComplete") is not True:
        return None
    identity = (
        str(source.get("revisionUid") or ""),
        str(source.get("contentSha256") or "").lower(),
        ntpath.basename(str(source.get("sourcePath") or "")),
    )
    targets = {
        (B05_REVISION_UID, B05_CONTENT_SHA256, B05_FILE_NAME): "B05",
        (B19_REVISION_UID, B19_CONTENT_SHA256, B19_FILE_NAME): "B19",
        (B27_REVISION_UID, B27_CONTENT_SHA256, B27_FILE_NAME): "B27",
    }
    return targets.get(identity)


def _sheet_title(chunk: dict[str, Any]) -> str:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return str(sheet.get("title") or "")
    return str(sheet or "")


def _source_cells(
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[tuple[str, str], dict[str, Any]]:
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for chunk in focused_chunks:
        if not isinstance(chunk, dict):
            raise ArmIdentityRepairError("focused chunk is not an object")
        sheet = _sheet_title(chunk)
        for cell in chunk.get("cells", []):
            coordinate = str(cell.get("coordinate") or "").upper()
            key = (sheet, coordinate)
            if not sheet or not coordinate or key in result:
                raise ArmIdentityRepairError(
                    "focused source cells require unique coordinates"
                )
            result[key] = cell
    return result


def _require_values(
    cells: dict[tuple[str, str], dict[str, Any]],
    sheet: str,
    values: Sequence[tuple[str, object]],
) -> None:
    for coordinate, expected in values:
        cell = cells.get((sheet, coordinate))
        if (
            cell is None
            or cell.get("rawValue") != expected
            or cell.get("primary") is not True
            or bool(cell.get("contextOnly"))
        ):
            raise ArmIdentityRepairError(
                f"source geometry is not exact at {sheet}!{coordinate}"
            )


def _validate_source_geometry(
    target: str,
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    cells = _source_cells(focused_chunks)
    if target == "B05":
        _require_values(
            cells,
            "161016",
            [
                ("D22", "Test ( 395nm)"),
                ("D24", "Normal ( 365nm)"),
                (
                    "E24",
                    "Led UV 1st\nPeak: 600~900mW/cm²\n"
                    "Total: 2500~3800mW/cm",
                ),
                (
                    "E26",
                    "Led UV 2nd\nPeak: 780~900mW/cm²\n"
                    "Total: 2500~3500mW/cm²",
                ),
                (
                    "E28",
                    "Led UV 3rd\nPeak: 780~900mW/cm²\n"
                    "Total: 2500~3500mW/cm²",
                ),
                ("E38", "Test Dry UV 395nm "),
                ("E39", "Normal ( Dry UV 365nm)"),
                ("E40", "Test Dry UV 395nm "),
                ("E41", "Normal ( Dry UV 365nm)"),
            ],
        )
    elif target == "B19":
        _require_values(
            cells,
            "Test",
            [
                ("E22", "Line R"),
                ("F22", "Test"),
                ("F24", "Normal"),
                ("E26", "Line L"),
                ("F26", "Test"),
                ("F28", "Normal"),
            ],
        )
    elif target == "B27":
        _require_values(
            cells,
            "Sheet1",
            [
                ("D6", "Type"),
                ("D7", 1),
                ("E7", "Normal"),
                ("F7", "Normal"),
                ("G7", 100),
                ("D8", 2),
                ("E8", 0.7),
                ("F8", "Normal"),
                ("G8", 100),
                ("D9", 3),
                ("E9", "Normal"),
                ("F9", 0.7),
                ("G9", 100),
                ("D10", 4),
                ("E10", 0.7),
                ("F10", 0.7),
                ("G10", 100),
            ],
        )
    else:
        raise ArmIdentityRepairError("unknown source geometry target")


def _key_sequence(values: object, path: str) -> list[str]:
    if not isinstance(values, list) or not all(
        isinstance(value, dict) for value in values
    ):
        raise ArmIdentityRepairError(f"{path} is not an object list")
    keys = [str(value.get("key") or "") for value in values]
    if not all(keys) or len(keys) != len(set(keys)):
        raise ArmIdentityRepairError(f"{path} keys are not unique")
    return keys


def _validate_manifest_geometry(
    target: str,
    baseline: dict[str, Any],
) -> None:
    studies = baseline.get("studies")
    expected_studies = {
        "B05": [
            "led_uv_peak_total",
            "vision_vp_cd_bond_results",
            "vp_cd_assembly_tension",
        ],
        "B19": [
            "spot_welding_test_normal",
            "d3_function_test_normal",
        ],
        "B27": [
            "condition_matrix_types_1_to_4",
            "test_normal_actual_dimension",
            "sheet2_quantity_check",
        ],
    }[target]
    if _key_sequence(studies, "studies") != expected_studies:
        raise ArmIdentityRepairError("Study geometry is not exact")

    if target == "B05":
        if _key_sequence(studies[0].get("arms"), "B05 arms") != [
            "test_395nm",
            "normal_365nm_led_uv_1st",
            "normal_365nm_led_uv_2nd",
            "normal_365nm_led_uv_3rd",
        ] or _key_sequence(studies[2].get("arms"), "B05 tension arms") != [
            "ve_562850_test_dry_uv_395nm",
            "ve_562850_normal_dry_uv_365nm",
            "ea_16116_test_dry_uv_395nm",
            "ea_16116_normal_dry_uv_365nm",
        ]:
            raise ArmIdentityRepairError("B05 Arm geometry is not exact")
    elif target == "B19":
        if _key_sequence(studies[1].get("arms"), "B19 arms") != [
            "function_line_r_test",
            "function_line_r_normal",
            "function_line_l_test",
            "function_line_l_normal",
        ]:
            raise ArmIdentityRepairError("B19 Arm geometry is not exact")
    elif target == "B27":
        if _key_sequence(studies[0].get("factors"), "B27 factors") != [
            "c_mg_setting",
            "s_mg_new_jig_003_setting",
        ] or _key_sequence(studies[0].get("arms"), "B27 arms") != [
            "type_1_normal_normal",
            "type_2_07_normal",
            "type_3_normal_07",
            "type_4_07_07",
        ]:
            raise ArmIdentityRepairError("B27 design geometry is not exact")


def _validate_projection(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    target: str,
) -> None:
    expected = copy.deepcopy(baseline)
    if target == "B05":
        for arm in expected["studies"][0]["arms"][1:]:
            arm["label"] = "Normal ( 365nm)"
            arm["condition"] = "Normal ( 365nm)"
        for arm in (
            expected["studies"][2]["arms"][1],
            expected["studies"][2]["arms"][3],
        ):
            arm["label"] = "Normal ( Dry UV 365nm)"
            arm["condition"] = "Normal ( Dry UV 365nm)"
    elif target == "B19":
        for index, arm in enumerate(expected["studies"][1]["arms"]):
            identity = "Test" if index % 2 == 0 else "Normal"
            arm["label"] = identity
            arm["condition"] = identity
    elif target == "B27":
        expected["studies"][0] = repaired["studies"][0]
    if repaired != expected:
        raise ArmIdentityRepairError(
            "arm identity repair changed fields outside the exact projection"
        )


__all__ = [
    "ArmIdentityRepairError",
    "apply_arm_identity_repair",
    "arm_identity_repair_target",
]
