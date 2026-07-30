"""Exact deterministic repairs for the two P14 B04/B08 draft failures.

The repairs in this module are intentionally source-specific and fail closed.
They require the proven Capture v2 revision, content SHA-256, source geometry,
and canonical manifest geometry.  No AI, workflow, database, or artifact write
is performed here.
"""

from __future__ import annotations

import copy
import hashlib
import json
import ntpath
from collections.abc import Sequence
from typing import Any


B04_B08_REPAIR_SCHEMA_VERSION = "b04-b08-exact-source-repair-v1"

B04_REVISION_UID = "capture_revision_a8c47a3e66c13d3c5a08619e"
B04_CONTENT_SHA256 = (
    "42396c93623366bcc300075e2f013b996c0c7a36364421850376e37679a0b3f8"
)
B04_FILE_NAME = (
    "008.MSU-20S15-07 Test MC check Height dimension 1 "
    "2025.01.25_1778470549_clean.xlsx"
)
B04_NUMERIC_VALIDATION_ERROR = (
    "studies[2].outcomes[1].observations[0].valueNumber=50 is not "
    "present in its cited Capture v2 cells"
)

B08_REVISION_UID = "capture_revision_9d052673885c1b42a080377e"
B08_CONTENT_SHA256 = (
    "45fd14a8e6b12065d291977d5b1f3efea3694e7a1facb0983896b23021411c79"
)
B08_FILE_NAME = (
    "01.MSU-20S15-07 Result test  frame new mold_1778470978_clean.xlsx"
)
B08_CONCLUSION_VALIDATION_ERROR = (
    "studies[0].conclusions[1] SOURCE_CONCLUSION requires directly cited "
    "captured narrative decision/conclusion text matching "
    "evidence.sourceText and supporting the claim; classify an AI synthesis "
    "as AI_DERIVED_DESCRIPTIVE or omit it"
)

_SHEET = "Sheet1"
_B04_REPAIR_KIND = "B04_PROSE_SAMPLE_SIZE_AND_CONTEXT"
_B08_REPAIR_KIND = "B08_ADJACENT_SOURCE_CONCLUSION"

_B04_SAMPLE_TEXT = "Check height of sample  ( 50pcs ) "
_B04_CONCLUSION_TEXT = (
    "6 Posistion of jig not same when setting sensor zero "
)
_B04_CONTEXT_CELLS = (
    ("B3", "CHECK SENSOR ZERO SET ON THE JIG ( 6 POSITION )"),
    ("B4", "Date"),
    ("C4", "Posistion on the jig "),
    ("I4", "Note"),
)
_B04_CONTEXT_NOTE = (
    "Queryable source context for the cited conclusion; not a separate "
    "conclusion."
)

_B08_LINE_1 = (
    "- Lot test frame new mold check SPK OK -> Continue move modul  test "
)
_B08_LINE_2 = " => Can use"
_B08_JOINED_SOURCE_TEXT = (
    "- Lot test frame new mold check SPK OK -> Continue move modul  test; "
    "=> Can use"
)


class B04B08RepairError(RuntimeError):
    """Raised when an input is outside the two exact safe projections."""


def _canonical_sha256(value: object) -> str:
    payload = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _error_matches(actual: str, expected: str) -> bool:
    normalized = str(actual or "").strip()
    return normalized in {expected, f"ValueError: {expected}"}


def _source_identity(
    baseline: dict[str, Any],
) -> tuple[str, str, str]:
    source = baseline.get("source")
    if not isinstance(source, dict) or source.get("contentComplete") is not True:
        raise B04B08RepairError(
            "repair requires a content-complete canonical source"
        )
    return (
        str(source.get("revisionUid") or ""),
        str(source.get("contentSha256") or "").lower(),
        ntpath.basename(str(source.get("sourcePath") or "")),
    )


def _sheet_title(chunk: dict[str, Any]) -> str:
    sheet = chunk.get("sheet")
    if isinstance(sheet, dict):
        return str(sheet.get("title") or "")
    return str(sheet or chunk.get("sheetTitle") or "")


def _source_cells(
    focused_chunks: Sequence[dict[str, Any]],
    *,
    revision_uid: str,
    content_sha256: str,
) -> dict[str, dict[str, Any]]:
    cells: dict[str, dict[str, Any]] = {}
    target_chunks = 0
    for chunk in focused_chunks:
        if not isinstance(chunk, dict):
            raise B04B08RepairError("focused chunk is not an object")
        if _sheet_title(chunk) != _SHEET:
            continue
        target_chunks += 1
        revision = chunk.get("sourceRevision")
        if not isinstance(revision, dict) or (
            str(revision.get("revisionUid") or "") != revision_uid
            or str(revision.get("contentSha256") or "").lower()
            != content_sha256
        ):
            raise B04B08RepairError(
                "focused chunk source identity does not match the manifest"
            )
        for cell in chunk.get("cells", []):
            if not isinstance(cell, dict):
                raise B04B08RepairError("source cell is not an object")
            coordinate = str(cell.get("coordinate") or "").strip().upper()
            if not coordinate:
                raise B04B08RepairError("source cell has no coordinate")
            if coordinate in cells:
                raise B04B08RepairError(
                    f"duplicate source coordinate {_SHEET}!{coordinate}"
                )
            cells[coordinate] = cell
    if not target_chunks:
        raise B04B08RepairError("focused chunks lack the target sheet")
    return cells


def _cell(
    cells: dict[str, dict[str, Any]],
    coordinate: str,
) -> dict[str, Any]:
    result = cells.get(coordinate.upper())
    if result is None:
        raise B04B08RepairError(
            f"required source cell {_SHEET}!{coordinate} is missing"
        )
    return result


def _require_cell(
    cells: dict[str, dict[str, Any]],
    coordinate: str,
    *,
    revision_uid: str,
    raw_value: object,
    data_type: str,
    merge_range: str | None = None,
) -> dict[str, Any]:
    cell = _cell(cells, coordinate)
    expected_source_key = f"{revision_uid}:1:{coordinate.upper()}"
    expected_merge_role = "anchor" if merge_range else "none"
    if (
        cell.get("rawValue") != raw_value
        or str(cell.get("dataType") or "") != data_type
        or str(cell.get("mergeRange") or "").upper()
        != str(merge_range or "").upper()
        or str(cell.get("mergeRole") or "none").casefold()
        != expected_merge_role
        or cell.get("primary") is not True
        or bool(cell.get("contextOnly"))
        or str(cell.get("sourceCellKey") or "") != expected_source_key
    ):
        raise B04B08RepairError(
            f"source geometry is not exact at {_SHEET}!{coordinate}"
        )
    return cell


def _evidence(
    coordinate: str,
    source_text: str,
    *,
    note: str = "",
) -> dict[str, Any]:
    return {
        "sheet": _SHEET,
        "range": coordinate,
        "role": "SOURCE",
        "sourceText": source_text,
        "note": note,
    }


def _unique_keyed(
    values: object,
    path: str,
) -> dict[str, dict[str, Any]]:
    if not isinstance(values, list):
        raise B04B08RepairError(f"{path} is not a list")
    result: dict[str, dict[str, Any]] = {}
    for index, value in enumerate(values):
        if not isinstance(value, dict):
            raise B04B08RepairError(f"{path}[{index}] is not an object")
        key = str(value.get("key") or "")
        if not key or key in result:
            raise B04B08RepairError(
                f"{path} has a missing or duplicate key"
            )
        result[key] = value
    return result


def _b04_context_evidence() -> list[dict[str, Any]]:
    return [
        _evidence(
            coordinate,
            source_text,
            note=_B04_CONTEXT_NOTE,
        )
        for coordinate, source_text in _B04_CONTEXT_CELLS
    ]


def _validate_b04_source_geometry(
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    cells = _source_cells(
        focused_chunks,
        revision_uid=B04_REVISION_UID,
        content_sha256=B04_CONTENT_SHA256,
    )
    _require_cell(
        cells,
        "B3",
        revision_uid=B04_REVISION_UID,
        raw_value=_B04_CONTEXT_CELLS[0][1],
        data_type="s",
    )
    _require_cell(
        cells,
        "B4",
        revision_uid=B04_REVISION_UID,
        raw_value=_B04_CONTEXT_CELLS[1][1],
        data_type="s",
        merge_range="B4:B5",
    )
    _require_cell(
        cells,
        "C4",
        revision_uid=B04_REVISION_UID,
        raw_value=_B04_CONTEXT_CELLS[2][1],
        data_type="s",
        merge_range="C4:H4",
    )
    _require_cell(
        cells,
        "I4",
        revision_uid=B04_REVISION_UID,
        raw_value=_B04_CONTEXT_CELLS[3][1],
        data_type="s",
        merge_range="I4:K5",
    )
    _require_cell(
        cells,
        "I6",
        revision_uid=B04_REVISION_UID,
        raw_value=_B04_CONCLUSION_TEXT,
        data_type="s",
        merge_range="I6:K6",
    )
    _require_cell(
        cells,
        "D23",
        revision_uid=B04_REVISION_UID,
        raw_value=_B04_SAMPLE_TEXT,
        data_type="s",
        merge_range="D23:M23",
    )

    for index, column in enumerate("DEFGHIJKLM", start=1):
        _require_cell(
            cells,
            f"{column}24",
            revision_uid=B04_REVISION_UID,
            raw_value=index,
            data_type="n",
        )
    for row_identity, row in enumerate(range(25, 30), start=1):
        _require_cell(
            cells,
            f"C{row}",
            revision_uid=B04_REVISION_UID,
            raw_value=row_identity,
            data_type="n",
        )
        for column in "DEFGHIJKLM":
            cell = _cell(cells, f"{column}{row}")
            value = cell.get("rawValue")
            if (
                isinstance(value, bool)
                or not isinstance(value, (int, float))
                or str(cell.get("dataType") or "") != "n"
                or str(cell.get("mergeRange") or "")
                or str(cell.get("mergeRole") or "none").casefold()
                != "none"
                or cell.get("primary") is not True
                or str(cell.get("sourceCellKey") or "")
                != (
                    f"{B04_REVISION_UID}:1:"
                    f"{column}{row}"
                )
            ):
                raise B04B08RepairError(
                    "B04 5-by-10 measurement geometry is not exact"
                )


def _validate_b08_source_geometry(
    focused_chunks: Sequence[dict[str, Any]],
) -> None:
    cells = _source_cells(
        focused_chunks,
        revision_uid=B08_REVISION_UID,
        content_sha256=B08_CONTENT_SHA256,
    )
    _require_cell(
        cells,
        "B25",
        revision_uid=B08_REVISION_UID,
        raw_value=_B08_LINE_1,
        data_type="s",
        merge_range="B25:T25",
    )
    _require_cell(
        cells,
        "B26",
        revision_uid=B08_REVISION_UID,
        raw_value=_B08_LINE_2,
        data_type="s",
    )


def _b04_state(
    baseline: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any], bool]:
    studies = baseline.get("studies")
    if (
        not isinstance(studies, list)
        or len(studies) != 3
        or any(not isinstance(study, dict) for study in studies)
    ):
        raise B04B08RepairError("B04 Study geometry is not exact")
    if [study.get("key") for study in studies] != [
        "sensor-zero-set-by-jig-position",
        "sample-height-ng-summary",
        "sample-height-50pcs-measurements",
    ]:
        raise B04B08RepairError("B04 Study identities are not exact")

    conclusions = _unique_keyed(
        studies[0].get("conclusions"),
        "studies[0].conclusions",
    )
    conclusion = conclusions.get("six-jig-positions-not-same")
    original_evidence = [
        _evidence("I6:K6", _B04_CONCLUSION_TEXT)
    ]
    repaired_evidence = original_evidence + _b04_context_evidence()
    if (
        conclusion is None
        or conclusion.get("text") != _B04_CONCLUSION_TEXT
        or conclusion.get("claimType") != "SOURCE_CONCLUSION"
        or conclusion.get("causalStrength") != "DESCRIPTIVE"
    ):
        raise B04B08RepairError(
            "B04 source conclusion identity is not exact"
        )

    outcomes = _unique_keyed(
        studies[2].get("outcomes"),
        "studies[2].outcomes",
    )
    outcome = outcomes.get("height-measurement-sample-size")
    if (
        outcome is None
        or outcome.get("originalLabel") != _B04_SAMPLE_TEXT
        or outcome.get("metricType") != "sample_size"
        or outcome.get("unit") != "pcs"
        or outcome.get("evidence")
        != [_evidence("D23:M23", _B04_SAMPLE_TEXT)]
    ):
        raise B04B08RepairError(
            "B04 sample-size Outcome geometry is not exact"
        )
    observations = _unique_keyed(
        outcome.get("observations"),
        "B04 sample-size observations",
    )
    observation = observations.get("height-measurement-sample-size-50")
    if (
        observation is None
        or observation.get("arm") != "height-measurement-50pcs"
        or observation.get("valueText") != "50pcs"
        or observation.get("sampleSize") != 50
        or observation.get("evidence")
        != [_evidence("D23:M23", _B04_SAMPLE_TEXT)]
    ):
        raise B04B08RepairError(
            "B04 sample-size Observation geometry is not exact"
        )

    arms = _unique_keyed(studies[2].get("arms"), "studies[2].arms")
    arm = arms.get("height-measurement-50pcs")
    series = _unique_keyed(
        studies[2].get("measurementSeries"),
        "studies[2].measurementSeries",
    ).get("sample-height-raw-matrix")
    if (
        arm is None
        or arm.get("sampleSize") != 50
        or arm.get("evidence")
        != [_evidence("D23:M23", _B04_SAMPLE_TEXT)]
        or series is None
        or series.get("seriesRole") != "RAW"
        or series.get("outcome") != "sample-height"
        or series.get("arm") != "height-measurement-50pcs"
        or series.get("sheet") != _SHEET
        or series.get("headerRange") != "D24:M24"
        or series.get("valueRange") != "D25:M29"
        or series.get("rowIdentityRange") != "C25:C29"
        or series.get("axisSource") != "HEADER"
    ):
        raise B04B08RepairError(
            "B04 Arm or 5-by-10 RAW series geometry is not exact"
        )

    value_number = observation.get("valueNumber")
    evidence_state = conclusion.get("evidence")
    is_unrepaired = (
        value_number == 50 and evidence_state == original_evidence
    )
    is_repaired = (
        value_number is None and evidence_state == repaired_evidence
    )
    if not (is_unrepaired or is_repaired):
        raise B04B08RepairError(
            "B04 repair is partial, conflicting, or already altered"
        )
    return observation, conclusion, is_repaired


def _b08_state(
    baseline: dict[str, Any],
) -> tuple[dict[str, Any], bool]:
    studies = baseline.get("studies")
    if (
        not isinstance(studies, list)
        or not studies
        or not isinstance(studies[0], dict)
    ):
        raise B04B08RepairError("B08 Study geometry is not exact")
    conclusions = _unique_keyed(
        studies[0].get("conclusions"),
        "studies[0].conclusions",
    )
    if list(conclusions) != [
        "new-mold-lot-spk-disposition",
        "new-mold-can-use-disposition",
    ]:
        raise B04B08RepairError("B08 conclusion identities are not exact")
    context = conclusions["new-mold-lot-spk-disposition"]
    target = conclusions["new-mold-can-use-disposition"]
    if (
        context.get("text") != _B08_LINE_1
        or context.get("claimType") != "SOURCE_CONCLUSION"
        or context.get("causalStrength") != "DESCRIPTIVE"
        or context.get("evidence")
        != [_evidence("B25", _B08_LINE_1)]
        or target.get("text") != _B08_LINE_2
        or target.get("claimType") != "SOURCE_CONCLUSION"
        or target.get("causalStrength") != "DESCRIPTIVE"
    ):
        raise B04B08RepairError(
            "B08 source conclusion geometry is not exact"
        )
    original = [_evidence("B26", _B08_LINE_2)]
    repaired = [_evidence("B25:B26", _B08_JOINED_SOURCE_TEXT)]
    evidence_state = target.get("evidence")
    if evidence_state == original:
        return target, False
    if evidence_state == repaired:
        return target, True
    raise B04B08RepairError(
        "B08 repair is partial, conflicting, or already altered"
    )


def _repair_b04(baseline: dict[str, Any]) -> dict[str, Any]:
    repaired = copy.deepcopy(baseline)
    observation, conclusion, is_repaired = _b04_state(repaired)
    if is_repaired:
        return repaired
    observation["valueNumber"] = None
    conclusion["evidence"].extend(_b04_context_evidence())
    _b04_state(repaired)
    return repaired


def _repair_b08(baseline: dict[str, Any]) -> dict[str, Any]:
    repaired = copy.deepcopy(baseline)
    target, is_repaired = _b08_state(repaired)
    if is_repaired:
        return repaired
    target["evidence"][0]["range"] = "B25:B26"
    target["evidence"][0]["sourceText"] = _B08_JOINED_SOURCE_TEXT
    _b08_state(repaired)
    return repaired


def _repair_kind(
    validation_error: str,
    baseline: dict[str, Any],
) -> str:
    revision_uid, content_sha256, file_name = _source_identity(baseline)
    if (
        revision_uid == B04_REVISION_UID
        and content_sha256 == B04_CONTENT_SHA256
        and file_name == B04_FILE_NAME
        and _error_matches(
            validation_error,
            B04_NUMERIC_VALIDATION_ERROR,
        )
    ):
        return _B04_REPAIR_KIND
    if (
        revision_uid == B08_REVISION_UID
        and content_sha256 == B08_CONTENT_SHA256
        and file_name == B08_FILE_NAME
        and _error_matches(
            validation_error,
            B08_CONCLUSION_VALIDATION_ERROR,
        )
    ):
        return _B08_REPAIR_KIND
    raise B04B08RepairError(
        "source identity and validation error are not an exact repair target"
    )


def b04_b08_repair_target(
    validation_error: str,
    baseline: dict[str, Any],
    focused_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Return a protected exact repair target, or ``None`` if unproven."""

    try:
        if not isinstance(baseline, dict):
            raise B04B08RepairError("baseline is not an object")
        kind = _repair_kind(validation_error, baseline)
        if kind == _B04_REPAIR_KIND:
            _validate_b04_source_geometry(focused_chunks)
            _b04_state(baseline)
            repaired = _repair_b04(baseline)
        else:
            _validate_b08_source_geometry(focused_chunks)
            _b08_state(baseline)
            repaired = _repair_b08(baseline)
        return {
            "schemaVersion": B04_B08_REPAIR_SCHEMA_VERSION,
            "repairKind": kind,
            "revisionUid": baseline["source"]["revisionUid"],
            "contentSha256": baseline["source"]["contentSha256"],
            "baselineProjectionSha256": _canonical_sha256(baseline),
            "repairedProjectionSha256": _canonical_sha256(repaired),
        }
    except (
        B04B08RepairError,
        KeyError,
        TypeError,
        ValueError,
    ):
        return None


def _validate_target(target: dict[str, Any]) -> str:
    if not isinstance(target, dict):
        raise B04B08RepairError("repair target is not an object")
    kind = str(target.get("repairKind") or "")
    expected_identity = {
        _B04_REPAIR_KIND: (B04_REVISION_UID, B04_CONTENT_SHA256),
        _B08_REPAIR_KIND: (B08_REVISION_UID, B08_CONTENT_SHA256),
    }.get(kind)
    if (
        target.get("schemaVersion")
        != B04_B08_REPAIR_SCHEMA_VERSION
        or expected_identity is None
        or target.get("revisionUid") != expected_identity[0]
        or str(target.get("contentSha256") or "").lower()
        != expected_identity[1]
    ):
        raise B04B08RepairError("repair target is not exact")
    for field in (
        "baselineProjectionSha256",
        "repairedProjectionSha256",
    ):
        value = str(target.get(field) or "")
        if len(value) != 64 or any(
            character not in "0123456789abcdef"
            for character in value
        ):
            raise B04B08RepairError(
                f"repair target {field} is invalid"
            )
    return kind


def apply_b04_b08_repair(
    baseline: dict[str, Any],
    target: dict[str, Any],
) -> dict[str, Any]:
    """Apply only one protected B04/B08 projection."""

    kind = _validate_target(target)
    current_hash = _canonical_sha256(baseline)
    if current_hash not in {
        target["baselineProjectionSha256"],
        target["repairedProjectionSha256"],
    }:
        raise B04B08RepairError(
            "baseline changed within the protected repair projection"
        )
    if kind == _B04_REPAIR_KIND:
        repaired = _repair_b04(baseline)
    else:
        repaired = _repair_b08(baseline)
    if _canonical_sha256(repaired) != target["repairedProjectionSha256"]:
        raise B04B08RepairError(
            "repair produced an unexpected protected projection"
        )
    return repaired


def validate_b04_b08_repair(
    baseline: dict[str, Any],
    repaired: dict[str, Any],
    target: dict[str, Any],
) -> None:
    """Reject every mutation outside the exact source-backed projection."""

    expected = apply_b04_b08_repair(baseline, target)
    if repaired != expected:
        raise B04B08RepairError(
            "repair changed fields outside the exact B04/B08 projection"
        )


__all__ = [
    "B04_B08_REPAIR_SCHEMA_VERSION",
    "B04_CONTENT_SHA256",
    "B04_FILE_NAME",
    "B04_NUMERIC_VALIDATION_ERROR",
    "B04_REVISION_UID",
    "B08_CONCLUSION_VALIDATION_ERROR",
    "B08_CONTENT_SHA256",
    "B08_FILE_NAME",
    "B08_REVISION_UID",
    "B04B08RepairError",
    "apply_b04_b08_repair",
    "b04_b08_repair_target",
    "validate_b04_b08_repair",
]
