"""Read-only Codex runner for append-only Study fragment v2 outputs."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import uuid
from pathlib import Path
from typing import Any, Callable, Sequence

from inference_data_ai_staged_draft_v2 import (
    STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
    StagedDraftV2Error,
    bytes_sha256,
    normalize_fragment_evidence_dispositions,
    normalize_fragment_complete_dispositions,
    normalize_fragment_missing_observation_arms,
    normalize_fragment_observation_replicate_evidence,
    normalize_fragment_multi_arm_series_rows,
    normalize_fragment_required_fields_and_series_headers,
    normalize_fragment_record_ids,
    normalize_fragment_unsupported_text_numeric_claims,
    validate_fragment_v2,
)


RunCommand = Callable[..., subprocess.CompletedProcess[str]]


def fragment_output_schema_v2() -> dict[str, Any]:
    evidence = {
        "type": "object",
        "properties": {
            "sheet": {"type": "string"},
            "range": {"type": "string"},
            "role": {"type": "string"},
            "sourceText": {"type": "string"},
            "note": {"type": "string"},
        },
        "required": [
            "sheet",
            "range",
            "role",
            "sourceText",
            "note",
        ],
        "additionalProperties": False,
    }
    record = {
        "type": "object",
        "properties": {
            "recordType": {
                "type": "string",
                "enum": [
                    "STUDY_PATCH",
                    "ENTITY_DECLARATION",
                    "OBSERVATION_APPEND",
                    "SERIES_SEGMENT_APPEND",
                    "COMPARISON_LINK_INTENT",
                    "CONCLUSION_APPEND",
                    "LIMITATION_APPEND",
                ],
            },
            "recordId": {"type": "string"},
            "logicalStudyId": {"type": "string"},
            "identityCellKeys": {
                "type": "array",
                "items": {"type": "string"},
                "minItems": 1,
            },
            "exactSourceLabel": {"type": "string"},
            "payloadJson": {"type": "string"},
            "evidence": {
                "type": "array",
                "items": evidence,
            },
        },
        "required": [
            "recordType",
            "recordId",
            "logicalStudyId",
            "identityCellKeys",
            "exactSourceLabel",
            "payloadJson",
            "evidence",
        ],
        "additionalProperties": False,
    }
    disposition = {
        "type": "object",
        "properties": {
            "sourceCellKey": {"type": "string"},
            "disposition": {
                "type": "string",
                "enum": [
                    "RECORD_EVIDENCE",
                    "CONTEXT_ONLY",
                    "NO_SEMANTIC_RECORD",
                ],
            },
            "recordIds": {
                "type": "array",
                "items": {"type": "string"},
            },
            "reason": {"type": "string"},
        },
        "required": [
            "sourceCellKey",
            "disposition",
            "recordIds",
            "reason",
        ],
        "additionalProperties": False,
    }
    return {
        "type": "object",
        "properties": {
            "schemaVersion": {
                "type": "string",
                "const": STUDY_DRAFT_FRAGMENT_V2_SCHEMA_VERSION,
            },
            "source": {
                "type": "object",
                "properties": {
                    "revisionUid": {"type": "string"},
                    "contentSha256": {"type": "string"},
                    "contentComplete": {
                        "type": "boolean",
                        "const": False,
                    },
                },
                "required": [
                    "revisionUid",
                    "contentSha256",
                    "contentComplete",
                ],
                "additionalProperties": False,
            },
            "planId": {"type": "string"},
            "partId": {"type": "string"},
            "inputEnvelopeSha256": {"type": "string"},
            "records": {
                "type": "array",
                "items": record,
            },
            "coverageDispositions": {
                "type": "array",
                "items": disposition,
            },
        },
        "required": [
            "schemaVersion",
            "source",
            "planId",
            "partId",
            "inputEnvelopeSha256",
            "records",
            "coverageDispositions",
        ],
        "additionalProperties": False,
    }


def _decode_fragment_transport(
    fragment: object,
) -> dict[str, Any]:
    """Decode strict transport payload strings into internal payload objects."""

    if not isinstance(fragment, dict):
        raise StagedDraftV2Error(
            "Fragment transport output must be a JSON object"
        )
    records = fragment.get("records")
    if not isinstance(records, list):
        raise StagedDraftV2Error(
            "Fragment transport records must be an array"
        )

    def unique_object(
        pairs: list[tuple[str, Any]],
    ) -> dict[str, Any]:
        value: dict[str, Any] = {}
        for key, item in pairs:
            if key in value:
                raise ValueError(f"duplicate property {key!r}")
            value[key] = item
        return value

    def reject_constant(value: str) -> None:
        raise ValueError(f"non-JSON numeric constant {value!r}")

    for index, record in enumerate(records):
        if not isinstance(record, dict):
            raise StagedDraftV2Error(
                f"Fragment transport records[{index}] must be an object"
            )
        if "payload" in record:
            raise StagedDraftV2Error(
                f"Fragment transport records[{index}] must use payloadJson"
            )
        payload_json = record.pop("payloadJson", None)
        if not isinstance(payload_json, str):
            raise StagedDraftV2Error(
                f"Fragment transport records[{index}].payloadJson "
                "must be a string"
            )
        try:
            payload = json.loads(
                payload_json,
                object_pairs_hook=unique_object,
                parse_constant=reject_constant,
            )
        except (json.JSONDecodeError, ValueError) as exc:
            raise StagedDraftV2Error(
                f"Fragment transport records[{index}].payloadJson "
                "must contain valid JSON"
            ) from exc
        if not isinstance(payload, dict):
            raise StagedDraftV2Error(
                f"Fragment transport records[{index}].payloadJson "
                "must encode an object"
            )
        record["payload"] = payload
    return fragment


def _normalize_current_fragment(
    *,
    fragment: dict[str, Any],
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
    rebind_identity: bool = False,
) -> dict[str, Any]:
    current = json.loads(json.dumps(fragment))
    if rebind_identity:
        current["planId"] = envelope["planId"]
        current["partId"] = envelope["partId"]
        current["inputEnvelopeSha256"] = envelope[
            "inputEnvelopeSha256"
        ]
    normalized = normalize_fragment_missing_observation_arms(
        fragment=current,
        envelope=envelope,
        all_selected_chunks=all_selected_chunks,
    )
    normalized = normalize_fragment_observation_replicate_evidence(
        fragment=normalized,
        envelope=envelope,
        all_selected_chunks=all_selected_chunks,
    )
    normalized = normalize_fragment_required_fields_and_series_headers(
        fragment=normalized,
        all_selected_chunks=all_selected_chunks,
    )
    normalized = normalize_fragment_multi_arm_series_rows(
        fragment=normalized,
        all_selected_chunks=all_selected_chunks,
    )
    normalized = normalize_fragment_record_ids(
        fragment=normalized,
        envelope=envelope,
    )
    normalized = normalize_fragment_evidence_dispositions(
        fragment=normalized,
        all_selected_chunks=all_selected_chunks,
    )
    normalized = normalize_fragment_complete_dispositions(
        fragment=normalized,
        envelope=envelope,
        all_selected_chunks=all_selected_chunks,
    )
    return normalize_fragment_unsupported_text_numeric_claims(
        fragment=normalized,
        envelope=envelope,
        all_selected_chunks=all_selected_chunks,
    )


def _reusable_prior_fragment(
    *,
    target: Path,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Promote an older fragment only after exact current validation."""

    owned = {
        str(value)
        for value in envelope.get("ownedSourceCellKeys", [])
    }
    expected_source = envelope.get("source", {})
    candidates = {
        *target.parent.glob(
            "study-draft-part-v2_*.fragment.json"
        ),
        *target.parent.glob(
            "study-draft-part-v2_*.fragment.rejected.json"
        ),
    }
    for path in sorted(
        candidates,
        key=lambda item: item.stat().st_mtime_ns,
        reverse=True,
    ):
        if path == target or not path.is_file():
            continue
        try:
            candidate = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if not isinstance(candidate, dict):
            continue
        source = candidate.get("source")
        if not isinstance(source, dict) or any(
            str(source.get(field) or "").casefold()
            != str(expected_source.get(field) or "").casefold()
            for field in ("revisionUid", "contentSha256")
        ):
            continue
        dispositions = candidate.get("coverageDispositions")
        if not isinstance(dispositions, list):
            continue
        candidate_owned = {
            str(item.get("sourceCellKey") or "")
            for item in dispositions
            if isinstance(item, dict)
            and str(item.get("sourceCellKey") or "")
        }
        if candidate_owned != owned:
            continue
        try:
            normalized = _normalize_current_fragment(
                fragment=candidate,
                envelope=envelope,
                all_selected_chunks=all_selected_chunks,
                rebind_identity=True,
            )
            return validate_fragment_v2(
                fragment=normalized,
                envelope=envelope,
                all_selected_chunks=all_selected_chunks,
            )
        except Exception:
            continue
    return None


def _codex_command(command: Sequence[str] | None) -> list[str]:
    if command:
        return [str(value) for value in command]
    executable = shutil.which("codex.cmd" if os.name == "nt" else "codex")
    if not executable:
        raise StagedDraftV2Error("Codex CLI executable was not found")
    return [executable]


def run_codex_study_fragment_v2(
    *,
    envelope: dict[str, Any],
    all_selected_chunks: Sequence[dict[str, Any]],
    output_path: str | Path,
    model: str | None = None,
    reasoning_effort: str | None = None,
    timeout_seconds: int = 1800,
    codex_command: Sequence[str] | None = None,
    run_command: RunCommand = subprocess.run,
    ai_call_observer: Callable[[], None] | None = None,
) -> dict[str, Any]:
    """Execute exactly the prompt hashed in the fragment input envelope."""

    prompt = str(envelope.get("promptText") or "")
    expected_prompt_sha = str(
        envelope.get("inputHashes", {}).get("promptSha256") or ""
    )
    if not prompt or bytes_sha256(prompt.encode("utf-8")) != expected_prompt_sha:
        raise StagedDraftV2Error(
            "Fragment runner prompt does not match its provenance hash"
        )
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    reusable = _reusable_prior_fragment(
        target=target,
        envelope=envelope,
        all_selected_chunks=all_selected_chunks,
    )
    if reusable is not None:
        rejected_path = target.with_name(
            target.stem + ".rejected" + target.suffix
        )
        rejected_path.unlink(missing_ok=True)
        target.write_text(
            json.dumps(reusable, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return reusable
    transport_token = uuid.uuid4().hex
    transport_prefix = f".{target.name}.{transport_token}"
    schema_path = target.parent / (
        transport_prefix + ".fragment.schema.json"
    )
    output_message = target.parent / (
        transport_prefix + ".last-message.json"
    )
    try:
        schema_path.write_text(
            json.dumps(fragment_output_schema_v2(), ensure_ascii=False),
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
            str(output_message),
        ]
        if reasoning_effort:
            command.extend(
                ["-c", f'model_reasoning_effort="{reasoning_effort}"']
            )
        if model:
            command.extend(["--model", model])
        command.append("-")
        if ai_call_observer is not None:
            ai_call_observer()
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
            raise StagedDraftV2Error(
                "Codex fragment v2 failed with exit code "
                f"{completed.returncode}: {detail[-2000:]}"
            )
        if not output_message.is_file():
            raise StagedDraftV2Error(
                "Codex fragment v2 did not produce an output message"
            )
        try:
            transport_fragment = json.loads(
                output_message.read_text(encoding="utf-8")
            )
        except json.JSONDecodeError as exc:
            raise StagedDraftV2Error(
                "Codex fragment v2 output is invalid JSON"
            ) from exc
    finally:
        for transport_path in (output_message, schema_path):
            transport_path.unlink(missing_ok=True)
    raw_fragment = _decode_fragment_transport(transport_fragment)
    normalized = _normalize_current_fragment(
        fragment=raw_fragment,
        envelope=envelope,
        all_selected_chunks=all_selected_chunks,
    )
    rejected_path = target.with_name(
        target.stem + ".rejected" + target.suffix
    )
    try:
        validated = validate_fragment_v2(
            fragment=normalized,
            envelope=envelope,
            all_selected_chunks=all_selected_chunks,
        )
    except Exception:
        rejected_path.write_text(
            json.dumps(normalized, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        raise
    rejected_path.unlink(missing_ok=True)
    target.write_text(
        json.dumps(validated, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return validated


__all__ = [
    "fragment_output_schema_v2",
    "run_codex_study_fragment_v2",
]
