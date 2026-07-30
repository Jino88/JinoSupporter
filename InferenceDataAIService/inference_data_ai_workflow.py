"""Durable single-XLSX ingestion into the canonical evidence database.

The workflow is intentionally source-read-only and leaves every AI-produced
study in NEEDS_REVIEW (or EXCLUDED for terminal tabular states). Each stage is
journaled so the same source fingerprint can resume without repeating valid
locator or draft artifacts.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import sqlite3
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager, nullcontext
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, ContextManager, Iterator, Sequence

from inference_data_ai_schema import (
    ensure_knowledge_schema,
    stable_uid,
    validate_analysis_integrity,
)
from inference_data_ai_content_coverage import (
    augment_exact_source_conclusions,
    build_content_coverage_inventory,
    validate_content_manifest_coverage,
)
from inference_data_ai_formula_derivation import (
    FORMULA_DERIVATION_SCHEMA_VERSION,
    FORMULA_EVALUATOR_VERSION,
    apply_formula_overlay_to_chunks,
    derive_formula_overlay,
    validate_formula_overlay,
)
from inference_data_ai_semantic_ai import (
    BATCH_LOCATOR_PROMPT_VERSION,
    LOCATOR_PROMPT_VERSION,
    STUDY_DRAFT_PROMPT_VERSION,
    build_study_draft_prompt,
    run_codex_locator_batch,
    run_codex_study_draft,
    validate_ai_study_draft,
    validate_locator_result,
)
from inference_data_ai_semantic_packets import (
    build_semantic_source_packets_from_db,
    packet_json_bytes,
)
from inference_data_ai_source_ingest import (
    CAPTURE_CONTRACT,
    COM_CAPTURE_CONTRACT,
    bridge_capture_to_canonical_source,
    ensure_capture_v2_schema,
    extract_workbook,
    import_capture,
    reconcile_capture_sheet_counts,
    verify_capture_revision,
)
from inference_data_ai_staged_draft import (
    assess_one_call_budget,
    audit_no_candidate_source_inventory,
    audit_unselected_source_inventory,
    build_deterministic_acoustic_matrix_fragment_v2,
    build_deterministic_error_axis_tail_fragment_v2,
    build_deterministic_fo_fragment_v2,
    build_deterministic_function_fragment_v2,
    build_deterministic_function_grid_fragment_v2,
    build_deterministic_mask_fragment_v2,
    build_deterministic_nti_f0_fragment_v2,
    build_deterministic_nti_horizontal_matrix_fragment_v2,
    build_deterministic_result_table_fragment_v2,
    build_fragment_envelope,
    build_monolithic_request,
    build_study_registry_v2,
    chunks_for_part_v2,
    final_provenance_v2,
    final_provenance_v2_matches,
    finalize_fragment_envelope,
    fragment_artifact_paths,
    locators_for_part_v2,
    merge_fragment_records,
    part_provenance_v2,
    part_provenance_v2_matches,
    plan_study_draft_v2,
    promote_required_source_locator_sections,
    project_canonical_manifest,
    registry_for_part,
    select_draft_universe,
    validate_fragment_v2,
)
from inference_data_ai_staged_runner_v2 import (
    run_codex_study_fragment_v2,
)
from inference_data_ai_study_contract import validate_study_manifest
from inference_data_ai_study_import import (
    AnalysisQuarantineError,
    import_study_manifest,
    make_database_evidence_checker,
    quarantine_canonical_analysis,
    resolve_manifest_revision,
    unsupported_rate_pair_observation_paths,
    validate_comparison_representation_alignment,
    validate_conclusion_evidence,
    validate_factor_and_arm_evidence,
    validate_numeric_observation_evidence,
)


WORKFLOW_SCHEMA_VERSION = "incremental-xlsx-ingest-v1"
DRAFT_PROVENANCE_SCHEMA_VERSION = "study-draft-provenance-v1"
TERMINAL_WORKBOOK_STATUSES = {"EMPTY_WORKBOOK", "NO_TABULAR_EVIDENCE"}
TERMINAL_ANALYSIS_STATUSES = {
    "EMPTY_WORKBOOK",
    "NO_TABULAR_EVIDENCE",
    "NO_SEMANTIC_CANDIDATE",
}
STAGE_ORDER = ("CAPTURE", "PACKET", "LOCATOR", "DRAFT", "IMPORT", "VERIFY")


class IncrementalIngestError(RuntimeError):
    """Raised when a source cannot safely complete the incremental workflow."""


class IncrementalIngestBusyError(IncrementalIngestError):
    """Raised when the same source fingerprint is already being processed."""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _safe_name(value: object) -> str:
    text = "".join(
        char if char.isalnum() or char in "._-" else "_"
        for char in str(value)
    ).strip("._")
    return text[:120] or "item"


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _atomic_write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def _atomic_write_bytes(path: Path, value: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_bytes(value)
    temporary.replace(path)


def _reset_journal_for_retry(
    journal: dict[str, Any],
    *,
    started_at: str,
) -> None:
    """Start one observable attempt without mixing prior stage results."""

    previous_attempt = {
        "attempt": int(journal.get("attempt") or 0),
        "semanticContracts": copy.deepcopy(
            journal.get("semanticContracts")
        ),
        "status": str(journal.get("status") or ""),
        "currentStage": str(journal.get("currentStage") or ""),
        "stages": copy.deepcopy(journal.get("stages") or {}),
        "result": copy.deepcopy(journal.get("result")),
        "updatedAt": str(journal.get("updatedAt") or ""),
        "finishedAt": str(journal.get("finishedAt") or ""),
    }
    history = journal.get("attemptHistory")
    if not isinstance(history, list):
        history = []
    history.append(previous_attempt)
    journal["attemptHistory"] = history
    journal["attempt"] = previous_attempt["attempt"] + 1
    journal["status"] = "RUNNING"
    journal["currentStage"] = ""
    journal["stages"] = {
        name: {"status": "PENDING"}
        for name in STAGE_ORDER
    }
    journal["result"] = None
    journal["updatedAt"] = started_at
    journal["finishedAt"] = ""


def _invalidate_downstream_journal_stages(
    journal: dict[str, Any],
    stage_name: str,
) -> None:
    """Clear terminal state before a stage is executed or re-executed."""

    try:
        stage_index = STAGE_ORDER.index(stage_name)
    except ValueError as exc:
        raise IncrementalIngestError(
            f"Unknown workflow stage: {stage_name}"
        ) from exc
    stages = journal.setdefault("stages", {})
    for downstream in STAGE_ORDER[stage_index + 1 :]:
        stages[downstream] = {"status": "PENDING"}
    journal["result"] = None
    journal["finishedAt"] = ""


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _draft_provenance_matches(
    provenance_path: Path,
    manifest_path: Path,
    source: dict[str, Any],
    expected_details: dict[str, Any] | None = None,
) -> bool:
    if not provenance_path.is_file() or not manifest_path.is_file():
        return False
    try:
        provenance = json.loads(
            provenance_path.read_text(encoding="utf-8")
        )
    except (OSError, json.JSONDecodeError):
        return False
    expected = provenance.get("source", {})
    matches = bool(
        provenance.get("schemaVersion")
        == DRAFT_PROVENANCE_SCHEMA_VERSION
        and provenance.get("promptVersion")
        == STUDY_DRAFT_PROMPT_VERSION
        and expected.get("revisionUid") == source.get("revisionUid")
        and str(expected.get("contentSha256") or "").lower()
        == str(source.get("contentSha256") or "").lower()
        and provenance.get("manifestSha256")
        == _file_sha256(manifest_path)
    )
    if not matches:
        return False
    if expected_details is not None:
        actual_details = provenance.get("details")
        return isinstance(actual_details, dict) and all(
            actual_details.get(key) == value
            for key, value in expected_details.items()
        )
    return True


def _write_draft_provenance(
    provenance_path: Path,
    manifest_path: Path,
    source: dict[str, Any],
    details: dict[str, Any] | None = None,
) -> None:
    _atomic_write_json(
        provenance_path,
        {
            "schemaVersion": DRAFT_PROVENANCE_SCHEMA_VERSION,
            "promptVersion": STUDY_DRAFT_PROMPT_VERSION,
            "source": {
                "revisionUid": source["revisionUid"],
                "contentSha256": source["contentSha256"],
            },
            "manifestPath": str(manifest_path),
            "manifestSha256": _file_sha256(manifest_path),
            "generatedAt": utc_now_iso(),
            "imagesAnalyzed": False,
            "details": copy.deepcopy(details or {}),
        },
    )


def _pid_exists(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except (OSError, SystemError):
        # Microsoft Store Python can surface an invalid/nonexistent Windows
        # PID as SystemError instead of OSError. Treat both as a stale owner.
        return False
    return True


@contextmanager
def _workflow_lock(path: Path) -> Iterator[None]:
    path.parent.mkdir(parents=True, exist_ok=True)
    while True:
        try:
            descriptor = os.open(
                path,
                os.O_CREAT | os.O_EXCL | os.O_WRONLY,
            )
            os.write(descriptor, str(os.getpid()).encode("ascii"))
            break
        except FileExistsError as exc:
            try:
                owner_pid = int(path.read_text(encoding="ascii").strip())
            except (OSError, ValueError):
                owner_pid = 0
            if owner_pid and _pid_exists(owner_pid):
                raise IncrementalIngestBusyError(
                    f"Incremental ingestion is already running for this source (PID {owner_pid})."
                ) from exc
            try:
                path.unlink()
            except FileNotFoundError:
                pass
    try:
        yield
    finally:
        os.close(descriptor)
        try:
            path.unlink()
        except FileNotFoundError:
            pass


@contextmanager
def _connect_rw(database_path: Path) -> Iterator[sqlite3.Connection]:
    if not database_path.is_file():
        raise IncrementalIngestError(
            f"Canonical database is not initialized: {database_path}"
        )
    connection = sqlite3.connect(str(database_path), timeout=60)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout=60000")
    connection.execute("PRAGMA foreign_keys=ON")
    try:
        if (
            connection.execute(
                """
                SELECT 1
                FROM sqlite_master
                WHERE type='table' AND name='schema_migrations'
                """
            ).fetchone()
            is None
        ):
            raise IncrementalIngestError(
                "Canonical database must be initialized before incremental ingestion."
            )
        yield connection
    finally:
        connection.close()


def _quarantine_invalidated_unverified_analyses(
    database_path: Path,
    *,
    source: dict[str, Any],
    reason: str,
    now_iso: Callable[[], str],
) -> dict[str, Any]:
    """Fail closed when a previously imported draft is no longer reusable."""

    quarantined: list[str] = []
    already_stale: list[str] = []
    protected: list[dict[str, str]] = []
    with _connect_rw(database_path) as connection:
        revision = resolve_manifest_revision(connection, source)
        rows = connection.execute(
            """
            SELECT public_analysis_id, analysis_status,
                   verification_status
            FROM workbook_analyses
            WHERE revision_id=?
              AND analyzer_name='canonical-study-import'
            ORDER BY workbook_analysis_id
            """,
            (int(revision["revision_id"]),),
        ).fetchall()
        for row in rows:
            public_analysis_id = str(row["public_analysis_id"])
            analysis_status = str(row["analysis_status"]).upper()
            verification_status = str(
                row["verification_status"]
            ).upper()
            if (
                analysis_status == "STALE"
                or verification_status == "STALE"
            ):
                already_stale.append(public_analysis_id)
                continue
            if verification_status == "VERIFIED":
                protected.append(
                    {
                        "publicAnalysisId": public_analysis_id,
                        "reason": "VERIFIED analysis is protected",
                    }
                )
                continue
            try:
                quarantine_canonical_analysis(
                    connection,
                    public_analysis_id=public_analysis_id,
                    reason=reason,
                    now_iso=now_iso,
                )
            except AnalysisQuarantineError as exc:
                protected.append(
                    {
                        "publicAnalysisId": public_analysis_id,
                        "reason": str(exc),
                    }
                )
            else:
                quarantined.append(public_analysis_id)
        connection.commit()
    return {
        "quarantined": quarantined,
        "alreadyStale": already_stale,
        "protected": protected,
    }


def _source_identity(
    packet_set: dict[str, Any],
    dataset: str,
) -> dict[str, Any]:
    revision = packet_set["inventory"]["sourceRevision"]
    return {
        "dataset": dataset,
        "sourcePath": str(revision["sourcePath"]),
        "revisionUid": str(revision["revisionUid"]),
        "contentSha256": str(revision["contentSha256"]),
    }


def _workbook_summary(packet_set: dict[str, Any]) -> dict[str, Any]:
    inventory = packet_set["inventory"]
    return {
        **inventory["workbook"],
        "semanticCellCoverageComplete": bool(
            inventory["semanticCellCoverageComplete"]
        ),
        "contentCompleteForManifest": bool(
            inventory["contentCompleteForManifest"]
        ),
        "coverage": inventory["coverage"],
        "sheets": [
            {
                "sheetIndex": sheet["sheetIndex"],
                "title": sheet["title"],
                "sheetState": sheet["sheetState"],
                "status": sheet["status"],
                "hasTabularEvidence": sheet["hasTabularEvidence"],
                "nonEmptyCellCount": sheet["nonEmptyCellCount"],
                "formulaCellCount": sheet["formulaCellCount"],
                "mergeCount": sheet["mergeCount"],
                "contentBounds": sheet["contentBounds"],
                "sections": sheet["sections"],
            }
            for sheet in inventory["sheets"]
        ],
    }


def _terminal_manifest(
    packet_set: dict[str, Any],
    dataset: str,
    *,
    status_override: str | None = None,
    extra_limitation: str | None = None,
    analysis_key_override: str | None = None,
) -> dict[str, Any]:
    source = _source_identity(packet_set, dataset)
    revision = packet_set["inventory"]["sourceRevision"]
    capture_status = str(packet_set["inventory"]["workbook"]["status"])
    analysis_status = status_override or capture_status
    limitations = [
        f"Capture v2 workbook status: {capture_status}.",
        "Only tabular cell evidence is in scope; embedded images were not analyzed.",
    ]
    if extra_limitation:
        limitations.append(extra_limitation)
    return validate_study_manifest(
        {
            "schemaVersion": "canonical-study-manifest-v1",
            "source": {
                **source,
                "contentComplete": False,
            },
            "workbookAnalysis": {
                "key": analysis_key_override
                or (
                    "tabular-evidence-"
                    + _safe_name(revision["revisionUid"]).lower()
                ),
                "title": str(revision["fileName"]),
                "summary": (
                    "No reviewable tabular Study was accepted from this source. "
                    "Visual content is outside the configured scope."
                ),
                "status": analysis_status,
                "verificationStatus": "EXCLUDED",
                "limitations": limitations,
                "evidence": [],
            },
            "studies": [],
        }
    )


def _draft_has_labels_but_no_reviewable_results(
    manifest: dict[str, Any],
) -> bool:
    """Detect outcome labels that have no observations or conclusions."""

    studies = manifest.get("studies", [])
    outcome_count = sum(
        len(study.get("outcomes", []))
        for study in studies
    )
    observation_count = sum(
        len(outcome.get("observations", []))
        for study in studies
        for outcome in study.get("outcomes", [])
    )
    conclusion_count = sum(
        len(study.get("conclusions", []))
        for study in studies
    )
    return (
        bool(studies)
        and outcome_count > 0
        and observation_count == 0
        and conclusion_count == 0
    )


def _has_primary_text(chunk: dict[str, Any]) -> bool:
    for cell in chunk.get("cells", []):
        value = (
            cell.get("displayValue")
            if cell.get("displayValue") is not None
            else cell.get("rawValue")
        )
        if isinstance(value, str) and value.strip():
            return True
    return False


def _deterministic_locator(
    source: dict[str, Any],
    chunk: dict[str, Any],
) -> dict[str, Any]:
    return {
        "schemaVersion": "semantic-locator-v1",
        "promptVersion": LOCATOR_PROMPT_VERSION,
        "revisionUid": source["revisionUid"],
        "contentSha256": source["contentSha256"],
        "chunkId": str(chunk["chunkId"]),
        "status": "NO_CANDIDATE",
        "candidates": [],
        "notes": [
            "Deterministic numeric/formula continuation retained for on-demand evidence retrieval."
        ],
    }


def _partition_jobs(
    jobs: list[tuple[dict[str, Any], Path]],
    *,
    batch_size: int,
    batch_max_bytes: int,
) -> list[list[tuple[dict[str, Any], Path]]]:
    if batch_size < 1 or batch_max_bytes < 1:
        raise ValueError("Locator batch limits must be positive.")
    batches: list[list[tuple[dict[str, Any], Path]]] = []
    current: list[tuple[dict[str, Any], Path]] = []
    current_bytes = 0
    for job in jobs:
        size = len(
            json.dumps(
                job[0],
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode("utf-8")
        )
        if current and (
            len(current) >= batch_size
            or current_bytes + size > batch_max_bytes
        ):
            batches.append(current)
            current = []
            current_bytes = 0
        current.append(job)
        current_bytes += size
    if current:
        batches.append(current)
    return batches


def _run_locator_batch_with_singleton_retry(
    batch: list[tuple[dict[str, Any], Path]],
    runner: Callable[
        [list[tuple[dict[str, Any], Path]]],
        list[dict[str, Any]],
    ],
) -> tuple[list[dict[str, Any]], int]:
    """Retry each chunk once in isolation after a failed batch validation."""

    try:
        return runner(batch), 1
    except Exception as batch_error:
        recovered: list[dict[str, Any]] = []
        failures: list[str] = []
        for job in batch:
            chunk_id = str(job[0]["chunkId"])
            try:
                recovered.extend(runner([job]))
            except Exception as exc:
                failures.append(f"{chunk_id}: {exc}")
        if failures:
            raise IncrementalIngestError(
                "Locator batch validation failed and isolated chunk retry "
                "did not recover every chunk. Initial error: "
                f"{batch_error}. Retry errors: {' | '.join(failures)}"
            ) from batch_error
        return recovered, 1 + len(batch)


def _load_locator_results(
    packet_set: dict[str, Any],
    locator_dir: Path,
    source: dict[str, Any],
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for chunk in packet_set["chunks"]:
        path = locator_dir / f"{_safe_name(chunk['chunkId'])}.locator.json"
        if not path.is_file():
            raise IncrementalIngestError(
                f"Missing locator result for source chunk {chunk['chunkId']}."
            )
        try:
            result = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise IncrementalIngestError(
                f"Invalid locator JSON for source chunk {chunk['chunkId']}."
            ) from exc
        results.append(
            validate_locator_result(
                result,
                revision_uid=source["revisionUid"],
                content_sha256=source["contentSha256"],
                chunk=chunk,
            )
        )
    return results


def ingest_workbook(
    *,
    database_path: str | Path,
    source_path: str | Path,
    artifact_root: str | Path,
    dataset: str = "InputDataFinish",
    resume: bool = True,
    max_cells: int = 400,
    max_rows: int = 50,
    empty_row_gap: int = 3,
    locator_workers: int = 3,
    locator_batch_size: int = 6,
    locator_batch_max_bytes: int = 240_000,
    draft_monolithic_max_bytes: int = 400_000,
    draft_fragment_max_chunks: int = 8,
    draft_fragment_max_cells: int = 2_000,
    draft_fragment_max_bytes: int = 400_000,
    draft_fragment_workers: int = 3,
    derive_formula_values: bool = False,
    repair_rejected_draft: bool = False,
    repair_unselected_source: bool = False,
    model: str | None = None,
    reasoning_effort: str | None = "medium",
    locator_timeout_seconds: int = 900,
    draft_timeout_seconds: int = 1800,
    capture_backend: str = "openxml",
    covered_cell_mode: str = "blank",
    include_hidden_sheets: bool = True,
    inspect_auth_dialog: bool = False,
    dismiss_auth_dialog: bool = False,
    auth_dialog_title: str = "",
    auth_dialog_class: str = "",
    auth_dialog_button: str = "",
    auth_dialog_timeout_seconds: float = 30.0,
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    pipeline_gate: (
        Callable[[str], ContextManager[None]] | None
    ) = None,
    now_iso: Callable[[], str] = utc_now_iso,
    locator_batch_runner: Callable[..., list[dict[str, Any]]] = run_codex_locator_batch,
    draft_runner: Callable[..., dict[str, Any]] = run_codex_study_draft,
    fragment_runner: Callable[..., dict[str, Any]] = (
        run_codex_study_fragment_v2
    ),
) -> dict[str, Any]:
    """Ingest one Excel source and return a review-gated terminal result."""

    database = Path(database_path).expanduser().resolve()
    source_file = Path(source_path).expanduser().resolve()
    artifacts = Path(artifact_root).expanduser().resolve()
    normalized_capture_backend = capture_backend.strip().lower()
    if normalized_capture_backend not in {"openxml", "com"}:
        raise ValueError("capture_backend must be openxml or com.")
    requested_capture_contract = (
        COM_CAPTURE_CONTRACT
        if normalized_capture_backend == "com"
        else CAPTURE_CONTRACT
    )
    if not source_file.is_file():
        raise FileNotFoundError(source_file)
    if (
        normalized_capture_backend == "openxml"
        and source_file.suffix.lower() != ".xlsx"
    ):
        raise IncrementalIngestError("Incremental ingestion accepts .xlsx sources only.")
    if (
        normalized_capture_backend == "com"
        and source_file.suffix.lower()
        not in {".xlsx", ".xlsm", ".xlsb", ".xls"}
    ):
        raise IncrementalIngestError(
            "Excel COM ingestion accepts .xlsx, .xlsm, .xlsb, or .xls."
        )
    if draft_fragment_workers < 1:
        raise ValueError("draft_fragment_workers must be positive.")
    source_sha256 = _sha256_file(source_file)
    run_id_parts = [
        "ingest-run",
        dataset,
        str(source_file),
        source_sha256,
        WORKFLOW_SCHEMA_VERSION,
    ]
    if normalized_capture_backend != "openxml":
        run_id_parts.extend(
            [normalized_capture_backend, requested_capture_contract]
        )
    if derive_formula_values:
        run_id_parts.extend(
            [
                FORMULA_DERIVATION_SCHEMA_VERSION,
                FORMULA_EVALUATOR_VERSION,
            ]
        )
    run_id = stable_uid(*run_id_parts)
    run_dir = artifacts / _safe_name(run_id)
    journal_path = run_dir / "journal.json"
    packet_path = run_dir / "semantic-source-packet.json"
    formula_overlay_path = run_dir / "formula-derivation.overlay.json"
    locator_dir = run_dir / "locators"
    manifest_path = run_dir / "canonical-study-manifest.json"
    draft_provenance_path = run_dir / "study-draft.provenance.json"
    draft_plan_path = run_dir / "study-draft-plan.json"
    draft_registry_path = run_dir / "study-registry-v2.json"
    staged_final_provenance_path = (
        run_dir / "study-draft-final-v2.provenance.json"
    )
    staged_merged_path = run_dir / "study-draft-merged-v2.json"
    lock_path = run_dir / ".workflow.lock"
    requested_formula_contract = {
        "enabled": bool(derive_formula_values),
        "schemaVersion": (
            FORMULA_DERIVATION_SCHEMA_VERSION
            if derive_formula_values
            else ""
        ),
        "evaluatorVersion": (
            FORMULA_EVALUATOR_VERSION
            if derive_formula_values
            else ""
        ),
    }
    requested_capture_semantics = {
        "backend": normalized_capture_backend,
        "captureContract": requested_capture_contract,
        "coveredCellMode": covered_cell_mode,
        "includeHiddenSheets": bool(include_hidden_sheets),
        "inspectAuthDialog": bool(inspect_auth_dialog),
        "dismissAuthDialog": bool(dismiss_auth_dialog),
        "authDialogTitle": auth_dialog_title,
        "authDialogClass": auth_dialog_class,
        "authDialogButton": auth_dialog_button,
        "authDialogTimeoutSeconds": float(
            auth_dialog_timeout_seconds
        ),
        "sourceReadOnly": True,
        "sourceSaved": False,
    }

    def pipeline_scope(name: str) -> ContextManager[None]:
        return (
            pipeline_gate(name)
            if pipeline_gate is not None
            else nullcontext()
        )

    def emit_progress(
        stage: str,
        status: str,
        detail: str = "",
    ) -> None:
        if progress_callback is None:
            return
        try:
            progress_callback(
                {
                    "schemaVersion": "ingest-progress-v1",
                    "runId": run_id,
                    "stage": stage,
                    "status": status,
                    "detail": detail,
                    "sourcePath": str(source_file),
                    "timestamp": now_iso(),
                }
            )
        except Exception:
            # Progress reporting must never corrupt or abort durable ingest.
            pass

    with _workflow_lock(lock_path):
        if resume and journal_path.is_file():
            try:
                journal = json.loads(journal_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                raise IncrementalIngestError(
                    f"Existing workflow journal is invalid: {journal_path}"
                ) from exc
            if (
                journal.get("schemaVersion") != WORKFLOW_SCHEMA_VERSION
                or journal.get("source", {}).get("contentSha256")
                != source_sha256
            ):
                raise IncrementalIngestError(
                    "Existing workflow journal does not match the source fingerprint."
                )
            existing_formula_contract = (
                journal.get("semanticContracts", {}).get(
                    "formulaDerivation"
                )
            )
            if existing_formula_contract is None:
                existing_formula_contract = {
                    "enabled": False,
                    "schemaVersion": "",
                    "evaluatorVersion": "",
                }
            comparable_existing_contract = {
                key: existing_formula_contract.get(key)
                for key in requested_formula_contract
            }
            if comparable_existing_contract != requested_formula_contract:
                raise IncrementalIngestError(
                    "Existing workflow journal formula-derivation contract "
                    "does not match this request."
                )
            existing_capture_semantics = (
                journal.get("semanticContracts", {}).get("capture")
            )
            if existing_capture_semantics is None:
                existing_capture_semantics = {
                    **requested_capture_semantics,
                    "backend": "openxml",
                    "captureContract": CAPTURE_CONTRACT,
                }
            comparable_existing_capture = {
                key: existing_capture_semantics.get(key)
                for key in requested_capture_semantics
            }
            if comparable_existing_capture != requested_capture_semantics:
                raise IncrementalIngestError(
                    "Existing workflow journal capture contract does not "
                    "match this request."
                )
            _reset_journal_for_retry(
                journal,
                started_at=now_iso(),
            )
        else:
            journal = {
                "schemaVersion": WORKFLOW_SCHEMA_VERSION,
                "runId": run_id,
                "attempt": 1,
                "dataset": dataset,
                "databasePath": str(database),
                "artifactDirectory": str(run_dir),
                "source": {
                    "sourcePath": str(source_file),
                    "fileName": source_file.name,
                    "contentSha256": source_sha256,
                    "imagesAnalyzed": False,
                },
                "semanticContracts": {
                    "capture": copy.deepcopy(
                        requested_capture_semantics
                    ),
                    "locatorPromptVersion": LOCATOR_PROMPT_VERSION,
                    "batchLocatorPromptVersion": (
                        BATCH_LOCATOR_PROMPT_VERSION
                    ),
                    "studyDraftPromptVersion": (
                        STUDY_DRAFT_PROMPT_VERSION
                    ),
                    "formulaDerivation": copy.deepcopy(
                        requested_formula_contract
                    ),
                },
                "status": "PENDING",
                "currentStage": "",
                "stages": {
                    name: {"status": "PENDING"}
                    for name in STAGE_ORDER
                },
                "createdAt": now_iso(),
            }
        journal["semanticContracts"] = {
            "capture": copy.deepcopy(requested_capture_semantics),
            "locatorPromptVersion": LOCATOR_PROMPT_VERSION,
            "batchLocatorPromptVersion": BATCH_LOCATOR_PROMPT_VERSION,
            "studyDraftPromptVersion": STUDY_DRAFT_PROMPT_VERSION,
            "formulaDerivation": copy.deepcopy(
                requested_formula_contract
            ),
        }
        journal["status"] = "RUNNING"
        journal["updatedAt"] = now_iso()
        journal["finishedAt"] = ""
        journal["result"] = None
        _atomic_write_json(journal_path, journal)

        def execute_stage(
            name: str,
            action: Callable[[], Any],
            summarize: Callable[[Any], dict[str, Any]],
            *,
            gate: str | None = None,
        ) -> Any:
            _invalidate_downstream_journal_stages(journal, name)
            journal["currentStage"] = name
            journal["stages"][name] = {
                "status": "RUNNING",
                "startedAt": now_iso(),
            }
            journal["updatedAt"] = now_iso()
            _atomic_write_json(journal_path, journal)
            emit_progress(name, "RUNNING")
            try:
                with (
                    pipeline_scope(gate)
                    if gate is not None
                    else nullcontext()
                ):
                    value = action()
            except Exception as exc:
                failed_at = now_iso()
                journal["status"] = "FAILED"
                journal["stages"][name] = {
                    **journal["stages"][name],
                    "status": "FAILED",
                    "finishedAt": failed_at,
                    "error": f"{type(exc).__name__}: {exc}",
                }
                journal["result"] = None
                journal["updatedAt"] = failed_at
                journal["finishedAt"] = failed_at
                _atomic_write_json(journal_path, journal)
                emit_progress(
                    name,
                    "FAILED",
                    f"{type(exc).__name__}: {exc}",
                )
                raise
            journal["stages"][name] = {
                **journal["stages"][name],
                "status": "COMPLETED",
                "finishedAt": now_iso(),
                "result": summarize(value),
            }
            journal["updatedAt"] = now_iso()
            _atomic_write_json(journal_path, journal)
            emit_progress(name, "COMPLETED")
            return value

        def capture_action() -> dict[str, Any]:
            with pipeline_scope("DB"):
                with _connect_rw(database) as connection:
                    ensure_knowledge_schema(connection, now_iso)
                    ensure_capture_v2_schema(connection)
                    existing = connection.execute(
                        """
                        SELECT
                            capture.revision_id AS capture_revision_id,
                            capture.revision_uid,
                            capture.document_id AS capture_document_id,
                            workbook.workbook_status,
                            canonical.revision_id AS canonical_revision_id,
                            document.document_id AS canonical_document_id,
                            document.document_uid
                        FROM capture_v2_documents capture_document
                        JOIN capture_v2_revisions capture
                          ON capture.document_id=capture_document.document_id
                         AND capture.is_current=1
                        JOIN capture_v2_workbooks workbook
                          ON workbook.revision_id=capture.revision_id
                        JOIN source_revisions canonical
                          ON canonical.capture_v2_revision_id=capture.revision_id
                         AND canonical.is_current=1
                        JOIN source_documents document
                          ON document.document_id=canonical.document_id
                        WHERE capture_document.source_path=?
                          AND capture.content_sha256=?
                          AND capture.capture_contract=?
                          AND document.dataset=?
                          AND document.source_path=?
                        """,
                        (
                            str(source_file),
                            source_sha256,
                            requested_capture_contract,
                            dataset,
                            str(source_file),
                        ),
                    ).fetchone()
                    if existing is not None:
                        repair = reconcile_capture_sheet_counts(
                            connection,
                            int(existing["capture_revision_id"]),
                        )
                        repaired_sheet_metadata = bool(
                            repair["repairedSheetCount"]
                        )
                        verification = verify_capture_revision(
                            connection,
                            int(existing["capture_revision_id"]),
                            verify_source_sha256=True,
                        )
                        if verification["ok"] and repaired_sheet_metadata:
                            connection.commit()
                        if not verification["ok"]:
                            raise IncrementalIngestError(
                                "Existing Capture v2 revision verification "
                                "failed: "
                                + "; ".join(verification["errors"])
                            )
                        return {
                            "payload": None,
                            "capture": {
                                "action": (
                                    "REPAIRED_CURRENT"
                                    if repaired_sheet_metadata
                                    else "REUSED_CURRENT"
                                ),
                                "documentId": int(
                                    existing["capture_document_id"]
                                ),
                                "revisionId": int(
                                    existing["capture_revision_id"]
                                ),
                                "contentSha256": source_sha256,
                                "captureContract": requested_capture_contract,
                            },
                            "bridge": {
                                "documentId": int(
                                    existing["canonical_document_id"]
                                ),
                                "documentUid": str(existing["document_uid"]),
                                "revisionId": int(
                                    existing["canonical_revision_id"]
                                ),
                                "revisionUid": str(existing["revision_uid"]),
                                "captureV2RevisionId": int(
                                    existing["capture_revision_id"]
                                ),
                                "workbookStatus": str(
                                    existing["workbook_status"]
                                ),
                            },
                            "verification": verification,
                        }

            if normalized_capture_backend == "com":
                from inference_data_ai_com_capture import (
                    extract_workbook_com,
                )

                with pipeline_scope("COM"):
                    payload = extract_workbook_com(
                        source_file,
                        covered_cell_mode=covered_cell_mode,
                        include_hidden=include_hidden_sheets,
                        inspect_auth_dialog=inspect_auth_dialog,
                        dismiss_auth_dialog=dismiss_auth_dialog,
                        auth_dialog_title=auth_dialog_title,
                        auth_dialog_class=auth_dialog_class,
                        auth_dialog_button=auth_dialog_button,
                        auth_dialog_timeout_seconds=(
                            auth_dialog_timeout_seconds
                        ),
                    )
            else:
                with pipeline_scope("PACKET"):
                    payload = extract_workbook(source_file)
            if payload["source"]["contentSha256"] != source_sha256:
                raise IncrementalIngestError(
                    "Source changed while Capture v2 was reading it."
                )

            with pipeline_scope("DB"):
                with _connect_rw(database) as connection:
                    ensure_knowledge_schema(connection, now_iso)
                    ensure_capture_v2_schema(connection)
                    captured = import_capture(
                        connection,
                        payload,
                        captured_at=now_iso(),
                    )
                    bridge = bridge_capture_to_canonical_source(
                        connection,
                        dataset=dataset,
                        payload=payload,
                        capture_result=captured,
                        captured_at=now_iso(),
                    )
                    repair = reconcile_capture_sheet_counts(
                        connection,
                        int(captured["revisionId"]),
                    )
                    if repair["repairedSheetCount"]:
                        captured["action"] = "REPAIRED_CURRENT"
                    verification = verify_capture_revision(
                        connection,
                        int(captured["revisionId"]),
                        verify_source_sha256=True,
                    )
                    if not verification["ok"]:
                        raise IncrementalIngestError(
                            "Capture v2 verification failed: "
                            + "; ".join(verification["errors"])
                        )
                    connection.commit()
            return {
                "payload": payload,
                "capture": captured,
                "bridge": bridge,
                "verification": verification,
            }

        capture_state = execute_stage(
            "CAPTURE",
            capture_action,
            lambda value: {
                "action": value["capture"]["action"],
                "captureRevisionId": value["capture"]["revisionId"],
                "canonicalRevisionId": value["bridge"]["revisionId"],
                "revisionUid": value["bridge"]["revisionUid"],
                "workbookStatus": value["bridge"]["workbookStatus"],
                "verified": value["verification"]["ok"],
            },
        )

        formula_overlay: dict[str, Any] | None = None

        def packet_action() -> dict[str, Any]:
            nonlocal formula_overlay
            packet_set = build_semantic_source_packets_from_db(
                database,
                revision_id=int(capture_state["capture"]["revisionId"]),
                max_cells=max_cells,
                max_rows=max_rows,
                empty_row_gap=empty_row_gap,
            )
            if derive_formula_values and packet_set["chunks"]:
                derived_overlay = derive_formula_overlay(
                    packet_set["chunks"],
                    expected_revision_uid=str(
                        packet_set["inventory"]["sourceRevision"][
                            "revisionUid"
                        ]
                    ),
                    expected_content_sha256=source_sha256,
                )
                if resume and formula_overlay_path.is_file():
                    try:
                        existing_overlay = json.loads(
                            formula_overlay_path.read_text(
                                encoding="utf-8"
                            )
                        )
                    except (OSError, json.JSONDecodeError) as exc:
                        raise IncrementalIngestError(
                            "Saved formula overlay is unreadable; resume "
                            "cannot safely reuse downstream artifacts."
                        ) from exc
                    validate_formula_overlay(
                        packet_set["chunks"],
                        existing_overlay,
                    )
                    if existing_overlay != derived_overlay:
                        raise IncrementalIngestError(
                            "Saved formula overlay differs from fresh "
                            "deterministic derivation."
                        )
                    formula_overlay = existing_overlay
                else:
                    _atomic_write_json(
                        formula_overlay_path,
                        derived_overlay,
                    )
                    try:
                        formula_overlay = json.loads(
                            formula_overlay_path.read_text(
                                encoding="utf-8"
                            )
                        )
                    except (OSError, json.JSONDecodeError) as exc:
                        raise IncrementalIngestError(
                            "Formula overlay could not be verified after "
                            "its atomic write."
                        ) from exc
                    validate_formula_overlay(
                        packet_set["chunks"],
                        formula_overlay,
                    )
                packet_set["chunks"] = apply_formula_overlay_to_chunks(
                    packet_set["chunks"],
                    formula_overlay,
                )
                unresolved_formula_count = sum(
                    1
                    for chunk in packet_set["chunks"]
                    for cell in chunk.get("cells", [])
                    if str(cell.get("formula") or "")
                    and cell.get("cachedValue") in (None, "")
                    and cell.get("displayValue") in (None, "")
                )
                if unresolved_formula_count:
                    raise IncrementalIngestError(
                        "Formula projection left unresolved source formulas: "
                        f"{unresolved_formula_count}"
                    )
                formula_summary = {
                    "enabled": True,
                    "schemaVersion": formula_overlay["schemaVersion"],
                    "evaluatorVersion": formula_overlay[
                        "evaluatorVersion"
                    ],
                    "overlaySha256": formula_overlay["overlaySha256"],
                    "overlayPath": str(formula_overlay_path),
                    "formulaCount": formula_overlay["formulaCount"],
                    "numericCount": formula_overlay["numericCount"],
                    "errorCount": formula_overlay["errorCount"],
                    "errorsByCode": formula_overlay["errorsByCode"],
                    "unresolvedFormulaCellCount": 0,
                    "captureMutated": False,
                }
            else:
                formula_summary = {
                    **requested_formula_contract,
                    "overlaySha256": "",
                    "overlayPath": "",
                    "formulaCount": 0,
                    "numericCount": 0,
                    "errorCount": 0,
                    "errorsByCode": {},
                    "unresolvedFormulaCellCount": 0,
                    "captureMutated": False,
                }
            packet_set["inventory"][
                "formulaDerivation"
            ] = copy.deepcopy(formula_summary)
            packet_set["formulaDerivation"] = copy.deepcopy(
                formula_summary
            )
            _atomic_write_bytes(packet_path, packet_json_bytes(packet_set))
            return packet_set

        packet_set = execute_stage(
            "PACKET",
            packet_action,
            lambda value: {
                "packetPath": str(packet_path),
                "workbookStatus": value["inventory"]["workbook"]["status"],
                "chunks": len(value["chunks"]),
                "terminalPackets": len(value["terminalPackets"]),
                "coverage": value["inventory"]["coverage"],
                "formulaDerivation": value["inventory"][
                    "formulaDerivation"
                ],
            },
            gate="PACKET",
        )
        journal["semanticContracts"]["formulaDerivation"] = copy.deepcopy(
            packet_set["inventory"]["formulaDerivation"]
        )
        journal["updatedAt"] = now_iso()
        _atomic_write_json(journal_path, journal)
        source = _source_identity(packet_set, dataset)
        workbook = _workbook_summary(packet_set)
        workbook_status = str(packet_set["inventory"]["workbook"]["status"])

        def locator_action() -> dict[str, Any]:
            locator_dir.mkdir(parents=True, exist_ok=True)
            if workbook_status in TERMINAL_WORKBOOK_STATUSES:
                return {
                    "results": [],
                    "aiCalls": 0,
                    "aiSubmitted": 0,
                    "deterministic": 0,
                    "skipped": 0,
                }
            jobs: list[tuple[dict[str, Any], Path]] = []
            skipped = 0
            deterministic = 0
            for chunk in packet_set["chunks"]:
                output = (
                    locator_dir
                    / f"{_safe_name(chunk['chunkId'])}.locator.json"
                )
                if resume and output.is_file():
                    try:
                        existing = json.loads(output.read_text(encoding="utf-8"))
                        validate_locator_result(
                            existing,
                            revision_uid=source["revisionUid"],
                            content_sha256=source["contentSha256"],
                            chunk=chunk,
                        )
                        skipped += 1
                        continue
                    except (OSError, ValueError, RuntimeError, json.JSONDecodeError):
                        pass
                if not _has_primary_text(chunk):
                    result = _deterministic_locator(source, chunk)
                    validate_locator_result(
                        result,
                        revision_uid=source["revisionUid"],
                        content_sha256=source["contentSha256"],
                        chunk=chunk,
                    )
                    _atomic_write_json(output, result)
                    deterministic += 1
                    continue
                jobs.append((chunk, output))
            batches = _partition_jobs(
                jobs,
                batch_size=locator_batch_size,
                batch_max_bytes=locator_batch_max_bytes,
            )

            def run_batch(
                batch: list[tuple[dict[str, Any], Path]],
            ) -> list[dict[str, Any]]:
                with pipeline_scope("AI"):
                    results = locator_batch_runner(
                        source=source,
                        workbook=workbook,
                        chunks=[chunk for chunk, _ in batch],
                        output_paths={
                            str(chunk["chunkId"]): output
                            for chunk, output in batch
                        },
                        model=model,
                        reasoning_effort=reasoning_effort,
                        timeout_seconds=locator_timeout_seconds,
                    )
                if not isinstance(results, list) or len(results) != len(
                    batch
                ):
                    raise IncrementalIngestError(
                        "Locator batch runner returned incomplete results."
                    )
                validated: list[dict[str, Any]] = []
                for result, (chunk, _output) in zip(
                    results,
                    batch,
                    strict=True,
                ):
                    validated.append(
                        validate_locator_result(
                            result,
                            revision_uid=source["revisionUid"],
                            content_sha256=source["contentSha256"],
                            chunk=chunk,
                        )
                    )
                return validated

            failures: list[str] = []
            ai_calls = 0
            with ThreadPoolExecutor(
                max_workers=max(1, locator_workers)
            ) as executor:
                futures = {
                    executor.submit(
                        _run_locator_batch_with_singleton_retry,
                        batch,
                        run_batch,
                    ): batch
                    for batch in batches
                }
                for future in as_completed(futures):
                    try:
                        _results, calls = future.result()
                        ai_calls += calls
                    except Exception as exc:
                        chunk_ids = ", ".join(
                            str(chunk["chunkId"])
                            for chunk, _ in futures[future]
                        )
                        failures.append(f"{chunk_ids}: {exc}")
            if failures:
                raise IncrementalIngestError(
                    "Semantic locator failed: " + " | ".join(failures)
                )
            results = _load_locator_results(
                packet_set,
                locator_dir,
                source,
            )
            return {
                "results": results,
                "aiCalls": ai_calls,
                "aiSubmitted": len(jobs),
                "deterministic": deterministic,
                "skipped": skipped,
            }

        locator_state = execute_stage(
            "LOCATOR",
            locator_action,
            lambda value: {
                "locatorDirectory": str(locator_dir),
                "results": len(value["results"]),
                "aiCalls": value["aiCalls"],
                "aiSubmitted": value["aiSubmitted"],
                "deterministicNoCandidate": value["deterministic"],
                "skipped": value["skipped"],
            },
        )

        def validate_manifest_content(
            manifest: dict[str, Any],
            *,
            content_complete: bool,
            coverage_chunks: Sequence[dict[str, Any]] | None = None,
            coverage_locator_results: (
                Sequence[dict[str, Any]] | None
            ) = None,
            expected_source_cell_keys: Sequence[str] | None = None,
            require_source_content_coverage: bool = False,
            precomputed_coverage_inventory: (
                dict[str, Any] | None
            ) = None,
        ) -> dict[str, Any]:
            coverage_inventory = precomputed_coverage_inventory
            manifest_to_validate = manifest
            if require_source_content_coverage:
                if (
                    coverage_chunks is None
                    or coverage_locator_results is None
                ):
                    raise IncrementalIngestError(
                        "Complete source-content coverage requires "
                        "focused chunks and locator results."
                    )
                if coverage_inventory is None:
                    coverage_inventory = (
                        build_content_coverage_inventory(
                            chunks=coverage_chunks,
                            locator_results=(
                                coverage_locator_results
                            ),
                            expected_source_cell_keys=(
                                expected_source_cell_keys
                            ),
                        )
                    )
                manifest_to_validate = (
                    augment_exact_source_conclusions(
                        manifest=manifest,
                        inventory=coverage_inventory,
                    )
                )
            with _connect_rw(database) as connection:
                revision = resolve_manifest_revision(connection, source)
                checker = make_database_evidence_checker(
                    connection,
                    revision,
                )
                validated = validate_ai_study_draft(
                    manifest_to_validate,
                    source=source,
                    content_complete=content_complete,
                    evidence_checker=checker,
                )
                validate_numeric_observation_evidence(
                    connection,
                    revision,
                    validated,
                    formula_overlay=formula_overlay,
                )
                validate_factor_and_arm_evidence(
                    connection,
                    revision,
                    validated,
                )
                validate_comparison_representation_alignment(
                    connection,
                    revision,
                    validated,
                    formula_overlay=formula_overlay,
                )
                validate_conclusion_evidence(
                    connection,
                    revision,
                    validated,
                )
                if coverage_inventory is not None:
                    validate_content_manifest_coverage(
                        manifest=validated,
                        inventory=coverage_inventory,
                        require_complete=True,
                    )
                return validated

        def validate_manifest(
            manifest: dict[str, Any],
            *,
            coverage_chunks: Sequence[dict[str, Any]] | None = None,
            coverage_locator_results: (
                Sequence[dict[str, Any]] | None
            ) = None,
            expected_source_cell_keys: Sequence[str] | None = None,
            require_source_content_coverage: bool = False,
            precomputed_coverage_inventory: (
                dict[str, Any] | None
            ) = None,
        ) -> dict[str, Any]:
            return validate_manifest_content(
                manifest,
                content_complete=bool(
                    packet_set["inventory"][
                        "contentCompleteForManifest"
                    ]
                ),
                coverage_chunks=coverage_chunks,
                coverage_locator_results=coverage_locator_results,
                expected_source_cell_keys=expected_source_cell_keys,
                require_source_content_coverage=(
                    require_source_content_coverage
                ),
                precomputed_coverage_inventory=(
                    precomputed_coverage_inventory
                ),
            )

        draft_provenance = {
            "aiExecutedThisAttempt": False,
            "artifactReused": False,
            "reusedPromptVersion": "",
            "staged": False,
            "mode": "",
            "partCount": 0,
            "partsExecuted": 0,
            "partsReused": 0,
            "requestPromptBytes": 0,
            "requestPromptSha256": "",
            "formulaDerivation": copy.deepcopy(
                packet_set["inventory"]["formulaDerivation"]
            ),
            "repairRejectedDraft": False,
            "requiredSourceLocatorPromotions": [],
            "invalidatedAnalyses": {
                "quarantined": [],
                "alreadyStale": [],
                "protected": [],
            },
        }

        def draft_action() -> dict[str, Any]:
            if workbook_status in TERMINAL_WORKBOOK_STATUSES:
                manifest = _terminal_manifest(packet_set, dataset)
                _atomic_write_json(manifest_path, manifest)
                return manifest
            locator_results = locator_state["results"]
            candidate_ids = {
                str(result["chunkId"])
                for result in locator_results
                if result["status"] in {"CANDIDATES", "NEEDS_REVIEW"}
                and result["candidates"]
            }
            if not candidate_ids:
                no_candidate_inventory = (
                    audit_no_candidate_source_inventory(
                        packet_set=packet_set,
                        locator_results=locator_results,
                    )
                )
                if no_candidate_inventory["requiredCells"]:
                    preview = ", ".join(
                        f"{item['sheet']}!{item['coordinate']}"
                        for item in no_candidate_inventory[
                            "requiredCells"
                        ][:20]
                    )
                    raise IncrementalIngestError(
                        "Locator coverage is unsafe: every chunk returned "
                        "NO_CANDIDATE while quantitative/formula source "
                        f"cells remain ({preview})."
                    )
                manifest = _terminal_manifest(
                    packet_set,
                    dataset,
                    status_override="NO_SEMANTIC_CANDIDATE",
                    extra_limitation=(
                        "Every source chunk completed semantic location without "
                        "a source-backed Study candidate."
                    ),
                )
                _atomic_write_json(manifest_path, manifest)
                return manifest

            universe = select_draft_universe(
                packet_set=packet_set,
                locator_results=locator_results,
            )
            selected_chunks = list(universe["selectedChunks"])
            selected_locators = list(
                universe["selectedLocatorResults"]
            )
            selected_source_cell_keys = list(
                universe["ownedSourceCellKeys"]
            )
            unselected_inventory = audit_unselected_source_inventory(
                packet_set=packet_set,
                locator_results=locator_results,
                selected_source_cell_keys=selected_source_cell_keys,
            )
            if (
                unselected_inventory["requiredCells"]
                and repair_unselected_source
            ):
                locator_results = (
                    promote_required_source_locator_sections(
                        locator_results=locator_results,
                        required_cells=unselected_inventory[
                            "requiredCells"
                        ],
                    )
                )
                locator_state["results"] = locator_results
                draft_provenance[
                    "requiredSourceLocatorPromotions"
                ] = [
                    copy.deepcopy(
                        result["deterministicCoveragePromotion"]
                    )
                    for result in locator_results
                    if isinstance(
                        result.get(
                            "deterministicCoveragePromotion"
                        ),
                        dict,
                    )
                ]
                universe = select_draft_universe(
                    packet_set=packet_set,
                    locator_results=locator_results,
                )
                selected_chunks = list(universe["selectedChunks"])
                selected_locators = list(
                    universe["selectedLocatorResults"]
                )
                selected_source_cell_keys = list(
                    universe["ownedSourceCellKeys"]
                )
                unselected_inventory = (
                    audit_unselected_source_inventory(
                        packet_set=packet_set,
                        locator_results=locator_results,
                        selected_source_cell_keys=(
                            selected_source_cell_keys
                        ),
                    )
                )
            if unselected_inventory["requiredCells"]:
                preview = ", ".join(
                    f"{item['sheet']}!{item['coordinate']}"
                    for item in unselected_inventory[
                        "requiredCells"
                    ][:20]
                )
                raise IncrementalIngestError(
                    "Candidate selection is incomplete: quantitative, "
                    "formula, or other required source results exist "
                    f"outside candidate-bearing sections ({preview})."
                )
            monolithic_prompt = build_study_draft_prompt(
                source=source,
                workbook=workbook,
                locator_results=selected_locators,
                focused_chunks=selected_chunks,
            )
            monolithic_request = build_monolithic_request(
                source=source,
                workbook=workbook,
                universe=universe,
                content_complete=bool(
                    packet_set["inventory"][
                        "contentCompleteForManifest"
                    ]
                ),
                prompt_text=monolithic_prompt,
            )
            budget = assess_one_call_budget(
                request=monolithic_request,
                max_prompt_bytes=draft_monolithic_max_bytes,
                max_source_cells=draft_fragment_max_cells,
            )
            draft_provenance["mode"] = budget["mode"]
            draft_provenance["requestPromptBytes"] = budget[
                "promptBytes"
            ]
            draft_provenance["requestSourceCellCount"] = budget[
                "sourceCellCount"
            ]
            draft_provenance["requestPromptSha256"] = budget[
                "promptSha256"
            ]
            registry: dict[str, Any] | None = None
            draft_plan: dict[str, Any] | None = None
            if budget["mode"] == "STAGED_V2":
                registry = build_study_registry_v2(
                    source=source,
                    universe=universe,
                )
                draft_plan = plan_study_draft_v2(
                    source=source,
                    workbook=workbook,
                    universe=universe,
                    registry=registry,
                    prompt_version=STUDY_DRAFT_PROMPT_VERSION,
                    max_chunks=draft_fragment_max_chunks,
                    max_cells=draft_fragment_max_cells,
                    max_serialized_bytes=draft_fragment_max_bytes,
                )

            if resume and manifest_path.is_file():
                try:
                    existing = json.loads(
                        manifest_path.read_text(encoding="utf-8")
                    )
                    existing = validate_manifest(
                        existing,
                        coverage_chunks=selected_chunks,
                        coverage_locator_results=selected_locators,
                        expected_source_cell_keys=(
                            selected_source_cell_keys
                        ),
                        require_source_content_coverage=bool(
                            existing.get("source", {}).get(
                                "contentComplete"
                            )
                        ),
                    )
                    if budget["mode"] == "MONOLITHIC":
                        matching_provenance = (
                            _draft_provenance_matches(
                                draft_provenance_path,
                                manifest_path,
                                source,
                                expected_details={
                                    "mode": "MONOLITHIC",
                                    "promptSha256": budget[
                                        "promptSha256"
                                    ],
                                    "envelopeSha256": budget[
                                        "envelopeSha256"
                                    ],
                                },
                            )
                        )
                    else:
                        if registry is None or draft_plan is None:
                            raise ValueError(
                                "Staged v2 resume lacks registry or plan"
                            )
                        staged_provenance = json.loads(
                            staged_final_provenance_path.read_text(
                                encoding="utf-8"
                            )
                        )
                        current_part_hashes = []
                        verified_part_fragments: list[
                            tuple[dict[str, Any], dict[str, Any]]
                        ] = []
                        for part in draft_plan["parts"]:
                            part_path, part_provenance_path = (
                                fragment_artifact_paths(run_dir, part)
                            )
                            if (
                                not part_path.is_file()
                                or not part_provenance_path.is_file()
                            ):
                                raise ValueError(
                                    "Staged v2 final references a missing part"
                                )
                            focused_chunks = chunks_for_part_v2(
                                universe,
                                part,
                            )
                            focused_locators = locators_for_part_v2(
                                universe,
                                part,
                            )
                            resume_envelope = finalize_fragment_envelope(
                                build_fragment_envelope(
                                    source=source,
                                    workbook=workbook,
                                    plan=draft_plan,
                                    part=part,
                                    focused_chunks=focused_chunks,
                                    locator_results=focused_locators,
                                    registry_slice=registry_for_part(
                                        registry,
                                        part,
                                    ),
                                )
                            )
                            part_output_sha256 = _file_sha256(
                                part_path
                            )
                            part_provenance = json.loads(
                                part_provenance_path.read_text(
                                    encoding="utf-8"
                                )
                            )
                            if not part_provenance_v2_matches(
                                provenance=part_provenance,
                                plan=draft_plan,
                                part=part,
                                envelope=resume_envelope,
                                output_sha256=part_output_sha256,
                                output_path=part_path,
                            ):
                                raise ValueError(
                                    "Staged v2 final contains an "
                                    "unverified part."
                                )
                            verified_fragment = validate_fragment_v2(
                                fragment=json.loads(
                                    part_path.read_text(
                                        encoding="utf-8"
                                    )
                                ),
                                envelope=resume_envelope,
                                all_selected_chunks=resume_envelope[
                                    "focusedChunks"
                                ],
                            )
                            verified_part_fragments.append(
                                (part, verified_fragment)
                            )
                            current_part_hashes.append(
                                {
                                    "partId": part["partId"],
                                    "outputSha256": part_output_sha256,
                                }
                            )
                        reconstructed_merged = merge_fragment_records(
                            plan=draft_plan,
                            fragments=verified_part_fragments,
                            selected_chunks=selected_chunks,
                        )
                        if not staged_merged_path.is_file():
                            raise ValueError(
                                "Staged v2 merged artifact is missing."
                            )
                        stored_merged = json.loads(
                            staged_merged_path.read_text(
                                encoding="utf-8"
                            )
                        )
                        if stored_merged != reconstructed_merged:
                            raise ValueError(
                                "Staged v2 merged artifact differs from "
                                "verified part outputs."
                            )
                        merged_output_sha256 = _file_sha256(
                            staged_merged_path
                        )
                        matching_provenance = bool(
                            final_provenance_v2_matches(
                                provenance=staged_provenance,
                                plan=draft_plan,
                                registry=registry,
                                final_sha256=_file_sha256(manifest_path),
                                ordered_part_hashes=current_part_hashes,
                                merged_path=staged_merged_path,
                                merged_sha256=merged_output_sha256,
                                final_path=manifest_path,
                            )
                        )
                    if not matching_provenance:
                        raise ValueError(
                            "Existing Study draft provenance does not match "
                            "the exact v2 request and contracts."
                        )
                    draft_provenance["artifactReused"] = True
                    draft_provenance["reusedPromptVersion"] = (
                        STUDY_DRAFT_PROMPT_VERSION
                    )
                    draft_provenance["staged"] = (
                        budget["mode"] == "STAGED_V2"
                    )
                    if draft_plan is not None:
                        draft_provenance["partCount"] = len(
                            draft_plan["parts"]
                        )
                    return existing
                except (
                    OSError,
                    ValueError,
                    RuntimeError,
                    json.JSONDecodeError,
                ) as exc:
                    invalidation = (
                        _quarantine_invalidated_unverified_analyses(
                            database,
                            source=source,
                            reason=(
                                "Existing canonical Study draft failed "
                                f"{STUDY_DRAFT_PROMPT_VERSION} reuse "
                                "validation and was fail-closed before "
                                f"regeneration: {type(exc).__name__}: "
                                f"{str(exc)[:500]}"
                            ),
                            now_iso=now_iso,
                        )
                    )
                    draft_provenance[
                        "invalidatedAnalyses"
                    ] = invalidation
                    journal.setdefault("safetyActions", {})[
                        "invalidatedDraftAnalyses"
                    ] = copy.deepcopy(invalidation)
                    journal["updatedAt"] = now_iso()
                    _atomic_write_json(journal_path, journal)

            def validate_canonical_claims(
                connection: sqlite3.Connection,
                revision: sqlite3.Row,
                draft: dict[str, Any],
            ) -> None:
                inventory = build_content_coverage_inventory(
                    chunks=selected_chunks,
                    locator_results=selected_locators,
                    expected_source_cell_keys=(
                        selected_source_cell_keys
                    ),
                )
                augmented = augment_exact_source_conclusions(
                    manifest=draft,
                    inventory=inventory,
                )
                draft.clear()
                draft.update(augmented)
                validate_numeric_observation_evidence(
                    connection,
                    revision,
                    draft,
                    formula_overlay=formula_overlay,
                )
                validate_factor_and_arm_evidence(
                    connection,
                    revision,
                    draft,
                )
                validate_comparison_representation_alignment(
                    connection,
                    revision,
                    draft,
                    formula_overlay=formula_overlay,
                )
                validate_conclusion_evidence(
                    connection,
                    revision,
                    draft,
                )
                validate_content_manifest_coverage(
                    manifest=draft,
                    inventory=inventory,
                    require_complete=True,
                )

            staged_part_hashes: list[dict[str, str]] = []
            if budget["mode"] == "MONOLITHIC":
                monolithic_coverage_inventory = (
                    build_content_coverage_inventory(
                        chunks=selected_chunks,
                        locator_results=selected_locators,
                        expected_source_cell_keys=(
                            selected_source_cell_keys
                        ),
                    )
                )
                with _connect_rw(database) as connection:
                    revision = resolve_manifest_revision(connection, source)
                    checker = make_database_evidence_checker(
                        connection,
                        revision,
                    )

                    def source_claim_validator(
                        draft: dict[str, Any],
                    ) -> None:
                        validate_canonical_claims(
                            connection,
                            revision,
                            draft,
                        )
                        validate_content_manifest_coverage(
                            manifest=draft,
                            inventory=(
                                monolithic_coverage_inventory
                            ),
                            require_complete=True,
                        )

                    if draft_runner is not run_codex_study_draft:
                        # Injected runners in tests or integrations represent
                        # an executed draft pass unless they explicitly
                        # replace this workflow contract.
                        draft_provenance["aiExecutedThisAttempt"] = True

                    def mark_study_ai_call() -> None:
                        draft_provenance[
                            "aiExecutedThisAttempt"
                        ] = True

                    def rate_pair_repair_paths(
                        draft: dict[str, Any],
                    ) -> list[tuple[int, int, int]]:
                        return unsupported_rate_pair_observation_paths(
                            connection,
                            revision,
                            draft,
                            formula_overlay=formula_overlay,
                        )

                    rebuilt_prompt = build_study_draft_prompt(
                        source=source,
                        workbook=workbook,
                        locator_results=selected_locators,
                        focused_chunks=selected_chunks,
                    )
                    if (
                        hashlib.sha256(
                            rebuilt_prompt.encode("utf-8")
                        ).hexdigest()
                        != budget["promptSha256"]
                    ):
                        raise IncrementalIngestError(
                            "Runner input differs from the budgeted "
                            "monolithic request."
                        )
                    rejected_manifest_path = manifest_path.with_name(
                        manifest_path.stem
                        + ".rejected"
                        + manifest_path.suffix
                    )
                    repair_current_rejected = bool(
                        repair_rejected_draft
                        and rejected_manifest_path.is_file()
                    )
                    draft_provenance["repairRejectedDraft"] = (
                        repair_current_rejected
                    )
                    with pipeline_scope("AI"):
                        manifest = draft_runner(
                            source=source,
                            workbook=workbook,
                            locator_results=selected_locators,
                            focused_chunks=selected_chunks,
                            content_complete=bool(
                                packet_set["inventory"][
                                    "contentCompleteForManifest"
                                ]
                            ),
                            output_path=manifest_path,
                            evidence_checker=checker,
                            additional_validator=source_claim_validator,
                            model=model,
                            reasoning_effort=reasoning_effort,
                            timeout_seconds=draft_timeout_seconds,
                            exact_prompt_text=(
                                None
                                if repair_current_rejected
                                else rebuilt_prompt
                            ),
                            expected_prompt_sha256=(
                                None
                                if repair_current_rejected
                                else budget["promptSha256"]
                            ),
                            ai_call_observer=mark_study_ai_call,
                            unsupported_rate_pair_paths=(
                                rate_pair_repair_paths
                                if draft_runner
                                is run_codex_study_draft
                                else None
                            ),
                        )
                manifest = validate_manifest(
                    manifest,
                    coverage_chunks=selected_chunks,
                    coverage_locator_results=selected_locators,
                    expected_source_cell_keys=(
                        selected_source_cell_keys
                    ),
                    require_source_content_coverage=True,
                )
            else:
                if registry is None or draft_plan is None:
                    raise IncrementalIngestError(
                        "Staged v2 requires a registry and exact plan."
                    )
                draft_provenance["staged"] = True
                _atomic_write_json(draft_plan_path, draft_plan)
                _atomic_write_json(draft_registry_path, registry)
                draft_parts = list(draft_plan["parts"])
                if not draft_parts:
                    raise IncrementalIngestError(
                        "Staged v2 produced no fragment parts."
                    )
                draft_provenance["partCount"] = len(draft_parts)
                completed_by_part: dict[
                    str,
                    tuple[dict[str, Any], str],
                ] = {}
                preflight_jobs: list[
                    tuple[
                        dict[str, Any],
                        dict[str, Any],
                        Path,
                        Path,
                    ]
                ] = []
                for part in draft_parts:
                    focused_chunks = chunks_for_part_v2(
                        universe,
                        part,
                    )
                    focused_locators = locators_for_part_v2(
                        universe,
                        part,
                    )
                    registry_slice = registry_for_part(
                        registry,
                        part,
                    )
                    envelope = finalize_fragment_envelope(
                        build_fragment_envelope(
                            source=source,
                            workbook=workbook,
                            plan=draft_plan,
                            part=part,
                            focused_chunks=focused_chunks,
                            locator_results=focused_locators,
                            registry_slice=registry_slice,
                        )
                    )
                    if (
                        envelope["promptBytes"]
                        != part.get("promptBytes")
                        or envelope["promptBytes"]
                        > draft_fragment_max_bytes
                    ):
                        raise IncrementalIngestError(
                            f"Fragment {part['partId']} exact prompt "
                            "preflight differs from its deterministic plan "
                            "or exceeds draft_fragment_max_bytes."
                        )
                    part_fragment_path, part_provenance_path = (
                        fragment_artifact_paths(run_dir, part)
                    )
                    preflight_jobs.append(
                        (
                            part,
                            envelope,
                            part_fragment_path,
                            part_provenance_path,
                        )
                    )
                pending_jobs: list[
                    tuple[
                        dict[str, Any],
                        dict[str, Any],
                        Path,
                        Path,
                    ]
                ] = []
                for (
                    part,
                    envelope,
                    part_fragment_path,
                    part_provenance_path,
                ) in preflight_jobs:
                    part_fragment_path.parent.mkdir(
                        parents=True,
                        exist_ok=True,
                    )
                    if (
                        resume
                        and part_fragment_path.is_file()
                        and part_provenance_path.is_file()
                    ):
                        try:
                            existing_fragment = json.loads(
                                part_fragment_path.read_text(
                                    encoding="utf-8"
                                )
                            )
                            existing_provenance = json.loads(
                                part_provenance_path.read_text(
                                    encoding="utf-8"
                                )
                            )
                            output_sha256 = _file_sha256(
                                part_fragment_path
                            )
                            if not part_provenance_v2_matches(
                                provenance=existing_provenance,
                                plan=draft_plan,
                                part=part,
                                envelope=envelope,
                                output_sha256=output_sha256,
                                output_path=part_fragment_path,
                            ):
                                raise ValueError(
                                    "Fragment v2 provenance mismatch."
                                )
                            existing_fragment = validate_fragment_v2(
                                fragment=existing_fragment,
                                envelope=envelope,
                                all_selected_chunks=envelope[
                                    "focusedChunks"
                                ],
                            )
                            completed_by_part[str(part["partId"])] = (
                                existing_fragment,
                                output_sha256,
                            )
                            draft_provenance["partsReused"] += 1
                            draft_provenance[
                                "reusedPromptVersion"
                            ] = STUDY_DRAFT_PROMPT_VERSION
                            continue
                        except (
                            OSError,
                            ValueError,
                            RuntimeError,
                            json.JSONDecodeError,
                        ):
                            pass
                    pending_jobs.append(
                        (
                            part,
                            envelope,
                            part_fragment_path,
                            part_provenance_path,
                        )
                    )

                def mark_fragment_ai_call() -> None:
                    draft_provenance[
                        "aiExecutedThisAttempt"
                    ] = True

                def execute_fragment_job(
                    job: tuple[
                        dict[str, Any],
                        dict[str, Any],
                        Path,
                        Path,
                    ],
                ) -> tuple[str, dict[str, Any], str]:
                    (
                        job_part,
                        job_envelope,
                        job_fragment_path,
                        job_provenance_path,
                    ) = job
                    try:
                        fragment = build_deterministic_mask_fragment_v2(
                            envelope=job_envelope,
                            all_selected_chunks=job_envelope[
                                "focusedChunks"
                            ],
                        )
                        if fragment is None:
                            fragment = build_deterministic_fo_fragment_v2(
                                envelope=job_envelope,
                                all_selected_chunks=job_envelope[
                                    "focusedChunks"
                                ],
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_function_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_function_grid_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_error_axis_tail_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_nti_f0_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_nti_horizontal_matrix_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_acoustic_matrix_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            fragment = (
                                build_deterministic_result_table_fragment_v2(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                )
                            )
                        if fragment is None:
                            with pipeline_scope("AI"):
                                fragment = fragment_runner(
                                    envelope=job_envelope,
                                    all_selected_chunks=job_envelope[
                                        "focusedChunks"
                                    ],
                                    output_path=job_fragment_path,
                                    model=model,
                                    reasoning_effort=reasoning_effort,
                                    timeout_seconds=draft_timeout_seconds,
                                    ai_call_observer=mark_fragment_ai_call,
                                )
                    except Exception as exc:
                        raise IncrementalIngestError(
                            "Staged fragment "
                            f"{job_part['partId']} failed: "
                            f"{type(exc).__name__}: {exc}"
                        ) from exc
                    fragment = validate_fragment_v2(
                        fragment=fragment,
                        envelope=job_envelope,
                        all_selected_chunks=job_envelope["focusedChunks"],
                    )
                    _atomic_write_json(
                        job_fragment_path,
                        fragment,
                    )
                    output_sha256 = _file_sha256(job_fragment_path)
                    _atomic_write_json(
                        job_provenance_path,
                        part_provenance_v2(
                            plan=draft_plan,
                            part=job_part,
                            envelope=job_envelope,
                            output_path=job_fragment_path,
                            output_sha256=output_sha256,
                            generated_at=now_iso(),
                        ),
                    )
                    return (
                        str(job_part["partId"]),
                        fragment,
                        output_sha256,
                    )

                if pending_jobs:
                    draft_provenance["partsExecuted"] = len(
                        pending_jobs
                    )
                    pending_job_iterator = iter(pending_jobs)
                    with ThreadPoolExecutor(
                        max_workers=min(
                            draft_fragment_workers,
                            len(pending_jobs),
                        )
                    ) as executor:
                        futures = {
                            executor.submit(
                                execute_fragment_job,
                                next(pending_job_iterator),
                            )
                            for _ in range(
                                min(
                                    draft_fragment_workers,
                                    len(pending_jobs),
                                )
                            )
                        }
                        try:
                            while futures:
                                future = next(as_completed(futures))
                                futures.remove(future)
                                (
                                    completed_part_id,
                                    completed_fragment,
                                    completed_sha256,
                                ) = future.result()
                                completed_by_part[completed_part_id] = (
                                    completed_fragment,
                                    completed_sha256,
                                )
                                try:
                                    next_job = next(
                                        pending_job_iterator
                                    )
                                except StopIteration:
                                    continue
                                futures.add(
                                    executor.submit(
                                        execute_fragment_job,
                                        next_job,
                                    )
                                )
                        except Exception:
                            for future in futures:
                                future.cancel()
                            raise

                part_fragments: list[
                    tuple[dict[str, Any], dict[str, Any]]
                ] = []
                for part in draft_parts:
                    part_id = str(part["partId"])
                    if part_id not in completed_by_part:
                        raise IncrementalIngestError(
                            f"Staged v2 part {part_id} did not complete."
                        )
                    fragment, output_sha256 = completed_by_part[
                        part_id
                    ]
                    part_fragments.append((part, fragment))
                    staged_part_hashes.append(
                        {
                            "partId": part_id,
                            "outputSha256": output_sha256,
                        }
                    )

                merged = merge_fragment_records(
                    plan=draft_plan,
                    fragments=part_fragments,
                    selected_chunks=selected_chunks,
                )
                _atomic_write_json(staged_merged_path, merged)
                final_coverage_inventory = (
                    build_content_coverage_inventory(
                        chunks=selected_chunks,
                        locator_results=selected_locators,
                        expected_source_cell_keys=(
                            selected_source_cell_keys
                        ),
                    )
                )
                manifest = project_canonical_manifest(
                    merged=merged,
                    registry=registry,
                    source=source,
                    workbook=workbook,
                    selected_chunks=selected_chunks,
                    semantic_inventory=final_coverage_inventory,
                )
                manifest = validate_manifest(
                    manifest,
                    coverage_chunks=selected_chunks,
                    coverage_locator_results=selected_locators,
                    expected_source_cell_keys=selected_source_cell_keys,
                    require_source_content_coverage=True,
                    precomputed_coverage_inventory=(
                        final_coverage_inventory
                    ),
                )
            if _draft_has_labels_but_no_reviewable_results(manifest):
                manifest = _terminal_manifest(
                    packet_set,
                    dataset,
                    status_override="NO_TABULAR_EVIDENCE",
                    extra_limitation=(
                        "The semantic draft contained outcome labels but no "
                        "source-backed observations or conclusions."
                    ),
                    analysis_key_override=str(
                        manifest["workbookAnalysis"]["key"]
                    ),
                )
            _atomic_write_json(manifest_path, manifest)
            if budget["mode"] == "MONOLITHIC":
                _write_draft_provenance(
                    draft_provenance_path,
                    manifest_path,
                    source,
                    details={
                        "mode": "MONOLITHIC",
                        "promptSha256": budget["promptSha256"],
                        "envelopeSha256": budget["envelopeSha256"],
                        "promptBytes": budget["promptBytes"],
                    },
                )
            else:
                if registry is None or draft_plan is None:
                    raise IncrementalIngestError(
                        "Staged v2 final provenance lacks registry/plan."
                    )
                _atomic_write_json(
                    staged_final_provenance_path,
                    final_provenance_v2(
                        plan=draft_plan,
                        registry=registry,
                        ordered_part_hashes=staged_part_hashes,
                        merged_path=staged_merged_path,
                        merged_sha256=_file_sha256(
                            staged_merged_path
                        ),
                        final_path=manifest_path,
                        final_sha256=_file_sha256(manifest_path),
                        generated_at=now_iso(),
                    ),
                )
            return manifest

        manifest = execute_stage(
            "DRAFT",
            draft_action,
            lambda value: {
                "manifestPath": str(manifest_path),
                "verificationStatus": value["workbookAnalysis"][
                    "verificationStatus"
                ],
                "studies": len(value["studies"]),
                "aiExecuted": draft_provenance["aiExecutedThisAttempt"],
                "artifactReused": draft_provenance["artifactReused"],
                "staged": draft_provenance["staged"],
                "mode": draft_provenance["mode"],
                "requestPromptBytes": draft_provenance[
                    "requestPromptBytes"
                ],
                "requestPromptSha256": draft_provenance[
                    "requestPromptSha256"
                ],
                "formulaDerivation": draft_provenance[
                    "formulaDerivation"
                ],
                "invalidatedAnalyses": draft_provenance[
                    "invalidatedAnalyses"
                ],
                "partCount": draft_provenance["partCount"],
                "partsExecuted": draft_provenance["partsExecuted"],
                "partsReused": draft_provenance["partsReused"],
                "planPath": (
                    str(draft_plan_path)
                    if draft_plan_path.is_file()
                    else ""
                ),
                "registryPath": (
                    str(draft_registry_path)
                    if draft_registry_path.is_file()
                    else ""
                ),
                "mergedPath": (
                    str(staged_merged_path)
                    if staged_merged_path.is_file()
                    else ""
                ),
                "studyDraftPromptVersion": (
                    STUDY_DRAFT_PROMPT_VERSION
                    if draft_provenance["aiExecutedThisAttempt"]
                    else draft_provenance["reusedPromptVersion"]
                ),
                "provenancePath": (
                    str(staged_final_provenance_path)
                    if staged_final_provenance_path.is_file()
                    else str(draft_provenance_path)
                    if draft_provenance_path.is_file()
                    else ""
                ),
            },
        )
        verification_status = str(
            manifest["workbookAnalysis"]["verificationStatus"]
        ).upper()
        if verification_status not in {"NEEDS_REVIEW", "EXCLUDED"}:
            raise IncrementalIngestError(
                "Incremental AI workflow cannot auto-verify a workbook analysis."
            )

        def import_action() -> dict[str, Any]:
            with _connect_rw(database) as connection:
                result = import_study_manifest(
                    connection,
                    manifest,
                    now_iso=now_iso,
                    formula_overlay=formula_overlay,
                    source_claims_prevalidated=True,
                )
                connection.commit()
                return result

        imported = execute_stage(
            "IMPORT",
            import_action,
            lambda value: value,
            gate="DB",
        )

        def verify_action() -> dict[str, Any]:
            if _sha256_file(source_file) != source_sha256:
                raise IncrementalIngestError(
                    "Source changed before incremental ingestion completed."
                )
            with _connect_rw(database) as connection:
                integrity = validate_analysis_integrity(
                    connection,
                    workbook_analysis_id=int(
                        imported["workbookAnalysisId"]
                    ),
                )
                revision = resolve_manifest_revision(connection, source)
                if (
                    not bool(revision["is_current"])
                    or str(revision["content_sha256"]).lower()
                    != source_sha256.lower()
                ):
                    raise IncrementalIngestError(
                        "Imported analysis is not attached to the current source fingerprint."
                    )
                if not integrity["ok"]:
                    raise IncrementalIngestError(
                        "Canonical database integrity verification failed."
                    )
                return integrity

        integrity = execute_stage(
            "VERIFY",
            verify_action,
            lambda value: {
                "ok": value["ok"],
                "foreignKeyErrors": value["foreignKeyErrors"],
                "invalidAggregationEffects": value[
                    "invalidAggregationEffects"
                ],
                "orphanEvidenceLinks": value["orphanEvidenceLinks"],
                "counts": value["counts"],
            },
            gate="DB",
        )
        final_status = (
            "EXCLUDED"
            if verification_status == "EXCLUDED"
            else "NEEDS_REVIEW"
        )
        result = {
            "schemaVersion": WORKFLOW_SCHEMA_VERSION,
            "runId": run_id,
            "status": final_status,
            "sourcePath": str(source_file),
            "contentSha256": source_sha256,
            "workbookStatus": workbook_status,
            "revisionUid": source["revisionUid"],
            "publicAnalysisId": imported["publicAnalysisId"],
            "studies": imported["studies"],
            "manifestPath": str(manifest_path),
            "journalPath": str(journal_path),
            "artifactDirectory": str(run_dir),
            "formulaDerivation": copy.deepcopy(
                packet_set["inventory"]["formulaDerivation"]
            ),
            "captureBackend": normalized_capture_backend,
            "captureContract": requested_capture_contract,
            "imagesAnalyzed": False,
            "integrityOk": integrity["ok"],
        }
        journal["status"] = final_status
        journal["currentStage"] = ""
        journal["result"] = result
        journal["updatedAt"] = now_iso()
        journal["finishedAt"] = now_iso()
        _atomic_write_json(journal_path, journal)
        emit_progress("WORKFLOW", "COMPLETED", final_status)
        return result


__all__ = [
    "IncrementalIngestBusyError",
    "IncrementalIngestError",
    "STAGE_ORDER",
    "WORKFLOW_SCHEMA_VERSION",
    "ingest_workbook",
    "utc_now_iso",
]
