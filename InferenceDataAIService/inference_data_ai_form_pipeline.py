"""End-to-end CLI orchestration for COM form review and corpus ingestion."""

from __future__ import annotations

import hashlib
import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from inference_data_ai_corpus_workflow import run_corpus_ingest
from inference_data_ai_form_preflight import (
    _atomic_write_json,
    run_form_preflight,
)
from inference_data_ai_form_registry import (
    analyze_form_family,
    decide_form_family,
    reclassify_form_preflight_report,
    write_form_group_review,
)


PIPELINE_SCHEMA_VERSION = "excel-form-pipeline-complete-v1"
CORPUS_MAX_RETRY_PASSES = 4
CORPUS_WORKBOOK_WORKERS = min(16, max(8, os.cpu_count() or 4))
CORPUS_PACKET_WORKERS = min(12, CORPUS_WORKBOOK_WORKERS)
CORPUS_AI_WORKERS = CORPUS_WORKBOOK_WORKERS
ProgressCallback = Callable[[dict[str, Any]], None]


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def choose_auto_form_decision(
    *,
    recommendation: str,
    validation_status: str,
    nearest_known_form_signature_id: str,
) -> str:
    """Choose the fail-closed terminal decision for CLI batch review."""

    normalized_recommendation = recommendation.strip().upper()
    if validation_status.strip().upper() != "PASSED":
        return "EXCLUDE"
    if (
        normalized_recommendation == "LINK_EXISTING"
        and nearest_known_form_signature_id.strip()
    ):
        return "LINK_EXISTING"
    if normalized_recommendation == "EXCLUDE":
        return "EXCLUDE"
    return "REGISTER_NEW"


def _emit(
    callback: ProgressCallback | None,
    *,
    stage: str,
    status: str,
    detail: str,
    source_path: str = "",
) -> None:
    if callback is None:
        return
    callback(
        {
            "schemaVersion": "ingest-progress-v1",
            "stage": stage,
            "status": status,
            "detail": detail,
            "sourcePath": source_path,
            "timestamp": utc_now_iso(),
        }
    )


def _read_json_object(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object: {path}")
    return value


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _manifest_relative_paths(
    manifest_path: Path,
    source_root: Path,
) -> list[str]:
    manifest = _read_json_object(manifest_path)
    workbooks = manifest.get("workbooks")
    if not isinstance(workbooks, list):
        raise ValueError("Form manifest must contain a workbooks array.")
    relative_paths: list[str] = []
    for index, workbook in enumerate(workbooks):
        if not isinstance(workbook, dict):
            raise ValueError(
                f"Form manifest workbooks[{index}] must be an object."
            )
        relative_path = str(workbook.get("relativePath") or "")
        if not relative_path.strip():
            raise ValueError(
                f"Form manifest workbooks[{index}].relativePath is required."
            )
        source = (source_root / Path(relative_path)).resolve()
        try:
            source.relative_to(source_root)
        except ValueError as exc:
            raise ValueError(
                "Form manifest path escapes the source root: "
                + relative_path
            ) from exc
        if not source.is_file():
            raise FileNotFoundError(source)
        expected = str(
            workbook.get("contentSha256") or ""
        ).strip().lower()
        if expected and _sha256_file(source).lower() != expected:
            raise ValueError(
                "Source changed after form preflight: " + relative_path
            )
        relative_paths.append(relative_path)
    return relative_paths


def _checkpoint(
    path: Path,
    *,
    started_at: str,
    stage: str,
    status: str,
    total_groups: int,
    outcomes: list[dict[str, Any]],
    errors: list[dict[str, Any]],
) -> None:
    _atomic_write_json(
        path,
        {
            "schemaVersion": PIPELINE_SCHEMA_VERSION,
            "startedAt": started_at,
            "updatedAt": utc_now_iso(),
            "stage": stage,
            "status": status,
            "totalGroups": total_groups,
            "completedGroups": len(outcomes),
            "errorGroups": len(errors),
            "outcomes": outcomes,
            "errors": errors,
        },
    )


def _run_corpus(
    *,
    database_path: Path,
    source_root: Path,
    output_root: Path,
    manifest_path: Path,
    dataset: str,
    draft_monolithic_max_bytes: int,
    progress_callback: ProgressCallback | None,
) -> dict[str, Any]:
    artifact_root = (
        output_root
        / "incremental-com-corpus"
        / "form-approved"
    )
    artifact_root.mkdir(parents=True, exist_ok=True)
    journal_path = artifact_root / "corpus-journal.json"
    include_relative_paths = _manifest_relative_paths(
        manifest_path,
        source_root,
    )
    pass_summaries: list[dict[str, Any]] = []
    result: dict[str, Any] = {}
    for pass_number in range(1, CORPUS_MAX_RETRY_PASSES + 1):
        if pass_number > 1:
            _emit(
                progress_callback,
                stage="CORPUS_RETRY",
                status="RUNNING",
                detail=(
                    f"검증 실패 draft 자동 복구 {pass_number}/"
                    f"{CORPUS_MAX_RETRY_PASSES}"
                ),
            )
        result = run_corpus_ingest(
            database_path=database_path,
            source_root=source_root,
            artifact_root=artifact_root,
            journal_path=journal_path,
            dataset=dataset,
            resume=True,
            retry_failed=True,
            inventory_only=False,
            include_relative_paths=include_relative_paths,
            offset=0,
            limit=0,
            workbook_workers=CORPUS_WORKBOOK_WORKERS,
            workbook_retry_attempts=3,
            com_workers=1,
            packet_workers=CORPUS_PACKET_WORKERS,
            ai_workers=CORPUS_AI_WORKERS,
            db_workers=1,
            ingest_options={
                "max_cells": 400,
                "max_rows": 50,
                "empty_row_gap": 3,
                "locator_workers": 3,
                "locator_batch_size": 6,
                "locator_batch_max_bytes": 240_000,
                "draft_monolithic_max_bytes": (
                    draft_monolithic_max_bytes
                ),
                "draft_fragment_max_chunks": 8,
                "draft_fragment_max_cells": 2_000,
                "draft_fragment_max_bytes": 400_000,
                "draft_fragment_workers": 3,
                "derive_formula_values": False,
                "repair_rejected_draft": True,
                "repair_unselected_source": True,
                "model": None,
                "reasoning_effort": "medium",
                "locator_timeout_seconds": 900,
                "draft_timeout_seconds": 1800,
                "capture_backend": "com",
                "covered_cell_mode": "blank",
                "include_hidden_sheets": True,
                "inspect_auth_dialog": False,
                "dismiss_auth_dialog": False,
                "auth_dialog_title": "",
                "auth_dialog_class": "",
                "auth_dialog_button": "",
                "auth_dialog_timeout_seconds": 30.0,
                "progress_callback": progress_callback,
            },
        )
        summary = dict(result.get("summary") or {})
        status_counts = dict(
            summary.get("currentStatusCounts") or {}
        )
        failed_count = int(
            status_counts.get("FAILED")
            or summary.get("failedThisRun")
            or 0
        )
        pass_summaries.append(
            {
                "pass": pass_number,
                "status": str(result.get("status") or ""),
                "attempted": int(summary.get("attempted") or 0),
                "completedThisRun": int(
                    summary.get("completedThisRun") or 0
                ),
                "failedThisRun": int(
                    summary.get("failedThisRun") or 0
                ),
                "remainingFailed": failed_count,
            }
        )
        if failed_count == 0:
            break
    result["retryPasses"] = pass_summaries
    return result


def run_form_pipeline_complete(
    *,
    database_path: str | Path,
    source_root: str | Path,
    output_root: str | Path,
    reviewer: str,
    dataset: str = "InputDataFinish",
    analysis_workers: int = 2,
    reasoning_effort: str = "low",
    analysis_timeout_seconds: int = 900,
    com_timeout_seconds: float = 300.0,
    codex_executable: str | None = None,
    exclude_on_analysis_error: bool = False,
    run_corpus: bool = True,
    max_families: int = 0,
    draft_monolithic_max_bytes: int = 400_000,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, Any]:
    """Run preflight, batch review, decisions, and selected corpus ingest."""

    if not reviewer.strip():
        raise ValueError("reviewer is required.")
    if analysis_workers < 1:
        raise ValueError("analysis_workers must be positive.")
    if draft_monolithic_max_bytes < 1:
        raise ValueError(
            "draft_monolithic_max_bytes must be positive."
        )
    database = Path(database_path).expanduser().resolve()
    source = Path(source_root).expanduser().resolve()
    output = Path(output_root).expanduser().resolve()
    if not database.is_file():
        raise FileNotFoundError(database)
    if not source.is_dir():
        raise FileNotFoundError(source)
    output.mkdir(parents=True, exist_ok=True)
    preflight_directory = output / "form-preflight"
    preflight_directory.mkdir(parents=True, exist_ok=True)
    report_path = preflight_directory / "latest.json"
    review_path = preflight_directory / "group-review.latest.json"
    checkpoint_path = preflight_directory / "pipeline.checkpoint.json"
    result_path = preflight_directory / "pipeline.result.json"
    contracts_directory = preflight_directory / "contracts"
    contracts_directory.mkdir(parents=True, exist_ok=True)
    started_at = utc_now_iso()
    outcomes: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []

    _emit(
        progress_callback,
        stage="FORM_PREFLIGHT",
        status="RUNNING",
        detail="전체 보관함 COM 사전 분석을 재개합니다.",
    )
    report = run_form_preflight(
        database_path=database,
        source_root=source,
        output_path=report_path,
        dataset=dataset,
        com_timeout_seconds=com_timeout_seconds,
        progress_callback=progress_callback,
    )
    if str(report.get("status") or "") != "COMPLETED":
        raise RuntimeError(
            "Form preflight did not complete: "
            + str(report.get("status") or "")
        )
    review = write_form_group_review(
        database_path=database,
        report_path=report_path,
        output_path=review_path,
    )
    groups = [
        group
        for group in review.get("groups") or []
        if str(group.get("decisionStatus") or "")
        in {"PENDING", "ANALYZED_PENDING_APPROVAL"}
    ]
    if max_families > 0:
        groups = groups[:max_families]
    total_groups = len(groups)
    _checkpoint(
        checkpoint_path,
        started_at=started_at,
        stage="FORM_FAMILY_REVIEW",
        status="RUNNING",
        total_groups=total_groups,
        outcomes=outcomes,
        errors=errors,
    )

    def decide(
        group: dict[str, Any],
        *,
        recommendation: str,
        validation_status: str,
    ) -> None:
        family_id = str(group["familyId"])
        automatic_decision = choose_auto_form_decision(
            recommendation=recommendation,
            validation_status=validation_status,
            nearest_known_form_signature_id=str(
                group.get("nearestKnownFormSignatureId") or ""
            ),
        )
        linked_signature = (
            str(group.get("nearestKnownFormSignatureId") or "")
            if automatic_decision == "LINK_EXISTING"
            else ""
        )
        decision_result = decide_form_family(
            database_path=database,
            report_path=report_path,
            family_id=family_id,
            decision=automatic_decision,
            reviewer=reviewer,
            display_name=str(group.get("displayName") or ""),
            linked_form_signature_id=linked_signature,
            group_snapshot=group,
            notes=(
                "CLI 전체 자동 처리: "
                f"AI recommendation={recommendation or 'ERROR'}, "
                f"validation={validation_status or 'ERROR'}"
            ),
        )
        outcomes.append(
            {
                "familyId": family_id,
                "memberCount": int(group.get("memberCount") or 0),
                "decision": decision_result["status"],
                "recommendation": recommendation,
                "validationStatus": validation_status,
            }
        )
        _emit(
            progress_callback,
            stage="FORM_FAMILY_REVIEW",
            status=decision_result["status"],
            detail=(
                f"{len(outcomes) + len(errors)}/{total_groups} · "
                f"{family_id} · {decision_result['status']}"
            ),
            source_path=str(group.get("representativeSource") or ""),
        )
        _checkpoint(
            checkpoint_path,
            started_at=started_at,
            stage="FORM_FAMILY_REVIEW",
            status="RUNNING",
            total_groups=total_groups,
            outcomes=outcomes,
            errors=errors,
        )

    analyzed = [
        group
        for group in groups
        if str(group.get("decisionStatus") or "")
        == "ANALYZED_PENDING_APPROVAL"
    ]
    for group in analyzed:
        decide(
            group,
            recommendation=str(group.get("recommendation") or ""),
            validation_status=str(
                group.get("validationStatus") or ""
            ),
        )

    pending = [
        group
        for group in groups
        if str(group.get("decisionStatus") or "") == "PENDING"
    ]

    def analyze(group: dict[str, Any]) -> dict[str, Any]:
        family_id = str(group["familyId"])
        return analyze_form_family(
            database_path=database,
            report_path=report_path,
            family_id=family_id,
            output_path=contracts_directory / f"{family_id}.json",
            codex_executable=codex_executable,
            reasoning_effort=reasoning_effort,
            timeout_seconds=analysis_timeout_seconds,
            group_snapshot=group,
        )

    with ThreadPoolExecutor(max_workers=analysis_workers) as executor:
        futures = {
            executor.submit(analyze, group): group
            for group in pending
        }
        for future in as_completed(futures):
            group = futures[future]
            family_id = str(group["familyId"])
            try:
                analysis = future.result()
                decide(
                    group,
                    recommendation=str(
                        analysis.get("recommendation") or ""
                    ),
                    validation_status=str(
                        analysis.get("validationStatus") or ""
                    ),
                )
            except Exception as exc:
                error = {
                    "familyId": family_id,
                    "memberCount": int(group.get("memberCount") or 0),
                    "error": f"{type(exc).__name__}: {exc}",
                }
                errors.append(error)
                _emit(
                    progress_callback,
                    stage="FORM_FAMILY_REVIEW",
                    status="FAILED",
                    detail=(
                        f"{len(outcomes) + len(errors)}/{total_groups} · "
                        f"{family_id} · {error['error']}"
                    ),
                    source_path=str(
                        group.get("representativeSource") or ""
                    ),
                )
                if exclude_on_analysis_error:
                    try:
                        decide(
                            group,
                            recommendation="EXCLUDE",
                            validation_status="ERROR",
                        )
                    except Exception as decision_exc:
                        error["decisionError"] = (
                            f"{type(decision_exc).__name__}: "
                            f"{decision_exc}"
                        )
                _checkpoint(
                    checkpoint_path,
                    started_at=started_at,
                    stage="FORM_FAMILY_REVIEW",
                    status="RUNNING",
                    total_groups=total_groups,
                    outcomes=outcomes,
                    errors=errors,
                )

    report = reclassify_form_preflight_report(
        database_path=database,
        report_path=report_path,
    )
    review = write_form_group_review(
        database_path=database,
        report_path=report_path,
        output_path=review_path,
    )
    pending_count = int(review["summary"]["pendingCount"])
    limited = max_families > 0 and total_groups < int(
        review["summary"]["groupCount"]
    )
    corpus_result: dict[str, Any] | None = None
    if run_corpus and pending_count == 0 and not limited:
        _emit(
            progress_callback,
            stage="CORPUS_INGEST",
            status="RUNNING",
            detail=(
                "승인된 manifest 전체 처리를 시작합니다: "
                f"{report['summary']['knownForms']}개"
            ),
        )
        corpus_result = _run_corpus(
            database_path=database,
            source_root=source,
            output_root=output,
            manifest_path=Path(report["knownFormManifestPath"]),
            dataset=dataset,
            draft_monolithic_max_bytes=draft_monolithic_max_bytes,
            progress_callback=progress_callback,
        )

    status = (
        "COMPLETED_WITH_ERRORS"
        if errors or pending_count > 0
        else str(corpus_result.get("status") or "COMPLETED")
        if corpus_result is not None
        else "COMPLETED"
    )
    result = {
        "schemaVersion": PIPELINE_SCHEMA_VERSION,
        "status": status,
        "startedAt": started_at,
        "finishedAt": utc_now_iso(),
        "databasePath": str(database),
        "sourceRoot": str(source),
        "reportPath": str(report_path),
        "reviewPath": str(review_path),
        "manifestPath": str(report["knownFormManifestPath"]),
        "checkpointPath": str(checkpoint_path),
        "preflightSummary": report["summary"],
        "reviewSummary": review["summary"],
        "outcomes": outcomes,
        "errors": errors,
        "corpus": corpus_result,
    }
    _atomic_write_json(result_path, result)
    _checkpoint(
        checkpoint_path,
        started_at=started_at,
        stage="COMPLETE",
        status=status,
        total_groups=total_groups,
        outcomes=outcomes,
        errors=errors,
    )
    return {**result, "resultPath": str(result_path)}


__all__ = [
    "PIPELINE_SCHEMA_VERSION",
    "choose_auto_form_decision",
    "run_form_pipeline_complete",
]
