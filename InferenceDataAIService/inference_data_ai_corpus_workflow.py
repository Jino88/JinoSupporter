"""Durable orchestration for a corpus of DRM-free ``.xlsx`` workbooks.

The corpus layer never opens Excel and never analyzes images.  It discovers a
deterministic inventory, retains one journal record per source fingerprint,
and delegates each selected workbook to :func:`ingest_workbook`.  Completed
fingerprints are resumable without being processed again, while failed
fingerprints remain visible and can be retried explicitly.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import threading
import sqlite3
from collections import Counter
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from contextlib import closing, contextmanager
from pathlib import Path
from typing import Any, Callable, Iterator, Mapping, Sequence

from inference_data_ai_schema import (
    stable_uid,
    validate_knowledge_integrity,
)
from inference_data_ai_workflow import ingest_workbook, utc_now_iso


CORPUS_WORKFLOW_SCHEMA_VERSION = "full-corpus-ingest-v1"
COMPLETED_STATUS = "COMPLETED"
FAILED_STATUS = "FAILED"
PENDING_STATUS = "PENDING"
RUNNING_STATUS = "RUNNING"
INTERRUPTED_STATUS = "INTERRUPTED"
RECORD_STATUSES = {
    COMPLETED_STATUS,
    FAILED_STATUS,
    PENDING_STATUS,
    RUNNING_STATUS,
    INTERRUPTED_STATUS,
}


class CorpusWorkflowError(RuntimeError):
    """Raised when a corpus batch cannot be initialized or resumed safely."""


class CorpusWorkflowBusyError(CorpusWorkflowError):
    """Raised when another process owns the same corpus journal."""


class CorpusPipelineGates:
    """Process-local bounded gates shared by every workbook workflow."""

    def __init__(
        self,
        *,
        com_workers: int,
        packet_workers: int,
        ai_workers: int,
        db_workers: int,
    ) -> None:
        limits = {
            "COM": com_workers,
            "PACKET": packet_workers,
            "AI": ai_workers,
            "DB": db_workers,
        }
        invalid = {
            name: value
            for name, value in limits.items()
            if value < 1
        }
        if invalid:
            raise ValueError(
                "Pipeline worker limits must be positive: "
                + ", ".join(
                    f"{name}={value}"
                    for name, value in sorted(invalid.items())
                )
            )
        self.limits = limits
        self._semaphores = {
            name: threading.BoundedSemaphore(value)
            for name, value in limits.items()
        }

    @contextmanager
    def acquire(self, stage: str) -> Iterator[None]:
        normalized = stage.strip().upper()
        try:
            semaphore = self._semaphores[normalized]
        except KeyError as exc:
            raise ValueError(
                f"Unknown corpus pipeline stage gate: {stage}"
            ) from exc
        semaphore.acquire()
        try:
            yield
        finally:
            semaphore.release()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _path_sort_key(path: Path) -> tuple[str, str]:
    text = str(path)
    return (text.casefold(), text)


def _normalized_workbook_filename(source_path: str | Path) -> str:
    """Return the filename identity used for legacy DRM-release matching.

    The legacy preparation/DRM pipeline can append one or more
    ``_<numeric id>`` suffixes and then append ``_clean``.  Any of those
    generated trailing suffixes do not belong to the user's original
    workbook name.
    """

    source = Path(source_path)
    stem = re.sub(r"_clean$", "", source.stem, flags=re.IGNORECASE)
    stem = re.sub(r"(?:_[0-9]{9,13})+$", "", stem)
    return (stem.strip() + source.suffix).casefold()


def _record_sort_key(record: Mapping[str, Any]) -> tuple[str, str, str]:
    source_path = str(record.get("sourcePath") or "")
    return (
        source_path.casefold(),
        source_path,
        str(record.get("contentSha256") or ""),
    )


def discover_excel_sources(
    source_root: str | Path,
    *,
    capture_backend: str = "openxml",
) -> list[Path]:
    """Return a recursively discovered deterministic Excel inventory.

    Excel owner/lock files whose names start with ``~$`` are deliberately
    excluded.  The function only inspects paths; workbook bytes are not
    modified and Excel/COM is never started.
    """

    normalized_backend = capture_backend.strip().lower()
    if normalized_backend not in {"openxml", "com"}:
        raise ValueError("capture_backend must be openxml or com.")
    extensions = (
        {".xlsx"}
        if normalized_backend == "openxml"
        else {".xlsx", ".xlsm", ".xlsb", ".xls"}
    )
    root = Path(source_root).expanduser().resolve()
    if root.is_file():
        candidates: Sequence[Path] = [root]
    elif root.is_dir():
        candidates = root.rglob("*")
    else:
        raise FileNotFoundError(root)
    files = [
        path.resolve()
        for path in candidates
        if path.is_file()
        and path.suffix.lower() in extensions
        and not path.name.startswith("~$")
    ]
    return sorted(files, key=_path_sort_key)


def discover_xlsx_sources(source_root: str | Path) -> list[Path]:
    """Backward-compatible DRM-free OpenXML inventory helper."""

    return discover_excel_sources(
        source_root,
        capture_backend="openxml",
    )


def _atomic_write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(
        f".{path.name}.{os.getpid()}.{threading.get_ident()}.tmp"
    )
    encoded = (
        json.dumps(value, ensure_ascii=False, indent=2) + "\n"
    ).encode("utf-8")
    try:
        with temporary.open("wb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass


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
def _batch_lock(path: Path) -> Iterator[None]:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor: int | None = None
    while descriptor is None:
        try:
            descriptor = os.open(
                path,
                os.O_CREAT | os.O_EXCL | os.O_WRONLY,
            )
            os.write(descriptor, str(os.getpid()).encode("ascii"))
        except FileExistsError as exc:
            try:
                owner_pid = int(path.read_text(encoding="ascii").strip())
            except (OSError, ValueError):
                owner_pid = 0
            if owner_pid and _pid_exists(owner_pid):
                raise CorpusWorkflowBusyError(
                    f"Corpus ingestion is already running (PID {owner_pid})."
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


def _relative_path(source: Path, root: Path) -> str:
    if root.is_dir():
        try:
            return str(source.relative_to(root))
        except ValueError:
            pass
    return source.name


def _new_record(
    *,
    source: Path,
    root: Path,
    content_sha256: str,
    discovered_at: str,
    size_bytes: int = 0,
    mtime_ns: int = 0,
    error: str = "",
) -> dict[str, Any]:
    record_id = stable_uid(
        "corpus-source",
        str(source),
        content_sha256 or "inventory-error",
    )
    return {
        "recordId": record_id,
        "sourcePath": str(source),
        "relativePath": _relative_path(source, root),
        "contentSha256": content_sha256,
        "sizeBytes": size_bytes,
        "mtimeNs": mtime_ns,
        "status": FAILED_STATUS if error else PENDING_STATUS,
        "attempts": 0,
        "result": None,
        "error": error,
        "imagesAnalyzed": False,
        "presentInLatestDiscovery": True,
        "firstDiscoveredAt": discovered_at,
        "lastDiscoveredAt": discovered_at,
        "lastStartedAt": "",
        "lastFinishedAt": discovered_at if error else "",
        "updatedAt": discovered_at,
    }


def _validate_loaded_journal(
    journal: Mapping[str, Any],
    *,
    source_root: Path,
    database_path: Path,
    artifact_root: Path,
    dataset: str,
) -> None:
    if journal.get("schemaVersion") != CORPUS_WORKFLOW_SCHEMA_VERSION:
        raise CorpusWorkflowError(
            "Existing corpus journal has an unsupported schema version."
        )
    expected = {
        "sourceRoot": str(source_root),
        "databasePath": str(database_path),
        "artifactRoot": str(artifact_root),
        "dataset": dataset,
    }
    mismatches = [
        name
        for name, value in expected.items()
        if str(journal.get(name) or "") != value
    ]
    if mismatches:
        raise CorpusWorkflowError(
            "Existing corpus journal belongs to a different "
            + ", ".join(mismatches)
            + "."
        )
    records = journal.get("records")
    if not isinstance(records, list):
        raise CorpusWorkflowError(
            "Existing corpus journal does not contain a record inventory."
        )
    seen: set[str] = set()
    for record in records:
        if not isinstance(record, dict):
            raise CorpusWorkflowError("Corpus journal contains an invalid record.")
        record_id = str(record.get("recordId") or "")
        if not record_id or record_id in seen:
            raise CorpusWorkflowError(
                "Corpus journal contains a missing or duplicate recordId."
            )
        seen.add(record_id)
        if record.get("status") not in RECORD_STATUSES:
            raise CorpusWorkflowError(
                f"Corpus journal record has invalid status: {record_id}"
            )


def _load_or_create_journal(
    *,
    journal_path: Path,
    source_root: Path,
    database_path: Path,
    artifact_root: Path,
    dataset: str,
    resume: bool,
    now: str,
) -> dict[str, Any]:
    if journal_path.is_file():
        if not resume:
            raise CorpusWorkflowError(
                "Corpus journal already exists; use resume=True or a new journal path."
            )
        try:
            journal = json.loads(journal_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise CorpusWorkflowError(
                f"Existing corpus journal is invalid JSON: {journal_path}"
            ) from exc
        _validate_loaded_journal(
            journal,
            source_root=source_root,
            database_path=database_path,
            artifact_root=artifact_root,
            dataset=dataset,
        )
        for record in journal["records"]:
            if record["status"] == RUNNING_STATUS:
                record["status"] = INTERRUPTED_STATUS
                record["error"] = (
                    "Previous corpus process ended before this attempt completed."
                )
                record["lastFinishedAt"] = now
                record["updatedAt"] = now
        for run in journal.get("runs", []):
            if run.get("status") == RUNNING_STATUS:
                run["status"] = INTERRUPTED_STATUS
                run["finishedAt"] = now
        return journal
    return {
        "schemaVersion": CORPUS_WORKFLOW_SCHEMA_VERSION,
        "batchId": stable_uid(
            "corpus-batch",
            dataset,
            str(source_root),
            str(database_path),
            CORPUS_WORKFLOW_SCHEMA_VERSION,
        ),
        "sourceRoot": str(source_root),
        "databasePath": str(database_path),
        "artifactRoot": str(artifact_root),
        "dataset": dataset,
        "status": PENDING_STATUS,
        "imagesAnalyzed": False,
        "createdAt": now,
        "updatedAt": now,
        "records": [],
        "runs": [],
    }


def _refresh_inventory(
    journal: dict[str, Any],
    sources: Sequence[Path],
    *,
    source_root: Path,
    now: str,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = journal["records"]
    by_id = {str(record["recordId"]): record for record in records}
    for record in records:
        record["presentInLatestDiscovery"] = False
    current: list[dict[str, Any]] = []
    for source in sources:
        try:
            stat = source.stat()
            content_sha256 = _sha256_file(source)
            size_bytes = int(stat.st_size)
            mtime_ns = int(stat.st_mtime_ns)
            inventory_error = ""
        except OSError as exc:
            content_sha256 = ""
            size_bytes = 0
            mtime_ns = 0
            inventory_error = f"{type(exc).__name__}: {exc}"
        candidate = _new_record(
            source=source,
            root=source_root,
            content_sha256=content_sha256,
            discovered_at=now,
            size_bytes=size_bytes,
            mtime_ns=mtime_ns,
            error=inventory_error,
        )
        record = by_id.get(candidate["recordId"])
        if record is None:
            record = candidate
            records.append(record)
            by_id[str(record["recordId"])] = record
        else:
            record["sourcePath"] = str(source)
            record["relativePath"] = _relative_path(source, source_root)
            record["sizeBytes"] = size_bytes
            record["mtimeNs"] = mtime_ns
            record["lastDiscoveredAt"] = now
            record["updatedAt"] = now
            if inventory_error:
                record["status"] = FAILED_STATUS
                record["error"] = inventory_error
                record["lastFinishedAt"] = now
        record["presentInLatestDiscovery"] = True
        current.append(record)
    records.sort(key=_record_sort_key)
    return sorted(current, key=_record_sort_key)


def _accounting(
    journal: Mapping[str, Any],
    *,
    current_records: Sequence[Mapping[str, Any]],
    selected_records: Sequence[Mapping[str, Any]],
    actions: Mapping[str, str],
) -> dict[str, Any]:
    tracked = list(journal["records"])
    current_statuses = Counter(str(item["status"]) for item in current_records)
    tracked_statuses = Counter(str(item["status"]) for item in tracked)
    result_statuses = Counter(
        str((item.get("result") or {}).get("status") or "NONE")
        for item in current_records
    )
    action_counts = Counter(actions.values())
    return {
        "discoveredSources": len(current_records),
        "trackedFingerprintRecords": len(tracked),
        "historicalOrMissingFingerprintRecords": sum(
            not bool(item.get("presentInLatestDiscovery"))
            for item in tracked
        ),
        "selectedSources": len(selected_records),
        "notSelectedSources": len(current_records) - len(selected_records),
        "attempted": sum(
            action_counts[name]
            for name in ("ATTEMPTED", "COMPLETED", "FAILED")
        ),
        "completedThisRun": action_counts["COMPLETED"],
        "failedThisRun": action_counts["FAILED"],
        "skippedCompleted": action_counts["SKIPPED_COMPLETED"],
        "skippedFailed": action_counts["SKIPPED_FAILED"],
        "skippedInventoryError": action_counts["SKIPPED_INVENTORY_ERROR"],
        "reconciledExisting": action_counts["RECONCILED_EXISTING"],
        "currentStatusCounts": dict(sorted(current_statuses.items())),
        "trackedStatusCounts": dict(sorted(tracked_statuses.items())),
        "currentResultStatusCounts": dict(sorted(result_statuses.items())),
    }


def _reconcile_existing_analyses(
    journal: dict[str, Any],
    *,
    database_path: Path,
    now: str,
) -> set[str]:
    """Mark workbooks with an answer-visible canonical analysis as ingested.

    A captured source revision alone is not semantic corpus completion.  Such
    a revision must remain eligible so the workbook workflow can continue
    through packet generation, AI analysis, import, and verification.
    """

    connection = sqlite3.connect(
        f"{database_path.as_uri()}?mode=ro",
        uri=True,
    )
    try:
        tables = {
            str(row[0])
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }
        required = {
            "source_documents",
            "source_revisions",
            "workbook_analyses",
            "knowledge_studies",
        }
        if not required <= tables:
            return set()
        rows = connection.execute(
            """
            SELECT
                document.source_path,
                document.original_file_name,
                revision.content_sha256,
                revision.revision_uid,
                analysis.public_analysis_id,
                analysis.analysis_status,
                analysis.verification_status,
                COUNT(study.study_id)
            FROM source_documents document
            JOIN source_revisions revision
              ON revision.document_id=document.document_id
             AND revision.is_current=1
            LEFT JOIN workbook_analyses analysis
              ON analysis.document_id=document.document_id
             AND analysis.revision_id=revision.revision_id
             AND analysis.analyzer_name='canonical-study-import'
             AND analysis.verification_status IN (
                 'VERIFIED','NEEDS_REVIEW','EXCLUDED'
             )
            LEFT JOIN knowledge_studies study
              ON study.workbook_analysis_id=analysis.workbook_analysis_id
            WHERE document.lifecycle_status='ACTIVE'
            GROUP BY
                document.document_id,
                revision.revision_id,
                analysis.workbook_analysis_id
            ORDER BY
                document.source_path,
                CASE WHEN analysis.public_analysis_id IS NULL THEN 1 ELSE 0 END,
                analysis.workbook_analysis_id
            """
        ).fetchall()
    finally:
        connection.close()
    existing: dict[tuple[str, str], Any] = {}
    existing_by_filename: dict[str, Any] = {}
    for row in rows:
        if not str(row[4] or "").strip():
            continue
        existing.setdefault(
            (str(row[0]), str(row[2]).lower()),
            row,
        )
        for filename_source in (row[0], row[1]):
            normalized = _normalized_workbook_filename(
                str(filename_source or "")
            )
            if normalized:
                existing_by_filename.setdefault(normalized, row)
    _downgrade_completed_without_current_analysis(
        journal,
        current_analysis_keys=set(existing),
        current_analysis_filenames=set(existing_by_filename),
        now=now,
    )
    reconciled: set[str] = set()
    for record in journal["records"]:
        was_completed = record["status"] == COMPLETED_STATUS
        row = existing.get(
            (
                str(record["sourcePath"]),
                str(record["contentSha256"]).lower(),
            )
        )
        match_kind = "PATH_AND_CONTENT"
        if row is None:
            row = existing_by_filename.get(
                _normalized_workbook_filename(
                    str(record["sourcePath"])
                )
            )
            match_kind = "NORMALIZED_FILENAME"
        if row is None:
            continue
        record["status"] = COMPLETED_STATUS
        record["result"] = {
            "schemaVersion": "existing-canonical-analysis-v2",
            "status": str(row[6]),
            "sourcePath": str(record["sourcePath"]),
            "contentSha256": str(record["contentSha256"]),
            "revisionUid": str(row[3]),
            "publicAnalysisId": str(row[4]),
            "analysisStatus": str(row[5]),
            "studies": int(row[7]),
            "imagesAnalyzed": False,
            "integrityOk": True,
            "reconciledExisting": True,
            "sourceOnlyDuplicate": False,
            "duplicateMatchKind": match_kind,
            "normalizedFileName": _normalized_workbook_filename(
                str(record["sourcePath"])
            ),
            "matchedSourcePath": str(row[0]),
            "matchedOriginalFileName": str(row[1] or ""),
            "matchedContentSha256": str(row[2]),
        }
        record["error"] = ""
        record["lastFinishedAt"] = now
        record["updatedAt"] = now
        if not was_completed:
            reconciled.add(str(record["recordId"]))
    return reconciled


def _downgrade_completed_without_current_analysis(
    journal: dict[str, Any],
    *,
    current_analysis_keys: set[tuple[str, str]],
    current_analysis_filenames: set[str] | None = None,
    now: str,
) -> int:
    """Requeue a corpus result without a current canonical analysis.

    This also migrates durable journal entries created by the former
    source-only reconciliation behavior.  Unrelated runner/test results that
    never claimed canonical or source-only DB reconciliation are preserved.
    """

    downgraded = 0
    for record in journal["records"]:
        result = record.get("result")
        claimed_database_completion = (
            isinstance(result, dict)
            and (
                bool(str(result.get("publicAnalysisId") or "").strip())
                or result.get("sourceOnlyDuplicate") is True
                or result.get("schemaVersion")
                == "existing-database-source-v1"
            )
        )
        if (
            record.get("status") != COMPLETED_STATUS
            or not isinstance(result, dict)
            or not claimed_database_completion
        ):
            continue
        key = (
            str(record.get("sourcePath") or ""),
            str(record.get("contentSha256") or "").lower(),
        )
        if (
            key in current_analysis_keys
            or _normalized_workbook_filename(key[0])
            in (current_analysis_filenames or set())
        ):
            continue
        record["status"] = PENDING_STATUS
        record["result"] = None
        record["error"] = (
            "The previously completed canonical analysis is no longer "
            "current or answer-visible; semantic re-ingestion is required."
        )
        record["lastFinishedAt"] = now
        record["updatedAt"] = now
        downgraded += 1
    return downgraded


def _prepare_database_for_parallel_ingest(database_path: Path) -> None:
    """Allow long-lived analysis readers to coexist with the DB writer."""

    connection = sqlite3.connect(str(database_path), timeout=60)
    try:
        connection.execute("PRAGMA busy_timeout=60000")
        mode = str(
            connection.execute("PRAGMA journal_mode=WAL").fetchone()[0]
        ).lower()
        if mode != "wal":
            raise CorpusWorkflowError(
                "Canonical database did not enter WAL journal mode."
            )
    finally:
        connection.close()


def run_corpus_ingest(
    *,
    database_path: str | Path,
    source_root: str | Path,
    artifact_root: str | Path,
    journal_path: str | Path | None = None,
    dataset: str = "InputDataFinish",
    resume: bool = True,
    retry_failed: bool = False,
    inventory_only: bool = False,
    include_relative_paths: Sequence[str] | None = None,
    offset: int = 0,
    limit: int = 0,
    workbook_workers: int = 1,
    workbook_retry_attempts: int = 1,
    com_workers: int = 1,
    packet_workers: int = 3,
    ai_workers: int = 3,
    db_workers: int = 1,
    ingest_options: Mapping[str, Any] | None = None,
    now_iso: Callable[[], str] = utc_now_iso,
    ingest_runner: Callable[..., dict[str, Any]] = ingest_workbook,
) -> dict[str, Any]:
    """Process a deterministic slice of a workbook corpus.

    ``offset`` and ``limit`` apply to the complete current inventory before
    resume/retry filtering, so repeated chunk invocations select the same
    source paths.  A non-positive ``limit`` means no upper bound.
    """

    if offset < 0 or limit < 0:
        raise ValueError("offset and limit must be non-negative.")
    if workbook_workers < 1:
        raise ValueError("workbook_workers must be positive.")
    if workbook_retry_attempts < 1:
        raise ValueError("workbook_retry_attempts must be positive.")
    pipeline_gates = CorpusPipelineGates(
        com_workers=com_workers,
        packet_workers=packet_workers,
        ai_workers=ai_workers,
        db_workers=db_workers,
    )
    database = Path(database_path).expanduser().resolve()
    root = Path(source_root).expanduser().resolve()
    artifacts = Path(artifact_root).expanduser().resolve()
    journal_file = (
        Path(journal_path).expanduser().resolve()
        if journal_path is not None
        else artifacts / "corpus-journal.json"
    )
    if not database.is_file():
        raise CorpusWorkflowError(
            f"Canonical database is not initialized: {database}"
        )
    if not inventory_only and workbook_workers > 1:
        _prepare_database_for_parallel_ingest(database)
    options = dict(ingest_options or {})
    reserved = {
        "database_path",
        "source_path",
        "artifact_root",
        "dataset",
        "resume",
        "pipeline_gate",
    }
    overlap = sorted(reserved.intersection(options))
    if overlap:
        raise ValueError(
            "ingest_options cannot override corpus-owned arguments: "
            + ", ".join(overlap)
        )
    sources = discover_excel_sources(
        root,
        capture_backend=str(options.get("capture_backend") or "openxml"),
    )
    lock_path = journal_file.with_name(journal_file.name + ".lock")
    with _batch_lock(lock_path):
        now = now_iso()
        journal = _load_or_create_journal(
            journal_path=journal_file,
            source_root=root,
            database_path=database,
            artifact_root=artifacts,
            dataset=dataset,
            resume=resume,
            now=now,
        )
        current = _refresh_inventory(
            journal,
            sources,
            source_root=root,
            now=now,
        )
        reconciled = _reconcile_existing_analyses(
            journal,
            database_path=database,
            now=now,
        )
        if include_relative_paths is not None:
            requested_order = list(
                dict.fromkeys(
                    str(Path(value))
                    for value in include_relative_paths
                    if str(value).strip()
                )
            )
            requested = set(requested_order)
            by_relative = {
                str(Path(str(record["relativePath"]))): record
                for record in current
            }
            missing = sorted(requested - set(by_relative))
            if missing:
                raise CorpusWorkflowError(
                    "Requested corpus paths are not in the current inventory: "
                    + ", ".join(missing)
                )
            filtered = [
                by_relative[relative]
                for relative in requested_order
            ]
        else:
            filtered = current
        selected = [] if inventory_only else filtered[offset:]
        if not inventory_only and limit > 0:
            selected = selected[:limit]
        run_id = stable_uid(
            "corpus-run",
            journal["batchId"],
            len(journal.get("runs", [])) + 1,
            now,
        )

        def emit_corpus_progress(
            status: str,
            detail: Mapping[str, Any],
        ) -> None:
            callback = options.get("progress_callback")
            if not callable(callback):
                return
            try:
                callback(
                    {
                        "schemaVersion": "ingest-progress-v1",
                        "runId": run_id,
                        "stage": "CORPUS",
                        "status": status,
                        "detail": json.dumps(
                            dict(detail),
                            ensure_ascii=False,
                            sort_keys=True,
                            separators=(",", ":"),
                        ),
                        "sourcePath": str(root),
                        "timestamp": now_iso(),
                    }
                )
            except Exception:
                # Progress telemetry must never abort durable corpus work.
                pass

        run = {
            "runId": run_id,
            "status": RUNNING_STATUS,
            "startedAt": now,
            "finishedAt": "",
            "options": {
                "offset": offset,
                "limit": limit,
                "workbookWorkers": workbook_workers,
                "workbookRetryAttempts": workbook_retry_attempts,
                "pipelineWorkers": dict(pipeline_gates.limits),
                "retryFailed": retry_failed,
                "resume": resume,
                "inventoryOnly": inventory_only,
                "includeRelativePathCount": (
                    len(include_relative_paths)
                    if include_relative_paths is not None
                    else 0
                ),
                "deriveFormulaValues": bool(
                    options.get("derive_formula_values", False)
                ),
                "repairRejectedDraft": bool(
                    options.get("repair_rejected_draft", False)
                ),
                "repairUnselectedSource": bool(
                    options.get("repair_unselected_source", False)
                ),
                "imagesAnalyzed": False,
            },
            "selectedRecordIds": [
                str(record["recordId"]) for record in selected
            ],
        }
        journal.setdefault("runs", []).append(run)
        journal["status"] = RUNNING_STATUS
        journal["updatedAt"] = now
        _atomic_write_json(journal_file, journal)

        actions: dict[str, str] = {}
        eligible: list[dict[str, Any]] = []
        requested_formula_mode = bool(
            options.get("derive_formula_values", False)
        )
        for record in selected:
            record_id = str(record["recordId"])
            prior_result = record.get("result")
            prior_formula_mode = bool(
                isinstance(prior_result, dict)
                and isinstance(
                    prior_result.get("formulaDerivation"),
                    dict,
                )
                and prior_result["formulaDerivation"].get("enabled")
            )
            if record_id in reconciled:
                actions[record_id] = "RECONCILED_EXISTING"
            elif not record.get("contentSha256"):
                actions[record_id] = "SKIPPED_INVENTORY_ERROR"
            elif (
                record["status"] == COMPLETED_STATUS
                and prior_formula_mode == requested_formula_mode
            ):
                actions[record_id] = "SKIPPED_COMPLETED"
            elif record["status"] == FAILED_STATUS and not retry_failed:
                actions[record_id] = "SKIPPED_FAILED"
            else:
                actions[record_id] = "ATTEMPTED"
                eligible.append(record)

        emit_corpus_progress(
            "RUNNING",
            {
                "selected": len(selected),
                "eligible": len(eligible),
                "skippedCompleted": sum(
                    action == "SKIPPED_COMPLETED"
                    for action in actions.values()
                ),
                "skippedFailed": sum(
                    action == "SKIPPED_FAILED"
                    for action in actions.values()
                ),
                "reconciledExisting": sum(
                    action == "RECONCILED_EXISTING"
                    for action in actions.values()
                ),
                "pipelineWorkers": dict(pipeline_gates.limits),
            },
        )

        workbook_artifacts = artifacts / "workbooks"

        def process(record: Mapping[str, Any]) -> dict[str, Any]:
            source = Path(str(record["sourcePath"]))
            current_sha256 = _sha256_file(source)
            if current_sha256.lower() != str(
                record["contentSha256"]
            ).lower():
                raise CorpusWorkflowError(
                    "Source fingerprint changed after corpus discovery."
                )
            result: dict[str, Any] | None = None
            for workflow_attempt in range(
                1,
                workbook_retry_attempts + 1,
            ):
                try:
                    result = ingest_runner(
                        database_path=database,
                        source_path=source,
                        artifact_root=workbook_artifacts,
                        dataset=dataset,
                        resume=True,
                        pipeline_gate=pipeline_gates.acquire,
                        **options,
                    )
                except Exception:
                    if workflow_attempt >= workbook_retry_attempts:
                        raise
                    continue
                result["workflowRetryAttempts"] = workflow_attempt
                break
            if result is None:
                raise CorpusWorkflowError(
                    "Workbook ingest exhausted retry attempts without a result."
                )
            if result.get("imagesAnalyzed") is not False:
                raise CorpusWorkflowError(
                    "Workbook ingest did not preserve imagesAnalyzed=false."
                )
            result_formula = result.get("formulaDerivation")
            returned_formula_mode = bool(
                isinstance(result_formula, dict)
                and result_formula.get("enabled")
            )
            if returned_formula_mode != requested_formula_mode:
                raise CorpusWorkflowError(
                    "Workbook ingest returned a different formula-derivation "
                    "mode than the corpus run requested."
                )
            if str(result.get("contentSha256") or "").lower() != (
                current_sha256.lower()
            ):
                raise CorpusWorkflowError(
                    "Workbook ingest returned a different source fingerprint."
                )
            if _sha256_file(source).lower() != current_sha256.lower():
                raise CorpusWorkflowError(
                    "Source changed while corpus ingestion was running."
                )
            return result

        futures: dict[Future[dict[str, Any]], dict[str, Any]] = {}
        with ThreadPoolExecutor(max_workers=workbook_workers) as executor:
            for record in eligible:
                started = now_iso()
                record["status"] = RUNNING_STATUS
                record["attempts"] = int(record.get("attempts") or 0) + 1
                record["lastStartedAt"] = started
                record["lastFinishedAt"] = ""
                record["updatedAt"] = started
                record["error"] = ""
                future = executor.submit(process, record)
                futures[future] = record
                journal["updatedAt"] = started
                _atomic_write_json(journal_file, journal)
            for future in as_completed(futures):
                record = futures[future]
                record_id = str(record["recordId"])
                finished = now_iso()
                try:
                    result = future.result()
                except Exception as exc:
                    record["status"] = FAILED_STATUS
                    record["result"] = None
                    record["error"] = f"{type(exc).__name__}: {exc}"
                    actions[record_id] = "FAILED"
                else:
                    record["status"] = COMPLETED_STATUS
                    record["result"] = result
                    record["error"] = ""
                    actions[record_id] = "COMPLETED"
                record["lastFinishedAt"] = finished
                record["updatedAt"] = finished
                journal["updatedAt"] = finished
                _atomic_write_json(journal_file, journal)

        summary = _accounting(
            journal,
            current_records=current,
            selected_records=selected,
            actions=actions,
        )
        with closing(
            sqlite3.connect(str(database), timeout=60)
        ) as connection:
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA busy_timeout=60000")
            connection.execute("PRAGMA foreign_keys=ON")
            integrity = validate_knowledge_integrity(connection)
        summary["integrityOk"] = bool(integrity["ok"])
        final_status = (
            "COMPLETED_WITH_ERRORS"
            if summary["failedThisRun"]
            or summary["skippedFailed"]
            or summary["skippedInventoryError"]
            or not integrity["ok"]
            else COMPLETED_STATUS
        )
        finished = now_iso()
        run["status"] = final_status
        run["finishedAt"] = finished
        run["summary"] = summary
        journal["status"] = final_status
        journal["updatedAt"] = finished
        journal["lastSummary"] = summary
        _atomic_write_json(journal_file, journal)
        emit_corpus_progress(
            (
                "FAILED"
                if final_status == "COMPLETED_WITH_ERRORS"
                else "COMPLETED"
            ),
            summary,
        )

        return {
            "schemaVersion": CORPUS_WORKFLOW_SCHEMA_VERSION,
            "batchId": journal["batchId"],
            "runId": run_id,
            "status": final_status,
            "sourceRoot": str(root),
            "databasePath": str(database),
            "artifactDirectory": str(artifacts),
            "journalPath": str(journal_file),
            "pipelineWorkers": dict(pipeline_gates.limits),
            "imagesAnalyzed": False,
            "integrity": integrity,
            "summary": summary,
            "items": [
                {
                    "recordId": record["recordId"],
                    "sourcePath": record["sourcePath"],
                    "relativePath": record["relativePath"],
                    "contentSha256": record["contentSha256"],
                    "sizeBytes": record["sizeBytes"],
                    "mtimeNs": record["mtimeNs"],
                    "status": record["status"],
                    "attempts": record["attempts"],
                    "action": actions.get(
                        str(record["recordId"]),
                        "NOT_SELECTED",
                    ),
                    "result": record.get("result"),
                    "error": record.get("error") or "",
                    "imagesAnalyzed": False,
                }
                for record in selected
            ],
        }


__all__ = [
    "COMPLETED_STATUS",
    "CORPUS_WORKFLOW_SCHEMA_VERSION",
    "CorpusPipelineGates",
    "CorpusWorkflowBusyError",
    "CorpusWorkflowError",
    "FAILED_STATUS",
    "INTERRUPTED_STATUS",
    "PENDING_STATUS",
    "RUNNING_STATUS",
    "discover_excel_sources",
    "discover_xlsx_sources",
    "run_corpus_ingest",
]
