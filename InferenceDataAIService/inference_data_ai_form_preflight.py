"""Read-only Excel COM capture and deterministic form-family preflight.

The preflight deliberately stops before semantic locator/draft AI.  Captures
are imported into Capture v2 so the later full workflow can reuse them without
opening the same DRM workbook through Excel COM a second time.
"""

from __future__ import annotations

import hashlib
import ctypes
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import time
import uuid
from contextlib import closing
from ctypes import wintypes
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable

from inference_data_ai_schema import ensure_knowledge_schema
from inference_data_ai_source_ingest import (
    COM_CAPTURE_CONTRACT,
    bridge_capture_to_canonical_source,
    ensure_capture_v2_schema,
    import_capture,
    sha256_file,
    verify_capture_revision,
)


SCHEMA_VERSION = "excel-form-preflight-v1"
CLASSIFIER_VERSION = "deterministic-form-family-v1"
KNOWN_FORM_THRESHOLD = 0.82
SIMILAR_FORM_THRESHOLD = 0.60
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
TOKEN_PATTERN = re.compile(r"[A-Z][A-Z0-9+/%._-]{1,}|[가-힣]{2,}")
DEFAULT_COM_TIMEOUT_SECONDS = 300.0
COM_WORKER_PATH = Path(__file__).with_name(
    "inference_data_ai_com_worker.py"
)


class FormPreflightCancelled(RuntimeError):
    """Raised when the caller requests a cooperative COM preflight stop."""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _atomic_write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(
        f".{path.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp"
    )
    try:
        temporary.write_text(
            json.dumps(value, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        for attempt in range(12):
            try:
                os.replace(temporary, path)
                break
            except PermissionError:
                if attempt == 11:
                    raise
                time.sleep(min(0.05 * (2**attempt), 0.5))
    finally:
        temporary.unlink(missing_ok=True)


def _terminate_recorded_excel(state_path: Path) -> None:
    """Terminate only the dedicated Excel PID recorded by our worker."""

    if sys.platform != "win32" or not state_path.is_file():
        return
    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
        if state.get("schemaVersion") != "isolated-excel-com-state-v1":
            return
        process_id = int(state.get("excelProcessId") or 0)
        if process_id <= 0:
            return
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.argtypes = [
            wintypes.DWORD,
            wintypes.BOOL,
            wintypes.DWORD,
        ]
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.QueryFullProcessImageNameW.argtypes = [
            wintypes.HANDLE,
            wintypes.DWORD,
            wintypes.LPWSTR,
            ctypes.POINTER(wintypes.DWORD),
        ]
        kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
        kernel32.TerminateProcess.argtypes = [
            wintypes.HANDLE,
            wintypes.UINT,
        ]
        kernel32.TerminateProcess.restype = wintypes.BOOL
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL
        process = kernel32.OpenProcess(
            0x0001 | 0x1000,
            False,
            process_id,
        )
        if not process:
            return
        try:
            size = ctypes.c_ulong(32768)
            image_path = ctypes.create_unicode_buffer(size.value)
            if not kernel32.QueryFullProcessImageNameW(
                process,
                0,
                image_path,
                ctypes.byref(size),
            ):
                return
            if Path(image_path.value).name.casefold() != "excel.exe":
                return
            kernel32.TerminateProcess(process, 1)
        finally:
            kernel32.CloseHandle(process)
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return


def extract_workbook_isolated(
    source: str | Path,
    *,
    scratch_root: str | Path,
    timeout_seconds: float = DEFAULT_COM_TIMEOUT_SECONDS,
    covered_cell_mode: str = "blank",
    include_hidden: bool = True,
    inspect_auth_dialog: bool = False,
    dismiss_auth_dialog: bool = False,
    auth_dialog_title: str = "",
    auth_dialog_class: str = "",
    auth_dialog_button: str = "",
    cancel_file: str | Path | None = None,
) -> dict[str, Any]:
    """Run one COM extraction in a disposable Python/Excel process pair."""

    root = Path(scratch_root).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(
        prefix="excel-com-",
        dir=root,
    ) as temporary:
        temporary_root = Path(temporary)
        output_path = temporary_root / "capture.json"
        state_path = temporary_root / "state.json"
        arguments = [
            sys.executable,
            str(COM_WORKER_PATH),
            "--input",
            str(Path(source).expanduser().resolve()),
            "--out",
            str(output_path),
            "--state",
            str(state_path),
            "--covered-cell-mode",
            covered_cell_mode,
        ]
        if not include_hidden:
            arguments.append("--exclude-hidden")
        if inspect_auth_dialog:
            arguments.append("--inspect-auth-dialog")
        if dismiss_auth_dialog:
            arguments.append("--dismiss-auth-dialog")
            arguments.extend(
                [
                    "--auth-dialog-title",
                    auth_dialog_title,
                    "--auth-dialog-class",
                    auth_dialog_class,
                    "--auth-dialog-button",
                    auth_dialog_button,
                ]
            )
        process_options = {
            "cwd": Path(__file__).parent,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "creationflags": (
                subprocess.CREATE_NO_WINDOW
                if sys.platform == "win32"
                else 0
            ),
        }
        try:
            if cancel_file is None:
                completed = subprocess.run(
                    arguments,
                    capture_output=True,
                    timeout=timeout_seconds,
                    check=False,
                    **process_options,
                )
            else:
                completed = _run_cancellable_worker(
                    arguments,
                    cancel_file=Path(cancel_file),
                    timeout_seconds=timeout_seconds,
                    **process_options,
                )
        except subprocess.TimeoutExpired as exc:
            _terminate_recorded_excel(state_path)
            raise RuntimeError(
                "Excel COM extraction timed out after "
                f"{timeout_seconds:g} seconds."
            ) from exc
        finally:
            # DispatchEx creates a dedicated Excel instance.  Clean up an
            # orphan if the worker or Excel terminated before normal Quit.
            _terminate_recorded_excel(state_path)
        if completed.returncode != 0:
            raw_detail = (completed.stderr or completed.stdout).strip()
            detail_lines = [
                line.strip()
                for line in raw_detail.splitlines()
                if line.strip()
                and not line.lstrip().startswith(
                    ("File ", "Traceback ", "^^^^")
                )
            ]
            detail = (
                detail_lines[-1]
                if detail_lines
                else "worker terminated without an error message"
            )
            if len(detail) > 420:
                detail = detail[:417] + "..."
            raise RuntimeError(
                f"Isolated Excel COM worker failed "
                f"(exit {completed.returncode}): {detail}"
            )
        if not output_path.is_file():
            raise RuntimeError(
                "Isolated Excel COM worker returned no capture payload."
            )
        return json.loads(output_path.read_text(encoding="utf-8"))


def _run_cancellable_worker(
    arguments: list[str],
    *,
    cancel_file: Path,
    timeout_seconds: float,
    **process_options: Any,
) -> subprocess.CompletedProcess[str]:
    process = subprocess.Popen(
        arguments,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        **process_options,
    )
    started = time.monotonic()
    try:
        while True:
            if cancel_file.is_file():
                process.kill()
                process.communicate()
                raise FormPreflightCancelled(
                    "Excel COM preflight stop requested."
                )
            elapsed = time.monotonic() - started
            remaining = timeout_seconds - elapsed
            if remaining <= 0:
                process.kill()
                process.communicate()
                raise subprocess.TimeoutExpired(
                    arguments,
                    timeout_seconds,
                )
            try:
                stdout, stderr = process.communicate(
                    timeout=min(0.25, remaining)
                )
                return subprocess.CompletedProcess(
                    arguments,
                    int(process.returncode or 0),
                    stdout,
                    stderr,
                )
            except subprocess.TimeoutExpired:
                continue
    finally:
        if process.poll() is None:
            process.kill()
            process.communicate()


def _bucket(value: int) -> str:
    if value <= 0:
        return "0"
    if value <= 3:
        return "1-3"
    if value <= 8:
        return "4-8"
    if value <= 16:
        return "9-16"
    if value <= 32:
        return "17-32"
    if value <= 64:
        return "33-64"
    if value <= 128:
        return "65-128"
    if value <= 256:
        return "129-256"
    return "257+"


def _portable_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, dict) and isinstance(value.get("value"), str):
        return str(value["value"])
    return ""


def _tokens(values: Iterable[Any]) -> list[str]:
    result: set[str] = set()
    for value in values:
        text = _portable_text(value).upper()
        for match in TOKEN_PATTERN.finditer(text):
            token = match.group(0).strip("._-")
            if len(token) >= 2 and not token.isdigit():
                result.add(token)
    return sorted(result)[:120]


def _bounds_shape(bounds: dict[str, Any] | None) -> tuple[int, int]:
    if not bounds:
        return 0, 0
    return int(bounds.get("rowCount") or 0), int(
        bounds.get("columnCount") or 0
    )


def signature_from_payload(payload: dict[str, Any]) -> dict[str, Any]:
    workbook = payload["workbook"]
    profiles: list[dict[str, Any]] = []
    for sheet in workbook.get("sheets", []):
        rows, columns = _bounds_shape(sheet.get("contentBounds") or sheet.get("usedBounds"))
        min_row = int((sheet.get("usedBounds") or {}).get("minRow") or 1)
        header_values = [
            cell.get("displayValue")
            if cell.get("displayValue") is not None
            else cell.get("rawValue")
            for cell in sheet.get("cells", [])
            if int(cell.get("row") or 0) <= min_row + 79
        ]
        profiles.append(
            {
                "tabular": bool(sheet.get("hasTabularEvidence")),
                "rowBucket": _bucket(rows),
                "columnBucket": _bucket(columns),
                "rows": rows,
                "columns": columns,
                "mergeBucket": _bucket(int(sheet.get("mergeCount") or 0)),
                "formulaBucket": _bucket(int(sheet.get("formulaCellCount") or 0)),
                "tokens": _tokens(
                    [sheet.get("title") or "", *header_values]
                ),
            }
        )
    return _finalize_signature(
        str(workbook.get("status") or ""),
        int(workbook.get("sheetCount") or len(profiles)),
        int(workbook.get("tabularSheetCount") or 0),
        profiles,
    )


def signature_from_database(
    connection: sqlite3.Connection,
    capture_revision_id: int,
) -> dict[str, Any]:
    workbook = connection.execute(
        """
        SELECT workbook_status, sheet_count, tabular_sheet_count
        FROM capture_v2_workbooks
        WHERE revision_id=?
        """,
        (capture_revision_id,),
    ).fetchone()
    if workbook is None:
        raise ValueError(
            f"Capture workbook is missing: {capture_revision_id}"
        )
    profiles: list[dict[str, Any]] = []
    sheets = connection.execute(
        """
        SELECT sheet_id, title, has_tabular_evidence, formula_cell_count,
               merge_count, used_bounds_json, content_bounds_json
        FROM capture_v2_sheets
        WHERE revision_id=?
        ORDER BY sheet_index
        """,
        (capture_revision_id,),
    ).fetchall()
    for sheet in sheets:
        used = json.loads(sheet["used_bounds_json"]) if sheet["used_bounds_json"] else None
        content = json.loads(sheet["content_bounds_json"]) if sheet["content_bounds_json"] else None
        rows, columns = _bounds_shape(content or used)
        min_row = int((used or {}).get("minRow") or 1)
        cell_values: list[Any] = []
        for cell in connection.execute(
            """
            SELECT display_value_json, raw_value_json
            FROM capture_v2_cells
            WHERE sheet_id=? AND row_index<=?
            ORDER BY row_index, column_index
            """,
            (int(sheet["sheet_id"]), min_row + 79),
        ):
            raw = (
                cell["display_value_json"]
                if cell["display_value_json"] is not None
                else cell["raw_value_json"]
            )
            if raw is None:
                continue
            try:
                cell_values.append(json.loads(raw))
            except json.JSONDecodeError:
                continue
        profiles.append(
            {
                "tabular": bool(sheet["has_tabular_evidence"]),
                "rowBucket": _bucket(rows),
                "columnBucket": _bucket(columns),
                "rows": rows,
                "columns": columns,
                "mergeBucket": _bucket(int(sheet["merge_count"])),
                "formulaBucket": _bucket(int(sheet["formula_cell_count"])),
                "tokens": _tokens([sheet["title"], *cell_values]),
            }
        )
    return _finalize_signature(
        str(workbook["workbook_status"]),
        int(workbook["sheet_count"]),
        int(workbook["tabular_sheet_count"]),
        profiles,
    )


def _finalize_signature(
    workbook_status: str,
    sheet_count: int,
    tabular_sheet_count: int,
    profiles: list[dict[str, Any]],
) -> dict[str, Any]:
    family_value = {
        "workbookStatus": workbook_status,
        "sheetCount": sheet_count,
        "tabularSheetCount": tabular_sheet_count,
        "sheets": [
            {
                "tabular": profile["tabular"],
                "rowBucket": profile["rowBucket"],
                "columnBucket": profile["columnBucket"],
                "mergeBucket": profile["mergeBucket"],
                "formulaBucket": profile["formulaBucket"],
                "tokens": profile["tokens"],
            }
            for profile in profiles
        ],
    }
    canonical = json.dumps(
        family_value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return {
        **family_value,
        "formSignatureId": "form-" + hashlib.sha256(
            canonical.encode("utf-8")
        ).hexdigest()[:16],
        "sheetProfiles": profiles,
    }


def _ratio(left: int, right: int) -> float:
    if left == right:
        return 1.0
    if left <= 0 or right <= 0:
        return 0.0
    return min(left, right) / max(left, right)


def _jaccard(left: Iterable[str], right: Iterable[str]) -> float:
    a, b = set(left), set(right)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _sheet_similarity(
    candidate: dict[str, Any],
    known: dict[str, Any],
) -> float:
    shape = (
        0.45 * _ratio(int(candidate["columns"]), int(known["columns"]))
        + 0.20 * _ratio(int(candidate["rows"]), int(known["rows"]))
        + 0.10 * (
            1.0
            if candidate["mergeBucket"] == known["mergeBucket"]
            else 0.0
        )
        + 0.10 * (
            1.0
            if candidate["formulaBucket"] == known["formulaBucket"]
            else 0.0
        )
        + 0.15 * (
            1.0 if candidate["tabular"] == known["tabular"] else 0.0
        )
    )
    tokens = _jaccard(candidate["tokens"], known["tokens"])
    return 0.48 * shape + 0.52 * tokens


def form_similarity(
    candidate: dict[str, Any],
    known: dict[str, Any],
) -> float:
    candidate_sheets = list(candidate.get("sheetProfiles") or [])
    known_sheets = list(known.get("sheetProfiles") or [])
    if not candidate_sheets and not known_sheets:
        sheet_score = 1.0
    elif not candidate_sheets or not known_sheets:
        sheet_score = 0.0
    else:
        remaining = set(range(len(known_sheets)))
        scores: list[float] = []
        for sheet in candidate_sheets:
            if not remaining:
                scores.append(0.0)
                continue
            best = max(
                remaining,
                key=lambda index: _sheet_similarity(
                    sheet,
                    known_sheets[index],
                ),
            )
            scores.append(
                _sheet_similarity(sheet, known_sheets[best])
            )
            remaining.remove(best)
        scores.extend(0.0 for _ in remaining)
        sheet_score = sum(scores) / max(
            len(candidate_sheets),
            len(known_sheets),
        )
    workbook_score = (
        0.45
        * _ratio(
            int(candidate.get("sheetCount") or 0),
            int(known.get("sheetCount") or 0),
        )
        + 0.40
        * _ratio(
            int(candidate.get("tabularSheetCount") or 0),
            int(known.get("tabularSheetCount") or 0),
        )
        + 0.15
        * (
            1.0
            if candidate.get("workbookStatus")
            == known.get("workbookStatus")
            else 0.0
        )
    )
    return round(0.20 * workbook_score + 0.80 * sheet_score, 4)


def classify_form(
    signature: dict[str, Any],
    known_forms: list[dict[str, Any]],
) -> dict[str, Any]:
    if not known_forms:
        return {
            "status": "NEW_FORM",
            "similarity": 0.0,
            "nearestKnownSource": "",
            "nearestKnownFormSignatureId": "",
            "reason": "비교할 기존 분석 양식이 없습니다.",
        }
    nearest = max(
        known_forms,
        key=lambda item: form_similarity(signature, item["signature"]),
    )
    similarity = form_similarity(signature, nearest["signature"])
    if similarity >= KNOWN_FORM_THRESHOLD:
        status = "KNOWN_FORM"
        reason = "기존 분석 양식과 구조·헤더가 충분히 유사합니다."
    elif similarity >= SIMILAR_FORM_THRESHOLD:
        status = "SIMILAR_FORM_REVIEW"
        reason = "기존 양식과 유사하지만 전체 처리 전 검토가 필요합니다."
    else:
        status = "NEW_FORM"
        reason = "기존 분석 양식과 구조 차이가 커서 AI 전체 처리를 보류합니다."
    return {
        "status": status,
        "similarity": similarity,
        "nearestKnownSource": nearest["sourcePath"],
        "nearestKnownFormSignatureId": nearest["signature"][
            "formSignatureId"
        ],
        "reason": reason,
    }


def load_known_forms(connection: sqlite3.Connection) -> list[dict[str, Any]]:
    rows = connection.execute(
        """
        SELECT DISTINCT capture.revision_id, document.source_path
        FROM workbook_analyses analysis
        JOIN knowledge_studies study
          ON study.workbook_analysis_id=analysis.workbook_analysis_id
        JOIN source_revisions revision
          ON revision.revision_id=analysis.revision_id
        JOIN source_documents document
          ON document.document_id=revision.document_id
        JOIN capture_v2_revisions capture
          ON capture.revision_id=revision.capture_v2_revision_id
        WHERE revision.is_current=1
          AND capture.is_current=1
          AND study.analysis_status<>'FAILED'
        ORDER BY document.source_path
        """
    ).fetchall()
    return [
        {
            "captureRevisionId": int(row["revision_id"]),
            "sourcePath": str(row["source_path"]),
            "signature": signature_from_database(
                connection,
                int(row["revision_id"]),
            ),
        }
        for row in rows
    ]


def _existing_capture(
    connection: sqlite3.Connection,
    source_path: Path,
    content_sha256: str,
) -> int | None:
    row = connection.execute(
        """
        SELECT revision.revision_id
        FROM capture_v2_documents document
        JOIN capture_v2_revisions revision
          ON revision.document_id=document.document_id
         AND revision.is_current=1
        WHERE document.source_path=?
          AND revision.content_sha256=?
          AND revision.capture_contract=?
        """,
        (str(source_path), content_sha256, COM_CAPTURE_CONTRACT),
    ).fetchone()
    return int(row["revision_id"]) if row is not None else None


def discover_excel_files(source_root: str | Path) -> list[Path]:
    root = Path(source_root).expanduser().resolve()
    if not root.is_dir():
        raise FileNotFoundError(root)
    return sorted(
        (
            path.resolve()
            for path in root.iterdir()
            if path.is_file()
            and path.suffix.lower() in EXCEL_EXTENSIONS
            and not path.name.startswith("~$")
        ),
        key=lambda value: str(value).casefold(),
    )


def run_form_preflight(
    *,
    database_path: str | Path,
    source_root: str | Path,
    output_path: str | Path,
    dataset: str = "InputDataFinish",
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    extractor: Callable[..., dict[str, Any]] | None = None,
    com_timeout_seconds: float = DEFAULT_COM_TIMEOUT_SECONDS,
    inspect_auth_dialog: bool = False,
    dismiss_auth_dialog: bool = False,
    auth_dialog_title: str = "",
    auth_dialog_class: str = "",
    auth_dialog_button: str = "",
    cancel_file: str | Path | None = None,
    retry_failed_captures: bool = False,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    root = Path(source_root).expanduser().resolve()
    report_path = Path(output_path).expanduser().resolve()
    cancel_path = (
        Path(cancel_file).expanduser().resolve()
        if cancel_file is not None
        else None
    )
    prior_failed: dict[str, dict[str, Any]] = {}
    if report_path.is_file() and not retry_failed_captures:
        try:
            prior_report = json.loads(
                report_path.read_text(encoding="utf-8")
            )
            for prior_item in prior_report.get("items") or []:
                if (
                    isinstance(prior_item, dict)
                    and str(prior_item.get("status") or "")
                    == "CAPTURE_FAILED"
                    and str(prior_item.get("sourcePath") or "")
                ):
                    prior_failed[
                        str(prior_item["sourcePath"]).casefold()
                    ] = dict(prior_item)
        except (OSError, ValueError, json.JSONDecodeError):
            prior_failed = {}
    files = discover_excel_files(root)
    cancelled = False
    with closing(sqlite3.connect(database)) as connection:
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys=ON")
        ensure_knowledge_schema(connection, utc_now_iso)
        ensure_capture_v2_schema(connection)
        from inference_data_ai_form_registry import (
            apply_registry_decision,
            family_descriptor,
            load_family_decisions,
        )

        known_forms = load_known_forms(connection)
        registry_decisions = load_family_decisions(connection)
        items: list[dict[str, Any]] = []
        for index, source in enumerate(files, start=1):
            if cancel_path is not None and cancel_path.is_file():
                cancelled = True
                break
            source_sha256 = sha256_file(source)
            if progress_callback is not None:
                progress_callback(
                    {
                        "schemaVersion": "ingest-progress-v1",
                        "stage": "FORM_PREFLIGHT",
                        "status": "RUNNING",
                        "detail": f"{index}/{len(files)} · {source.name}",
                        "sourcePath": str(source),
                        "timestamp": utc_now_iso(),
                    }
                )
            prior_failure = prior_failed.get(
                str(source).casefold()
            )
            if (
                prior_failure is not None
                and str(
                    prior_failure.get("contentSha256") or ""
                ).lower()
                == source_sha256.lower()
            ):
                item = {
                    **prior_failure,
                    "sourcePath": str(source),
                    "relativePath": source.relative_to(root).as_posix(),
                    "fileName": source.name,
                    "contentSha256": source_sha256,
                    "sizeBytes": source.stat().st_size,
                    "captureAction": "REUSED_FAILED",
                }
                items.append(item)
                _write_report(
                    report_path,
                    database,
                    root,
                    known_forms,
                    items,
                    complete=index == len(files),
                )
                if progress_callback is not None:
                    progress_callback(
                        {
                            "schemaVersion": "ingest-progress-v1",
                            "stage": "FORM_PREFLIGHT",
                            "status": "CAPTURE_FAILED",
                            "detail": (
                                f"{index}/{len(files)} · "
                                f"CAPTURE_FAILED 재사용 · {source.name}"
                            ),
                            "sourcePath": str(source),
                            "timestamp": utc_now_iso(),
                        }
                    )
                continue
            try:
                revision_id = _existing_capture(
                    connection,
                    source,
                    source_sha256,
                )
                if revision_id is None:
                    capture_options = {
                        "covered_cell_mode": "blank",
                        "include_hidden": True,
                        "inspect_auth_dialog": inspect_auth_dialog,
                        "dismiss_auth_dialog": dismiss_auth_dialog,
                        "auth_dialog_title": auth_dialog_title,
                        "auth_dialog_class": auth_dialog_class,
                        "auth_dialog_button": auth_dialog_button,
                    }
                    if extractor is not None:
                        payload = extractor(
                            source,
                            **capture_options,
                        )
                    else:
                        payload = extract_workbook_isolated(
                            source,
                            scratch_root=report_path.parent
                            / ".com-workers",
                            timeout_seconds=com_timeout_seconds,
                            cancel_file=cancel_path,
                            **capture_options,
                        )
                    capture = import_capture(
                        connection,
                        payload,
                        captured_at=utc_now_iso(),
                    )
                    bridge_capture_to_canonical_source(
                        connection,
                        dataset=dataset,
                        payload=payload,
                        capture_result=capture,
                        captured_at=utc_now_iso(),
                    )
                    revision_id = int(capture["revisionId"])
                    verification = verify_capture_revision(
                        connection,
                        revision_id,
                        verify_source_sha256=True,
                    )
                    if not verification["ok"]:
                        raise RuntimeError(
                            "Capture verification failed: "
                            + "; ".join(verification["errors"])
                        )
                    capture_action = "CAPTURED_COM"
                    signature = signature_from_payload(payload)
                else:
                    capture_action = "REUSED_CAPTURE"
                    signature = signature_from_database(
                        connection,
                        revision_id,
                    )
                family = family_descriptor(signature)
                classification = apply_registry_decision(
                    classify_form(
                        signature,
                        known_forms,
                    ),
                    family=family,
                    decisions=registry_decisions,
                )
                item = {
                    "sourcePath": str(source),
                    "relativePath": source.relative_to(root).as_posix(),
                    "fileName": source.name,
                    "contentSha256": source_sha256,
                    "sizeBytes": source.stat().st_size,
                    "captureAction": capture_action,
                    "captureRevisionId": revision_id,
                    **classification,
                    "formSignatureId": signature["formSignatureId"],
                    "formFamilyId": family["familyId"],
                }
                connection.commit()
            except FormPreflightCancelled:
                connection.rollback()
                cancelled = True
                break
            except Exception as exc:
                connection.rollback()
                item = {
                    "sourcePath": str(source),
                    "relativePath": source.relative_to(root).as_posix(),
                    "fileName": source.name,
                    "contentSha256": source_sha256,
                    "sizeBytes": source.stat().st_size,
                    "captureAction": "FAILED",
                    "captureRevisionId": 0,
                    "status": "CAPTURE_FAILED",
                    "similarity": 0.0,
                    "nearestKnownSource": "",
                    "nearestKnownFormSignatureId": "",
                    "reason": f"{type(exc).__name__}: {exc}",
                    "formSignatureId": "",
                }
            items.append(item)
            _write_report(
                report_path,
                database,
                root,
                known_forms,
                items,
                complete=index == len(files),
            )
            if progress_callback is not None:
                progress_callback(
                    {
                        "schemaVersion": "ingest-progress-v1",
                        "stage": "FORM_PREFLIGHT",
                        "status": item["status"],
                        "detail": (
                            f"{index}/{len(files)} · "
                            f"{item['status']} · {source.name}"
                        ),
                        "sourcePath": str(source),
                        "timestamp": utc_now_iso(),
                    }
                )
        if cancelled and progress_callback is not None:
            progress_callback(
                {
                    "schemaVersion": "ingest-progress-v1",
                    "stage": "FORM_PREFLIGHT",
                    "status": "CANCELLED",
                    "detail": (
                        f"사용자 중지 · 완료 {len(items)}/{len(files)}"
                    ),
                    "sourcePath": "",
                    "timestamp": utc_now_iso(),
                }
            )
    return _write_report(
        report_path,
        database,
        root,
        known_forms,
        items,
        complete=not cancelled,
        cancelled=cancelled,
    )


def _write_report(
    report_path: Path,
    database: Path,
    root: Path,
    known_forms: list[dict[str, Any]],
    items: list[dict[str, Any]],
    *,
    complete: bool,
    cancelled: bool = False,
) -> dict[str, Any]:
    counts = {
        status: sum(item["status"] == status for item in items)
        for status in (
            "KNOWN_FORM",
            "SIMILAR_FORM_REVIEW",
            "NEW_FORM",
            "EXCLUDED_FORM",
            "CAPTURE_FAILED",
        )
    }
    manifest_path = report_path.with_name(
        report_path.stem + ".known-forms.manifest.json"
    )
    manifest = {
        "schemaVersion": "excel-form-preflight-manifest-v1",
        "sourceRoot": str(root),
        "workbooks": [
            {
                "relativePath": item["relativePath"],
                "contentSha256": item["contentSha256"],
                "formSignatureId": item["formSignatureId"],
                "preflightStatus": item["status"],
                "formFamilyId": str(
                    item.get("formFamilyId") or ""
                ),
                "registryDecision": str(
                    item.get("registryDecision") or ""
                ),
                **(
                    {
                        "extractionContract": item[
                            "extractionContract"
                        ]
                    }
                    if item.get("extractionContract")
                    else {}
                ),
            }
            for item in items
            if item["status"] == "KNOWN_FORM"
        ],
    }
    _atomic_write_json(manifest_path, manifest)
    report = {
        "schemaVersion": SCHEMA_VERSION,
        "classifierVersion": CLASSIFIER_VERSION,
        "status": (
            "CANCELLED"
            if cancelled
            else "COMPLETED"
            if complete
            else "RUNNING"
        ),
        "generatedAt": utc_now_iso(),
        "databasePath": str(database),
        "sourceRoot": str(root),
        "knownCatalogCount": len(known_forms),
        "summary": {
            "total": len(items),
            "knownForms": counts["KNOWN_FORM"],
            "similarReview": counts["SIMILAR_FORM_REVIEW"],
            "newForms": counts["NEW_FORM"],
            "excludedForms": counts["EXCLUDED_FORM"],
            "captureFailed": counts["CAPTURE_FAILED"],
            "fullProcessingAllowed": (
                complete
                and not cancelled
                and counts["KNOWN_FORM"] > 0
            ),
        },
        "knownFormManifestPath": str(manifest_path),
        "items": items,
    }
    _atomic_write_json(report_path, report)
    return report


__all__ = [
    "CLASSIFIER_VERSION",
    "FormPreflightCancelled",
    "KNOWN_FORM_THRESHOLD",
    "SCHEMA_VERSION",
    "SIMILAR_FORM_THRESHOLD",
    "classify_form",
    "discover_excel_files",
    "extract_workbook_isolated",
    "form_similarity",
    "load_known_forms",
    "run_form_preflight",
    "signature_from_database",
    "signature_from_payload",
]
