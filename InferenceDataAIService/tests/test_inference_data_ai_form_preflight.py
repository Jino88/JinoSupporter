from __future__ import annotations

import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from inference_data_ai_form_preflight import (
    _atomic_write_json,
    classify_form,
    extract_workbook_isolated,
    FormPreflightCancelled,
    form_similarity,
    run_form_preflight,
    signature_from_payload,
)


def payload(
    *,
    columns: int = 6,
    rows: int = 20,
    tokens: tuple[str, ...] = ("TEST", "MIN", "MAX", "RESULT"),
) -> dict:
    cells = [
        {
            "row": 1,
            "column": index + 1,
            "displayValue": value,
            "rawValue": value,
        }
        for index, value in enumerate(tokens)
    ]
    return {
        "workbook": {
            "status": "CAPTURED",
            "sheetCount": 1,
            "tabularSheetCount": 1,
            "sheets": [
                {
                    "title": "Report",
                    "hasTabularEvidence": True,
                    "usedBounds": {
                        "minRow": 1,
                        "rowCount": rows,
                        "columnCount": columns,
                    },
                    "contentBounds": {
                        "minRow": 1,
                        "rowCount": rows,
                        "columnCount": columns,
                    },
                    "mergeCount": 2,
                    "formulaCellCount": 0,
                    "cells": cells,
                }
            ],
        }
    }


class FormPreflightTests(unittest.TestCase):
    def test_unchanged_previous_failure_is_reused_without_com(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            archive = root / "archive"
            archive.mkdir()
            source = archive / "failed.xlsx"
            source.write_bytes(b"fixture")
            output = root / "latest.json"
            import hashlib

            digest = hashlib.sha256(b"fixture").hexdigest()
            output.write_text(
                json.dumps(
                    {
                        "items": [
                            {
                                "sourcePath": str(source.resolve()),
                                "relativePath": source.name,
                                "fileName": source.name,
                                "contentSha256": digest,
                                "sizeBytes": source.stat().st_size,
                                "captureAction": "FAILED",
                                "captureRevisionId": 0,
                                "status": "CAPTURE_FAILED",
                                "similarity": 0.0,
                                "nearestKnownSource": "",
                                "nearestKnownFormSignatureId": "",
                                "reason": "previous COM timeout",
                                "formSignatureId": "",
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )
            extractor = mock.Mock()

            result = run_form_preflight(
                database_path=root / "canonical.sqlite",
                source_root=archive,
                output_path=output,
                extractor=extractor,
            )

        extractor.assert_not_called()
        self.assertEqual("COMPLETED", result["status"])
        self.assertEqual(1, result["summary"]["captureFailed"])
        self.assertEqual(
            "REUSED_FAILED",
            result["items"][0]["captureAction"],
        )

    def test_existing_cancel_marker_stops_before_com(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            archive = root / "archive"
            archive.mkdir()
            (archive / "one.xlsx").write_bytes(b"fixture")
            cancel_file = root / "cancel.request"
            cancel_file.write_text("stop", encoding="utf-8")
            extractor = mock.Mock()

            result = run_form_preflight(
                database_path=root / "canonical.sqlite",
                source_root=archive,
                output_path=root / "latest.json",
                extractor=extractor,
                cancel_file=cancel_file,
            )

        self.assertEqual("CANCELLED", result["status"])
        self.assertEqual(0, result["summary"]["total"])
        extractor.assert_not_called()

    def test_cancel_kills_current_isolated_worker(self) -> None:
        class FakeProcess:
            def __init__(self) -> None:
                self.killed = False
                self.returncode: int | None = None

            def kill(self) -> None:
                self.killed = True
                self.returncode = -9

            def communicate(
                self,
                timeout: float | None = None,
            ) -> tuple[str, str]:
                return "", ""

            def poll(self) -> int | None:
                return self.returncode

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "fixture.xlsx"
            source.write_bytes(b"fixture")
            cancel_file = root / "cancel.request"
            cancel_file.write_text("stop", encoding="utf-8")
            process = FakeProcess()
            with mock.patch(
                "inference_data_ai_form_preflight.subprocess.Popen",
                return_value=process,
            ):
                with self.assertRaises(FormPreflightCancelled):
                    extract_workbook_isolated(
                        source,
                        scratch_root=root / "workers",
                        cancel_file=cancel_file,
                    )

        self.assertTrue(process.killed)

    def test_atomic_report_write_retries_transient_windows_lock(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output_path = Path(temporary) / "latest.json"
            real_replace = __import__("os").replace
            attempts = 0

            def transient_replace(
                source: str | Path,
                destination: str | Path,
            ) -> None:
                nonlocal attempts
                attempts += 1
                if attempts < 3:
                    raise PermissionError("temporary reader lock")
                real_replace(source, destination)

            with (
                mock.patch(
                    "inference_data_ai_form_preflight.os.replace",
                    side_effect=transient_replace,
                ),
                mock.patch(
                    "inference_data_ai_form_preflight.time.sleep"
                ),
            ):
                _atomic_write_json(
                    output_path,
                    {"status": "RUNNING"},
                )

            self.assertEqual(3, attempts)
            self.assertEqual(
                {"status": "RUNNING"},
                json.loads(output_path.read_text(encoding="utf-8")),
            )

    def test_isolated_com_worker_returns_payload_file(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "fixture.xlsx"
            source.write_bytes(b"fixture")
            expected = payload()

            def fake_run(arguments: list[str], **_: object):
                output_path = Path(
                    arguments[arguments.index("--out") + 1]
                )
                output_path.write_text(
                    json.dumps(expected),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess(
                    arguments,
                    0,
                    "",
                    "",
                )

            with mock.patch(
                "inference_data_ai_form_preflight.subprocess.run",
                side_effect=fake_run,
            ):
                actual = extract_workbook_isolated(
                    source,
                    scratch_root=root / "workers",
                    timeout_seconds=1,
                )

        self.assertEqual(expected, actual)

    def test_isolated_com_timeout_is_one_capture_error(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "fixture.xlsx"
            source.write_bytes(b"fixture")
            timeout = subprocess.TimeoutExpired(
                ["python", "worker"],
                1,
            )
            with (
                mock.patch(
                    "inference_data_ai_form_preflight.subprocess.run",
                    side_effect=timeout,
                ),
                mock.patch(
                    "inference_data_ai_form_preflight"
                    "._terminate_recorded_excel"
                ) as terminate,
            ):
                with self.assertRaisesRegex(
                    RuntimeError,
                    "timed out",
                ):
                    extract_workbook_isolated(
                        source,
                        scratch_root=root / "workers",
                        timeout_seconds=1,
                    )
            self.assertGreaterEqual(terminate.call_count, 1)

    def test_isolated_worker_failure_uses_short_last_error(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "fixture.xlsx"
            source.write_bytes(b"fixture")
            failure = subprocess.CompletedProcess(
                ["python", "worker"],
                1,
                "",
                "Traceback (most recent call last):\n"
                '  File "worker.py", line 3, in main\n'
                "    payload = capture()\n"
                "              ^^^^^^^^^\n"
                "IndexError: list index out of range\n",
            )
            with mock.patch(
                "inference_data_ai_form_preflight.subprocess.run",
                return_value=failure,
            ):
                with self.assertRaisesRegex(
                    RuntimeError,
                    r"IndexError: list index out of range$",
                ) as captured:
                    extract_workbook_isolated(
                        source,
                        scratch_root=root / "workers",
                    )

        self.assertNotIn("Traceback", str(captured.exception))

    def test_identical_form_is_known(self) -> None:
        known_signature = signature_from_payload(payload())
        result = classify_form(
            signature_from_payload(payload(rows=24)),
            [
                {
                    "sourcePath": "known.xlsx",
                    "signature": known_signature,
                }
            ],
        )
        self.assertEqual("KNOWN_FORM", result["status"])
        self.assertGreaterEqual(result["similarity"], 0.82)

    def test_different_shape_and_headers_are_new(self) -> None:
        known_signature = signature_from_payload(payload())
        candidate = signature_from_payload(
            payload(
                columns=30,
                rows=200,
                tokens=("VISION", "DEFECT", "POSITION", "IMAGE"),
            )
        )
        result = classify_form(
            candidate,
            [
                {
                    "sourcePath": "known.xlsx",
                    "signature": known_signature,
                }
            ],
        )
        self.assertEqual("NEW_FORM", result["status"])
        self.assertLess(result["similarity"], 0.60)

    def test_extra_sheet_reduces_similarity(self) -> None:
        known = signature_from_payload(payload())
        candidate_payload = payload()
        candidate_payload["workbook"]["sheetCount"] = 2
        candidate_payload["workbook"]["tabularSheetCount"] = 2
        candidate_payload["workbook"]["sheets"].append(
            {
                **candidate_payload["workbook"]["sheets"][0],
                "title": "Other",
            }
        )
        candidate = signature_from_payload(candidate_payload)
        self.assertLess(form_similarity(candidate, known), 0.82)


if __name__ == "__main__":
    unittest.main()
