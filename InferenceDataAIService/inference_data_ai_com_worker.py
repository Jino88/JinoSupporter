"""One-workbook Excel COM worker used as a native-crash boundary.

The parent process launches one worker per workbook.  If Excel or pywin32
hangs or terminates this worker, the parent can mark only that workbook as
failed and continue the archive preflight.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import traceback
import uuid
from pathlib import Path
from typing import Any

from inference_data_ai_com_capture import extract_workbook_com


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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--state", required=True)
    parser.add_argument(
        "--covered-cell-mode",
        choices=("blank", "anchor", "raw"),
        default="blank",
    )
    parser.add_argument("--exclude-hidden", action="store_true")
    parser.add_argument("--inspect-auth-dialog", action="store_true")
    parser.add_argument("--dismiss-auth-dialog", action="store_true")
    parser.add_argument("--auth-dialog-title", default="")
    parser.add_argument("--auth-dialog-class", default="")
    parser.add_argument("--auth-dialog-button", default="")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    source = Path(args.input).expanduser().resolve()
    output_path = Path(args.out).expanduser().resolve()
    state_path = Path(args.state).expanduser().resolve()

    def record_excel_process(process_id: int) -> None:
        _atomic_write_json(
            state_path,
            {
                "schemaVersion": "isolated-excel-com-state-v1",
                "workerProcessId": os.getpid(),
                "excelProcessId": process_id,
                "sourcePath": str(source),
            },
        )

    try:
        payload: dict[str, Any] = extract_workbook_com(
            source,
            covered_cell_mode=args.covered_cell_mode,
            include_hidden=not args.exclude_hidden,
            inspect_auth_dialog=args.inspect_auth_dialog,
            dismiss_auth_dialog=args.dismiss_auth_dialog,
            auth_dialog_title=args.auth_dialog_title,
            auth_dialog_class=args.auth_dialog_class,
            auth_dialog_button=args.auth_dialog_button,
            excel_process_callback=record_excel_process,
        )
        _atomic_write_json(output_path, payload)
        return 0
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
