from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import time
from pathlib import Path
from typing import Any, Callable


try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


COM_CALL_REJECTED_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED: "Call was rejected by callee."
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
}

OFFICE_BLOCKING_WINDOW_MARKERS = (
    "microsoft office",
    "microsoft office activation wizard",
    "office activation wizard",
)


def office_blocking_windows(close: bool = False) -> list[str]:
    try:
        import ctypes
    except Exception:
        return []

    user32 = ctypes.windll.user32
    found: list[str] = []

    enum_proc_type = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

    def callback(hwnd: int, _: int) -> bool:
        try:
            if not user32.IsWindowVisible(hwnd):
                return True

            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return True

            buffer = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buffer, length + 1)
            title = buffer.value.strip()
            lowered = title.lower()
            if any(marker in lowered for marker in OFFICE_BLOCKING_WINDOW_MARKERS):
                found.append(title)
                if close:
                    user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
        except Exception:
            pass
        return True

    user32.EnumWindows(enum_proc_type(callback), 0)
    return found


def com_hresult(exc: BaseException) -> int | None:
    value = getattr(exc, "hresult", None)
    if isinstance(value, int):
        return value
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def is_retryable_com_error(exc: BaseException) -> bool:
    hresult = com_hresult(exc)
    if hresult in COM_CALL_REJECTED_HRESULTS:
        return True

    message = str(exc).lower()
    return (
        "call was rejected by callee" in message
        or "servercall_retrylater" in message
        or "rpc_e_call_rejected" in message
    )


def pump_com_messages() -> None:
    try:
        import pythoncom

        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def com_retry(label: str, func: Callable[[], Any], attempts: int = 5, delay_seconds: float = 0.25) -> Any:
    last: BaseException | None = None
    activation_blocked = False
    for attempt in range(1, attempts + 1):
        try:
            return func()
        except Exception as exc:
            retryable = is_retryable_com_error(exc)
            # When Office's activation wizard is modal, pywin32 can surface
            # an RPC rejection as AttributeError (for example
            # ``__call__.Name``) while resolving a returned worksheet proxy.
            # Treat that as transient only when the blocking window is
            # actually present, then allow the close operation a few retries
            # to take effect.
            blocking = office_blocking_windows(close=True)
            if blocking:
                activation_blocked = True
            if not retryable and not activation_blocked:
                raise
            last = exc
            if blocking:
                print(
                    json.dumps(
                        {
                            "status": "warn",
                            "message": "Office activation dialog is blocking Excel COM; closing it and retrying.",
                            "windows": blocking,
                            "while": label,
                        },
                        ensure_ascii=False,
                    ),
                    file=sys.stderr,
                    flush=True,
                )
            pump_com_messages()
            time.sleep(min(2.0, delay_seconds * attempt))

    hresult = com_hresult(last) if last is not None else None
    blocking = office_blocking_windows(close=False)
    if blocking:
        raise RuntimeError(
            "Microsoft Office activation dialog is blocking Excel COM. "
            "Activate Office or close the dialog, then run extraction again. "
            f"blockingWindows={blocking}, while={label}, hresult={hresult}"
        ) from last

    raise RuntimeError(f"Excel COM call was rejected after {attempts} attempts while {label}. hresult={hresult}") from last


def com_get(obj: Any, attr: str, label: str | None = None) -> Any:
    return com_retry(label or attr, lambda: getattr(obj, attr))


def com_set(obj: Any, attr: str, value: Any, label: str | None = None) -> None:
    def assign() -> None:
        setattr(obj, attr, value)

    com_retry(label or attr, assign)


def com_call(label: str, func: Callable[[], Any]) -> Any:
    return com_retry(label, func)


def excel_col_label(number: int) -> str:
    label = ""
    while number > 0:
        number, rem = divmod(number - 1, 26)
        label = chr(65 + rem) + label
    return label


def cell_address(row: int, column: int) -> str:
    return f"{excel_col_label(column)}{row}"


def range_address(top: int, left: int, bottom: int, right: int) -> str:
    if top == bottom and left == right:
        return cell_address(top, left)
    return f"{cell_address(top, left)}:{cell_address(bottom, right)}"


def json_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def value_text(value: Any) -> str:
    value = json_value(value)
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def to_matrix(raw: Any, row_count: int, col_count: int) -> list[list[Any]]:
    if row_count <= 0 or col_count <= 0:
        return []

    if row_count == 1 and col_count == 1:
        return [[raw]]

    if row_count == 1:
        row = list(raw) if isinstance(raw, tuple) else [raw]
        return [row + [None] * max(0, col_count - len(row))]

    if col_count == 1:
        rows = list(raw) if isinstance(raw, tuple) else [raw]
        return [[item] for item in rows] + [[None] for _ in range(max(0, row_count - len(rows)))]

    rows = []
    raw_rows = list(raw) if isinstance(raw, tuple) else []
    for r_idx in range(row_count):
        if r_idx < len(raw_rows):
            raw_row = raw_rows[r_idx]
            row = list(raw_row) if isinstance(raw_row, tuple) else [raw_row]
        else:
            row = []
        rows.append(row + [None] * max(0, col_count - len(row)))
    return rows


def visible_sheet(ws: Any) -> bool:
    try:
        return int(com_get(ws, "Visible", "read worksheet visibility")) == -1
    except Exception:
        return True


def range_is_definitely_unmerged(value: Any) -> bool:
    """Return True only when Excel explicitly reports no merged cells.

    Excel COM returns ``None``/VT_NULL for a range that contains a mixture of
    merged and ordinary cells.  That is an unknown/mixed state, not evidence
    that the range has no merges, so callers must still scan the individual
    cells in that case.
    """

    return value is False or (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and value == 0
    )


def read_merge_map(ws: Any, used: Any, values: list[list[Any]], covered_cell_mode: str) -> tuple[list[dict[str, Any]], dict[tuple[int, int], dict[str, Any]]]:
    merges: list[dict[str, Any]] = []
    merge_map: dict[tuple[int, int], dict[str, Any]] = {}
    seen: set[tuple[int, int, int, int]] = set()

    try:
        has_merges = com_get(used, "MergeCells", "read UsedRange.MergeCells")
    except Exception:
        has_merges = None

    if range_is_definitely_unmerged(has_merges):
        return merges, merge_map

    top = int(com_get(used, "Row", "read UsedRange.Row"))
    left = int(com_get(used, "Column", "read UsedRange.Column"))
    used_rows = com_get(used, "Rows", "read UsedRange.Rows")
    used_cols = com_get(used, "Columns", "read UsedRange.Columns")
    row_count = int(com_get(used_rows, "Count", "read UsedRange.Rows.Count"))
    col_count = int(com_get(used_cols, "Count", "read UsedRange.Columns.Count"))

    for r_offset in range(row_count):
        abs_row = top + r_offset
        for c_offset in range(col_count):
            abs_col = left + c_offset
            try:
                cell = com_call(f"read cell {abs_row},{abs_col}", lambda r=abs_row, c=abs_col: ws.Cells(r, c))
                if not com_get(cell, "MergeCells", f"read cell {abs_row},{abs_col}.MergeCells"):
                    continue
                area = com_get(cell, "MergeArea", f"read cell {abs_row},{abs_col}.MergeArea")
                m_top = int(com_get(area, "Row", "read merge area Row"))
                m_left = int(com_get(area, "Column", "read merge area Column"))
                area_rows = com_get(area, "Rows", "read merge area Rows")
                area_cols = com_get(area, "Columns", "read merge area Columns")
                m_bottom = m_top + int(com_get(area_rows, "Count", "read merge area Rows.Count")) - 1
                m_right = m_left + int(com_get(area_cols, "Count", "read merge area Columns.Count")) - 1
            except Exception:
                continue

            key = (m_top, m_left, m_bottom, m_right)
            if key in seen:
                continue
            seen.add(key)

            anchor_r_offset = m_top - top
            anchor_c_offset = m_left - left
            anchor_value = None
            if 0 <= anchor_r_offset < len(values) and 0 <= anchor_c_offset < len(values[anchor_r_offset]):
                anchor_value = values[anchor_r_offset][anchor_c_offset]

            address = range_address(m_top, m_left, m_bottom, m_right)
            merge = {
                "address": address,
                "top": m_top,
                "left": m_left,
                "bottom": m_bottom,
                "right": m_right,
                "rowSpan": m_bottom - m_top + 1,
                "columnSpan": m_right - m_left + 1,
                "anchor": {"row": m_top, "column": m_left},
                "value": json_value(anchor_value),
            }
            merges.append(merge)

            for rr in range(m_top, m_bottom + 1):
                for cc in range(m_left, m_right + 1):
                    role = "anchor" if rr == m_top and cc == m_left else "covered"
                    merge_map[(rr, cc)] = {
                        "role": role,
                        "address": address,
                        "anchor": {"row": m_top, "column": m_left},
                        "anchorValue": json_value(anchor_value),
                        "coveredCellMode": covered_cell_mode,
                    }

    merges.sort(key=lambda item: (item["top"], item["left"], item["bottom"], item["right"]))
    return merges, merge_map


def extract_sheet(ws: Any, sheet_index: int, include_empty: bool, covered_cell_mode: str) -> dict[str, Any]:
    used = com_get(ws, "UsedRange", "read worksheet UsedRange")
    top = int(com_get(used, "Row", "read UsedRange.Row"))
    left = int(com_get(used, "Column", "read UsedRange.Column"))
    used_rows = com_get(used, "Rows", "read UsedRange.Rows")
    used_cols = com_get(used, "Columns", "read UsedRange.Columns")
    row_count = int(com_get(used_rows, "Count", "read UsedRange.Rows.Count"))
    col_count = int(com_get(used_cols, "Count", "read UsedRange.Columns.Count"))
    bottom = top + row_count - 1
    right = left + col_count - 1

    try:
        raw_values = com_get(used, "Value", "read UsedRange.Value")
    except Exception:
        raw_values = None

    values = to_matrix(raw_values, row_count, col_count)
    merges, merge_map = read_merge_map(ws, used, values, covered_cell_mode)

    rows: list[dict[str, Any]] = []
    non_empty = 0

    for r_offset in range(row_count):
        abs_row = top + r_offset
        cells: list[dict[str, Any]] = []
        row_non_empty = 0

        for c_offset in range(col_count):
            abs_col = left + c_offset
            raw_value = values[r_offset][c_offset] if r_offset < len(values) and c_offset < len(values[r_offset]) else None
            merge = merge_map.get((abs_row, abs_col), {"role": "none"})
            export_value = raw_value
            if merge.get("role") == "covered":
                if covered_cell_mode == "blank":
                    export_value = None
                elif covered_cell_mode == "anchor":
                    export_value = merge.get("anchorValue")

            text = value_text(export_value)
            if text:
                row_non_empty += 1

            if include_empty or text or merge.get("role") != "none":
                cells.append(
                    {
                        "row": abs_row,
                        "column": abs_col,
                        "colLabel": excel_col_label(abs_col),
                        "address": cell_address(abs_row, abs_col),
                        "value": json_value(export_value),
                        "rawValue": json_value(raw_value),
                        "merge": merge,
                    }
                )

        non_empty += row_non_empty
        rows.append(
            {
                "rowNumber": abs_row,
                "nonEmptyCount": row_non_empty,
                "cells": cells,
            }
        )

    return {
        "sheetIndex": sheet_index,
        "sheetName": str(com_get(ws, "Name", "read worksheet name")),
        "visible": visible_sheet(ws),
        "usedRange": {
            "top": top,
            "left": left,
            "bottom": bottom,
            "right": right,
            "rowCount": row_count,
            "columnCount": col_count,
            "address": range_address(top, left, bottom, right),
        },
        "nonEmptyCells": non_empty,
        "mergeCount": len(merges),
        "merges": merges,
        "rows": rows,
    }


def extract_workbook(path: Path, include_hidden: bool, include_empty: bool, covered_cell_mode: str) -> dict[str, Any]:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        office_blocking_windows(close=True)

        excel = com_call("start Excel", lambda: win32com.client.DispatchEx("Excel.Application"))
        com_set(excel, "Visible", False, "set Excel.Visible")
        com_set(excel, "DisplayAlerts", False, "set Excel.DisplayAlerts")
        com_set(excel, "EnableEvents", False, "set Excel.EnableEvents")
        try:
            com_set(excel, "ScreenUpdating", False, "set Excel.ScreenUpdating")
        except Exception:
            pass

        workbooks = com_get(excel, "Workbooks", "read Excel.Workbooks")
        workbook = com_call(
            "open workbook",
            lambda: workbooks.Open(
                str(path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            ),
        )
        sheets = []
        worksheets = com_get(workbook, "Worksheets", "read workbook.Worksheets")
        sheet_count = int(com_get(worksheets, "Count", "read workbook.Worksheets.Count"))
        for index in range(1, sheet_count + 1):
            ws = com_call(f"read worksheet {index}", lambda i=index: worksheets(i))
            if not include_hidden and not visible_sheet(ws):
                continue
            sheet_name = str(com_get(ws, "Name", "read worksheet name"))
            print(f"[sheet] {sheet_name}", file=sys.stderr, flush=True)
            sheets.append(extract_sheet(ws, index, include_empty, covered_cell_mode))

        return {
            "schemaVersion": "input-data-com-grid-v1",
            "extractedAt": dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
            "sourcePath": str(path),
            "fileName": path.name,
            "fileSize": path.stat().st_size,
            "mtimeNs": path.stat().st_mtime_ns,
            "coveredCellMode": covered_cell_mode,
            "includeEmptyCells": include_empty,
            "sheets": sheets,
            "totals": {
                "sheetCount": len(sheets),
                "rowCount": sum(int(sheet["usedRange"]["rowCount"]) for sheet in sheets),
                "cellCount": sum(int(sheet["usedRange"]["rowCount"]) * int(sheet["usedRange"]["columnCount"]) for sheet in sheets),
                "nonEmptyCells": sum(int(sheet["nonEmptyCells"]) for sheet in sheets),
                "mergeCount": sum(int(sheet["mergeCount"]) for sheet in sheets),
            },
        }
    finally:
        if workbook is not None:
            try:
                com_call("close workbook", lambda: workbook.Close(SaveChanges=False))
            except Exception:
                pass
        if excel is not None:
            try:
                com_call("quit Excel", lambda: excel.Quit())
            except Exception:
                pass
        pythoncom.CoUninitialize()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract Excel workbook grid and merge metadata through Excel COM.")
    parser.add_argument("--input", required=True, help="Source Excel workbook path.")
    parser.add_argument("--output", required=True, help="Output JSON path.")
    parser.add_argument("--include-hidden", action="store_true", help="Include hidden worksheets.")
    parser.add_argument("--sparse", action="store_true", help="Omit ordinary empty cells. Merged covered cells are still emitted.")
    parser.add_argument("--covered-cell-mode", choices=["blank", "anchor", "raw"], default="blank")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = Path(args.input).resolve()
    output = Path(args.output).resolve()
    if not source.exists():
        print(json.dumps({"status": "fail", "error": f"Input not found: {source}"}, ensure_ascii=False))
        return 2

    started = dt.datetime.utcnow()
    try:
        result = extract_workbook(
            source,
            include_hidden=args.include_hidden,
            include_empty=not args.sparse,
            covered_cell_mode=args.covered_cell_mode,
        )
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8-sig")
        elapsed = (dt.datetime.utcnow() - started).total_seconds()
        print(json.dumps({
            "status": "ok",
            "input": str(source),
            "output": str(output),
            "elapsed": elapsed,
            "sheetCount": result["totals"]["sheetCount"],
            "cellCount": result["totals"]["cellCount"],
            "mergeCount": result["totals"]["mergeCount"],
        }, ensure_ascii=False))
        return 0
    except Exception as exc:
        elapsed = (dt.datetime.utcnow() - started).total_seconds()
        print(json.dumps({
            "status": "fail",
            "input": str(source),
            "output": str(output),
            "elapsed": elapsed,
            "error": str(exc),
        }, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
