"""Read-only Excel COM capture for DRM/policy protected workbooks.

The extractor emits the same coordinate-stable Capture v2 payload consumed by
the canonical ingestion pipeline.  It never saves the source workbook.  A
dialog can only be dismissed when its Excel-owned title, window class, and
button caption all exactly match caller-supplied values.
"""

from __future__ import annotations

import hashlib
import json
import math
import sys
import threading
import time
from datetime import date, datetime, time as datetime_time
from pathlib import Path
from typing import Any, Callable

from inference_data_ai_source_ingest import (
    CAPTURE_SCHEMA_VERSION,
    COM_CAPTURE_CONTRACT,
    CaptureError,
    _bounds_from_coordinates,
    _bounds_payload,
    _compact,
    sha256_file,
)


EXTRACTOR_NAME = "inference_data_ai_com_capture"
EXTRACTOR_VERSION = "2.1"
SUPPORTED_EXCEL_EXTENSIONS = frozenset(
    {".xlsx", ".xlsm", ".xlsb", ".xls"}
)


def _a1(row: int, column: int) -> str:
    letters = ""
    value = column
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def _as_matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    def is_sequence(item: Any) -> bool:
        return isinstance(item, (tuple, list))

    def fixed_row(item: Any) -> list[Any]:
        values = list(item) if is_sequence(item) else [item]
        return (values + [None] * columns)[:columns]

    if rows == 1 and columns == 1:
        return [[value]]
    if not is_sequence(value):
        # Range-level properties such as NumberFormat can return one scalar
        # when every cell has the same value.  Preserve the fixed grid by
        # broadcasting that property to every coordinate.
        return [[value for _ in range(columns)] for _ in range(rows)]
    if rows == 1:
        if (
            len(value) == 1
            and is_sequence(value[0])
        ):
            return [fixed_row(value[0])]
        return [fixed_row(value)]
    if columns == 1:
        flattened = [
            item[0] if is_sequence(item) and item else item
            for item in list(value)
        ]
        flattened = (flattened + [None] * rows)[:rows]
        return [[item] for item in flattened]
    outer = list(value)
    if outer and not any(is_sequence(item) for item in outer):
        if len(outer) == rows * columns:
            return [
                fixed_row(
                    outer[
                        row_offset
                        * columns:(row_offset + 1)
                        * columns
                    ]
                )
                for row_offset in range(rows)
            ]
        return [
            fixed_row(outer if row_offset == 0 else [])
            for row_offset in range(rows)
        ]
    matrix = [
        fixed_row(outer[row_offset])
        if row_offset < len(outer)
        else fixed_row([])
        for row_offset in range(rows)
    ]
    return matrix


def _portable_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if math.isnan(value):
            return {"type": "float", "value": "NaN"}
        if math.isinf(value):
            return {
                "type": "float",
                "value": "Infinity" if value > 0 else "-Infinity",
            }
        return value
    if isinstance(value, (datetime, date, datetime_time)):
        return {
            "type": type(value).__name__,
            "value": value.isoformat(),
        }
    return {"type": type(value).__name__, "value": str(value)}


def _data_type(value: Any, *, formula: bool = False) -> str:
    if formula:
        return "f"
    if value is None:
        return "n"
    if isinstance(value, bool):
        return "b"
    if isinstance(value, (int, float)):
        return "n"
    if isinstance(value, (datetime, date, datetime_time)):
        return "d"
    if isinstance(value, str) and value.startswith("#"):
        return "e"
    return "s"


def _safe_get(callable_value: Any, default: Any = None) -> Any:
    try:
        return callable_value()
    except Exception:
        return default


def _normalise_window_text(value: str) -> str:
    return " ".join(value.replace("&", "").split()).casefold()


class _AuthDialogMonitor:
    """Observe or click one exact dialog owned by the created Excel instance."""

    def __init__(
        self,
        excel_hwnd: int,
        timeout_seconds: float,
        inspect_only: bool,
        *,
        dialog_title: str = "",
        dialog_class: str = "",
        button_caption: str = "",
    ) -> None:
        self._excel_hwnd = int(excel_hwnd)
        self._timeout_seconds = max(0.0, float(timeout_seconds))
        self._inspect_only = inspect_only
        self._dialog_title = _normalise_window_text(dialog_title)
        self._dialog_class = dialog_class.casefold()
        self._button_caption = _normalise_window_text(button_caption)
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._handled_windows: set[int] = set()
        self.observed_dialogs: list[dict[str, Any]] = []
        self.close_request_count = 0

    def start(self) -> None:
        self._thread = threading.Thread(
            target=self._run,
            name="excel-auth-dialog-monitor",
            daemon=True,
        )
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2)

    def _run(self) -> None:
        try:
            import win32con
            import win32gui
            import win32process
        except ImportError:
            return
        try:
            _, excel_pid = win32process.GetWindowThreadProcessId(
                self._excel_hwnd
            )
        except Exception:
            return

        deadline = time.monotonic() + self._timeout_seconds
        while (
            not self._stop_event.is_set()
            and time.monotonic() < deadline
        ):
            dialogs = self._excel_owned_dialogs(
                win32con,
                win32gui,
                win32process,
                int(excel_pid),
            )
            for dialog in dialogs:
                hwnd = int(dialog["hwnd"])
                if hwnd in self._handled_windows:
                    continue
                if self._inspect_only:
                    self.observed_dialogs.append(dialog)
                    self._handled_windows.add(hwnd)
                    print(
                        "AUTH_DIALOG_JSON "
                        + json.dumps(
                            {"authDialog": dialog},
                            ensure_ascii=False,
                            separators=(",", ":"),
                        ),
                        file=sys.stderr,
                        flush=True,
                    )
                    continue
                button = self._matching_button(dialog)
                if button is None:
                    continue
                try:
                    win32gui.PostMessage(
                        int(button["hwnd"]),
                        win32con.BM_CLICK,
                        0,
                        0,
                    )
                except Exception:
                    continue
                self._handled_windows.add(hwnd)
                self.close_request_count += 1
            self._stop_event.wait(0.25)

    def _excel_owned_dialogs(
        self,
        win32con: Any,
        win32gui: Any,
        win32process: Any,
        excel_pid: int,
    ) -> list[dict[str, Any]]:
        dialogs: list[dict[str, Any]] = []

        def visit(window: int, _: Any) -> bool:
            try:
                if (
                    window == self._excel_hwnd
                    or not win32gui.IsWindowVisible(window)
                ):
                    return True
                _, window_pid = win32process.GetWindowThreadProcessId(window)
                root_owner = win32gui.GetAncestor(
                    window,
                    win32con.GA_ROOTOWNER,
                )
                if (
                    int(window_pid) != excel_pid
                    or int(root_owner) != self._excel_hwnd
                ):
                    return True
                dialogs.append(
                    {
                        "hwnd": int(window),
                        "pid": int(window_pid),
                        "rootOwner": int(root_owner),
                        "title": win32gui.GetWindowText(window),
                        "class": win32gui.GetClassName(window),
                        "buttons": self._buttons(win32gui, window),
                    }
                )
            except Exception:
                pass
            return True

        win32gui.EnumWindows(visit, None)
        return dialogs

    @staticmethod
    def _buttons(win32gui: Any, dialog: int) -> list[dict[str, Any]]:
        buttons: list[dict[str, Any]] = []

        def visit(window: int, _: Any) -> bool:
            try:
                if win32gui.GetClassName(window) == "Button":
                    buttons.append(
                        {
                            "hwnd": int(window),
                            "caption": win32gui.GetWindowText(window),
                            "enabled": bool(
                                win32gui.IsWindowEnabled(window)
                            ),
                        }
                    )
            except Exception:
                pass
            return True

        win32gui.EnumChildWindows(dialog, visit, None)
        return buttons

    def _matching_button(
        self,
        dialog: dict[str, Any],
    ) -> dict[str, Any] | None:
        if (
            _normalise_window_text(str(dialog["title"]))
            != self._dialog_title
            or str(dialog["class"]).casefold() != self._dialog_class
        ):
            return None
        for button in dialog["buttons"]:
            if (
                button["enabled"]
                and _normalise_window_text(str(button["caption"]))
                == self._button_caption
            ):
                return button
        return None


def _merge_inventory(
    sheet: Any,
    *,
    top: int,
    left: int,
    row_count: int,
    column_count: int,
) -> tuple[
    list[dict[str, Any]],
    dict[tuple[int, int], tuple[str, str, tuple[int, int]]],
]:
    bottom = top + row_count - 1
    right = left + column_count - 1
    merged: dict[tuple[int, int, int, int], dict[str, Any]] = {}
    membership: dict[
        tuple[int, int],
        tuple[str, str, tuple[int, int]],
    ] = {}
    for row in range(top, bottom + 1):
        for column in range(left, right + 1):
            cell = sheet.Cells(row, column)
            if not bool(_safe_get(lambda: cell.MergeCells, False)):
                continue
            area = cell.MergeArea
            min_row = int(area.Row)
            min_column = int(area.Column)
            max_row = min_row + int(area.Rows.Count) - 1
            max_column = min_column + int(area.Columns.Count) - 1
            key = (min_row, min_column, max_row, max_column)
            if key not in merged:
                address = (
                    _a1(min_row, min_column)
                    if key[0] == key[2] and key[1] == key[3]
                    else (
                        f"{_a1(min_row, min_column)}:"
                        f"{_a1(max_row, max_column)}"
                    )
                )
                merged[key] = {
                    **_bounds_payload(
                        min_row,
                        min_column,
                        max_row,
                        max_column,
                    ),
                    "anchor": _a1(min_row, min_column),
                }
            address = str(merged[key]["address"])
            for covered_row in range(min_row, max_row + 1):
                for covered_column in range(
                    min_column,
                    max_column + 1,
                ):
                    role = (
                        "anchor"
                        if (covered_row, covered_column)
                        == (min_row, min_column)
                        else "covered"
                    )
                    membership[(covered_row, covered_column)] = (
                        address,
                        role,
                        (min_row, min_column),
                    )
    ordered = [merged[key] for key in sorted(merged)]
    return ordered, membership


def _cell_style(cell: Any) -> dict[str, Any]:
    return _compact(
        {
            "styleName": _safe_get(lambda: str(cell.Style), ""),
            "horizontalAlignment": _safe_get(
                lambda: int(cell.HorizontalAlignment),
                None,
            ),
            "verticalAlignment": _safe_get(
                lambda: int(cell.VerticalAlignment),
                None,
            ),
            "wrapText": _safe_get(lambda: bool(cell.WrapText), None),
            "font": _compact(
                {
                    "name": _safe_get(lambda: str(cell.Font.Name), ""),
                    "size": _safe_get(lambda: float(cell.Font.Size), None),
                    "bold": _safe_get(lambda: bool(cell.Font.Bold), None),
                    "italic": _safe_get(
                        lambda: bool(cell.Font.Italic),
                        None,
                    ),
                }
            ),
            "interiorColor": _safe_get(
                lambda: int(cell.Interior.Color),
                None,
            ),
        }
    )


def _style_id(style: dict[str, Any]) -> int:
    """Return a deterministic positive ID for one exact COM style payload."""

    if not style:
        return 0
    encoded = json.dumps(
        style,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    value = int.from_bytes(
        hashlib.sha256(encoded).digest()[:8],
        "big",
    ) & 0x7FFF_FFFF_FFFF_FFFF
    return value or 1


def _captured_formula_count(cells: list[dict[str, Any]]) -> int:
    return sum(
        1
        for cell in cells
        if str(cell.get("formula") or "").startswith("=")
    )


def _extract_sheet(
    sheet: Any,
    sheet_index: int,
    *,
    covered_cell_mode: str,
) -> dict[str, Any]:
    used = sheet.UsedRange
    top = int(used.Row)
    left = int(used.Column)
    row_count = max(1, int(used.Rows.Count))
    column_count = max(1, int(used.Columns.Count))
    values = _as_matrix(used.Value2, row_count, column_count)
    formulas_value = _safe_get(lambda: used.Formula2, None)
    if formulas_value is None:
        formulas_value = _safe_get(lambda: used.Formula, used.Value2)
    formulas = _as_matrix(
        formulas_value,
        row_count,
        column_count,
    )
    number_formats_value = _safe_get(
        lambda: used.NumberFormat,
        None,
    )
    number_formats = (
        _as_matrix(number_formats_value, row_count, column_count)
        if number_formats_value is not None
        else None
    )
    merges, merge_membership = _merge_inventory(
        sheet,
        top=top,
        left=left,
        row_count=row_count,
        column_count=column_count,
    )

    content_coordinates: set[tuple[int, int]] = set()
    cells: list[dict[str, Any]] = []
    for row_offset in range(row_count):
        for column_offset in range(column_count):
            row = top + row_offset
            column = left + column_offset
            value = values[row_offset][column_offset]
            formula_candidate = formulas[row_offset][column_offset]
            formula = (
                str(formula_candidate)
                if isinstance(formula_candidate, str)
                and formula_candidate.startswith("=")
                else None
            )
            merge_info = merge_membership.get((row, column))
            role = merge_info[1] if merge_info else "none"
            if role == "covered":
                if covered_cell_mode == "blank":
                    value = None
                elif covered_cell_mode == "anchor" and merge_info:
                    anchor_row, anchor_column = merge_info[2]
                    value = values[anchor_row - top][anchor_column - left]
                formula = None
            if formula or value is not None:
                content_coordinates.add((row, column))

            cell = sheet.Cells(row, column)
            number_format = (
                number_formats[row_offset][column_offset]
                if number_formats is not None
                else None
            )
            if number_format in (None, ""):
                number_format = _safe_get(
                    lambda: cell.NumberFormat,
                    "General",
                )
            portable_value = _portable_value(value)
            display_value: Any = portable_value
            if value is not None or formula:
                display_text = _safe_get(lambda: cell.Text, None)
                if display_text not in (None, ""):
                    display_value = _portable_value(display_text)
            style = (
                _cell_style(cell)
                if value is not None or formula or merge_info
                else {}
            )
            cells.append(
                {
                    "row": row,
                    "column": column,
                    "coordinate": _a1(row, column),
                    "rawValue": None if formula else portable_value,
                    "formula": formula,
                    "cachedValue": portable_value if formula else None,
                    "displayValue": display_value,
                    "dataType": _data_type(value, formula=bool(formula)),
                    "cachedDataType": (
                        _data_type(value) if formula else None
                    ),
                    "numberFormat": str(number_format or "General"),
                    "styleId": _style_id(style),
                    "style": style,
                    "mergeRange": merge_info[0] if merge_info else None,
                    "mergeRole": role,
                }
            )

    nonempty_rows = {row for row, _ in content_coordinates}
    nonempty_columns = {column for _, column in content_coordinates}
    has_tabular_evidence = (
        len(content_coordinates) >= 4
        and len(nonempty_rows) >= 2
        and len(nonempty_columns) >= 2
    )
    is_truly_empty = not content_coordinates and not merges
    status = (
        "EMPTY"
        if is_truly_empty
        else "CAPTURED"
        if has_tabular_evidence
        else "NO_TABULAR_EVIDENCE"
    )
    row_dimensions = []
    for row in range(top, top + row_count):
        row_range = sheet.Rows(row)
        hidden = bool(_safe_get(lambda: row_range.Hidden, False))
        height = _safe_get(lambda: float(row_range.RowHeight), None)
        if hidden or height is not None:
            row_dimensions.append(
                {"row": row, "height": height, "hidden": hidden}
            )
    column_dimensions = []
    for column in range(left, left + column_count):
        column_range = sheet.Columns(column)
        hidden = bool(_safe_get(lambda: column_range.Hidden, False))
        width = _safe_get(
            lambda: float(column_range.ColumnWidth),
            None,
        )
        if hidden or width is not None:
            column_dimensions.append(
                {
                    "key": _a1(1, column)[:-1],
                    "minColumn": column,
                    "maxColumn": column,
                    "width": width,
                    "hidden": hidden,
                }
            )
    return {
        "sheetIndex": sheet_index,
        "title": str(sheet.Name),
        "sheetState": (
            "visible" if int(sheet.Visible) == -1 else "hidden"
        ),
        "status": status,
        "isTrulyEmpty": is_truly_empty,
        "hasTabularEvidence": has_tabular_evidence,
        "usedBounds": _bounds_payload(
            top,
            left,
            top + row_count - 1,
            left + column_count - 1,
        ),
        "contentBounds": _bounds_from_coordinates(
            sorted(content_coordinates)
        ),
        "nonEmptyCellCount": len(content_coordinates),
        "structuralCellCount": len(cells) - len(content_coordinates),
        "capturedCellCount": len(cells),
        "formulaCellCount": _captured_formula_count(cells),
        "mergeCount": len(merges),
        "freezePanes": None,
        "autoFilter": _safe_get(
            lambda: str(sheet.AutoFilter.Range.Address(False, False)),
            None,
        ),
        "sheetMetadata": {
            "visibleCode": int(sheet.Visible),
            "gridMode": "FIXED_USED_RANGE_WITH_MERGE_PLACEHOLDERS",
        },
        "rowDimensions": row_dimensions,
        "columnDimensions": column_dimensions,
        "mergedRanges": merges,
        "cells": cells,
    }


def extract_workbook_com(
    source_path: str | Path,
    *,
    covered_cell_mode: str = "blank",
    include_hidden: bool = True,
    password: str | None = None,
    inspect_auth_dialog: bool = False,
    dismiss_auth_dialog: bool = False,
    auth_dialog_title: str = "",
    auth_dialog_class: str = "",
    auth_dialog_button: str = "",
    auth_dialog_timeout_seconds: float = 30.0,
    excel_process_callback: Callable[[int], None] | None = None,
) -> dict[str, Any]:
    """Capture one Excel source through a dedicated read-only COM instance."""

    source = Path(source_path).expanduser().resolve()
    if source.suffix.casefold() not in SUPPORTED_EXCEL_EXTENSIONS:
        raise CaptureError(
            "Excel COM capture accepts .xlsx, .xlsm, .xlsb, or .xls."
        )
    if not source.is_file():
        raise FileNotFoundError(source)
    if covered_cell_mode not in {"blank", "anchor", "raw"}:
        raise ValueError(
            "covered_cell_mode must be blank, anchor, or raw."
        )
    if inspect_auth_dialog and dismiss_auth_dialog:
        raise ValueError(
            "Dialog inspection and dismissal cannot be enabled together."
        )
    if dismiss_auth_dialog and not (
        auth_dialog_title.strip()
        and auth_dialog_class.strip()
        and auth_dialog_button.strip()
    ):
        raise ValueError(
            "Dialog dismissal requires exact title, class, and button."
        )

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise CaptureError(
            "pywin32 and Microsoft Excel are required for COM capture."
        ) from exc

    before = source.stat()
    content_sha256 = sha256_file(source)
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    monitor: _AuthDialogMonitor | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        if excel_process_callback is not None:
            import win32process

            _, excel_process_id = win32process.GetWindowThreadProcessId(
                int(excel.Hwnd)
            )
            excel_process_callback(int(excel_process_id))
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3

        if inspect_auth_dialog or dismiss_auth_dialog:
            monitor = _AuthDialogMonitor(
                int(excel.Hwnd),
                auth_dialog_timeout_seconds,
                inspect_auth_dialog,
                dialog_title=auth_dialog_title,
                dialog_class=auth_dialog_class,
                button_caption=auth_dialog_button,
            )
            monitor.start()
        open_arguments: dict[str, Any] = {
            "Filename": str(source),
            "UpdateLinks": 0,
            "ReadOnly": True,
            "IgnoreReadOnlyRecommended": True,
            "AddToMru": False,
        }
        if password:
            open_arguments["Password"] = password
        workbook = excel.Workbooks.Open(**open_arguments)
        sheets = []
        for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
            sheet = workbook.Worksheets(sheet_index)
            if not include_hidden and int(sheet.Visible) != -1:
                continue
            sheets.append(
                _extract_sheet(
                    sheet,
                    sheet_index,
                    covered_cell_mode=covered_cell_mode,
                )
            )

        tabular_sheet_count = sum(
            bool(sheet["hasTabularEvidence"]) for sheet in sheets
        )
        nonempty_sheet_count = sum(
            not bool(sheet["isTrulyEmpty"]) for sheet in sheets
        )
        workbook_status = (
            "EMPTY_WORKBOOK"
            if not nonempty_sheet_count
            else "NO_TABULAR_EVIDENCE"
            if not tabular_sheet_count
            else "CAPTURED"
        )
        payload = {
            "schemaVersion": CAPTURE_SCHEMA_VERSION,
            "captureContract": COM_CAPTURE_CONTRACT,
            "extractor": {
                "name": EXTRACTOR_NAME,
                "version": EXTRACTOR_VERSION,
                "formulaEvaluation": False,
                "imageHandling": "IGNORED",
                "cellTraversal": (
                    "FIXED_USED_RANGE_WITH_MERGE_PLACEHOLDERS"
                ),
                "coveredCellMode": covered_cell_mode,
                "tabularEvidenceRule": (
                    "AT_LEAST_4_VALUES_ACROSS_2_ROWS_AND_2_COLUMNS"
                ),
                "readOnly": True,
                "sourceSaved": False,
            },
            "source": {
                "sourcePath": str(source),
                "fileName": source.name,
                "extension": source.suffix.lower(),
                "sizeBytes": before.st_size,
                "mtimeNs": before.st_mtime_ns,
                "contentSha256": content_sha256,
            },
            "workbook": {
                "status": workbook_status,
                "isTrulyEmpty": nonempty_sheet_count == 0,
                "sheetCount": len(sheets),
                "nonEmptySheetCount": nonempty_sheet_count,
                "tabularSheetCount": tabular_sheet_count,
                "metadata": {
                    "name": str(workbook.Name),
                    "readOnly": bool(workbook.ReadOnly),
                    "includeHiddenSheets": include_hidden,
                    "dialogInspection": inspect_auth_dialog,
                    "dialogDismissal": dismiss_auth_dialog,
                    "observedAuthDialogs": (
                        len(monitor.observed_dialogs)
                        if monitor is not None
                        else 0
                    ),
                    "exactDialogCloseRequests": (
                        monitor.close_request_count
                        if monitor is not None
                        else 0
                    ),
                },
                "sheets": sheets,
            },
        }
    except Exception as exc:
        if isinstance(exc, (CaptureError, ValueError)):
            raise
        raise CaptureError(f"Excel COM capture failed: {exc}") from exc
    finally:
        if monitor is not None:
            monitor.stop()
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    after = source.stat()
    if (
        before.st_size != after.st_size
        or before.st_mtime_ns != after.st_mtime_ns
        or content_sha256 != sha256_file(source)
    ):
        raise CaptureError(
            "Source changed while Excel COM capture was reading it."
        )
    return payload


__all__ = [
    "EXTRACTOR_NAME",
    "EXTRACTOR_VERSION",
    "SUPPORTED_EXCEL_EXTENSIONS",
    "extract_workbook_com",
]
