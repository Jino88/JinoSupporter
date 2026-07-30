from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path
from unittest import mock


MODULE_PATH = Path(__file__).with_name("input_data_excel_com_extract.py")
SPEC = importlib.util.spec_from_file_location("input_data_excel_com_extract", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
extractor = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(extractor)


class Count:
    def __init__(self, count: int) -> None:
        self.Count = count


class MergeArea:
    def __init__(self, row: int, column: int, row_count: int, column_count: int) -> None:
        self.Row = row
        self.Column = column
        self.Rows = Count(row_count)
        self.Columns = Count(column_count)


class Cell:
    def __init__(self, merge_area: MergeArea | None = None) -> None:
        self.MergeCells = merge_area is not None
        self.MergeArea = merge_area


class UsedRange:
    def __init__(self, merge_cells: object) -> None:
        self.Row = 11
        self.Column = 3
        self.Rows = Count(2)
        self.Columns = Count(2)
        self.MergeCells = merge_cells
        self.Value = (("Merged header", None), (None, None))


class Worksheet:
    Name = "Fixture"
    Visible = -1

    def __init__(self, used: UsedRange, cells: dict[tuple[int, int], Cell]) -> None:
        self.UsedRange = used
        self._cells = cells

    def Cells(self, row: int, column: int) -> Cell:
        return self._cells[(row, column)]


def mixed_merge_fixture() -> Worksheet:
    area = MergeArea(11, 3, 2, 2)
    cells = {(row, column): Cell(area) for row in range(11, 13) for column in range(3, 5)}
    return Worksheet(UsedRange(None), cells)


class ReadMergeMapTests(unittest.TestCase):
    def test_mixed_used_range_scans_each_cell_and_preserves_merge_roles(self) -> None:
        worksheet = mixed_merge_fixture()

        merges, merge_map = extractor.read_merge_map(
            worksheet,
            worksheet.UsedRange,
            [["Merged header", None], [None, None]],
            "blank",
        )

        self.assertEqual(1, len(merges))
        self.assertEqual("C11:D12", merges[0]["address"])
        self.assertEqual("Merged header", merges[0]["value"])
        self.assertEqual("anchor", merge_map[(11, 3)]["role"])
        self.assertEqual("covered", merge_map[(11, 4)]["role"])
        self.assertEqual("covered", merge_map[(12, 3)]["role"])
        self.assertEqual("covered", merge_map[(12, 4)]["role"])
        self.assertEqual({"row": 11, "column": 3}, merge_map[(12, 4)]["anchor"])

    def test_explicit_false_skips_cell_scan(self) -> None:
        class ExplodingWorksheet:
            def Cells(self, *_: int) -> Cell:
                raise AssertionError("Cells must not be read when UsedRange is definitely unmerged")

        used = UsedRange(False)
        merges, merge_map = extractor.read_merge_map(ExplodingWorksheet(), used, [[None, None], [None, None]], "blank")

        self.assertEqual([], merges)
        self.assertEqual({}, merge_map)

    def test_extract_sheet_keeps_full_grid_and_blank_covered_cells(self) -> None:
        worksheet = mixed_merge_fixture()

        sheet = extractor.extract_sheet(worksheet, 1, include_empty=True, covered_cell_mode="blank")
        cells = {cell["address"]: cell for row in sheet["rows"] for cell in row["cells"]}

        self.assertEqual({"top": 11, "left": 3, "bottom": 12, "right": 4, "rowCount": 2, "columnCount": 2, "address": "C11:D12"}, sheet["usedRange"])
        self.assertEqual(4, len(cells))
        self.assertEqual("Merged header", cells["C11"]["value"])
        self.assertEqual("anchor", cells["C11"]["merge"]["role"])
        self.assertIsNone(cells["D11"]["value"])
        self.assertEqual("covered", cells["D11"]["merge"]["role"])

    def test_extract_sheet_repeats_anchor_only_in_anchor_mode(self) -> None:
        worksheet = mixed_merge_fixture()

        sheet = extractor.extract_sheet(worksheet, 1, include_empty=True, covered_cell_mode="anchor")
        cells = {cell["address"]: cell for row in sheet["rows"] for cell in row["cells"]}

        self.assertEqual("Merged header", cells["D12"]["value"])
        self.assertIsNone(cells["D12"]["rawValue"])


class ComRetryTests(unittest.TestCase):
    def test_retries_attribute_error_when_office_activation_window_is_present(self) -> None:
        calls = 0

        def flaky_read() -> str:
            nonlocal calls
            calls += 1
            if calls == 1:
                raise AttributeError("__call__.Name")
            return "Fixture"

        with (
            mock.patch.object(extractor, "office_blocking_windows", return_value=["Microsoft Office Activation Wizard"]),
            mock.patch.object(extractor.time, "sleep"),
        ):
            result = extractor.com_retry("read worksheet name", flaky_read)

        self.assertEqual("Fixture", result)
        self.assertEqual(2, calls)


if __name__ == "__main__":
    unittest.main()
