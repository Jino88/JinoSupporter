from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


UI_PATH = Path(__file__).parents[1] / "inference_data_ai_ui.py"
SPEC = importlib.util.spec_from_file_location("inference_data_ai_ui", UI_PATH)
assert SPEC is not None and SPEC.loader is not None
ui = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = ui
SPEC.loader.exec_module(ui)


class CommonPipelineUiTests(unittest.TestCase):
    def test_build_commands_uses_one_cli_and_absolute_output_paths(self) -> None:
        options = ui.PipelineOptions((r"D:\input\reports\a.xlsx",), "My Dataset", True, True, True)
        commands = ui.build_commands(options)
        self.assertEqual(["com-index", "quick-index"], [command[2] for command in commands[:2]])
        self.assertEqual(ui.ANALYSIS_RUNNER_PATH.name, Path(commands[2][1]).name)
        self.assertIn("--source", commands[2])
        self.assertIn("--verify-after-import", commands[0])
        self.assertIn("--covered-cell-mode", commands[0])
        self.assertIn("blank", commands[0])
        self.assertIn("--include-hidden", commands[0])
        self.assertIn("--html", commands[1])
        self.assertTrue(Path(commands[0][commands[0].index("--db") + 1]).is_absolute())
        self.assertTrue(Path(commands[1][commands[1].index("--html") + 1]).is_absolute())

    def test_build_commands_respects_selected_outputs(self) -> None:
        options = ui.PipelineOptions((r"D:\input\reports\a.xlsx",), "Dataset", True, False, False, False)
        commands = ui.build_commands(options)
        self.assertEqual(1, len(commands))
        self.assertEqual("com-index", commands[0][2])
        self.assertNotIn("--include-hidden", commands[0])

    def test_count_excel_files_ignores_office_temp_files(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            (root / "a.xlsx").write_bytes(b"")
            (root / "b.xlsm").write_bytes(b"")
            (root / "~$open.xlsx").write_bytes(b"")
            (root / "readme.txt").write_text("x", encoding="utf-8")
            nested = root / "nested"
            nested.mkdir()
            (nested / "c.xls").write_bytes(b"")
            self.assertEqual(3, len(ui.expand_excel_paths([str(root)])))

    def test_source_snapshot_changes_only_for_relevant_excel_files(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "a.xlsx"
            workbook.write_bytes(b"one")
            self.assertEqual([str(workbook.resolve())], ui.expand_excel_paths([str(root)]))
            (root / "notes.txt").write_text("not an Excel input", encoding="utf-8")
            self.assertEqual([str(workbook.resolve())], ui.expand_excel_paths([str(root)]))
            second = root / "b.xls"
            second.write_bytes(b"two")
            self.assertEqual([str(workbook.resolve()), str(second.resolve())], ui.expand_excel_paths([str(root)]))


if __name__ == "__main__":
    unittest.main()
