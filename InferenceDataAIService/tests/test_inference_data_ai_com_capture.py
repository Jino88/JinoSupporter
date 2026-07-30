from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import inference_data_ai_com_capture as com_capture
import inference_data_ai_source_ingest as source_ingest


class ComCaptureContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def workbook_path(self) -> Path:
        path = self.root / "fixture.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "Group"
        sheet["B1"] = "Value"
        sheet["A2"] = "A"
        sheet["B2"] = 2
        workbook.save(path)
        workbook.close()
        return path

    def test_com_contract_import_preserves_extractor_and_source_kind(
        self,
    ) -> None:
        payload = source_ingest.extract_workbook(
            self.workbook_path()
        )
        payload["captureContract"] = (
            source_ingest.COM_CAPTURE_CONTRACT
        )
        payload["extractor"] = {
            "name": com_capture.EXTRACTOR_NAME,
            "version": com_capture.EXTRACTOR_VERSION,
            "sourceSaved": False,
        }
        with sqlite3.connect(":memory:") as connection:
            connection.row_factory = sqlite3.Row
            imported = source_ingest.import_capture(
                connection,
                payload,
            )
            revision = connection.execute(
                """
                SELECT capture_contract, extractor_name, extractor_version
                FROM capture_v2_revisions
                WHERE revision_id=?
                """,
                (imported["revisionId"],),
            ).fetchone()
            document = connection.execute(
                "SELECT source_kind FROM capture_v2_documents"
            ).fetchone()
        self.assertEqual(
            source_ingest.COM_CAPTURE_CONTRACT,
            imported["captureContract"],
        )
        self.assertEqual(
            source_ingest.COM_CAPTURE_CONTRACT,
            revision["capture_contract"],
        )
        self.assertEqual(
            com_capture.EXTRACTOR_NAME,
            revision["extractor_name"],
        )
        self.assertEqual(
            com_capture.EXTRACTOR_VERSION,
            revision["extractor_version"],
        )
        self.assertEqual("XLSX", document["source_kind"])

    def test_dialog_match_requires_exact_title_class_and_button(
        self,
    ) -> None:
        monitor = com_capture._AuthDialogMonitor(
            1,
            1,
            False,
            dialog_title="Company Login",
            dialog_class="#32770",
            button_caption="Continue",
        )
        exact = {
            "title": "Company Login",
            "class": "#32770",
            "buttons": [
                {
                    "hwnd": 2,
                    "caption": "&Continue",
                    "enabled": True,
                }
            ],
        }
        wrong_title = {**exact, "title": "Other Login"}
        wrong_class = {**exact, "class": "OtherClass"}
        wrong_button = {
            **exact,
            "buttons": [
                {
                    "hwnd": 2,
                    "caption": "Cancel",
                    "enabled": True,
                }
            ],
        }
        self.assertEqual(
            2,
            monitor._matching_button(exact)["hwnd"],
        )
        self.assertIsNone(monitor._matching_button(wrong_title))
        self.assertIsNone(monitor._matching_button(wrong_class))
        self.assertIsNone(monitor._matching_button(wrong_button))

    def test_dialog_dismissal_rejects_incomplete_match_before_com_start(
        self,
    ) -> None:
        path = self.root / "protected.xlsx"
        path.write_bytes(b"not opened")
        with self.assertRaisesRegex(
            ValueError,
            "exact title, class, and button",
        ):
            com_capture.extract_workbook_com(
                path,
                dismiss_auth_dialog=True,
                auth_dialog_title="Company Login",
            )

    def test_fixed_grid_helpers_keep_coordinates_stable(self) -> None:
        self.assertEqual("A1", com_capture._a1(1, 1))
        self.assertEqual("AB9", com_capture._a1(9, 28))
        self.assertEqual(
            [[1, 2], [3, 4]],
            com_capture._as_matrix(((1, 2), (3, 4)), 2, 2),
        )
        self.assertEqual(
            [[1], [2]],
            com_capture._as_matrix(((1,), (2,)), 2, 1),
        )
        self.assertEqual(
            [
                ["General", "General", "General"],
                ["General", "General", "General"],
            ],
            com_capture._as_matrix("General", 2, 3),
        )
        self.assertEqual(
            [
                ["0.00", None, None],
                ["General", None, None],
            ],
            com_capture._as_matrix(
                (("0.00",), ("General",)),
                2,
                3,
            ),
        )

    def test_style_ids_are_stable_and_separate_conflicting_payloads(
        self,
    ) -> None:
        first = {"font": {"name": "Arial", "bold": True}}
        same_different_order = {
            "font": {"bold": True, "name": "Arial"}
        }
        second = {"font": {"name": "Arial", "bold": False}}
        self.assertEqual(0, com_capture._style_id({}))
        self.assertEqual(
            com_capture._style_id(first),
            com_capture._style_id(same_different_order),
        )
        self.assertNotEqual(
            com_capture._style_id(first),
            com_capture._style_id(second),
        )

    def test_formula_count_matches_formulas_persisted_on_cells(
        self,
    ) -> None:
        cells = [
            {"coordinate": "A1", "formula": "=SUM(B1:B2)"},
            {"coordinate": "A2", "formula": None},
            {"coordinate": "A3", "formula": ""},
            {"coordinate": "A4", "formula": "=1+1"},
        ]

        self.assertEqual(
            2,
            com_capture._captured_formula_count(cells),
        )


if __name__ == "__main__":
    unittest.main()
