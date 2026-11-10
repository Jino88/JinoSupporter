"""Tests of working_process_ng_cli.py: python -m unittest tools/test_working_process_ng_cli.py"""
from __future__ import annotations

import csv
import io
import json
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import working_process_ng_cli as cli  # noqa: E402

RATE = "NG rate (%)"
D3 = ["NG tràn keo", "NG thiếu keo", "NG Lệch"]

# The Daily NG Rate header of 2026-11-10 (a "" item is an empty merged column, rate=False has no NG rate column).
SHEET: list[tuple[str, list[str], bool]] = [
    ("1. BONDING ENCLOSURE \nBôi keo tổ hợp spk", ["NG"], True),
    ("2. AOI 1", D3, True),
    ("3. SPK ASSY\nTổ hợp spk", ["NG"], True),
    ("4. FIX BONDING  SPK 1\n Bôi keo cố định spk 1", ["NG"], True),
    ("5. AOI TEST 2", D3, True),
    ("6. FIX BONDING 2\nBôi keo tổ hợp plate", ["NG"], True),
    ("7. AOI TEST 3", ["NG tràn keo", "NG thiếu keo", "NG lệch"], True),
    ("8. PLATE ASSY\nTổ hợp plate", ["NG"], True),
    ("9. PBONDING FIX PLATE \n Bôi keo cố định Plate", ["NG"], True),
    ("10. AOI TEST 3", ["NG"], True),
    ("11. HEIGHT INSPECTION\nKiểm tra chiều cao", ["NG Cao"], True),
    ("12. BONDING FIX SPK 2nd\nBôi keo cố định spk 2 nd( UV)", ["Tràn keo", "Thiếu keo", "Lệch"], True),
    ("13.  AOI TEST", D3, True),
    ("14. COATING CREAM GLUING", ["NG"], True),
    ("15. SPOT WELDING\nHàn", ["Hàn non \nWeak solder)"], True),
    ("16. AI TEST\nKiểm tra spot welding", ["NG"], True),
    ("17. BONDING ENCLOSURE 2nd\nBôi keo tổ hợp lower ", D3, True),
    ("18. AOI TEST", ["NG"], True),
    ("19. QR MARKING", ["NG"], True),
    ("20. QR TEST ", ["NG"], True),
    (" 21.  LOWER ASSY\nTổ hợp lower ", ["NG"], True),
    ("22. GLUING DRAWING GUIDE\nBôi keo rãnh dẫn hướng", ["NG"], True),
    (" 23. VISION INSPECTION \nKiểm tra keo", D3, True),
    (" 24. BONDING LSR GRILL\nBôi keo phần tổ hợp grill", ["NG"], True),
    ("25. AOI TEST 7", ["NG"], True),
    ("26. TAPE ADHESION\nDán tape", ["NG"], True),
    ("27.VISION INSPECTION \nKiểm tra sau dán", ["NG", ""], True),
    ("28. GRILL ASSY\nTổ hợp grill", ["NG", ""], True),
    ("29. HEIGHT INSPECTION 2", ["NG", "Cao"], True),
    ("30. SPL TEST VS\nKiểm tra Sóng VS", ["NG"], True),
    ("31. AIR LEAK TEST", ["NG kênh Gril", "QR code"], True),
    ("32. BAKO TEST", ["THD", "THD + FRF", "RB"], True),
    ("33. Hearing", ["Noise", "Touch", "No sound"], True),
    ("34. TAPE ADHESION 2\nDán tape 2", ["NG damage"], True),
    ("35. Kiểm tra ANTENA", ["Xước ANTENA", "Gãy ANTENA", "Other"], False),
    ("36. VISION INSPECTION \nKiểm tra sau dán VS-TAPE-CUSHION", ["NG", ""], True),
    ("37. VISUAL\nNgoại quan", ["KEO Pt", "Tổn thương ANTENA", "Dị vật+TT GR", "Tổn thương UPPER"], True),
]


def sheet_text(sheet=SHEET, lead: int = 0) -> str:
    """The header rows as Excel copies them: merged cells as empty cells, tab separated, line breaks quoted."""
    head, items, ba = [""] * lead, [""] * lead, [""] * lead
    if lead:
        head[0], items[0] = "Date", ""
    for cell, ngs, rate in sheet:
        cols = ["Input"] + [c for n in ngs for c in (n, "")] + ([RATE, ""] if rate else [])
        head += [cell] + [""] * (len(cols) - 1)
        items += cols
        ba += [""] + ["Before", "After"] * ((len(cols) - 1) // 2)
    out = io.StringIO()
    csv.writer(out, delimiter="\t", lineterminator="\r\n").writerows([head, items, ba])
    return out.getvalue()


def make_db(path: Path, info: dict | None) -> None:
    conn = sqlite3.connect(path)
    conn.execute("CREATE TABLE WorkingModels (ModelName TEXT PRIMARY KEY, InfoJson TEXT NOT NULL DEFAULT '{}', "
                 "Revision INTEGER NOT NULL, UpdatedAt TEXT NOT NULL, UpdatedBy TEXT NOT NULL)")
    if info is not None:
        conn.execute("INSERT INTO WorkingModels VALUES ('MSM-G778B', ?, 4, '2026-11-10T00:00:00.0000000Z', 'admin')",
                     (json.dumps(info),))
    conn.commit()
    conn.close()


class ParseHeader(unittest.TestCase):
    def test_sheet_processes_and_ng_items(self):
        procs, warnings = cli.parse_header(sheet_text(lead=4))
        self.assertEqual(len(procs), 37)
        self.assertEqual(procs[0], {"name": "BONDING ENCLOSURE", "localName": "Bôi keo tổ hợp spk", "ngItems": ["NG"]})
        self.assertEqual(procs[1], {"name": "AOI 1", "localName": "", "ngItems": D3})
        self.assertEqual(procs[3]["name"], "FIX BONDING SPK 1")
        self.assertEqual(procs[11]["localName"], "Bôi keo cố định spk 2 nd( UV)")
        self.assertEqual(procs[14]["ngItems"], ["Hàn non Weak solder)"])
        self.assertEqual(procs[20]["name"], "LOWER ASSY")
        self.assertEqual(procs[26]["name"], "VISION INSPECTION")
        self.assertEqual(procs[26]["ngItems"], ["NG"])
        self.assertEqual(procs[34], {"name": "Kiểm tra ANTENA", "localName": "", "ngItems": ["Xước ANTENA", "Gãy ANTENA", "Other"]})
        self.assertEqual(procs[36]["ngItems"], ["KEO Pt", "Tổn thương ANTENA", "Dị vật+TT GR", "Tổn thương UPPER"])
        self.assertTrue(any("VISION INSPECTION appears 3 times" in w for w in warnings))

    def test_process_without_input_uses_the_columns(self):
        text = sheet_text([("1. A", ["x"], True), ("2. B", ["y", "z"], True)])
        rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
        rows[1][rows[1].index("Input", 1)] = ""  # B lost its Input cell
        out = io.StringIO()
        csv.writer(out, delimiter="\t").writerows(rows)
        procs, warnings = cli.parse_header(out.getvalue())
        self.assertEqual([p["ngItems"] for p in procs], [["x"], ["y", "z"]])
        self.assertIn("2 processes but 1 Input cells", warnings[0])

    def test_no_item_row(self):
        with self.assertRaises(ValueError):
            cli.parse_header("1. A\t2. B\r\n")


class MergeManual(unittest.TestCase):
    def test_keeps_materials_and_matching_processes(self):
        current = {"materials": ["UPPER"], "processes": [
            {"id": "keep1", "type": "MAIN", "name": "aoi 1", "localName": "old", "ngItems": ["x"],
             "materials": [{"material": "UPPER", "qtyPerUnit": 1}]},
            {"id": "gone", "type": "SUB", "name": "Old step", "localName": "", "ngItems": [], "materials": []},
        ]}
        g, dropped = cli.merge_manual(current, [{"name": "AOI 1", "localName": "", "ngItems": ["NG"]},
                                               {"name": "New", "localName": "Mới", "ngItems": []}])
        self.assertEqual(g["materials"], ["UPPER"])
        self.assertEqual(g["processes"][0], {"id": "keep1", "type": "MAIN", "name": "AOI 1", "localName": "old",
                                             "ngItems": ["NG"], "materials": [{"material": "UPPER", "qtyPerUnit": 1}]})
        self.assertEqual(g["processes"][1]["type"], "SUB")
        self.assertEqual(len(g["processes"][1]["id"]), 12)
        self.assertEqual(dropped, ["Old step"])


class Apply(unittest.TestCase):
    def run_cli(self, *args: str) -> str:
        out = io.StringIO()
        with redirect_stdout(out):
            code = cli.main(list(args))
        self.assertEqual(code, 0)
        return out.getvalue()

    def test_dry_run_then_apply(self):
        with tempfile.TemporaryDirectory() as d:
            db, src = Path(d) / "p.db", Path(d) / "h.txt"
            make_db(db, {"phase": "PV", "devPlan": {"events": []}})
            src.write_text(sheet_text(), encoding="utf-8")

            self.assertIn("Not saved", self.run_cli("MSM-G778B", "--file", str(src), "--db", str(db)))
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT Revision FROM WorkingModels").fetchone()[0], 4)
            conn.close()

            self.assertIn("revision 5", self.run_cli("MSM-G778B", "--file", str(src), "--db", str(db), "--apply"))
            conn = sqlite3.connect(db)
            info, by = conn.execute("SELECT InfoJson, UpdatedBy FROM WorkingModels").fetchone()
            conn.close()
            info = json.loads(info)
            self.assertEqual(info["phase"], "PV")
            self.assertEqual(info["devPlan"], {"events": []})
            self.assertEqual(info["bmesMaterial"], cli.MANUAL)
            self.assertEqual(len(info["manualGroup"]["processes"]), 37)
            self.assertEqual(info["manualGroup"]["processes"][31]["ngItems"], ["THD", "THD + FRF", "RB"])
            self.assertTrue(by)

    def test_unknown_model_is_not_created(self):
        with tempfile.TemporaryDirectory() as d:
            db, src = Path(d) / "p.db", Path(d) / "h.txt"
            make_db(db, None)
            src.write_text(sheet_text(), encoding="utf-8")
            with self.assertRaises(SystemExit), redirect_stdout(io.StringIO()):
                cli.main(["MSM-G778B", "--file", str(src), "--db", str(db), "--apply"])
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM WorkingModels").fetchone()[0], 0)
            conn.close()


if __name__ == "__main__":
    unittest.main()
