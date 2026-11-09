import { describe, expect, it } from "vitest";
import { headerColumns, parseSetupPaste, tsvRecords } from "./setup";

/** A row of the Setup Plan sheet as Excel copies it (cells joined with tabs). */
const row = (...cells: string[]) => cells.join("\t");

// The sheet pasted on 2026-11-09: A No, B No. Layout, C-D LINE (merged), E-H Process Name (merged),
// I-K Machine Line (Plan, State, Note), L-N SETUP (Plan, Finish, Status), O Note.
const header = [
  row("No", '"No.\nLayout"', "LINE", "", "Process Name", "", "", "", "Machine Line", "", "", "SETUP", "", "", "Note"),
  row("", "", "", "", "", "", "", "", "Plan", "State", "Note", "Plan", "Finish", "Status", ""),
];
const rows = [
  row("1", "1", "SUB 1", "", "CD LOADING", "", "", "", "20-Nov", "X", "", "27-Nov", "", "", ""),
  row("3", "3", "SUB 1", "", "CD ETCHING", "", "", "", "1-Nov", "X", "", "10-Nov", "", "", " - Not yet drawing laser "),
  row("4", "4", "SUB 1", "", "CD BENDING   ", "", "", "", "8-Nov", "O", "90.0%", "17-Nov", "3-Nov", "O", " - Check visual CD after bending OK"),
  row("18", "18", "SUB 1", "", "SP POSITION AI", "", "", "", "31-Oct", "Δ", "", "9-Nov", "", "", ""),
  row("24", "24", "", "", "VP LOADING", "", "", "", "17-Nov", "", "", "26-Nov", "", "", ""),
];
const want = [
  [1, "SUB 1", "CD LOADING", "2026-11-20", "X", "", "2026-11-27", "", "", ""],
  [3, "SUB 1", "CD ETCHING", "2026-11-01", "X", "", "2026-11-10", "", "", "Not yet drawing laser"],
  [4, "SUB 1", "CD BENDING", "2026-11-08", "O", "90.0%", "2026-11-17", "2026-11-03", "O", "Check visual CD after bending OK"],
  [18, "SUB 1", "SP POSITION AI", "2026-10-31", "Δ", "", "2026-11-09", "", "", ""],
  [24, "SUB 1", "VP LOADING", "2026-11-17", "", "", "2026-11-26", "", "", ""],
];
const fields = (text: string) => {
  const p = parseSetupPaste(text, "2026-11-09");
  expect(p.problem).toBe("");
  return p.rows.map((r) => [r.no, r.line, r.processName, r.machinePlan, r.machineState, r.machineNote, r.setupPlan, r.setupFinish, r.setupStatus, r.note]);
};

describe("Setup Plan paste from the real sheet", () => {
  it("finds the columns from the header rows when they are copied too", () => {
    expect(fields([...header, ...rows].join("\r\n") + "\r\n")).toEqual(want);
  });

  it("drops the empty columns merged LINE and Process Name cells leave when only rows are copied", () => {
    expect(fields(rows.join("\r\n"))).toEqual(want);
    expect(fields(row("87", "87", "MAIN ", "", "WINDING", "", "", "", "10-Oct", "", "", "24-Nov", "", "", "")))
      .toEqual([[87, "MAIN", "WINDING", "2026-10-10", "", "", "2026-11-24", "", "", ""]]);
  });

  it("reads the header: LINE, Process Name, Machine Line, SETUP and the last Note", () => {
    expect(headerColumns(tsvRecords(header[0])[0])).toEqual({
      no: 0, layout: 1, line: 2, process: 4, machinePlan: 8, machineState: 9, machineNote: 10,
      setupPlan: 11, setupFinish: 12, setupStatus: 13, note: 14,
    });
    expect(headerColumns(["1", "1", "SUB 1", "CD LOADING"])).toBeNull();
  });

  it("reads quoted cells as Excel writes them", () => {
    expect(tsvRecords('a\t"two\nlines"\t"say ""hi"""\r\nb\t\tc')).toEqual([["a", "two\nlines", 'say "hi"'], ["b", "", "c"]]);
  });
});
