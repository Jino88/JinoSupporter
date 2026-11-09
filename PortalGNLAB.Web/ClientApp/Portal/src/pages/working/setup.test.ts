import { describe, expect, it } from "vitest";
import { blankSetup, nextNo, parseSetupPaste, parseSheetDate, parseState, rate, rowsInOrder, setupProblem, setupSummary, shortDate } from "./setup";
import type { SetupFields, SetupRow } from "./setup";

const row = (id: string, f: Partial<SetupFields>, createdAt = "2026-11-09T00:00:00Z"): SetupRow => ({
  ...blankSetup(), processName: "P", ...f, id, modelName: "M", revision: 1, createdBy: "", createdAt, updatedBy: "", updatedAt: createdAt,
});

describe("setup plan", () => {
  it("summarises machines and setups per line in first-seen order, with a total", () => {
    const rows = [
      row("a", { no: 2, line: "SUB 1", machineState: "O", setupStatus: "O" }),
      row("b", { no: 1, line: "SUB 1", machineState: "X" }),
      row("c", { no: 3, line: "MAIN", machineState: "O", setupStatus: "Δ" }),
      row("d", { no: 4, line: "" }),
    ];
    const s = setupSummary(rows);
    expect(s.lines.map((l) => [l.line, l.total, l.machineFinish, l.setupFinish])).toEqual([
      ["SUB 1", 2, 1, 1], ["MAIN", 1, 1, 0], ["(No line)", 1, 0, 0],
    ]);
    expect(s.total).toMatchObject({ total: 4, machineFinish: 2, setupFinish: 1 });
    expect([rate(16, 35), rate(0, 0), rate(16, 136)]).toEqual(["45.7%", "0.0%", "11.8%"]);
  });

  it("orders rows by No., then by when they were added", () => {
    const r = rowsInOrder([row("x", { no: 2 }), row("y", { no: 1 }, "2026-11-09T02:00:00Z"), row("z", { no: 1 }, "2026-11-09T01:00:00Z")]);
    expect(r.map((x) => x.id)).toEqual(["z", "y", "x"]);
  });

  it("gives a new row the No. after the highest", () => {
    expect(nextNo([])).toBe(1);
    expect(nextNo([row("a", { no: 3 }), row("b", { no: 18 }), row("c", { no: 7 })])).toBe(19);
  });

  it("reads dates as typed in Excel", () => {
    const today = "2026-11-09";
    expect(["20-Nov", "8-Dec", "1-Nov-27", "2026-11-20", "2026/1/5", "11/20", "11/20/2027", "", "31-Feb", "soon", "13/40"].map((t) => parseSheetDate(t, today)))
      .toEqual(["2026-11-20", "2026-12-08", "2027-11-01", "2026-11-20", "2026-01-05", "2026-11-20", "2027-11-20", "", null, null, null]);
    expect([shortDate("2026-11-20", today), shortDate("2027-04-11", today), shortDate("", today)]).toEqual(["20-Nov", "11-Apr-27", ""]);
  });

  it("takes O / X / Δ look-alikes", () => {
    expect(["O", "○", "x", "×", "△", "Δ", "", "Y"].map(parseState)).toEqual(["O", "O", "X", "X", "Δ", "Δ", "", null]);
  });

  it("checks a row before it is saved", () => {
    expect(setupProblem({ ...blankSetup(), processName: "CD LOADING" })).toBe("");
    expect(setupProblem(blankSetup())).toMatch(/Process name/);
    expect(setupProblem({ ...blankSetup(), processName: "x", setupPlan: "27-Nov" })).toMatch(/Setup plan/);
    expect(setupProblem({ ...blankSetup(), processName: "x", no: 1.5 })).toMatch(/No\./);
  });

  it("parses rows pasted from the sheet: headers skipped, LINE filled down, the Note kept as the problem", () => {
    const text = [
      "No\tNo. Layout\tLINE\tProcess Name\tPlan\tState\tNote\tPlan\tFinish\tStatus\tNote",
      "1\t1\tSUB 1\tCD LOADING\t20-Nov\tX\t\t27-Nov\t\t\t",
      "3\t3\t\tCD ETCHING\t1-Nov\tX\t\t10-Nov\t\t\t- Not yet drawing laser",
      "4\t4\t\tCD BENDING\t8-Nov\tO\t90.0%\t17-Nov\t3-Nov\tO\t- Check visual CD after bending OK",
      "\t\t\t\t\t\t\t\t\t\t",
    ].join("\r\n");
    const p = parseSetupPaste(text, "2026-11-09");
    expect(p.problem).toBe("");
    expect(p.rows.map((r) => [r.no, r.line, r.processName, r.machinePlan, r.machineState, r.machineNote, r.setupFinish, r.setupStatus, r.note])).toEqual([
      [1, "SUB 1", "CD LOADING", "2026-11-20", "X", "", "", "", ""],
      [3, "SUB 1", "CD ETCHING", "2026-11-01", "X", "", "", "", "Not yet drawing laser"],
      [4, "SUB 1", "CD BENDING", "2026-11-08", "O", "90.0%", "2026-11-03", "O", "Check visual CD after bending OK"],
    ]);
  });

  it("refuses a pasted row with a bad state or date, and an empty paste", () => {
    expect(parseSetupPaste("1\t1\tSUB 1\tA\t20-Nov\tY", "2026-11-09").problem).toMatch(/Line 1: machine state/);
    expect(parseSetupPaste("1\t1\tSUB 1\tA\tsoon", "2026-11-09").problem).toMatch(/machine plan "soon"/);
    expect(parseSetupPaste("\n\n", "2026-11-09")).toMatchObject({ rows: [], problem: expect.stringMatching(/No rows/) });
  });
});
