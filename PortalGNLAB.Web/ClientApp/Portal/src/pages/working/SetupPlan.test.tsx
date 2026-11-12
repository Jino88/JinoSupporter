import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import SetupPlan from "./SetupPlan";
import type { SetupRow } from "./setup";

const row = (id: string, no: number, line: string, processName: string, machinePlan: string, setupPlan: string, setupStatus = ""): SetupRow =>
  ({ id, no, line, processName, machinePlan, setupPlan, setupStatus, setupFinish: "", layoutNo: "", machineState: "", machineNote: "",
    modelName: "MSM-MP", revision: 1, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "" }) as SetupRow;

const view = (rows: SetupRow[]) => renderToStaticMarkup(<SetupPlan modelName="MSM-MP" modelPhase="MP" entries={[]} today="2026-11-11"
  onChanged={() => {}} onError={() => {}} initial={rows} />);

describe("Mass Production Plan", () => {
  it("is called Plan and shows its lines on a calendar like Dev Plan, the processes of the first line under it", () => {
    const html = view([
      row("a", 1, "L1", "Bonding", "2026-11-01", "2026-11-03", "O"), row("b", 2, "L1", "Curing", "2026-11-04", "2026-11-06"),
      row("c", 3, "L2", "Testing", "2026-11-10", "2026-11-12"),
    ]);
    expect(html).toContain("<h2>Plan <span class=\"cnt\">3</span></h2>");
    expect(html).not.toContain("Setup Plan");
    // the lines on the calendar: a row each with its days and how many processes are set up
    const names = [...html.matchAll(/<td class="nm"><span class="tw">[^<]*<\/span><span class="dot"[^>]*><\/span><b[^>]*>([^<]+)<\/b><small>([^<]+)<\/small>/g)].map((m) => [m[1], m[2]]);
    expect(names.slice(0, 2)).toEqual([["L1", "11-01 – 11-06 · 1/2"], ["L2", "11-10 – 11-12 · 0/1"]]);
    expect(html).toContain('class="wk-gantt fill"');
    // the line picked (the first) under it: its schedule and only its processes in the table
    expect(html).toContain("<h3>L1</h3>");
    expect(html).toContain("Detailed schedule");
    expect(html).toMatch(/<b class="done">Bonding<\/b><small>11-01 – 11-03 · O<\/small>/);
    const table = html.slice(html.indexOf('class="wk-it su-t"'));
    expect(table).toContain(">Bonding<");
    expect(table).toContain(">Curing<");
    expect(table).not.toContain(">Testing<");
    expect(html).not.toContain("su-sum");
  });

  it("says there is no process yet without rows", () => {
    const html = view([]);
    expect(html).toContain("No process yet.");
    expect(html).not.toContain("wk-gantt");
  });
});
