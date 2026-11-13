import { describe, expect, it } from "vitest";
import { barAt, devPlans, hasOpen, labelSide, modelRow, mpIssueRows, mpPlans, planBrief, timelineMonths, timelineRange, treeModels, type ModelLoad } from "./homeSummary";
import type { SetupRow } from "./setup";
import type { Entry, ModelInfo } from "./types";

const today = "2026-11-11";
const task = (start: string, end: string, done = false) => ({ id: start + end, name: "t", start, end, owner: "", done, note: "" });
const plan = (...events: { name: string; start?: string; end?: string; tasks?: ReturnType<typeof task>[] }[]) =>
  ({ events: events.map((e, i) => ({ id: "e" + i, color: "#123456", rows: [], bomFrom: "", start: "", end: "", tasks: [], ...e })) });
const load = (name: string, info: ModelInfo, entries: Partial<Entry>[] = []): ModelLoad =>
  ({ name, path: ["MODULE", "MSM"], model: { name, info, revision: 1, updatedAt: "2026-11-10T09:00:00", updatedBy: "a" }, entries: entries as Entry[], error: "" });

describe("Home summary", () => {
  it("lists every model of the tree once, sub-groups before a group's own models", () => {
    const tree = [{ id: "1", name: "MODULE", total: 3, models: ["A"], children: [{ id: "2", name: "MSM", total: 2, models: ["B", "A"], children: [] }] }];
    expect(treeModels(tree)).toEqual([{ name: "B", path: ["MODULE", "MSM"] }, { name: "A", path: ["MODULE", "MSM"] }]);
  });

  it("tells each event's state against today, what is under way, what is next and the late tasks", () => {
    const p = planBrief("M", [], {
      phase: "PVT",
      devPlan: plan(
        { name: "ER", start: "2026-10-01", end: "2026-10-10", tasks: [task("2026-10-01", "2026-10-05", true)] },
        { name: "DR", start: "2026-10-20", end: "2026-11-05", tasks: [task("2026-10-20", "2026-10-30"), task("2026-11-01", "2026-11-02", true)] },
        { name: "DV", tasks: [task("2026-11-09", "2026-11-20")] },
        { name: "PV", start: "2026-12-10", end: "2026-12-20" },
        { name: "MP prep", start: "2026-12-01" },
        { name: "Idea" },
      ),
    }, today);
    expect(p.phase).toBe("PV");
    expect(p.events.map((e) => e.state)).toEqual(["done", "late", "now", "upcoming", "upcoming", "undated"]);
    expect(p.now.map((e) => e.name)).toEqual(["DV"]);
    expect(p.next?.name).toBe("MP prep");
    expect(p.lateTasks).toBe(1);
    expect(p.events[1]).toMatchObject({ done: 1, total: 2, span: ["2026-10-20", "2026-11-05"] });
    expect(p.events[1].tasks.map((x) => [x.span, x.done, x.late])).toEqual([[["2026-10-20", "2026-10-30"], false, true], [["2026-11-01", "2026-11-02"], true, false]]);
    expect(p.events[5].tasks).toEqual([]);
  });

  it("shows the plans of the models in development, the ones under way first, and the ones in development without a plan", () => {
    const loads = [
      load("Later", { phase: "DR", devPlan: plan({ name: "X", start: "2027-01-05", end: "2027-01-06" }) }),
      load("Mass", { phase: "MP", devPlan: plan({ name: "X", start: "2026-11-01", end: "2026-11-30" }) }),
      load("Busy", { phase: "ER", devPlan: plan({ name: "X", start: "2026-11-01", end: "2026-11-30" }) }),
      load("Soon", { devPlan: plan({ name: "X", start: "2026-12-01", end: "2026-12-02" }) }),
      load("Empty", { phase: "DV" }),
      load("Unknown", {}),
      { name: "Broken", path: [], model: null, entries: [], error: "HTTP 500" },
    ];
    const { plans, noPlan } = devPlans(loads, today);
    expect(plans.map((p) => p.name)).toEqual(["Busy", "Soon", "Later"]);
    expect(noPlan).toEqual(["Empty"]);
  });

  it("places the events on a timeline from the first plan day to the last, today always in", () => {
    const p = planBrief("M", [], { devPlan: plan({ name: "A", start: "2026-11-20", end: "2026-12-05" }) }, today);
    const range = timelineRange([p], today, 0)!;
    expect(range).toEqual({ from: "2026-11-11", to: "2026-12-05", days: 25 });
    expect(barAt(["2026-11-20", "2026-12-05"], range)).toEqual({ left: (9 / 25) * 100, width: (16 / 25) * 100 });
    expect(timelineMonths(range).map(([m, , w]) => [m, Math.round((w * 25) / 100)])).toEqual([["2026-11", 20], ["2026-12", 5]]);
    expect(timelineRange([planBrief("N", [], { devPlan: plan({ name: "B" }) }, today)], today)).toBeNull();
    // a task outside its event's days is on the timeline too
    const q = planBrief("M", [], { devPlan: plan({ name: "A", start: "2026-11-20", end: "2026-11-21", tasks: [task("2026-11-25", "2026-12-02")] }) }, today);
    expect(timelineRange([q], today, 0)?.to).toBe("2026-12-02");
  });

  it("writes the days of every bar beside it, on its left near the right edge, inside only when no side has room", () => {
    expect(labelSide({ left: 10, width: 30 })).toBe("right");
    expect(labelSide({ left: 14, width: 4 })).toBe("right");
    expect(labelSide({ left: 60, width: 1 })).toBe("right");
    expect(labelSide({ left: 80, width: 12 })).toBe("left");
    expect(labelSide({ left: 2, width: 88 })).toBe("right");
    expect(labelSide({ left: 10, width: 87 })).toBe("left");
    expect(labelSide({ left: 3, width: 94 })).toBe("in");
  });

  it("counts the open development and mass production issues and the late ones of a model", () => {
    const issue = (x: Partial<Entry>): Partial<Entry> => ({ kind: "issue", status: "Open", doneDate: "", due: "", phase: "DV", ...x });
    const row = modelRow(load("M", { phase: "DV", customer: "C", sopTarget: "2027-03-01" }, [
      issue({ id: "1" }), issue({ id: "2", due: "2026-11-01" }), issue({ id: "3", phase: "MP" }), issue({ id: "4", doneDate: "2026-11-02" }),
      issue({ id: "5", setupRowId: "r" }), { id: "6", kind: "dev" },
    ]), today);
    expect(row).toMatchObject({ phase: "DV", customer: "C", sopTarget: "2027-03-01", devOpen: 2, mpOpen: 2, lateIssues: 1, events: 0 });
  });

  it("makes the MP plans from the Setup Plans: a line per event, each process from machine plan to setup plan", () => {
    const row = (no: number, line: string, processName: string, machinePlan: string, setupPlan: string, setupStatus = "", setupFinish = "") =>
      ({ id: "r" + no, no, line, processName, machinePlan, setupPlan, setupStatus, setupFinish, layoutNo: "", machineState: "", machineNote: "",
         modelName: "M", revision: 1, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "" }) as SetupRow;
    const withSetup = (name: string, phase: string, setup: SetupRow[]): ModelLoad => ({ ...load(name, { phase }), setup });
    const plans = mpPlans([
      withSetup("MSM-MP", "MP", [
        row(2, "L2", "Assy", "2026-11-05", "2026-11-09"), row(1, "L1", "Bonding", "2026-11-01", "2026-11-03", "O"),
        row(3, "L1", "Aging", "", "2026-11-20"), row(4, "", "Packing", "2026-11-12", "", "", "2026-11-12"),
      ]),
      withSetup("MSM-DEV", "PV", []),
      load("MSM-NONE", { phase: "MP" }),
    ], today);
    expect(plans.map((p) => p.name)).toEqual(["MSM-MP"]);
    const [p] = plans;
    expect(p.events.map((e) => [e.name, e.tasks.map((t) => [t.name, t.span, t.done, t.late])])).toEqual([
      ["L1", [["Bonding", ["2026-11-01", "2026-11-03"], true, false], ["Aging", ["2026-11-20", "2026-11-20"], false, false]]],
      ["L2", [["Assy", ["2026-11-05", "2026-11-09"], false, true]]],
      ["(No line)", [["Packing", ["2026-11-12", "2026-11-12"], true, false]]],
    ]);
    expect(p.lateTasks).toBe(1);
    // a model's Setup Plan does not show in its development plan
    expect(devPlans([withSetup("MSM-DEV", "PV", [row(1, "L1", "X", "2026-11-01", "2026-11-02")])], today).plans).toEqual([]);
  });

  it("lists the MP issues of every model: open first, late first among them, then model order and number", () => {
    const issue = (x: Partial<Entry>): Partial<Entry> => ({ kind: "issue", status: "Open", doneDate: "", due: "", phase: "MP", seq: 1, ...x });
    const rows = mpIssueRows([
      load("A", { phase: "MP" }, [
        issue({ id: "a1", seq: 1, doneDate: "2026-11-05", status: "Closed" }), issue({ id: "a2", seq: 2 }), issue({ id: "a3", seq: 3, phase: "DV" }),
      ]),
      load("B", { phase: "PV" }, [issue({ id: "b1", seq: 1, phase: "PV", setupRowId: "r1", due: "2026-11-01" }), { id: "b2", kind: "dev" }]),
    ], today);
    expect(rows.map((r) => [r.model, r.issue.id, r.open, r.late])).toEqual([
      ["B", "b1", true, true], ["A", "a2", true, false], ["A", "a1", false, false],
    ]);
  });

  it("puts the sub-issues and the notes not done under their MP issue, numbered as the Issues tab does", () => {
    const note = (id: string, date: string, x: object = {}) => ({ id, date, createdAt: date, text: id, owner: "", due: "", doneDate: "", ...x });
    const issue = (x: Partial<Entry>): Partial<Entry> => ({ kind: "issue", status: "Open", doneDate: "", due: "", phase: "MP", seq: 1, logs: [], ...x });
    const rows = mpIssueRows([
      load("A", { phase: "MP" }, [
        issue({ id: "p", seq: 1, doneDate: "2026-11-05", logs: [note("n2", "2026-11-03"), note("n1", "2026-11-01", { doneDate: "2026-11-02" })] as never }),
        issue({ id: "k2", seq: 3, parentId: "p", logs: [note("m1", "2026-11-04", { due: "2026-11-10" })] as never }),
        issue({ id: "k1", seq: 2, parentId: "p", doneDate: "2026-11-06" }),
        issue({ id: "q", seq: 4, doneDate: "2026-11-06" }),
      ]),
    ], today);
    // a closed issue with an open sub-issue comes before an issue with nothing open
    expect(rows.map((r) => [r.issue.id, r.no, hasOpen(r)])).toEqual([["p", "1", true], ["q", "4", false]]);
    const [p] = rows;
    expect(p.openNotes.map((n) => [n.no, n.log.id])).toEqual([["1-2", "n2"]]);
    expect(p.subs.map((s) => [s.issue.id, s.no, s.open])).toEqual([["k1", "1.1", false], ["k2", "1.2", true]]);
    expect(p.subs[1].openNotes.map((n) => [n.no, n.late])).toEqual([["1.2-1", true]]);
  });
});
