import { describe, expect, it } from "vitest";
import { barAt, devPlans, labelSide, modelRow, planBrief, timelineMonths, timelineRange, treeModels, type ModelLoad } from "./homeSummary";
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
  });

  it("writes the days of a short bar beside it, on its left near the right edge, so they are not cut", () => {
    expect(labelSide({ left: 10, width: 30 })).toBe("in");
    expect(labelSide({ left: 14, width: 4 })).toBe("right");
    expect(labelSide({ left: 60, width: 1 })).toBe("right");
    expect(labelSide({ left: 80, width: 2 })).toBe("left");
  });

  it("counts the open development and mass production issues and the late ones of a model", () => {
    const issue = (x: Partial<Entry>): Partial<Entry> => ({ kind: "issue", status: "Open", doneDate: "", due: "", phase: "DV", ...x });
    const row = modelRow(load("M", { phase: "DV", customer: "C", sopTarget: "2027-03-01" }, [
      issue({ id: "1" }), issue({ id: "2", due: "2026-11-01" }), issue({ id: "3", phase: "MP" }), issue({ id: "4", doneDate: "2026-11-02" }),
      issue({ id: "5", setupRowId: "r" }), { id: "6", kind: "dev" },
    ]), today);
    expect(row).toMatchObject({ phase: "DV", customer: "C", sopTarget: "2027-03-01", devOpen: 2, mpOpen: 2, lateIssues: 1, events: 0 });
  });
});
