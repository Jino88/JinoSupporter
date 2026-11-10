import { describe, expect, it } from "vitest";
import { blankEvent, blankTask, calendarDays, eventSpan, inSpan, monthRuns, planProblem, progress, readDevPlan } from "./plan";
import type { PlanEvent, PlanTask } from "./plan";

const task = (start: string, end: string, f: Partial<PlanTask> = {}): PlanTask => ({ ...blankTask(), name: "t", start, end, ...f });
const ev = (start: string, end: string, tasks: PlanTask[] = []): PlanEvent => ({ ...blankEvent(0), name: "e", start, end, tasks });

describe("development plan", () => {
  it("spans an event over its own dates, else over its tasks", () => {
    expect(eventSpan(ev("2026-11-03", "2026-11-10"))).toEqual(["2026-11-03", "2026-11-10"]);
    expect(eventSpan(ev("2026-11-03", ""))).toEqual(["2026-11-03", "2026-11-03"]);
    expect(eventSpan(ev("", "", [task("2026-11-12", "2026-11-14"), task("2026-11-05", ""), task("", "")]))).toEqual(["2026-11-05", "2026-11-14"]);
    expect(eventSpan(ev("", ""))).toBeNull();
    expect(eventSpan(ev("2026-11-10", "2026-11-03"))).toEqual(["2026-11-03", "2026-11-10"]);
  });

  it("lays the calendar from the first day to the last with a margin, with the months over it", () => {
    const days = calendarDays([["2026-10-30", "2026-11-02"], null, ["2026-11-01", "2026-11-03"]], 1);
    expect(days).toEqual(["2026-10-29", "2026-10-30", "2026-10-31", "2026-11-01", "2026-11-02", "2026-11-03", "2026-11-04"]);
    expect(monthRuns(days)).toEqual([["2026-10", 3], ["2026-11", 4]]);
    expect(calendarDays([null])).toEqual([]);
    expect([inSpan("2026-11-02", ["2026-11-01", "2026-11-03"]), inSpan("2026-11-04", ["2026-11-01", "2026-11-03"]), inSpan("2026-11-01", null)])
      .toEqual([true, false, false]);
  });

  it("checks a name and the order of the dates", () => {
    expect(planProblem({ name: "DR build", start: "2026-11-01", end: "2026-11-05" })).toBe("");
    expect(planProblem({ name: " ", start: "", end: "" })).toMatch(/Name/);
    expect(planProblem({ name: "a", start: "2026-11-05", end: "2026-11-01" })).toMatch(/before/);
    expect(planProblem({ name: "a", start: "11/05", end: "" })).toMatch(/date/);
  });

  it("counts the tasks done and reads a saved plan leniently", () => {
    expect(progress(ev("", "", [task("", "", { done: true }), task("", "")]))).toEqual([1, 2]);
    const p = readDevPlan({ events: [{ id: "a", name: "ER", start: "2026-11-01", tasks: [{ name: "Jig", done: "yes" }, 3] }, null] });
    expect(p.events).toHaveLength(1);
    expect(p.events[0]).toMatchObject({ id: "a", name: "ER", start: "2026-11-01", end: "", color: "#4f46e5" });
    expect(p.events[0].tasks.map((t) => [t.name, t.done, t.owner])).toEqual([["Jig", false, ""]]);
    expect(readDevPlan(undefined)).toEqual({ events: [] });
  });
});
