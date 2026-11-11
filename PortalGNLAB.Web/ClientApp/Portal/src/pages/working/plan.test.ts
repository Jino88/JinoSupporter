import { describe, expect, it } from "vitest";
import {
  blankEvent, blankRow, blankTask, calendarDays, eventSpan, inSpan, monthDayText, monthRuns, moveTo, planProblem, progress, planDays, readDevPlan, readMonthDay, readQty, rowTotal, setQty,
} from "./plan";
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

  it("keeps a quantity per day on each row of an event and totals it", () => {
    let r = setQty(blankRow("SPK Assy"), "2026-11-19", 600);
    r = setQty(r, "2026-11-18", 600);
    expect(Object.keys(r.qty)).toEqual(["2026-11-18", "2026-11-19"]);
    expect(rowTotal(r)).toBe(1200);
    expect(setQty(r, "2026-11-19", undefined).qty).toEqual({ "2026-11-18": 600 });
    expect(setQty(r, "2026-11-19", 0).qty).toEqual({ "2026-11-18": 600 });
    expect([readQty("1,200"), readQty(" "), readQty("abc"), readQty("-5"), readQty("2.5")]).toEqual([1200, undefined, null, null, 2.5]);
  });

  it("lays the plan calendar over the events with 2 days on each side, and every day that has a quantity", () => {
    const e = { ...ev("2026-11-18", "2026-11-20"), rows: [setQty(blankRow("OQC"), "2026-11-25", 450)] };
    const days = planDays([e]);
    expect([days[0], days.at(-1), days.length]).toEqual(["2026-11-16", "2026-11-27", 12]);
    expect(planDays([ev("", "")])).toEqual([]);
    expect(planDays([{ ...ev("", ""), rows: [setQty(blankRow("a"), "2026-11-03", 1)] }])).toEqual(["2026-11-01", "2026-11-02", "2026-11-03", "2026-11-04", "2026-11-05"]);
  });

  it("reads the quantity rows of a saved event leniently", () => {
    const p = readDevPlan({ events: [{ name: "LPR", rows: [{ id: "r", name: "SPK Assy", qty: { "2026-11-19": 600, "2026-11-18": 300, bad: 1, "2026-11-20": "x", "2026-11-21": 0 } }, 7] }, { name: "MP" }] });
    expect(p.events[0].rows).toEqual([{ id: "r", name: "SPK Assy", qty: { "2026-11-18": 300, "2026-11-19": 600 } }]);
    expect(p.events[1].rows).toEqual([]);
  });

  it("reads a task date typed as month and day of the fixed year", () => {
    expect(readMonthDay("11/28", 2026)).toBe("2026-11-28");
    expect(readMonthDay(" 1/5 ", 2026)).toBe("2026-01-05");
    expect(readMonthDay("1128", 2026)).toBe("2026-11-28");
    expect(readMonthDay("12-03", 2026)).toBe("2026-12-03");
    expect(readMonthDay("2.29", 2028)).toBe("2028-02-29");
    expect(readMonthDay("", 2026)).toBe("");
    for (const bad of ["2/29", "13/01", "11/31", "0/5", "11/", "2026-11-28", "ab"]) expect(readMonthDay(bad, 2026)).toBeNull();
    expect(monthDayText("2026-11-28")).toBe("11/28");
    expect(monthDayText("")).toBe("");
  });

  it("moves a dragged task after the one dropped on going down and before it going up", () => {
    const l = ["a", "b", "c", "d"].map((id) => ({ id }));
    const ids = (x: { id: string }[]) => x.map((y) => y.id).join("");
    expect(ids(moveTo(l, "a", "c"))).toBe("bcad");
    expect(ids(moveTo(l, "d", "b"))).toBe("adbc");
    expect(ids(moveTo(l, "b", "b"))).toBe("abcd");
    expect(moveTo(l, "x", "b")).toBe(l);
    expect(ids(l)).toBe("abcd");
  });
});
