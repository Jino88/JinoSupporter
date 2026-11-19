import { describe, expect, it } from "vitest";
import {
  actualSpan, markDone, blankEvent, blankRow, blankTask, addBomRows, addSubTask, bomRowNames, calendarDays, eventSpan, groupSpan, inSpan, monthDayText, monthRuns, moveInGroup, moveTask, moveTo, planProblem,
  progress, planDays, readDevPlan, readMonthDay, readQty, readYearMonthDay, removeTask, rowTotal, groupQty, sampleByDay, sampleKind, setQty, setTaskQty, taskTotal, taskTree,
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
    expect(p.events[0].rows).toEqual([{ id: "r", name: "SPK Assy", group: "making", qty: { "2026-11-18": 300, "2026-11-19": 600 } }]);
    expect(p.events[1].rows).toEqual([]);
    expect(p.events[1].bomFrom).toBe("");
    const m = readDevPlan({ events: [{ name: "LPR", bomFrom: "MSM-G778B", rows: [{ id: "a", name: "MOTOR", group: "material", qty: {} }, { id: "b", name: "x", group: "odd", qty: {} }] }] });
    expect(m.events[0].bomFrom).toBe("MSM-G778B");
    expect(m.events[0].rows.map((r) => r.group)).toEqual(["material", "making"]);
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
    // the year typed in its own box
    expect(readYearMonthDay("2027", "1/5")).toBe("2027-01-05");
    expect(readYearMonthDay(" 2028 ", "2/29")).toBe("2028-02-29");
    expect(readYearMonthDay("2027", "2/29")).toBeNull();
    for (const bad of ["", "27", "20271", "abcd"]) expect(readYearMonthDay(bad, "1/5")).toBeNull();
    expect(readYearMonthDay("", " ")).toBe("");
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

  it("puts the BOM's direct materials into the Material Plan once, skipping ones it has", () => {
    const bom = [
      { level: 1, materialName: "Enclosure Upper", materialCode: "P-1" },
      { level: 2, materialName: "Screw", materialCode: "S-9" },
      { level: 1, materialName: "MOTOR", materialCode: "P-2" },
      { level: 1, materialName: "motor", materialCode: "P-2" },
      { level: 1, materialName: "", materialCode: "P-3" },
      { level: 1, materialName: "SPK Assy", materialCode: "P-4" },
    ];
    const rows = [blankRow("Motor", "material"), blankRow("SPK Assy")];
    expect(bomRowNames(rows, bom)).toEqual(["Enclosure Upper", "P-3", "SPK Assy"]);
    const e = addBomRows({ ...ev("", ""), rows }, ["Enclosure Upper"], "MSM-G778B");
    expect(e.bomFrom).toBe("MSM-G778B");
    expect(e.rows.map((r) => [r.name, r.group])).toEqual([["Motor", "material"], ["SPK Assy", "making"], ["Enclosure Upper", "material"]]);
  });

  it("moves a quantity row past the next row of its own group", () => {
    const r = [blankRow("m1", "material"), blankRow("k1"), blankRow("m2", "material"), blankRow("k2")];
    const names = (x: { name: string }[]) => x.map((y) => y.name).join(",");
    expect(names(moveInGroup(r, r[0].id, 1))).toBe("m2,k1,m1,k2");
    expect(names(moveInGroup(r, r[3].id, -1))).toBe("m1,k2,m2,k1");
    expect(moveInGroup(r, r[2].id, 1)).toBe(r);
    expect(moveInGroup(r, r[1].id, -1)).toBe(r);
  });

  describe("sub-tasks", () => {
    const t = (id: string, parentId?: string, start = "", end = ""): PlanTask => ({ ...task(start, end, { id, name: id }), ...(parentId ? { parentId } : {}) });
    const ids = (l: PlanTask[]) => l.map((x) => x.id).join(",");
    const list = [t("A"), t("a1", "A"), t("B"), t("a2", "A"), t("b1", "B")];

    it("shows each task followed by its sub-tasks; one whose parent is gone or is a sub-task becomes a task", () => {
      expect(ids(taskTree(list))).toBe("A,a1,a2,B,b1");
      const fixed = taskTree([t("A"), t("x", "gone"), t("a1", "A"), t("y", "a1")]);
      expect(fixed.map((x) => [x.id, x.parentId ?? ""])).toEqual([["A", ""], ["a1", "A"], ["x", ""], ["y", ""]]);
    });

    it("adds a sub-task after the last sub-task of its task", () => {
      expect(ids(addSubTask(taskTree(list), "A", t("a3")))).toBe("A,a1,a2,a3,B,b1");
      expect(ids(addSubTask(taskTree(list), "B", t("b2")))).toBe("A,a1,a2,B,b1,b2");
    });

    it("moves a task with its sub-tasks, a sub-task only within its task", () => {
      const tree = taskTree(list);
      expect(ids(moveTask(tree, "A", "B"))).toBe("B,b1,A,a1,a2");
      expect(ids(moveTask(tree, "B", "A"))).toBe("B,b1,A,a1,a2");
      expect(ids(moveTask(tree, "a2", "a1"))).toBe("A,a2,a1,B,b1");
      expect(moveTask(tree, "a1", "b1")).toBe(tree);
      expect(moveTask(tree, "a1", "B")).toBe(tree);
    });

    it("deletes a task with its sub-tasks", () => {
      expect(ids(removeTask(taskTree(list), "A"))).toBe("B,b1");
      expect(ids(removeTask(taskTree(list), "a1"))).toBe("A,a2,B,b1");
    });

    it("spans a task without dates over its sub-tasks", () => {
      const l = [t("A"), t("a1", "A", "2026-11-05", "2026-11-07"), t("a2", "A", "2026-11-03", ""), t("B", undefined, "2026-11-01", "2026-11-02"), t("b1", "B", "2026-11-09", "")];
      expect(groupSpan(l, l[0])).toEqual(["2026-11-03", "2026-11-07"]);
      expect(groupSpan(l, l[3])).toEqual(["2026-11-01", "2026-11-02"]);
      expect(groupSpan(l, l[1])).toEqual(["2026-11-05", "2026-11-07"]);
      expect(groupSpan([t("C")], t("C"))).toBeNull();
    });

    it("reads the sub-tasks of a saved plan under their task", () => {
      const p = readDevPlan({ events: [{ id: "e", tasks: [{ id: "A", name: "A" }, { id: "B", name: "B" }, { id: "a1", name: "a1", parentId: "A" }, { id: "z", name: "z", parentId: "nope" }] }] });
      expect(p.events[0].tasks.map((x) => [x.id, x.parentId ?? ""])).toEqual([["A", ""], ["a1", "A"], ["B", ""], ["z", ""]]);
      expect("parentId" in p.events[0].tasks[0]).toBe(false);
    });
  });
});

describe("actual start and end of a task", () => {
  const t = (f: Record<string, unknown>) => ({ ...blankTask(), name: "T", ...f });
  it("are kept when saved and read back, dropped when not dates", () => {
    const p = readDevPlan({ events: [{ name: "E", tasks: [{ id: "a", name: "A", actualStart: "2026-11-18", actualEnd: "2026-11-20" }, { id: "b", name: "B", actualStart: "x", actualEnd: "" }] }] });
    expect(p.events[0].tasks[0]).toMatchObject({ actualStart: "2026-11-18", actualEnd: "2026-11-20" });
    expect("actualStart" in p.events[0].tasks[1] || "actualEnd" in p.events[0].tasks[1]).toBe(false);
  });

  it("take the days the task really took: still going = up to today, only an end = that day, sub-tasks for a task without its own", () => {
    const today = "2026-11-16";
    expect(actualSpan([], t({ actualStart: "2026-11-10", actualEnd: "2026-11-12" }), today)).toEqual(["2026-11-10", "2026-11-12"]);
    expect(actualSpan([], t({ actualStart: "2026-11-10" }), today)).toEqual(["2026-11-10", "2026-11-16"]);
    expect(actualSpan([], t({ actualStart: "2026-11-20" }), today)).toEqual(["2026-11-20", "2026-11-20"]);
    expect(actualSpan([], t({ actualEnd: "2026-11-12" }), today)).toEqual(["2026-11-12", "2026-11-12"]);
    expect(actualSpan([], t({}), today)).toBeNull();
    const parent = t({ id: "p" });
    const kids = [t({ id: "k1", parentId: "p", actualStart: "2026-11-03", actualEnd: "2026-11-05" }), t({ id: "k2", parentId: "p", actualStart: "2026-11-04", actualEnd: "2026-11-09" })];
    expect(actualSpan([parent, ...kids], parent, today)).toEqual(["2026-11-03", "2026-11-09"]);
  });

  it("ticking done gives today as the actual end when there is none, never overwriting one", () => {
    expect(markDone(t({}), true, "2026-11-16")).toMatchObject({ done: true, actualEnd: "2026-11-16" });
    expect(markDone(t({ actualEnd: "2026-11-12" }), true, "2026-11-16").actualEnd).toBe("2026-11-12");
    expect(markDone(t({ done: true, actualEnd: "2026-11-12" }), false, "2026-11-16")).toMatchObject({ done: false, actualEnd: "2026-11-12" });
  });
});

describe("samples of a sample task counted day by day", () => {
  const t = (f: Partial<PlanTask>) => ({ ...blankTask(), name: "T", ...f });
  it("sets and clears the count of a day, in date order, the field gone with the last day", () => {
    let s = setTaskQty(t({ code: "sample-make" }), "2026-12-17", 400);
    s = setTaskQty(s, "2026-12-15", 100);
    expect(s.qty).toEqual({ "2026-12-15": 100, "2026-12-17": 400 });
    expect(Object.keys(s.qty!)).toEqual(["2026-12-15", "2026-12-17"]);
    expect(taskTotal(s)).toBe(500);
    s = setTaskQty(setTaskQty(s, "2026-12-15", undefined), "2026-12-17", 0);
    expect("qty" in s).toBe(false);
  });

  it("sums each day per sample task over the sample tasks only", () => {
    const tasks = [
      t({ code: "sample-make", qty: { "2026-12-15": 100, "2026-12-16": 200 } }),
      t({ code: "sample-make", parentId: "x", qty: { "2026-12-16": 50 } }),
      t({ code: "sample-send", qty: { "2026-12-16": 300 } }),
      t({ code: "setup", qty: { "2026-12-16": 999 } }),
    ];
    expect(sampleByDay(tasks)).toEqual({ "2026-12-15": { "sample-make": 100 }, "2026-12-16": { "sample-make": 250, "sample-send": 300 } });
  });

  it("counts a sub-task without a category of its own for the sample task it is under", () => {
    const tasks = [
      t({ id: "p", code: "sample-send", qty: { "2026-12-16": 10 } }),
      t({ id: "a", parentId: "p", qty: { "2026-12-16": 20, "2026-12-17": 5 } }),
      t({ id: "b", parentId: "p", code: "setup", qty: { "2026-12-16": 999 } }),
      t({ id: "c", parentId: "q", qty: { "2026-12-16": 999 } }),
    ];
    expect(tasks.map((x) => sampleKind(x, tasks) ?? "-")).toEqual(["sample-send", "sample-send", "-", "-"]);
    expect(sampleByDay(tasks)).toEqual({ "2026-12-16": { "sample-send": 30 }, "2026-12-17": { "sample-send": 5 } });
    expect(groupQty(tasks[0], tasks)).toEqual({ "2026-12-16": 30, "2026-12-17": 5 });
  });

  it("reads a task code through the list in use (its own codes for the built-in tasks)", () => {
    const kind = (c: string | undefined) => (c === "Make Sample" ? "sample-make" : c);
    expect(sampleKind(t({ code: "Make Sample" }), [], kind)).toBe("sample-make");
    expect(sampleKind(t({ code: "Make Sample" }), [])).toBeUndefined();
  });

  it("are kept when saved and read back, days not dates or counts not above 0 dropped", () => {
    const p = readDevPlan({ events: [{ name: "E", tasks: [
      { id: "a", name: "A", code: "sample-send", qty: { "2026-12-19": 300, x: 5, "2026-12-20": 0, "2026-12-21": "7" } },
      { id: "b", name: "B", code: "sample-send", qty: { x: 1 } },
    ] }] });
    expect(p.events[0].tasks[0].qty).toEqual({ "2026-12-19": 300 });
    expect("qty" in p.events[0].tasks[1]).toBe(false);
  });
});
