import { describe, expect, it } from "vitest";
import { readDevPlan, type DevPlan } from "./plan";
import { applyTidy, eventLabel, matchStage, matchTask, taskLabel, templateTasks, tidyRows } from "./planNames";
import { DEFAULT_VOCAB as V, readVocab } from "./planVocab";

const task = (id: string, name: string, f: Record<string, unknown> = {}) => ({ id, name, start: "", end: "", owner: "", done: false, note: "", ...f });
const event = (id: string, name: string, tasks: ReturnType<typeof task>[], f: Record<string, unknown> = {}) =>
  ({ id, name, start: "", end: "", color: "#4f46e5", tasks, rows: [], bomFrom: "", ...f });

describe("standard names of the Dev Plan", () => {
  it("makes event and task names of the standard name and the quantity or detail", () => {
    expect(eventLabel("ER", "1.2k")).toBe("ER (1.2k)");
    expect(eventLabel("MP", " ")).toBe("MP");
    expect(taskLabel("Sample shipment", "GN→AS: 300ea Bringup")).toBe("Sample shipment · GN→AS: 300ea Bringup");
    expect(taskLabel("Setup", "")).toBe("Setup");
  });

  it("finds the standard task in a written name by its longest other name, the rest kept as the detail", () => {
    const m = (n: string) => { const r = matchTask(n, V.tasks); return r && [r.item.code, r.detail]; };
    expect(m("GN 자재입고완료")).toEqual(["material-in", "GN"]);
    expect(m("GN 자재 도착")).toEqual(["material-in", "GN"]);
    expect(m("MAKE SAMPLE  (OK 800 SET)")).toEqual(["sample-make", "OK 800 SET"]);
    expect(m("GN→AS 샘플발송: 300ea Bringup")).toEqual(["sample-send", "GN→AS: 300ea Bringup"]);
    // "승인시료 발송" is longer than "시료 발송": the approval sample, not any sample shipment
    expect(m("150set 승인시료 발송")).toEqual(["approval-sample", "150set"]);
    expect(m("개발라인 준비")).toEqual(["line-ready", ""]);
    expect(m("1Kea _ DV1")).toBeNull();
  });

  it("finds the stage an event name starts with, as a whole word, the rest as its quantity", () => {
    const m = (n: string) => { const r = matchStage(n, V.stages); return r && [r.item.code, r.qty]; };
    expect(m("ER (1.2kpcs)")).toEqual(["ER", "1.2kpcs"]);
    expect(m("DV1 (1.2k)")).toEqual(["DV1", "1.2k"]);
    expect(m("LPR")).toEqual(["LPR", ""]);
    expect(m("ERP")).toBeNull();
    expect(m("DR build")).toBeNull();
  });

  it("lists the names of a plan not standard yet with their match, and renames only the ones given a code", () => {
    const plan: DevPlan = readDevPlan({ events: [
      event("e1", "ER (300 Set)", [task("a", "GN 자재 입고 완료"), task("b", "1Kea _ DV1"), task("c", "Setup", { code: "setup" })]),
      event("e2", "MP", [], { stage: "MP" }),
    ] });
    const rows = tidyRows(plan, V);
    expect(rows.map((r) => [r.kind, r.from, r.code, r.detail])).toEqual([
      ["event", "ER (300 Set)", "ER", "300 Set"],
      ["task", "GN 자재 입고 완료", "material-in", "GN"],
      ["task", "1Kea _ DV1", "", "1Kea _ DV1"],
    ]);
    const out = applyTidy(plan, rows, V);
    expect(out.events[0]).toMatchObject({ name: "ER (300 Set)", stage: "ER", qty: "300 Set" });
    expect(out.events[0].tasks.map((t) => [t.name, t.code, t.detail])).toEqual([
      ["Material in · GN", "material-in", "GN"], ["1Kea _ DV1", undefined, undefined], ["Setup", "setup", undefined],
    ]);
    // kept through a save and a read
    expect(readDevPlan(JSON.parse(JSON.stringify(out))).events[0].tasks[0]).toMatchObject({ code: "material-in", detail: "GN" });
  });

  it("gives an event of a stage the standard tasks it does not have yet, in order", () => {
    const ev = readDevPlan({ events: [event("e", "DV1", [task("x", "Line ready", { code: "line-ready" })], { stage: "DV1" })] }).events[0];
    expect(templateTasks(ev, V).map((t) => [t.code, t.name])).toEqual([
      ["material-in", "Material in"], ["sample-make", "Sample making"], ["sample-send", "Sample shipment"],
    ]);
    expect(templateTasks({ ...ev, stage: undefined }, V)).toEqual([]);
  });

  it("reads a saved list leniently: nothing saved is the default, unknown template codes dropped", () => {
    expect(readVocab(null)).toBe(V);
    const v = readVocab({ stages: [{ code: "ER", name: "ER" }, { code: "er", name: "again" }], tasks: [{ code: "a", name: "A", aliases: [" x ", "x", ""] }],
      templates: { ER: ["a", "zz", "a"] } });
    expect(v.stages).toHaveLength(1);
    expect(v.tasks[0].aliases).toEqual(["x"]);
    expect(v.templates).toEqual({ ER: ["a"] });
  });
});
