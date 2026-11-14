import { describe, expect, it, vi } from "vitest";
import type { ManualGroup } from "./manual";
import type { ProcessRow } from "./processList";
import { isWorkbook, jigSummary, needsReplaceConfirm, processSummary, readProcessAnswer, runJig, runProcess } from "./docsAi";

const proc = (id: string) => ({ id, type: "SUB", name: id, localName: "", ngItems: [], materials: [{ material: "M1", qtyPerUnit: 1 }] });
const group = (n: number): ManualGroup => ({ materials: [], processes: Array.from({ length: n }, (_, i) => proc(`p${i}`)) }) as unknown as ManualGroup;
const row = (no: string, jigs: string[]): ProcessRow =>
  ({ id: "", no, process: no, workers: null, stdTime: null, machine: "", jigs: jigs.map((name) => ({ name, drawingNo: "", ver: "", qty: "1" })), tester: "", material: "", note: "" });

describe("summaries", () => {
  it("words the results", () => {
    expect(processSummary({ processes: 3, materials: 5, note: "", model: "", saved: true })).toBe("Process: 3 processes, 5 materials");
    expect(jigSummary({ processes: 4, jigKinds: 2, note: "", unmatched: 0, saved: true })).toBe("JIG/MC List: 4 processes, 2 JIG kinds");
  });
  it("accepts only xlsx / xlsm", () => {
    expect(isWorkbook({ name: "a.XLSM" })).toBe(true);
    expect(isWorkbook({ name: "a.xls" })).toBe(false);
  });
  it("adds the materials of the answer and counts them", () => {
    const a = readProcessAnswer(group(0), { group: group(2), note: "n", model: "m" });
    expect(a.group.processes).toHaveLength(2);
    expect(a.group.materials).toEqual(["M1"]);
    expect(a.empty).toBe(false);
  });
  it("confirms only when there is something to replace", () => {
    expect(needsReplaceConfirm(true, false)).toBe(true);
    expect(needsReplaceConfirm(false, false)).toBe(false);
  });
});

describe("runProcess", () => {
  const steps = (cur: ManualGroup, ok: boolean, order: string[]) => ({
    hasInstructions: true, currentGroup: cur,
    ask: async () => { order.push("ask"); return { group: group(2), note: "", model: "m" }; },
    confirm: () => { order.push("confirm"); return ok; },
    save: async () => { order.push("save"); },
  });
  it("asks, confirms the replacement, then saves", async () => {
    const order: string[] = [];
    const r = await runProcess(steps(group(1), true, order));
    expect(order).toEqual(["ask", "confirm", "save"]);
    expect(r).toMatchObject({ processes: 2, materials: 1, saved: true });
  });
  it("saves without confirming when the table was empty", async () => {
    const order: string[] = [];
    await runProcess(steps(group(0), true, order));
    expect(order).toEqual(["ask", "save"]);
  });
  it("does not save when the replacement is declined", async () => {
    const order: string[] = [];
    const r = await runProcess(steps(group(1), false, order));
    expect(order).toEqual(["ask", "confirm"]);
    expect(r.saved).toBe(false);
  });
});

describe("runJig", () => {
  const rows = [row("1", ["A", "B"]), row("2", ["A"])];
  const mk = (cur: ProcessRow[], ok: boolean, order: string[], match = rows) => ({
    match: async () => { order.push("match"); return { rows: match, note: "x", unmatched: ["u"] }; },
    current: async () => { order.push("get"); return { rows: cur, revision: 7 }; },
    confirm: () => { order.push("confirm"); return ok; },
    save: vi.fn(async (_r: ProcessRow[], rev: number) => { order.push(`save@${rev}`); }),
  });
  it("matches, reads the revision, confirms, then saves with that revision", async () => {
    const order: string[] = [];
    const r = await runJig(mk([row("9", [])], true, order));
    expect(order).toEqual(["match", "get", "confirm", "save@7"]);
    expect(r).toMatchObject({ processes: 2, jigKinds: 2, unmatched: 1, saved: true });
  });
  it("skips the confirmation when the list is empty", async () => {
    const order: string[] = [];
    await runJig(mk([], true, order));
    expect(order).toEqual(["match", "get", "save@7"]);
  });
  it("keeps the old list when declined or when nothing matched", async () => {
    const a: string[] = [];
    expect((await runJig(mk([row("9", [])], false, a))).saved).toBe(false);
    expect(a).not.toContain("save@7");
    const b: string[] = [];
    await runJig(mk([], true, b, []));
    expect(b).toEqual(["match"]);
  });
});
