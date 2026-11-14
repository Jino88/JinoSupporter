import { afterEach, describe, expect, it, vi } from "vitest";
import type { ManualGroup } from "./manual";
import type { ProcessRow } from "./processList";
import { MAX_DOC_BYTES, docKind, docLinks, elapsedText, isWorkbook, jigSummary, needsReplaceConfirm, processSummary, readProcessAnswer, runJig, runProcess, sortDocs } from "./docsAi";

const proc = (id: string) => ({ id, type: "SUB", name: id, localName: "", ngItems: [], materials: [{ material: "M1", qtyPerUnit: 1 }] });
const group = (n: number): ManualGroup => ({ materials: [], processes: Array.from({ length: n }, (_, i) => proc(`p${i}`)) }) as unknown as ManualGroup;
const row = (no: string, jigs: string[]): ProcessRow =>
  ({ id: "", no, process: no, workers: null, stdTime: null, machine: "", jigs: jigs.map((name) => ({ name, drawingNo: "", ver: "", qty: "1" })), tester: "", material: "", note: "" });

describe("summaries", () => {
  it("words the results", () => {
    expect(processSummary({ processes: 3, materials: 5, note: "", model: "", saved: true })).toBe("Process: 3 processes, 5 materials");
    expect(jigSummary({ processes: 4, jigKinds: 2, note: "", unmatched: 0, saved: true })).toBe("JIG/MC List: 4 processes, 2 JIG kinds");
  });
  it("shows how long a step has run", () => {
    expect([elapsedText(0), elapsedText(9_999), elapsedText(65_000), elapsedText(-5)]).toEqual(["0:00", "0:09", "1:05", "0:00"]);
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
  it("asks with both counts before replacing", async () => {
    const asked: string[] = [];
    await runProcess({ ...steps(group(1), true, []), confirm: (t: string) => { asked.push(t); return true; } });
    expect(asked).toEqual(["Process: replace the 1 processes of the Process table with the 2 the AI made from the Work Instructions?"]);
    const jig: string[] = [];
    await runJig({
      match: async () => ({ rows: [row("1", []), row("2", [])], note: "", unmatched: [] }),
      current: async () => ({ rows: [row("9", [])], revision: 1 }),
      confirm: (t) => { jig.push(t); return false; }, save: async () => {},
    });
    expect(jig).toEqual(["JIG/MC List: replace its 1 rows with the 2 the AI made from the JIG / Machine / Tester list?"]);
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

describe("several documents at once", () => {
  const f = (name: string) => ({ name });
  it("uses the first Work Instructions and the first JIG list and says why the others are not used", () => {
    const d = sortDocs([
      { file: f("jig.xlsx"), kind: "jig-list" }, { file: f("pvr.xlsx"), kind: "instructions" },
      { file: f("old pvr.xlsx"), kind: "instructions" }, { file: f("note.xlsx"), kind: "unknown" },
      { file: f("drm.xlsx"), kind: "unknown", reason: "This workbook is encrypted (DRM)." },
    ]);
    expect([d.instructions?.name, d.jig?.name]).toEqual(["pvr.xlsx", "jig.xlsx"]);
    expect(d.skipped).toEqual(["old pvr.xlsx (another Work Instructions)", "note.xlsx (not recognized)", "drm.xlsx (This workbook is encrypted (DRM).)"]);
    expect(sortDocs([])).toEqual({ skipped: [] });
  });

  it("ties a JIG/MC row to the Process of the same name (any case and spacing) and keeps the ones in one only", () => {
    const g = { materials: [], processes: [
      { ...proc("a"), name: "Xếp  upper" }, { ...proc("b"), name: "Kiểm tra" },
    ] } as unknown as ManualGroup;
    const links = docLinks(g, [{ ...row("01", ["CAP"]), process: "Xếp UPPER" }, { ...row("02", []), process: "DRY" }]);
    expect(links.map((l) => [l.no, l.process, l.materials, l.jigs, l.inProcess, l.inJig])).toEqual([
      ["01", "Xếp UPPER", ["M1"], ["CAP"], true, true],
      ["02", "DRY", [], [], false, true],
      ["", "Kiểm tra", ["M1"], [], true, false],
    ]);
    expect(docLinks(undefined, undefined)).toEqual([]);
  });
});

describe("what a document is (server)", () => {
  afterEach(() => vi.unstubAllGlobals());
  const file = (name: string, size: number) => Object.defineProperty(new File(["x"], name), "size", { value: size });
  const reply = (status: number, body: unknown) => vi.fn(async () => new Response(JSON.stringify(body), { status }));

  it("reads the kind and why a workbook is not used", async () => {
    vi.stubGlobal("fetch", reply(200, { kind: "instructions", reason: "" }));
    expect(await docKind("M", file("pvr.xlsx", 76 * 1024 * 1024))).toEqual({ kind: "instructions" });
    vi.stubGlobal("fetch", reply(200, { kind: "unknown", reason: "This workbook is encrypted (DRM)." }));
    expect(await docKind("M", file("drm.xlsx", 10))).toEqual({ kind: "unknown", reason: "This workbook is encrypted (DRM)." });
    vi.stubGlobal("fetch", reply(400, { error: "Choose an Excel workbook." }));
    await expect(docKind("M", file("a.xlsx", 10))).rejects.toThrow("a.xlsx: Choose an Excel workbook.");
  });

  it("does not send a workbook too large, and names the file when the connection is cut", async () => {
    const send = vi.fn(async () => { throw new TypeError("Failed to fetch"); });
    vi.stubGlobal("fetch", send);
    expect(await docKind("M", file("huge.xlsx", MAX_DOC_BYTES + 1))).toEqual({ kind: "unknown", reason: "larger than 200 MB" });
    expect(send).not.toHaveBeenCalled();
    await expect(docKind("M", file("pvr.xlsx", 10))).rejects.toThrow(
      "pvr.xlsx could not be sent to the server (Failed to fetch). Check that the server is running, then try again.");
  });
});
