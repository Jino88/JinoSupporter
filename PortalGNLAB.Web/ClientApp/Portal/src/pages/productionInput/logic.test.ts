import { describe, expect, it } from "vitest";
import {
  actuals,
  addMaterial,
  addProcess,
  afterOverBefore,
  calc,
  chainPlan,
  deleteMaterial,
  linkMaterial,
  materialRows,
  moveProcess,
  ngRate,
  ngSheet,
  ngTotals,
  pct,
  planSheet,
  prevOutFrom,
  setInput,
  setPair,
  stageActuals,
  summarize,
  totalRows,
  updateProcess,
  updateProcessMaterial,
  weekStart,
  wipRows,
} from "./logic";
import { copySetup, x316bSetup } from "./template";
import { emptyDayData, emptyPlanData, type DayData, type DaySummary, type ModelSetup, type Pair } from "./types";

/** MSM-X316B PV2 on 2026-10-29, as typed in the current workbooks. */
function sample() {
  const setup = x316bSetup();
  const at = (no: number) => setup.processes[no - 1];
  let data: DayData = { ...emptyDayData(), startSet: { A: 1579, C: 1579 }, sendOqc: 400 };
  const put = (no: number, i: number, pair: Pair, key: "ng" | "am" | "pm" = "ng") => {
    const p = at(no);
    data = { ...data, [key]: { ...data[key], [p.id]: { ...data[key][p.id], [p.ngItems[i].id]: pair } } };
  };
  put(13, 1, [2, 0]); put(13, 2, [5, 5]); put(23, 0, [1, 1]); put(26, 0, [12, 12]);
  const am: Record<number, Pair[]> = { 28: [[1, 1]], 29: [[2, 2], [0, 0]], 30: [[14, 14], [12, 12], [8, 8], [0, 0], [0, 0]], 31: [[1, 1], [4, 4], [0, 0]], 32: [[1, 1], [0, 0]], 33: [[1, 1], [1, 1], [5, 5]] };
  const day: Record<number, Pair[]> = { 28: [[2, 2]], 29: [[4, 4], [0, 0]], 30: [[25, 25], [23, 23], [14, 14], [0, 0], [0, 0]], 31: [[2, 2], [7, 7], [0, 0]], 32: [[1, 1], [0, 0]], 33: [[2, 2], [2, 2], [7, 7]] };
  for (const no of Object.keys(day).map(Number))
    day[no].forEach((d, i) => { put(no, i, am[no][i], "am"); put(no, i, [d[0] - am[no][i][0], d[1] - am[no][i][1]], "pm"); });
  data = { ...data, amShare: { [at(28).id]: 735 } };
  data = { ...data, materials: { UPPER: { opening: 3145, in: 800, source: "GN receipt", out: null, counted: null } } };
  const previous = { segOut: { A: 1494, C: 1433 }, closing: {} } as unknown as DaySummary;
  return { setup, data, previous, at };
}

describe("production input calculation", () => {
  it("matches the workbook values of 2026-10-29", () => {
    const { setup, data, previous } = sample();
    const rows = calc(setup, data, prevOutFrom(previous));
    const a = actuals(setup, rows, data);
    expect(a).toEqual({ sub: 1579, main: 1579, func: 1343, vis: 1331, oqc: 400 });
    const t = ngTotals(setup, rows);
    expect(pct(t.processBefore)).toBe("2.27%");
    expect(pct(t.processAfter)).toBe("2.14%");
    expect(pct(t.functionAfter)).toBe("5.47%");
  });

  it("counts Sub actual at Sub Dry in (segment A output)", () => {
    const { setup, data, previous, at } = sample();
    const ng = setPair(data, at(7), "day", at(7).ngItems[0].id, 1, 4);
    const ngB = setPair(ng, at(7), "day", at(7).ngItems[0].id, 0, 4);
    expect(actuals(setup, calc(setup, ngB, prevOutFrom(previous)), ngB).sub).toBe(1575);
  });

  it("starts B and D from the previous day and carries NG to the next process", () => {
    const { setup, data, previous, at } = sample();
    const rows = calc(setup, data, prevOutFrom(previous));
    expect(rows[at(8).id]).toMatchObject({ input: 1494, mode: "prev" });
    expect(rows[at(23).id]).toMatchObject({ input: 1433, mode: "prev" });
    const edited = setPair(data, at(2), "day", at(2).ngItems[0].id, 1, 3);
    expect(calc(setup, edited, prevOutFrom(previous))[at(3).id].input).toBe(1576);
  });

  it("types the segment start, then an override for later processes", () => {
    const { setup, data, at } = sample();
    expect(setInput(setup, data, at(1), "day", 1600).startSet.A).toBe(1600);
    expect(setInput(setup, data, at(3), "day", 10).override[at(3).id]).toBe(10);
    expect(setInput(setup, data, at(28), "am", 700).amShare[at(28).id]).toBe(700);
  });

  it("takes a typed input for units taken out mid-line and carries it on", () => {
    const { setup, data, previous, at } = sample();
    const rows = calc(setup, data, prevOutFrom(previous));
    const taken = setInput(setup, data, at(5), "day", rows[at(5).id].input - 3, rows[at(5).id].autoInput);
    const after = calc(setup, taken, prevOutFrom(previous));
    expect(after[at(5).id]).toMatchObject({ input: 1576, autoInput: 1579, mode: "set" });
    expect(after[at(6).id].input).toBe(1576);
    const back = setInput(setup, taken, at(5), "day", 1579, after[at(5).id].autoInput);
    expect(back.override[at(5).id]).toBeUndefined();
  });

  it("keeps the function slots consistent with a typed day input", () => {
    const { setup, data, previous, at } = sample();
    const typed = { ...data, override: { ...data.override, [at(30).id]: 1400 } };
    const r = calc(setup, typed, prevOutFrom(previous))[at(30).id];
    expect(r.am!.input + r.pm!.input).toBe(1400);
  });

  it("flags After larger than Before", () => {
    const { setup, data, at } = sample();
    expect(afterOverBefore(setup, data)).toBe(0);
    expect(afterOverBefore(setup, setPair(data, at(1), "day", at(1).ngItems[0].id, 1, 9))).toBe(1);
  });

  it("takes material out from process input × qty per unit", () => {
    const { setup, data, previous } = sample();
    const rows = calc(setup, data, prevOutFrom(previous));
    const upper = materialRows(setup, rows, data, previous).find((m) => m.name === "UPPER")!;
    expect(upper).toMatchObject({ out: 1579, closing: 2366, outTyped: false });
    expect(upper.uses.map((u) => `${u.no}. ${u.process.name} × ${u.qtyPerUnit}`)).toEqual(["1. Bonding Enclosure × 1"]);
    const twice = updateProcessMaterial(setup, setup.processes[0].id, 0, { qtyPerUnit: 2 });
    const upper2 = materialRows(twice, calc(twice, data, prevOutFrom(previous)), data, previous).find((m) => m.name === "UPPER")!;
    expect(upper2).toMatchObject({ out: 3158, closing: 787 });
  });

  it("opens with the previous day's closing and saves it in the summary", () => {
    const { setup, previous } = sample();
    const data = emptyDayData();
    const prev = { ...previous, closing: { LOWER: 500, UPPER: 10 }, stock: { UPPER: 7 } };
    const rows = calc(setup, data, prevOutFrom(prev));
    const mats = materialRows(setup, rows, data, prev);
    expect(mats.find((m) => m.name === "LOWER")).toMatchObject({ opening: 500, openingTyped: false, closing: 500 });
    expect(mats.find((m) => m.name === "UPPER")!.opening).toBe(7);
    const s = summarize(setup, data, rows, prevOutFrom(prev), mats);
    expect(s.closing.LOWER).toBe(500);
    const counted = { ...data, materials: { LOWER: { opening: null, in: 0, source: "", out: null, counted: 480 } } };
    const s2 = summarize(setup, counted, rows, prevOutFrom(prev), materialRows(setup, rows, counted, prev));
    expect(s2.stock).toMatchObject({ LOWER: 480 });
    expect(s.prevOut).toEqual({ B: 1494, D: 1433 });
  });

  it("computes WIP system values", () => {
    const { setup, data, previous } = sample();
    const rows = calc(setup, data, prevOutFrom(previous));
    const w = Object.fromEntries(wipRows(setup, rows, data).map((r) => [r.key, r.system]));
    expect(w.func).toBe(0);
    expect(w.vis).toBe(0);
    expect(w.fin).toBe(rows[setup.processes.at(-1)!.id].out - 400);
  });
});

describe("setup edits", () => {
  it("adds a process at the end of its segment and keeps segments in order", () => {
    const setup = x316bSetup();
    const added = addProcess(setup, "A");
    expect(added.processes[7].name).toBe("New process");
    expect(added.processes[8].segment).toBe("B");
    const moved = updateProcess(setup, setup.processes[0].id, { segment: "C" });
    expect(moved.processes.filter((p) => p.segment === "C").at(-1)!.name).toBe("Bonding Enclosure");
  });

  it("moves a process inside its segment and into the next segment at the edge", () => {
    const setup = x316bSetup();
    expect(moveProcess(setup, setup.processes[1].id, -1).processes[0].name).toBe("AOI 1");
    const down = moveProcess(setup, setup.processes[6].id, 1);
    expect(down.processes[6]).toMatchObject({ name: "AOI 3", segment: "B" });
    const up = moveProcess(setup, setup.processes[7].id, -1);
    expect(up.processes[7]).toMatchObject({ name: "Fix Bonding 3", segment: "A" });
    expect(moveProcess(setup, setup.processes[0].id, -1)).toBe(setup);
  });

  it("links a dragged material once with 1 per unit", () => {
    const setup = x316bSetup();
    const id = setup.processes[1].id;
    const linked = linkMaterial(setup, id, "TRAY");
    expect(linked.processes[1].materials).toEqual([{ material: "TRAY", qtyPerUnit: 1 }]);
    expect(linkMaterial(linked, id, "TRAY").processes[1].materials).toHaveLength(1);
  });

  it("rejects duplicate materials and unlinks deleted ones", () => {
    const setup = x316bSetup();
    expect(() => addMaterial(setup, "upper")).toThrow("Material already exists.");
    expect(addMaterial(setup, "Glue 233").materials.at(-1)).toBe("Glue 233");
    const without = deleteMaterial(setup, "UPPER");
    expect(without.processes[0].materials).toEqual([]);
  });

  it("copies a setup with new ids", () => {
    const setup = x316bSetup();
    const copy: ModelSetup = copySetup(setup);
    expect(copy.processes.map((p) => p.name)).toEqual(setup.processes.map((p) => p.name));
    expect(copy.processes[0].id).not.toBe(setup.processes[0].id);
  });
});

describe("plan and sheets", () => {
  it("uses the live value of the day on screen for stages with a source", () => {
    const stage = { id: "s", name: "Function OK", source: "func" as const };
    const summaries = { "2026-10-28": { func: 1140 } as DaySummary, "2026-10-29": { func: 1 } as DaySummary };
    const live = { date: "2026-10-29", actuals: { sub: 0, main: 0, func: 1343, vis: 0, oqc: 0 } };
    expect(stageActuals(stage, emptyPlanData(), summaries, live)).toEqual({ "2026-10-28": 1140, "2026-10-29": 1343 });
    const typed = { id: "t", name: "Export", source: null };
    expect(stageActuals(typed, { ...emptyPlanData(), actuals: { t: { "2026-10-29": 900 } } }, summaries, live)).toEqual({ "2026-10-29": 900 });
  });

  it("fills following stages one day later, only after today", () => {
    const stages = x316bSetup().stages;
    const plan = { ...emptyPlanData(), values: { [stages[0].id]: { "2026-10-29": 800, "2026-10-30": 900 } } };
    const dates = ["2026-10-29", "2026-10-30", "2026-10-31"];
    const filled = chainPlan(stages, plan, dates, "2026-10-29", 1);
    expect(filled.values[stages[1].id]).toEqual({ "2026-10-30": 800, "2026-10-31": 900 });
    expect(filled.values[stages[2].id]).toEqual({ "2026-10-31": 800 });
    expect(filled.values[stages[5].id]).toBeUndefined();
  });

  it("builds the Daily NG Rate form with merged headers", () => {
    const { setup, data } = sample();
    const summary = { prevOut: { B: 1494, D: 1433 } } as DaySummary;
    const sheet = ngSheet("MSM-X316B", setup, [{ date: "2026-10-29", data, summary }]);
    expect(sheet.headerRows).toBe(3);
    const width = sheet.rows[0].length;
    expect(sheet.rows.every((r) => r.length === width)).toBe(true);
    expect(sheet.rows[0][4]).toBe("1. BONDING ENCLOSURE");
    expect(sheet.merges).toContainEqual({ row: 0, col: 4, rowSpan: 1, colSpan: 5 });
    expect(sheet.rows[3].slice(0, 5)).toEqual(["2026-10-29", "MSM-X316B", expect.any(Number), expect.any(Number), 1579]);
  });

  it("builds the Plan form with Plan and Actual rows", () => {
    const stages = x316bSetup().stages.slice(0, 1);
    const plan = { ...emptyPlanData(), values: { [stages[0].id]: { "2026-10-29": 800 } } };
    const sheet = planSheet("M", stages, plan, ["2026-10-29", "2026-10-30"], () => ({ "2026-10-29": 1579 }));
    expect(sheet.rows).toEqual([
      ["M", null, "Total", "10-29", "10-30"],
      ["SPK Assy (Sub 12h)", "Plan", 800, 800, null],
      [null, "Actual", 1579, 1579, null],
    ]);
  });

  it("starts the week on Monday", () => {
    expect(weekStart("2026-10-29")).toBe("2026-10-26");
    expect(weekStart("2026-11-01")).toBe("2026-10-26");
  });
});

describe("NG rate per NG item", () => {
  it("divides each item by the process input and the items add up to the process rate", () => {
    const items = [11, 22, 33].map((n) => ngRate(n, 1523));
    expect(items.map((r) => pct(r))).toEqual(["0.72%", "1.44%", "2.17%"]);
    expect(pct(ngRate(66, 1523))).toBe("4.33%");
    expect(items.reduce((a, b) => a + b, 0)).toBeCloseTo(ngRate(66, 1523), 12);
  });

  it("is 0 when nothing went into the process", () => {
    expect(ngRate(5, 0)).toBe(0);
  });
});

describe("Total of several days", () => {
  it("sums inputs and every NG item, and takes the rate over the summed input", () => {
    const { setup, data, at } = sample();
    const summary = { prevOut: { B: 1494, D: 1433 } } as DaySummary;
    const other = setPair({ ...data, startSet: { ...data.startSet, A: 1000 } }, at(2), "day", at(2).ngItems[0].id, 0, 10);
    const t = totalRows(setup, [{ date: "2026-10-28", data: other, summary }, { date: "2026-10-29", data, summary }]);
    expect(t[at(2).id]).toMatchObject({ input: 2579, before: 10, after: 0, out: 2579 });
    expect(t[at(2).id].pairs[0]).toEqual([10, 0]);
    expect(t[at(2).id].rateBefore).toBeCloseTo(10 / 2579, 12);
    expect(t[at(13).id].pairs).toEqual([[0, 0], [4, 0], [10, 10]]);
    expect(t[at(13).id].input).toBe(1494 * 2);
    expect(t[at(30).id].pairs[0]).toEqual([50, 50]);
  });

  it("gives the day rates back when every day is the same", () => {
    const { setup, data } = sample();
    const summary = { prevOut: { B: 1494, D: 1433 } } as DaySummary;
    const one = ngTotals(setup, calc(setup, data, summary.prevOut));
    const two = ngTotals(setup, totalRows(setup, [{ date: "2026-10-28", data, summary }, { date: "2026-10-29", data, summary }]));
    expect(two.processBefore).toBeCloseTo(one.processBefore, 12);
    expect(two.functionAfter).toBeCloseTo(one.functionAfter, 12);
  });

  it("is all zero without an entered day", () => {
    const { setup, at } = sample();
    expect(totalRows(setup, [])[at(2).id]).toMatchObject({ input: 0, rateBefore: 0, pairs: [[0, 0], [0, 0], [0, 0]] });
  });
});
