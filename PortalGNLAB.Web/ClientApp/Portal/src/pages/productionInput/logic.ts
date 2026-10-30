import type {
  Cell,
  DayData,
  DaySummary,
  ModelSetup,
  NgCounts,
  NgGroup,
  NgItem,
  Pair,
  PlanData,
  PlanStage,
  Process,
  ProcessType,
  Segment,
  Sheet,
  SheetMerge,
  Slot,
} from "./types";

export interface SegmentInfo {
  name: string;
  short: string;
  color: string;
  /** set = first input typed, prev = previous day's output of `from`, chain = previous process output. */
  start: "set" | "prev" | "chain";
  from?: Segment;
  why: string;
}

export const SEGMENTS: Record<Segment, SegmentInfo> = {
  A: { name: "SUB ASSY → SUB DRY 12h", short: "Sub assy → Sub Dry", color: "#6366f1", start: "set", why: "Input = UPPER loaded today" },
  B: { name: "MAIN — SPK SIDE", short: "Main — SPK side", color: "#0ea5e9", start: "prev", from: "A", why: "Input = previous day's Sub Dry output (12h dry)" },
  C: { name: "MAIN — LOWER / PT ASSY → MAIN DRY 12h", short: "Lower / PT → Main Dry", color: "#0ea5e9", start: "set", why: "Input = LOWER assy count" },
  D: { name: "AFTER MAIN DRY 12h", short: "After Main Dry", color: "#14b8a6", start: "prev", from: "C", why: "Input = previous day's Main Dry output (12h dry)" },
  E: { name: "FUNCTION", short: "Function", color: "#f59e0b", start: "chain", why: "Input carried from the previous process · can be split by time slot" },
  F: { name: "PACKING / OQC", short: "Packing / OQC", color: "#10b981", start: "chain", why: "" },
};
export const SEGMENT_KEYS = Object.keys(SEGMENTS) as Segment[];

export const PROCESS_TYPES: ProcessType[] = ["SUB", "MAIN", "FUNCTION", "VISUAL"];
export const ngGroup = (p: Pick<Process, "type">): NgGroup => (p.type === "FUNCTION" ? "F" : "P");
const SEGMENT_TYPE: Record<Segment, ProcessType> = { A: "SUB", B: "MAIN", C: "MAIN", D: "MAIN", E: "FUNCTION", F: "VISUAL" };
export const defaultType = (segment: Segment): ProcessType => SEGMENT_TYPE[segment];

/** Gives setups saved with the old P / F group a Process Type that keeps their NG totals. */
export function normalizeProcess(p: Process): Process {
  if (PROCESS_TYPES.includes(p.type)) return p;
  const { group, ...rest } = p;
  // A P process in the function block (Gauss, Visual) is a visual check.
  const type: ProcessType = group === "F" ? "FUNCTION" : group === "P" && p.segment === "E" ? "VISUAL" : defaultType(p.segment);
  return { ...rest, type };
}

export const sum = (values: number[]) => values.reduce((s, v) => s + v, 0);
const ZERO: Pair = [0, 0];

export function newId(): string {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
}

/** 1234 → "1,234"; null → "—". */
export const fmt = (v: number | null | undefined) => (v == null ? "—" : v.toLocaleString("en-US"));
export const signed = (v: number) => (v > 0 ? "+" : "") + v.toLocaleString("en-US");
export const pct = (v: number, digits = 2) => (v * 100).toFixed(digits) + "%";
/** NG count over the process input; 0 when nothing went in. */
export const ngRate = (ng: number, input: number) => (input ? ng / input : 0);
/** Rate as a percent number with 2 decimals, for the Excel forms. */
const pctNumber = (v: number) => Math.round(v * 10000) / 100;

export const firstOf = (setup: ModelSetup, seg: Segment) => setup.processes.find((p) => p.segment === seg);

const pairOf = (counts: NgCounts, p: Process, item: NgItem): Pair => counts[p.id]?.[item.id] ?? ZERO;

/** NG pairs of the day; a function process adds its two time slots. */
export function dayPairs(p: Process, data: DayData): Pair[] {
  if (p.segment !== "E") return p.ngItems.map((i) => pairOf(data.ng, p, i));
  return p.ngItems.map((i) => {
    const a = pairOf(data.am, p, i), b = pairOf(data.pm, p, i);
    return [a[0] + b[0], a[1] + b[1]];
  });
}

export function viewPairs(p: Process, data: DayData, slot: Slot): Pair[] {
  return p.segment !== "E" || slot === "day" ? dayPairs(p, data) : p.ngItems.map((i) => pairOf(data[slot], p, i));
}

export interface SlotRow {
  input: number;
  before: number;
  after: number;
  out: number;
}

export interface Row {
  input: number;
  /** Input the process would get without a typed override (carried, previous day or start count). */
  autoInput: number;
  mode: "set" | "prev" | "auto";
  before: number;
  after: number;
  out: number;
  rateBefore: number;
  rateAfter: number;
  am: SlotRow | null;
  pm: SlotRow | null;
}

export type Rows = Record<string, Row>;

/** Start inputs of the "prev" segments: the previous day's output of the segment they follow. */
export function prevOutFrom(previous: DaySummary | null | undefined): Partial<Record<Segment, number>> {
  const result: Partial<Record<Segment, number>> = {};
  for (const k of SEGMENT_KEYS) {
    const from = SEGMENTS[k].from;
    if (from) result[k] = previous?.segOut?.[from] ?? 0;
  }
  return result;
}

/** Input / NG / output of every process. Output = input − NG after rework, carried to the next process. */
export function calc(setup: ModelSetup, data: DayData, prevOut: Partial<Record<Segment, number>>): Rows {
  const rows: Rows = {};
  let carry = 0, carryAm = 0, carryPm = 0;
  for (const p of setup.processes) {
    const s = SEGMENTS[p.segment], first = firstOf(setup, p.segment) === p;
    let autoInput: number, mode: Row["mode"];
    if (first && s.start === "set") { autoInput = data.startSet[p.segment] ?? 0; mode = "set"; }
    else if (first && s.start === "prev") { autoInput = prevOut[p.segment] ?? 0; mode = "prev"; }
    else { autoInput = carry; mode = "auto"; }
    let input = autoInput;
    if (data.override[p.id] != null) { input = data.override[p.id]; mode = "set"; }
    const d = dayPairs(p, data);
    const before = sum(d.map((x) => x[0])), after = sum(d.map((x) => x[1]));
    const out = input - after;
    let am: SlotRow | null = null, pm: SlotRow | null = null;
    if (p.segment === "E") {
      const inAm = first ? (data.amShare[p.id] ?? Math.round(input / 2)) : carryAm;
      // a typed day input of a later function process keeps the 08–12h carry and moves the difference to 12–17h
      const inPm = first || data.override[p.id] != null ? input - inAm : carryPm;
      const slot = (counts: NgCounts, slotInput: number): SlotRow => {
        const pairs = p.ngItems.map((i) => pairOf(counts, p, i));
        const a = sum(pairs.map((x) => x[1]));
        return { input: slotInput, before: sum(pairs.map((x) => x[0])), after: a, out: slotInput - a };
      };
      am = slot(data.am, inAm);
      pm = slot(data.pm, inPm);
      carryAm = am.out;
      carryPm = pm.out;
    }
    rows[p.id] = {
      input, autoInput, mode, before, after, out,
      rateBefore: input ? before / input : 0,
      rateAfter: input ? after / input : 0,
      am, pm,
    };
    carry = out;
  }
  return rows;
}

export function segOut(setup: ModelSetup, rows: Rows, seg: Segment): number | null {
  const last = setup.processes.filter((p) => p.segment === seg).at(-1);
  return last ? rows[last.id].out : null;
}

export interface Actuals {
  sub: number | null;
  main: number | null;
  func: number | null;
  vis: number | null;
  oqc: number | null;
}

/** Day actuals of the plan stages: Sub Dry in, Main Dry in, last function test out, Packing in, sent to OQC. */
export function actuals(setup: ModelSetup, rows: Rows, data: DayData): Actuals {
  const fn = setup.processes.filter((p) => ngGroup(p) === "F").at(-1);
  const pk = firstOf(setup, "F");
  return {
    sub: segOut(setup, rows, "A"),
    main: segOut(setup, rows, "C"),
    func: fn ? rows[fn.id].out : null,
    vis: pk ? rows[pk.id].input : null,
    oqc: data.sendOqc,
  };
}

/** Sum of the stage NG rates of each NG group (the workbook's NG Total definition). */
export function ngTotals(setup: ModelSetup, rows: Record<string, { rateBefore: number; rateAfter: number }>) {
  const g = (grp: string, k: "rateBefore" | "rateAfter") =>
    sum(setup.processes.filter((p) => ngGroup(p) === grp).map((p) => rows[p.id][k]));
  return { processBefore: g("P", "rateBefore"), processAfter: g("P", "rateAfter"), functionBefore: g("F", "rateBefore"), functionAfter: g("F", "rateAfter") };
}

/** Number of NG cells where After is larger than Before (slots counted separately). */
export function afterOverBefore(setup: ModelSetup, data: DayData): number {
  let bad = 0;
  for (const p of setup.processes) {
    const lists = p.segment === "E" ? [data.am, data.pm] : [data.ng];
    for (const counts of lists) for (const i of p.ngItems) { const [b, a] = pairOf(counts, p, i); if (a > b) bad++; }
  }
  return bad;
}

export interface MaterialUse {
  no: number;
  process: Process;
  qtyPerUnit: number;
  used: number;
}

export function materialUses(setup: ModelSetup, rows: Rows, name: string): MaterialUse[] {
  const uses: MaterialUse[] = [];
  setup.processes.forEach((p, i) => p.materials.forEach((m) => {
    if (m.material === name) uses.push({ no: i + 1, process: p, qtyPerUnit: m.qtyPerUnit, used: rows[p.id].input * m.qtyPerUnit });
  }));
  return uses;
}

export interface MaterialRow {
  name: string;
  opening: number;
  openingTyped: boolean;
  in: number;
  source: string;
  suggested: number;
  out: number;
  outTyped: boolean;
  closing: number;
  counted: number | null;
  diff: number | null;
  uses: MaterialUse[];
}

/** Stock of each material: opening (previous day's stock unless typed) + in − out. */
export function materialRows(setup: ModelSetup, rows: Rows, data: DayData, previous: DaySummary | null | undefined): MaterialRow[] {
  return setup.materials.map((name) => {
    const m = data.materials[name];
    const uses = materialUses(setup, rows, name);
    const suggested = Math.round(sum(uses.map((u) => u.used)));
    const opening = m?.opening ?? previous?.stock?.[name] ?? previous?.closing?.[name] ?? 0;
    const inQty = m?.in ?? 0;
    const out = m?.out ?? suggested;
    const closing = opening + inQty - out;
    const counted = m?.counted ?? null;
    return {
      name, opening, openingTyped: m?.opening != null, in: inQty, source: m?.source ?? "",
      suggested, out, outTyped: m?.out != null, closing, counted,
      diff: counted == null ? null : counted - closing, uses,
    };
  });
}

export interface WipRow {
  key: string;
  name: string;
  hint: string;
  system: number | null;
  counted: number | null;
  diff: number | null;
}

export function wipRows(setup: ModelSetup, rows: Rows, data: DayData): WipRow[] {
  const procs = setup.processes;
  const visual = procs.findIndex((p) => p.segment === "E" && p.name.trim().toLowerCase() === "visual");
  const fnFirst = firstOf(setup, "E"), dOut = segOut(setup, rows, "D"), last = procs.at(-1);
  const list: [string, string, number | null, string][] = [
    ["sub", "Sub Dry 12h", segOut(setup, rows, "A"), "Today's Sub-assy output, waits 12h"],
    ["main", "Main Dry 12h", segOut(setup, rows, "C"), "Today's Lower / PT output, waits 12h"],
    ["low", "Waiting Lower assy", null, "Floor count only"],
    ["pt", "Waiting PT assy", null, "Floor count only"],
    ["func", "Waiting Function", fnFirst && dOut != null ? dOut - rows[fnFirst.id].input : null, "After-Main-Dry output − Function input"],
    ["vis", "Waiting Visual", visual > 0 ? rows[procs[visual - 1].id].out - rows[procs[visual].id].input : null, "Previous process output − Visual input"],
    ["fin", "Final OK (not yet sent)", last ? rows[last.id].out - data.sendOqc : null, "Last process output − sent to OQC"],
  ];
  return list.map(([key, name, system, hint]) => {
    const counted = data.wipCount[key] ?? null;
    return { key, name, hint, system, counted, diff: system == null || counted == null ? null : counted - system };
  });
}

export function summarize(setup: ModelSetup, data: DayData, rows: Rows, prevOut: Partial<Record<Segment, number>>, materials: MaterialRow[]): DaySummary {
  const segs: Partial<Record<Segment, number | null>> = {};
  for (const k of SEGMENT_KEYS) segs[k] = segOut(setup, rows, k);
  return {
    ...actuals(setup, rows, data),
    segOut: segs,
    prevOut,
    closing: Object.fromEntries(materials.map((m) => [m.name, m.closing])),
    stock: Object.fromEntries(materials.map((m) => [m.name, m.counted ?? m.closing])),
  };
}

/* ---------- day data edits (return a new object) ---------- */

export function setPair(data: DayData, p: Process, slot: Slot, itemId: string, k: 0 | 1, value: number): DayData {
  const key = p.segment !== "E" ? "ng" : slot === "pm" ? "pm" : "am";
  const counts = data[key];
  const pair: Pair = [...(counts[p.id]?.[itemId] ?? ZERO)];
  pair[k] = value;
  return { ...data, [key]: { ...counts, [p.id]: { ...counts[p.id], [itemId]: pair } } };
}

/**
 * Typed input: 08–12h share of the first function process, a segment start count, or an override
 * (units taken out between processes, e.g. for review). Typing the automatic value back clears the override.
 */
export function setInput(setup: ModelSetup, data: DayData, p: Process, slot: Slot, value: number, autoInput?: number): DayData {
  const first = firstOf(setup, p.segment) === p;
  if (p.segment === "E" && slot === "am" && first) return { ...data, amShare: { ...data.amShare, [p.id]: value } };
  if (first && SEGMENTS[p.segment].start === "set" && data.override[p.id] == null)
    return { ...data, startSet: { ...data.startSet, [p.segment]: value } };
  if (autoInput != null && value === autoInput) return clearOverride(data, p.id);
  return { ...data, override: { ...data.override, [p.id]: value } };
}

export function clearOverride(data: DayData, processId: string): DayData {
  const override = { ...data.override };
  delete override[processId];
  return { ...data, override };
}

/* ---------- setup edits ---------- */

const sortBySegment = (procs: Process[]) =>
  procs.map((p, i) => ({ p, i }))
    .sort((a, b) => SEGMENT_KEYS.indexOf(a.p.segment) - SEGMENT_KEYS.indexOf(b.p.segment) || a.i - b.i)
    .map((x) => x.p);

export function newProcess(segment: Segment, name = "New process"): Process {
  return { id: newId(), name, localName: "", segment, type: defaultType(segment), ngItems: [{ id: newId(), name: "NG" }], materials: [] };
}

export function addProcess(setup: ModelSetup, segment: Segment): ModelSetup {
  const procs = [...setup.processes];
  const last = procs.filter((x) => x.segment === segment).at(-1);
  procs.splice(last ? procs.indexOf(last) + 1 : procs.length, 0, newProcess(segment));
  return { ...setup, processes: sortBySegment(procs) };
}

export function deleteProcess(setup: ModelSetup, id: string): ModelSetup {
  return { ...setup, processes: setup.processes.filter((p) => p.id !== id) };
}

/** Moves a process inside its segment; at the edge of a segment it joins the neighbouring segment instead. */
export function moveProcess(setup: ModelSetup, id: string, dir: -1 | 1): ModelSetup {
  const procs = [...setup.processes];
  const k = procs.findIndex((p) => p.id === id), j = k + dir;
  if (k < 0) return setup;
  if (!procs[j] || procs[j].segment !== procs[k].segment) {
    const next = SEGMENT_KEYS[SEGMENT_KEYS.indexOf(procs[k].segment) + dir];
    if (!next) return setup;
    procs[k] = { ...procs[k], segment: next };
    return { ...setup, processes: procs };
  }
  [procs[k], procs[j]] = [procs[j], procs[k]];
  return { ...setup, processes: procs };
}

/** Changes a process; a new segment moves it to the end of that segment. */
export function updateProcess(setup: ModelSetup, id: string, change: Partial<Process>): ModelSetup {
  let procs = setup.processes.map((p) => (p.id === id ? { ...p, ...change } : p));
  if (change.segment) {
    const moved = procs.find((p) => p.id === id)!;
    procs = sortBySegment([...procs.filter((p) => p.id !== id), moved]);
  }
  return { ...setup, processes: procs };
}

const mapProcess = (setup: ModelSetup, id: string, f: (p: Process) => Process): ModelSetup =>
  ({ ...setup, processes: setup.processes.map((p) => (p.id === id ? f(p) : p)) });

export const addNgItem = (setup: ModelSetup, id: string) =>
  mapProcess(setup, id, (p) => ({ ...p, ngItems: [...p.ngItems, { id: newId(), name: "New NG" }] }));
export const renameNgItem = (setup: ModelSetup, id: string, itemId: string, name: string) =>
  mapProcess(setup, id, (p) => ({ ...p, ngItems: p.ngItems.map((i) => (i.id === itemId ? { ...i, name: name.trim() || "NG" } : i)) }));
export const deleteNgItem = (setup: ModelSetup, id: string, itemId: string) =>
  mapProcess(setup, id, (p) => ({ ...p, ngItems: p.ngItems.filter((i) => i.id !== itemId) }));
/** Links a material to a process (1 per unit); a material already linked stays as it is. */
export const linkMaterial = (setup: ModelSetup, id: string, material: string) =>
  mapProcess(setup, id, (p) => (p.materials.some((m) => m.material === material) ? p : { ...p, materials: [...p.materials, { material, qtyPerUnit: 1 }] }));
export const updateProcessMaterial = (setup: ModelSetup, id: string, index: number, change: { material?: string; qtyPerUnit?: number }) =>
  mapProcess(setup, id, (p) => ({ ...p, materials: p.materials.map((m, i) => (i === index ? { ...m, ...change } : m)) }));
export const deleteProcessMaterial = (setup: ModelSetup, id: string, index: number) =>
  mapProcess(setup, id, (p) => ({ ...p, materials: p.materials.filter((_, i) => i !== index) }));

/** Adds a material to the model; throws when the name is blank or already used. */
export function addMaterial(setup: ModelSetup, name: string): ModelSetup {
  const n = name.trim();
  if (!n) throw new Error("Type a material name.");
  if (setup.materials.some((m) => m.toLowerCase() === n.toLowerCase())) throw new Error("Material already exists.");
  return { ...setup, materials: [...setup.materials, n] };
}

/** Removes a material and its links from every process. */
export function deleteMaterial(setup: ModelSetup, name: string): ModelSetup {
  return {
    ...setup,
    materials: setup.materials.filter((m) => m !== name),
    processes: setup.processes.map((p) => ({ ...p, materials: p.materials.filter((m) => m.material !== name) })),
  };
}

export const addStage = (setup: ModelSetup): ModelSetup =>
  ({ ...setup, stages: [...setup.stages, { id: newId(), name: "New stage", source: null }] });
export const updateStage = (setup: ModelSetup, id: string, change: Partial<PlanStage>): ModelSetup =>
  ({ ...setup, stages: setup.stages.map((s) => (s.id === id ? { ...s, ...change } : s)) });
export const deleteStage = (setup: ModelSetup, id: string): ModelSetup =>
  ({ ...setup, stages: setup.stages.filter((s) => s.id !== id) });

/* ---------- dates ---------- */

export function addDays(iso: string, n: number): string {
  const d = new Date(iso + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

export const weekday = (iso: string) => new Date(iso + "T00:00:00Z").getUTCDay();
export const isWeekend = (iso: string) => weekday(iso) === 0 || weekday(iso) === 6;
export const WEEKDAYS = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];

/** Monday of the week of the date. */
export const weekStart = (iso: string) => addDays(iso, -((weekday(iso) + 6) % 7));

export const dateRange = (start: string, count: number) => Array.from({ length: count }, (_, i) => addDays(start, i));

export function localToday(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/* ---------- plan ---------- */

/**
 * Actual of a stage by date. Stages with a source take the saved day summaries and, for the day
 * on screen, the live value; the others use what was typed in the plan.
 */
export function stageActuals(
  stage: PlanStage,
  plan: PlanData,
  summaries: Record<string, DaySummary>,
  live: { date: string; actuals: Actuals } | null,
): Record<string, number> {
  if (!stage.source) return { ...(plan.actuals[stage.id] ?? {}) };
  const result: Record<string, number> = {};
  for (const [date, s] of Object.entries(summaries)) {
    const v = s[stage.source];
    if (v != null) result[date] = v;
  }
  if (live) {
    const v = live.actuals[stage.source];
    if (v != null) result[live.date] = v;
    else delete result[live.date];
  }
  return result;
}

export function setPlanCell(plan: PlanData, kind: "values" | "actuals", stageId: string, date: string, value: number | null): PlanData {
  const row = { ...(plan[kind][stageId] ?? {}) };
  if (value == null) delete row[date];
  else row[date] = value;
  return { ...plan, [kind]: { ...plan[kind], [stageId]: row } };
}

export function setPlanEvent(plan: PlanData, date: string, text: string): PlanData {
  const events = { ...plan.events };
  if (text.trim()) events[date] = text.trim();
  else delete events[date];
  return { ...plan, events };
}

/**
 * Fills the plan of each following stage with a source from the stage before it, `lead` days later.
 * Only dates after `today` change.
 */
export function chainPlan(stages: PlanStage[], plan: PlanData, dates: string[], today: string, lead: number): PlanData {
  const chained = stages.filter((s) => s.source);
  const values = { ...plan.values };
  for (let s = 1; s < chained.length; s++) {
    const prev = values[chained[s - 1].id] ?? {};
    const cur = { ...(values[chained[s].id] ?? {}) };
    for (const d of dates) {
      if (d <= today) continue;
      const v = prev[addDays(d, -lead)];
      if (v != null) cur[d] = v;
    }
    values[chained[s].id] = cur;
  }
  return { ...plan, values };
}

/* ---------- Excel-form sheets ---------- */

export interface DayEntry {
  date: string;
  data: DayData;
  summary: DaySummary;
}

export interface TotalRow {
  input: number;
  before: number;
  after: number;
  out: number;
  /** Σ before / Σ input, so a busy day weighs more than a quiet one. */
  rateBefore: number;
  rateAfter: number;
  /** Σ per NG item, in the order of the process NG items. */
  pairs: Pair[];
}

/** Daily rows of an entered day (each day starts from its own saved previous-day outputs). */
export const entryRows = (setup: ModelSetup, e: DayEntry): Rows => calc(setup, e.data, e.summary?.prevOut ?? {});

/** Σ of the entered days per process: inputs, outputs and every NG item; rates over the summed input. */
export function totalRows(setup: ModelSetup, entries: DayEntry[]): Record<string, TotalRow> {
  const result: Record<string, TotalRow> = {};
  for (const p of setup.processes)
    result[p.id] = { input: 0, before: 0, after: 0, out: 0, rateBefore: 0, rateAfter: 0, pairs: p.ngItems.map((): Pair => [0, 0]) };
  for (const e of entries) {
    const rows = entryRows(setup, e);
    for (const p of setup.processes) {
      const r = rows[p.id], t = result[p.id];
      t.input += r.input; t.before += r.before; t.after += r.after; t.out += r.out;
      dayPairs(p, e.data).forEach(([b, a], i) => { t.pairs[i] = [t.pairs[i][0] + b, t.pairs[i][1] + a]; });
    }
  }
  for (const t of Object.values(result)) {
    t.rateBefore = ngRate(t.before, t.input);
    t.rateAfter = ngRate(t.after, t.input);
  }
  return result;
}

/** Daily NG Rate in the workbook's wide layout: one row per day, per process Input + NG item B/A + NG rate. */
export function ngSheet(modelName: string, setup: ModelSetup, entries: DayEntry[]): Sheet {
  const procs = setup.processes;
  const width = 4 + sum(procs.map((p) => 3 + p.ngItems.length * 2));
  const blank = (): Cell[] => Array<Cell>(width).fill(null);
  const h0 = blank(), h1 = blank(), h2 = blank();
  const merges: SheetMerge[] = [
    { row: 0, col: 0, rowSpan: 3, colSpan: 1 },
    { row: 0, col: 1, rowSpan: 3, colSpan: 1 },
    { row: 0, col: 2, rowSpan: 2, colSpan: 2 },
  ];
  h0[0] = "Date"; h0[1] = "Model"; h0[2] = "NG Total %"; h2[2] = "Before"; h2[3] = "After";
  let c = 4;
  procs.forEach((p, i) => {
    const w = 3 + p.ngItems.length * 2;
    h0[c] = `${i + 1}. ${p.name.toUpperCase()}`;
    merges.push({ row: 0, col: c, rowSpan: 1, colSpan: w });
    h1[c] = "Input";
    merges.push({ row: 1, col: c, rowSpan: 2, colSpan: 1 });
    p.ngItems.forEach((item, j) => {
      const col = c + 1 + j * 2;
      h1[col] = item.name; h2[col] = "B"; h2[col + 1] = "A";
      merges.push({ row: 1, col, rowSpan: 1, colSpan: 2 });
    });
    const rc = c + 1 + p.ngItems.length * 2;
    h1[rc] = "NG rate %"; h2[rc] = "B"; h2[rc + 1] = "A";
    merges.push({ row: 1, col: rc, rowSpan: 1, colSpan: 2 });
    c += w;
  });
  const body = entries.map((e) => {
    const rows = calc(setup, e.data, e.summary.prevOut ?? {});
    const line: Cell[] = [e.date, modelName,
      pctNumber(sum(procs.map((p) => rows[p.id].rateBefore))), pctNumber(sum(procs.map((p) => rows[p.id].rateAfter)))];
    for (const p of procs) {
      const r = rows[p.id];
      line.push(r.input);
      for (const [b, a] of dayPairs(p, e.data)) line.push(b, a);
      line.push(pctNumber(r.rateBefore), pctNumber(r.rateAfter));
    }
    return line;
  });
  return { name: "Daily NG Rate", rows: [h0, h1, h2, ...body], merges, headerRows: 3 };
}

/** Output daily "Plan": Plan and Actual rows per stage, one column per date. */
export function planSheet(modelName: string, stages: PlanStage[], plan: PlanData, dates: string[], actualOf: (stage: PlanStage) => Record<string, number>): Sheet {
  const head: Cell[] = [modelName, null, "Total", ...dates.map((d) => d.slice(5))];
  const merges: SheetMerge[] = [{ row: 0, col: 0, rowSpan: 1, colSpan: 2 }];
  const rows: Cell[][] = [head];
  const total = (values: Record<string, number>) => sum(dates.map((d) => values[d] ?? 0));
  for (const stage of stages) {
    const planned = plan.values[stage.id] ?? {}, act = actualOf(stage);
    merges.push({ row: rows.length, col: 0, rowSpan: 2, colSpan: 1 });
    rows.push([stage.name, "Plan", total(planned), ...dates.map((d) => planned[d] ?? null)]);
    rows.push([null, "Actual", total(act), ...dates.map((d) => act[d] ?? null)]);
  }
  return { name: "Plan", rows, merges, headerRows: 1 };
}
